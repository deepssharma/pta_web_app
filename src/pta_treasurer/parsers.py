"""
parsers.py
All file parsing functions for the PTA Treasurer Report Generator.
Handles QuickBooks Transaction Detail CSV, Givebacks CSV, and Chase PDF.
"""

import csv
import io
import re
import pdfplumber
from pathlib import Path
from datetime import datetime
from collections import defaultdict
from itertools import combinations

def _read_csv_robust(path: Path) -> list:
    """
    Reads a CSV file trying multiple encodings.
    Handles NUL bytes, mixed line endings, and BOM characters.
    Returns list of rows.
    """
    for encoding in ['utf-8-sig', 'utf-16', 'utf-16-le', 'utf-16-be',
                     'utf-8', 'latin-1', 'windows-1252']:
        try:
            with open(path, 'rb') as f:
                raw = f.read()
            content = raw.replace(b'\x00', b'').decode(encoding)
            content = content.replace('\r\n', '\n').replace('\r', '\n')
            reader  = csv.reader(content.splitlines())
            lines   = [row for row in reader if any(c.strip() for c in row)]
            if lines:
                return lines, encoding
        except (UnicodeDecodeError, UnicodeError):
            continue
    raise ValueError(f'Could not read {path.name} with any known encoding')


def _parse_amount(val: str) -> float:
    """Parse a dollar amount string to float."""
    try:
        return float(str(val).replace('$', '').replace(',', '')
                              .replace('"', '').strip())
    except (ValueError, AttributeError):
        return 0.0


def _money_str(val) -> str:
    """Format float as dollar string."""
    try:
        return f'${float(val):,.2f}'
    except:
        return '$0.00'


# ── 1. PARSE QUICKBOOKS TRANSACTION DETAIL ────────────────────────────────────

def parse_quickbooks_detail(folder: Path, input_month: str = '',
                             fiscal_year: str = '',
                             pass_through_categories: set | None = None,
                             pass_through_fund_name: str = '') -> dict:
    """
    Parses QuickBooks Transaction Detail by Account CSV export.

    Args:
        folder:                   Path to the month folder (e.g. input/July_2025/)
        input_month:              Month name e.g. 'July' — used to find the right file
        fiscal_year:              Year string e.g. '2025'
        pass_through_categories:  Optional set of uppercased QuickBooks category
                                   names (e.g. {'READTHON', 'READTHON-EXPENSES'})
                                   treated as a pass-through fund — money that
                                   moves through the org's bank account but
                                   isn't the org's own. Excluded from
                                   income/expenses/net_income entirely, tracked
                                   separately via the pass_through_* totals
                                   below. None/empty = no pass-through fund
                                   configured, today's exact behavior.
        pass_through_fund_name:   The fund's display name (e.g. 'READTHON'),
                                   used only to warn about a category that
                                   looks fund-related but isn't an exact match
                                   in pass_through_categories (QuickBooks can
                                   split a fund across several similarly-named
                                   sections — this has caught a real missed
                                   variant in production before).

    Returns dict with keys:
        period, income, income_total, expenses, expense_total,
        net_income, pass_through_income_total, pass_through_expense_total,
        pass_through_net, transactions

    Transactions under a pass_through_categories section are tagged
    is_pass_through=True and diverted away from income/expenses.
    """
    pass_through_categories = pass_through_categories or set()
    # Find matching file
    all_csv = sorted(folder.glob('*.csv'))
    if not all_csv:
        raise FileNotFoundError(
            f'No CSV files found in {folder}\n'
            f'Expected a QuickBooks Transaction Detail export'
        )

    # Prefer file with month name in it
    if input_month:
        matching = [f for f in all_csv
                    if input_month.lower() in f.name.lower()
                    and 'givebacks' not in f.name.lower()]
    else:
        matching = [f for f in all_csv
                    if 'givebacks' not in f.name.lower()]

    if not matching:
        raise FileNotFoundError(
            f'No QuickBooks file found for {input_month} in {folder}\n'
            f'Files available: {[f.name for f in all_csv]}'
        )

    path = matching[-1]
    print(f'  QuickBooks: {path.name}')

    lines, encoding = _read_csv_robust(path)
    print(f'  Encoding  : {encoding}')

    data = {
        'period':                    '',
        'income':                    {},
        'income_total':              0.0,
        'expenses':                  {},
        'expense_total':             0.0,
        'net_income':                0.0,
        'pass_through_income_total':  0.0,
        'pass_through_expense_total': 0.0,
        'pass_through_net':           0.0,
        'transactions':              [],
    }

    # Period from line 3
    if len(lines) > 2:
        data['period'] = lines[2][0].strip().strip('"')

    # Find header row
    header_row_idx = None
    for i, row in enumerate(lines):
        if any('transaction date' in c.lower() for c in row):
            header_row_idx = i
            break

    if header_row_idx is None:
        raise ValueError(
            f'Could not find header row in {path.name}\n'
            f'Expected a "Transaction Detail by Account" export'
        )

    headers = lines[header_row_idx]

    def col(name):
        for i, h in enumerate(headers):
            if name.lower() in h.lower():
                return i
        return None

    idx_date   = col('Transaction date') or 1
    idx_type   = col('Transaction type') or 2
    idx_num    = col('Num')              or 3
    idx_name   = col('Name')             or 4
    idx_desc   = col('Description')      or 5
    idx_amount = col('Amount')           or 7

    # Parse category sections
    # Skip the "Checking (XXXX)" section — use category sections only
    current_category = None
    in_checking      = False

    PARENT_KEYWORDS = {
        'fundraising-net', 'program-income', 'program expense',
         'admin/general', 'total',
        'accrual basis',
    }

    for row in lines[header_row_idx + 1:]:
        if len(row) < 2:
            continue

        first  = row[0].strip().strip('"')
        second = row[1].strip().strip('"') if len(row) > 1 else ''

        # Section header — non-empty first col, empty second col
        if first and not second:
            first_lower = first.lower()

            if 'checking' in first_lower and ('(' in first_lower or '4346' in first_lower):
                in_checking = True
                current_category = None
                continue

            if first_lower.startswith('total for checking'):
                in_checking = False
                continue

            if first_lower.startswith('total'):
                continue

            if any(kw in first_lower for kw in PARENT_KEYWORDS):
                continue

            if 'accrual basis' in first_lower:
                continue

            # Real category section
            in_checking = False
            current_category = first
            continue

        if in_checking or not current_category:
            continue

        if first.lower().startswith('total'):
            continue

        # Parse transaction
        date_val   = row[idx_date].strip()   if len(row) > idx_date   else ''
        type_val   = row[idx_type].strip()   if len(row) > idx_type   else ''
        num_val    = row[idx_num].strip()    if len(row) > idx_num    else ''
        name_val   = row[idx_name].strip()   if len(row) > idx_name   else ''
        desc_val   = row[idx_desc].strip()   if len(row) > idx_desc   else ''
        amt_val    = row[idx_amount].strip() if len(row) > idx_amount else ''

        if not date_val or not amt_val:
            continue

        amount = _parse_amount(amt_val)
        if amount == 0.0:
            continue

        # Clean description
        clean_desc = desc_val
        if 'ORIG CO NAME' in desc_val or 'GB Payout' in desc_val:
            clean_desc = 'MemberHub/Givebacks Deposit'
        elif desc_val.upper().startswith('CHECK #'):
            clean_desc = desc_val

        # Classify by transaction type
        is_deposit      = type_val.lower() == 'deposit'
        is_expense      = type_val.lower() in ('check', 'expense', 'bill payment', 'bill')
        is_pass_through = current_category.strip().upper() in pass_through_categories

        if (pass_through_fund_name and not is_pass_through
                and pass_through_fund_name.lower() in current_category.lower()):
            print(f'  WARNING: category "{current_category}" looks like the '
                  f'{pass_through_fund_name} pass-through fund but is not in '
                  f'{pass_through_categories} — check QuickBooks naming or add '
                  f'it if this is a new variant')

        transaction = {
            'date':            date_val,
            'type':            type_val,
            'check_no':        num_val,
            'payee':           name_val,
            'description':     clean_desc,
            'category':        current_category,
            'amount':          abs(amount),
            'is_income':       is_deposit,
            'is_pass_through': is_pass_through,
        }
        data['transactions'].append(transaction)

        if is_pass_through:
            if is_deposit:
                data['pass_through_income_total'] += abs(amount)
            elif is_expense:
                data['pass_through_expense_total'] += abs(amount)
        elif is_deposit:
            data['income'][current_category] = (
                data['income'].get(current_category, 0.0) + abs(amount))
        elif is_expense:
            data['expenses'][current_category] = (
                data['expenses'].get(current_category, 0.0) + abs(amount))

    data['income_total']    = sum(data['income'].values())
    data['expense_total']   = sum(data['expenses'].values())
    data['net_income']      = data['income_total'] - data['expense_total']
    data['pass_through_net'] = (data['pass_through_income_total']
                                 - data['pass_through_expense_total'])

    print(f'  Period    : {data["period"]}')
    print(f'  Income    : ${data["income_total"]:,.2f}  ({len(data["income"])} categories)')
    print(f'  Expenses  : ${data["expense_total"]:,.2f}  ({len(data["expenses"])} categories)')
    print(f'  Transactions: {len(data["transactions"])}')
    if data['pass_through_income_total'] or data['pass_through_expense_total']:
        print(f'  Pass-through: +${data["pass_through_income_total"]:,.2f} '
              f'-${data["pass_through_expense_total"]:,.2f} '
              f'(net ${data["pass_through_net"]:,.2f})')
    return data


# ── 2. PARSE GIVEBACKS (multiple files) ──────────────────────────────────────

def parse_givebacks_files(file_info_list: list) -> list:
    """
    Merges all Givebacks CSV files for the month.
    
    Args:
        file_info_list: List of (Path, month_label, fiscal_idx) tuples
                        as produced by Cell 3 detection logic.
    
    Returns list of dicts with keys:
        item, category, count, total, source_file
    """
    if not file_info_list:
        raise ValueError('No Givebacks files provided')

    merged = {}
    for fpath, lbl, idx in file_info_list:
        try:
            lines, encoding = _read_csv_robust(fpath)
        except Exception as e:
            print(f'  WARNING: Could not read {fpath.name}: {e}')
            continue

        # Find header row
        header_idx = None
        for i, row in enumerate(lines):
            if any('item' in c.lower() for c in row):
                header_idx = i
                break

        if header_idx is None:
            # Try reading as DictReader directly
            try:
                with open(fpath, newline='', encoding='utf-8-sig') as f:
                    reader = csv.DictReader(f)
                    for r in reader:
                        _process_givebacks_row(r, fpath.name, merged)
            except Exception as e:
                print(f'  WARNING: Could not parse {fpath.name}: {e}')
            continue

        headers = lines[header_idx]
        for row in lines[header_idx + 1:]:
            if len(row) < len(headers):
                continue
            r = dict(zip(headers, row))
            _process_givebacks_row(r, fpath.name, merged)

    rows = list(merged.values())
    for r in rows:
        r['source_file'] = ', '.join(sorted(r['source_files']))
        del r['source_files']

    total = sum(r['total'] for r in rows)
    print(f'  Givebacks : {len(rows)} items  total=${total:,.2f}')
    return rows


def _process_givebacks_row(r: dict, filename: str, merged: dict):
    """Helper to process a single Givebacks CSV row into the merged dict."""
    item = r.get('Item', '').strip()
    if not item:
        return

    try:
        amt = float(str(r.get('Total', '0')).replace('$', '').replace(',', '').strip())
    except (ValueError, AttributeError):
        amt = 0.0

    try:
        cnt = int(str(r.get('No. of Transactions', '0')).strip() or 0)
    except (ValueError, AttributeError):
        cnt = 0

    if item in merged:
        merged[item]['count']  += cnt
        merged[item]['total']  += amt
        merged[item]['source_files'].add(filename)
    else:
        merged[item] = {
            'item':         item,
            'category':     r.get('Categories', '').strip(),
            'count':        cnt,
            'total':        amt,
            'source_files': {filename},
        }


# ── 3. PARSE CHASE BANK PDF ───────────────────────────────────────────────────

def parse_chase_pdf(bank_file: Path) -> dict:
    """
    Parses a Chase Bank statement PDF.
    
    Args:
        bank_file: Path to the Chase PDF statement
    
    Returns dict with keys:
        period, account, beginning_balance, ending_balance,
        total_deposits, total_checks, total_withdrawals, total_fees,
        deposits, checks, withdrawals, fees, daily_balances, source_file
    """
    bank = {
        'period':             '',
        'account':            '',
        'beginning_balance':  0.0,
        'ending_balance':     0.0,
        'total_deposits':     0.0,
        'total_checks':       0.0,
        'total_withdrawals':  0.0,
        'total_fees':         0.0,
        'deposits':           [],
        'checks':             [],
        'withdrawals':        [],
        'fees':               [],
        'daily_balances':     {},
        'source_file':        bank_file.name,
    }

    try:
        with pdfplumber.open(bank_file) as pdf:
            text = '\n'.join(p.extract_text() or '' for p in pdf.pages)
    except Exception as e:
        raise ValueError(f'Could not read PDF {bank_file.name}: {e}')

    def grab(pattern):
        m = re.search(pattern, text)
        try:
            return float(m.group(1).replace(',', '')) if m else 0.0
        except (ValueError, AttributeError):
            return 0.0

    # Period
    m = re.search(
        r'(January|February|March|April|May|June|July|August|'
        r'September|October|November|December)\s+\d+,\s+\d{4}'
        r'\s*through\s*\S+\s+\d+,\s+\d{4}', text)
    if m:
        bank['period'] = m.group(0)

    # Account number
    m = re.search(r'Account Number:\s+([\d]+)', text)
    if m:
        bank['account'] = m.group(1)

    # Summary balances
    bank['beginning_balance'] = grab(r'Beginning Balance\s+\$?([\d,]+\.\d{2})')
    bank['ending_balance']    = grab(r'Ending Balance\s+\d+\s+\$?([\d,]+\.\d{2})')
    bank['total_deposits']    = grab(r'Total Deposits and Additions\s+\$?([\d,]+\.\d{2})')
    bank['total_checks']      = grab(r'Total Checks Paid\s+\$?([\d,]+\.\d{2})')
    bank['total_fees']        = grab(r'Total Fees\s+\$?([\d,]+\.\d{2})')
    bank['total_withdrawals'] = grab(
        r'Total Electronic Withdrawals\s+\$?([\d,]+\.\d{2})'
    )
    if bank['total_withdrawals'] == 0.0:
        bank['total_withdrawals'] = grab(
            r'Total Other Withdrawals\s+\$?([\d,]+\.\d{2})'
        )
    if bank['total_withdrawals'] == 0.0:
        bank['total_withdrawals'] = grab(
            r'Electronic Withdrawals\s+\d+\s+-?([\d,]+\.\d{2})'
        )
    # Deposits
    dep = re.search(r'DEPOSITS AND ADDITIONS(.*?)CHECKS PAID', text, re.DOTALL)
    if dep:
        for m in re.finditer(
                r'(\d{2}/\d{2})\s+(.+?)\s+\$?([\d,]+\.\d{2})', dep.group(1)):
            bank['deposits'].append({
                'date':        m.group(1),
                'description': m.group(2).strip(),
                'amount':      float(m.group(3).replace(',', '')),
            })

    # Checks
    # Chase marks non-sequential ("gap") check numbers with a leading '*'
    # before the '^' - e.g. "1257 *^ 05/12 4,280.30" - which the old \^?-only
    # pattern didn't account for, silently dropping every such check.
    for m in re.finditer(
            r'(\d{4})\s+\*?\^?\s+(\d{2}/\d{2})\s+\$?([\d,]+\.\d{2})', text):
        bank['checks'].append({
            'check_no': m.group(1),
            'date':     m.group(2),
            'amount':   float(m.group(3).replace(',', '')),
        })

    # Other Withdrawals
    # No IGNORECASE here on purpose: Chase statements repeat "Electronic
    # Withdrawals" in title case as a one-line summary near the top (e.g.
    # "Electronic Withdrawals   2   -7,027.00") before the real itemized
    # section header, which is consistently ALL CAPS further down. Matching
    # case-insensitively grabs that earlier summary line instead - wrong
    # slice of text, wrong (empty) transaction list.
    wd = re.search(
        r'(?:OTHER WITHDRAWALS|ELECTRONIC WITHDRAWALS)(.*?)'
        r'(?:FEES|DAILY ENDING BALANCE|CHECKS PAID)',
        text, re.DOTALL)
    if wd:
        # $ (when present) sits directly against the digits with no space
        # of its own, e.g. "...CO Entry $37.00" - without \$? here, \s+
        # can never land right before the digit run whenever a $ prefix is
        # present, silently dropping that entry entirely (unlike deposits/
        # checks above, which already handle this with \$?).
        for m in re.finditer(
                r'(\d{2}/\d{2})\s+(.+?)\s+\$?([\d,]+\.\d{2})', wd.group(1)):
            bank['withdrawals'].append({
                'date':        m.group(1),
                'description': m.group(2).strip(),
                'amount':      float(m.group(3).replace(',', '')),
            })
        if bank['total_withdrawals'] == 0.0 and bank['withdrawals']:
            bank['total_withdrawals'] = sum(
                w['amount'] for w in bank['withdrawals'])

    # Fees
    fee = re.search(r'FEES(.*?)DAILY ENDING BALANCE', text, re.DOTALL)
    if fee:
        for m in re.finditer(
                r'(\d{2}/\d{2})\s+(.+?)\s+\$?([\d,]+\.\d{2})', fee.group(1)):
            bank['fees'].append({
                'date':        m.group(1),
                'description': m.group(2).strip(),
                'amount':      float(m.group(3).replace(',', '')),
            })

    # Daily balances
    bal = re.search(r'DAILY ENDING BALANCE(.*)$', text, re.DOTALL)
    if bal:
        for m in re.finditer(r'(\d{2}/\d{2})\s+([\d,]+\.\d{2})', bal.group(1)):
            bank['daily_balances'][m.group(1)] = float(m.group(2).replace(',', ''))

    print(f'  Bank      : {bank_file.name}')
    print(f'  Period    : {bank["period"]}')
    print(f'  Beginning : ${bank["beginning_balance"]:,.2f}  '
          f'Ending: ${bank["ending_balance"]:,.2f}')
    print(f'  Deposits      : {len(bank["deposits"])}  total=${bank["total_deposits"]:,.2f}')
    print(f'  Checks        : {len(bank["checks"])}  total=${bank["total_checks"]:,.2f}')
    print(f'  Withdrawals   : ${bank["total_withdrawals"]:,.2f}')
    print(f'  Fees          : ${bank["total_fees"]:,.2f}')
    return bank


# ── 4. MATCH TRANSACTION TO BANK STATEMENT MONTH ──────────────────────────────

def find_bank_statement_month(txn: dict, month_label: str,
                               bank_by_month: dict, fiscal_order: list):
    """
    A QuickBooks transaction is recorded under the month it was entered/
    categorized (`month_label`), but the actual bank deposit/check can clear
    in a *different* calendar month's Chase statement (e.g. a Givebacks
    payout QuickBooks logs in July may not post to the bank until August).
    This finds which month's statement actually contains it.

    Args:
        txn:           a transaction dict from parse_quickbooks_detail()['transactions']
                       (needs 'date' as MM/DD/YYYY, 'amount', 'is_income')
        month_label:   the QB-recorded month, e.g. 'July 2025'
        bank_by_month: {month_label: parse_chase_pdf() result} for every
                       processed month
        fiscal_order:  month_labels in true chronological order, used to find
                       "the next month" after month_label

    Returns:
        The matched month_label, or None if unmatched (render as "not yet
        reconciled" rather than silently assuming it's the QB month).

    Match strategy: exact date+amount match within the same month first
    (bank dates are 'MM/DD' only - the year is inferred from that month's own
    label). If nothing matches there, falls back to an amount-only match in
    the next chronological month, since the whole point of the lag case is
    that the date itself shifted into the next month - there's no reliable
    date to match against there. This means two unrelated transactions in
    adjacent months sharing the exact same dollar amount could in theory
    match incorrectly; for a treasurer's real transaction volume this is an
    acceptable tradeoff, not a guarantee.
    """
    try:
        txn_date = datetime.strptime(txn['date'], '%m/%d/%Y')
    except (ValueError, TypeError):
        return None
    amount = round(txn['amount'], 2)
    is_credit = txn.get('is_income', False)

    def bank_entries(bank):
        return bank.get('deposits', []) if is_credit else (
            bank.get('checks', []) + bank.get('withdrawals', []))

    def exact_match(bank, bank_month_label):
        try:
            year = int(bank_month_label.split()[1])
        except (IndexError, ValueError):
            return False
        for e in bank_entries(bank):
            try:
                e_date = datetime.strptime(f"{e['date']}/{year}", '%m/%d/%Y')
            except (ValueError, KeyError):
                continue
            if abs(e['amount'] - amount) < 0.01 and e_date.date() == txn_date.date():
                return True
        return False

    def amount_only_match(bank):
        return any(abs(e['amount'] - amount) < 0.01 for e in bank_entries(bank))

    # 1. Same month, exact date + amount
    bank = bank_by_month.get(month_label)
    if bank and exact_match(bank, month_label):
        return month_label

    # 2. Next chronological month, amount only (the lag case)
    if month_label in fiscal_order:
        idx = fiscal_order.index(month_label)
        if idx + 1 < len(fiscal_order):
            next_label = fiscal_order[idx + 1]
            next_bank = bank_by_month.get(next_label)
            if next_bank and amount_only_match(next_bank):
                return next_label

    return None


def match_credits_to_bank_statement(credits: list, month_label: str,
                                     bank_by_month: dict, fiscal_order: list):
    """
    Unlike debits (each check clears the bank as its own line - a natural
    1:1 relationship with find_bank_statement_month above), a single bank
    deposit is very often the sum of several QuickBooks-recorded income
    transactions that hit the bank the same day (e.g. one ACH Givebacks
    payout that QuickBooks splits across several income categories).
    Matching each credit transaction individually against bank deposits
    therefore misses almost everything - confirmed against real data: every
    QuickBooks credit transaction sharing an exact date summed to exactly
    one same-date bank deposit, with no exceptions in the months checked.

    Groups `credits` (transactions from parse_quickbooks_detail()['transactions']
    where is_income is True) by exact date, matches each group's total
    against a single bank deposit (same month first, then next month for the
    lag case - same search strategy as find_bank_statement_month), and
    mutates every transaction dict in the group with the result.

    Args:
        credits:       list of credit transaction dicts, mutated in place
        month_label:   the QB-recorded month, e.g. 'July 2025'
        bank_by_month: {month_label: parse_chase_pdf() result} for every
                       processed month
        fiscal_order:  month_labels in true chronological order
    """
    by_date = {}
    for t in credits:
        by_date.setdefault(t['date'], []).append(t)

    def exact_match(bank, bank_month_label, target_date, target_amount):
        try:
            year = int(bank_month_label.split()[1])
        except (IndexError, ValueError):
            return False
        for d in bank.get('deposits', []):
            try:
                d_date = datetime.strptime(f"{d['date']}/{year}", '%m/%d/%Y')
            except (ValueError, KeyError):
                continue
            if abs(d['amount'] - target_amount) < 0.01 and d_date.date() == target_date.date():
                return True
        return False

    def amount_only_match(bank, target_amount):
        return any(abs(d['amount'] - target_amount) < 0.01
                   for d in bank.get('deposits', []))

    for date_str, group in by_date.items():
        try:
            group_date = datetime.strptime(date_str, '%m/%d/%Y')
        except (ValueError, TypeError):
            for t in group:
                t['bank_statement_month'] = None
            continue

        group_sum = round(sum(t['amount'] for t in group), 2)
        matched_month = None

        bank = bank_by_month.get(month_label)
        if bank and exact_match(bank, month_label, group_date, group_sum):
            matched_month = month_label
        elif month_label in fiscal_order:
            idx = fiscal_order.index(month_label)
            if idx + 1 < len(fiscal_order):
                next_label = fiscal_order[idx + 1]
                next_bank = bank_by_month.get(next_label)
                if next_bank and amount_only_match(next_bank, group_sum):
                    matched_month = next_label

        for t in group:
            t['bank_statement_month'] = matched_month


# ── 5. CONSOLIDATE GIVEBACKS PAYOUTS INTO CREDITS ─────────────────────────────

def consolidate_givebacks_payouts(credit_txns: list, payouts: list):
    """
    QuickBooks splits a single Givebacks/MemberHub bank deposit across
    several income-category line items - Credits should show one row per
    actual payout, not one row per QuickBooks category split.

    Matching is by date-group total, not description or category: a real
    payout's QuickBooks entries almost always land on one date, but only a
    fraction of them carry a description tagging them as a Givebacks
    deposit - the rest of that same payout shows up as ordinary
    membership-dues category lines with other descriptions. Pre-filtering
    to the tagged description alone excludes most of a payout's real
    entries, so the date-group total almost never equals the payout total
    and consolidation silently fails - confirmed against real data, where
    every payout's full-day credit total matched its payout total to the
    cent, but the tagged-only subtotal did not. So every credit transaction
    on a date is a match candidate; the exact-to-the-cent total match
    against a specific real payout amount is itself precise enough to rule
    out a coincidental sweep-in of unrelated same-day income.

    Args:
        credit_txns: this month's credit transactions - must already have
                     'bank_statement_month' set (from
                     match_credits_to_bank_statement) before calling this
        payouts:     list of {'total': float, 'items': [...], 'source_file': str},
                     one per Givebacks CSV file for this month (parsed
                     individually, not merged - see parse_givebacks_files)

    A payout's QuickBooks entries usually land on one date, but two other
    cases show up in real data:
    - a date mixes a payout's entries with OTHER, unrelated income entered
      the same day (e.g. a $420 payout's 5 entries sharing a date with two
      unrelated $170 entries from something else) - the whole-date total
      then doesn't match anything, even though the payout's own entries are
      right there. Phase 1b looks for a subset of one date's entries that
      matches, leaving the rest of that date as-is.
    - a payout's entries are split across a couple of *different* dates -
      Phase 2 tries summing combinations of remaining whole date-groups (up
      to 5 at a time; a payout split across more dates than that is
      vanishingly unlikely and not worth the combinatorial cost).
    Every phase only accepts a match that's *unique* - if two different
    subsets/combinations both work, that's genuinely ambiguous and the
    payout is left unmatched rather than guessed at. A payout that matches
    nothing anywhere most commonly means it simply isn't recorded in
    QuickBooks at all yet - no amount of matching logic can fix a
    transaction that was never entered.

    Returns:
        (consolidated_credits, payout_dates)
        consolidated_credits: credit_txns with matched entries replaced by
            one consolidated row per payout; any entries that never matched
            a payout (individually, as part of a within-date subset, or in
            cross-date combination) pass through unchanged
        payout_dates: {payout_index: date_str_or_None} - None means no
            match was found anywhere for that payout (render as
            unreconciled in MemberHub_Summary, don't drop it)
    """
    by_date = {}
    for t in credit_txns:
        by_date.setdefault(t['date'], []).append(t)

    # Mutable "what's left to match" per date - phases remove entries as
    # they consume them, so later phases only ever see what's still
    # unclaimed (e.g. Phase 1b can claim 5 of a date's 7 entries and leave
    # the other 2 available for a different match).
    remaining = {date: list(group) for date, group in by_date.items()}
    payout_dates = {}
    consolidated_rows = []

    def add_consolidated_row(date, txns, total):
        consolidated_rows.append({
            'date':                 date,
            'category':             'MemberHub/Givebacks Deposit',
            'amount':               total,
            'is_income':            True,
            'bank_statement_month': txns[0].get('bank_statement_month'),
        })

    def add_consolidated_row_multi(dates, total):
        combined_txns = [t for d in dates for t in remaining[d]]
        earliest = min(dates, key=lambda d: datetime.strptime(d, '%m/%d/%Y'))
        consolidated_rows.append({
            'date':                 earliest,
            'category':             'MemberHub/Givebacks Deposit',
            'amount':               total,
            'is_income':            True,
            'bank_statement_month': combined_txns[0].get('bank_statement_month'),
        })
        return earliest

    # Phase 1: exact whole-date match.
    unmatched = []
    for i, payout in enumerate(payouts):
        payout_total = round(payout['total'], 2)
        matched_date = None
        for date, txns in remaining.items():
            if not txns:
                continue
            if abs(round(sum(t['amount'] for t in txns), 2) - payout_total) < 0.01:
                matched_date = date
                break
        if matched_date:
            add_consolidated_row(matched_date, remaining[matched_date], payout_total)
            payout_dates[i] = matched_date
            remaining[matched_date] = []
        else:
            payout_dates[i] = None
            unmatched.append(i)

    # Phase 1b: within-date subset match - a payout's own entries mixed in
    # with other, unrelated income on the same date. Capped at 12 remaining
    # entries per date to keep the subset search (2^n) cheap; a date with
    # more than that is rare enough not to be worth the combinatorial cost.
    still_unmatched = []
    for i in unmatched:
        payout_total = round(payouts[i]['total'], 2)
        found = None
        ambiguous = False
        for date, txns in remaining.items():
            n = len(txns)
            if n < 2 or n > 12:
                continue
            amounts = [t['amount'] for t in txns]
            local_matches = []
            for r in range(1, n):
                for combo in combinations(range(n), r):
                    if abs(round(sum(amounts[j] for j in combo), 2) - payout_total) < 0.01:
                        local_matches.append(combo)
                        if len(local_matches) > 1:
                            break
                if len(local_matches) > 1:
                    break
            if len(local_matches) == 1:
                if found is not None:
                    ambiguous = True
                    break
                found = (date, local_matches[0])
            elif len(local_matches) > 1:
                ambiguous = True
                break
        if found and not ambiguous:
            date, idxs = found
            idx_set = set(idxs)
            subset_txns = [t for j, t in enumerate(remaining[date]) if j in idx_set]
            add_consolidated_row(date, subset_txns, payout_total)
            payout_dates[i] = date
            remaining[date] = [t for j, t in enumerate(remaining[date]) if j not in idx_set]
        else:
            still_unmatched.append(i)

    # Phase 2: cross-date combination match for whatever's still unmatched
    # (a single payout whose QuickBooks entries landed on different dates).
    for i in still_unmatched:
        payout_total = round(payouts[i]['total'], 2)
        available = [d for d, txns in remaining.items() if txns]
        date_sums = {d: round(sum(t['amount'] for t in remaining[d]), 2) for d in available}
        matches = []
        for r in range(2, min(len(available), 5) + 1):
            for combo in combinations(available, r):
                if abs(round(sum(date_sums[d] for d in combo), 2) - payout_total) < 0.01:
                    matches.append(combo)
                    if len(matches) > 1:
                        break
            if len(matches) > 1:
                break
        if len(matches) == 1:
            combo = matches[0]
            payout_dates[i] = add_consolidated_row_multi(list(combo), payout_total)
            for d in combo:
                remaining[d] = []

    # Entries that never matched a payout (individually, as a within-date
    # subset, or in cross-date combination) pass through unchanged - safe
    # fallback, still visible, just not consolidated.
    for date, txns in remaining.items():
        consolidated_rows.extend(txns)

    return consolidated_rows, payout_dates


def extract_payout_id(filename: str):
    """
    Extracts the Givebacks payout ID from a filename like
    'givebacks_august_po_1Rql024TnYF4pDk8eQ4Q3ZZ8.csv' -> '1Rql024TnYF4pDk8eQ4Q3ZZ8'.

    The same payout file has been found copied into more than one month's
    input/{Month}_{Year}/givebacks/ folder in practice (e.g. a payout
    downloaded/placed under both July and August) - since each month is
    processed independently, that silently double-counts the payout across
    two months rather than raising an error. This lets a caller build a
    cross-month registry to catch it. Returns None for a filename that
    doesn't match the expected pattern (e.g. hand-renamed) - treat that as
    "can't check," not as an error.
    """
    m = re.search(r'_po_([A-Za-z0-9]+)\.csv$', filename, re.IGNORECASE)
    return m.group(1) if m else None


