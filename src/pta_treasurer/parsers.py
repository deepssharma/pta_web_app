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
                             fiscal_year: str = '') -> dict:
    """
    Parses QuickBooks Transaction Detail by Account CSV export.
    
    Args:
        folder:      Path to the month folder (e.g. input/July_2025/)
        input_month: Month name e.g. 'July' — used to find the right file
        fiscal_year: Year string e.g. '2025'
    
    Returns dict with keys:
        period, income, income_total, expenses, expense_total,
        net_income, transactions
    """
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
        'period':        '',
        'income':        {},
        'income_total':  0.0,
        'expenses':      {},
        'expense_total': 0.0,
        'net_income':    0.0,
        'transactions':  [],
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
        is_deposit = type_val.lower() == 'deposit'
        is_expense = type_val.lower() in ('check', 'expense', 'bill payment', 'bill')

        transaction = {
            'date':        date_val,
            'type':        type_val,
            'check_no':    num_val,
            'payee':       name_val,
            'description': clean_desc,
            'category':    current_category,
            'amount':      abs(amount),
            'is_income':   is_deposit,
        }
        data['transactions'].append(transaction)

        if is_deposit:
            data['income'][current_category] = (
                data['income'].get(current_category, 0.0) + abs(amount))
        elif is_expense:
            data['expenses'][current_category] = (
                data['expenses'].get(current_category, 0.0) + abs(amount))

    data['income_total']  = sum(data['income'].values())
    data['expense_total'] = sum(data['expenses'].values())
    data['net_income']    = data['income_total'] - data['expense_total']

    print(f'  Period    : {data["period"]}')
    print(f'  Income    : ${data["income_total"]:,.2f}  ({len(data["income"])} categories)')
    print(f'  Expenses  : ${data["expense_total"]:,.2f}  ({len(data["expenses"])} categories)')
    print(f'  Transactions: {len(data["transactions"])}')
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
    for m in re.finditer(
            r'(\d{4})\s+\^?\s+(\d{2}/\d{2})\s+\$?([\d,]+\.\d{2})', text):
        bank['checks'].append({
            'check_no': m.group(1),
            'date':     m.group(2),
            'amount':   float(m.group(3).replace(',', '')),
        })

    # Other Withdrawals
    wd = re.search(
        r'(?:OTHER WITHDRAWALS|ELECTRONIC WITHDRAWALS)(.*?)'
        r'(?:FEES|DAILY ENDING BALANCE|CHECKS PAID)',
        text, re.DOTALL | re.IGNORECASE)
    if wd:
        for m in re.finditer(
                r'(\d{2}/\d{2})\s+(.+?)\s+([\d,]+\.\d{2})', wd.group(1)):
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


