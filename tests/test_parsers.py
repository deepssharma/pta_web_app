"""
Tests for parsers.py helper functions
"""
import pytest
from pathlib import Path
from unittest.mock import MagicMock, patch

from pta_treasurer.parsers import _parse_amount, _money_str
from pta_treasurer.parsers import parse_chase_pdf
from pta_treasurer.parsers import parse_givebacks_files
from pta_treasurer.parsers import find_bank_statement_month, match_credits_to_bank_statement
from pta_treasurer.parsers import consolidate_givebacks_payouts, extract_payout_id

FISCAL_ORDER = ['July 2025', 'August 2025', 'September 2025']


# ── Tests for _parse_amount ───────────────────────────────────────────────────

def test_parse_amount_basic():
    assert _parse_amount('$1,234.56') == 1234.56

def test_parse_amount_no_dollar_sign():
    assert _parse_amount('1234.56') == 1234.56

def test_parse_amount_with_comma():
    assert _parse_amount('$10,000.00') == 10000.0

def test_parse_amount_zero():
    assert _parse_amount('$0.00') == 0.0

def test_parse_amount_empty_string():
    assert _parse_amount('') == 0.0

def test_parse_amount_negative():
    assert _parse_amount('-$500.00') == -500.0

def test_parse_amount_invalid():
    assert _parse_amount('not a number') == 0.0

def test_parse_amount_with_quotes():
    assert _parse_amount('"$1,234.56"') == 1234.56

# ── Tests for _money_str ──────────────────────────────────────────────────────
def test_money_str_basic():
    assert _money_str(1234.56) == '$1,234.56'

def test_money_str_zero():
    assert _money_str(0.0) == '$0.00'

def test_money_str_large():
    assert _money_str(100000.00) == '$100,000.00'

# ── Tests for fiscal month logic ──────────────────────────────────────────────
# We'll test the fiscal index logic directly
# July=0, Aug=1, ... June=11
FISCAL_START_MONTH = 7

def calendar_to_fiscal(cal_month_num):
    return (cal_month_num - FISCAL_START_MONTH) % 12

def test_fiscal_july_is_zero():
    assert calendar_to_fiscal(7) == 0

def test_fiscal_august_is_one():
    assert calendar_to_fiscal(8) == 1

def test_fiscal_june_is_eleven():
    assert calendar_to_fiscal(6) == 11

def test_fiscal_january_is_six():
    assert calendar_to_fiscal(1) == 6

def test_fiscal_december_is_five():
    assert calendar_to_fiscal(12) == 5


# ── Tests for CSV robust reader ───────────────────────────────────────────────
import tempfile, os
from pta_treasurer.parsers import _read_csv_robust

def test_read_csv_utf8():
    with tempfile.NamedTemporaryFile(mode='wb', suffix='.csv', delete=False) as f:
        f.write(b'Item,Amount\nBook Fair,$100.00\n')
        tmp = f.name
    try:
        lines, encoding = _read_csv_robust(Path(tmp))
        assert len(lines) == 2
        assert lines[0] == ['Item', 'Amount']
        assert lines[1] == ['Book Fair', '$100.00']
        assert 'utf' in encoding.lower()
    finally:
        os.unlink(tmp)

def test_read_csv_with_nul_bytes():
    with tempfile.NamedTemporaryFile(mode='wb', suffix='.csv', delete=False) as f:
        # Simulate QuickBooks export with NUL bytes
        f.write(b'Item\x00,Amount\x00\nBook Fair\x00,$100.00\x00\n')
        tmp = f.name
    try:
        lines, encoding = _read_csv_robust(Path(tmp))
        assert len(lines) == 2
        assert lines[0][0] == 'Item'
    finally:
        os.unlink(tmp)

def test_read_csv_mixed_line_endings():
    with tempfile.NamedTemporaryFile(mode='wb', suffix='.csv', delete=False) as f:
        f.write(b'Item,Amount\r\nBook Fair,$100.00\r\n')
        tmp = f.name
    try:
        lines, encoding = _read_csv_robust(Path(tmp))
        assert len(lines) == 2
    finally:
        os.unlink(tmp)

def test_read_csv_file_not_found():
    import pytest
    with pytest.raises(Exception):
        _read_csv_robust(Path('nonexistent_file.csv'))


# ── Tests for Givebacks row processing ───────────────────────────────────────
from pta_treasurer.parsers import _process_givebacks_row

def test_givebacks_row_basic():
    merged = {}
    row = {'Item': 'Book Fair', 'Categories': 'Fundraiser',
           'No. of Transactions': '5', 'Total': '$250.00'}
    _process_givebacks_row(row, 'test.csv', merged)
    assert 'Book Fair' in merged
    assert merged['Book Fair']['total'] == 250.0
    assert merged['Book Fair']['count'] == 5

def test_givebacks_row_merges_duplicates():
    merged = {}
    row1 = {'Item': 'Book Fair', 'Categories': 'Fundraiser',
            'No. of Transactions': '3', 'Total': '$150.00'}
    row2 = {'Item': 'Book Fair', 'Categories': 'Fundraiser',
            'No. of Transactions': '2', 'Total': '$100.00'}
    _process_givebacks_row(row1, 'file1.csv', merged)
    _process_givebacks_row(row2, 'file2.csv', merged)
    assert merged['Book Fair']['total'] == 250.0
    assert merged['Book Fair']['count'] == 5
    assert len(merged['Book Fair']['source_files']) == 2

def test_givebacks_row_empty_item_skipped():
    merged = {}
    row = {'Item': '', 'Categories': 'Fundraiser',
           'No. of Transactions': '1', 'Total': '$50.00'}
    _process_givebacks_row(row, 'test.csv', merged)
    assert len(merged) == 0

def test_givebacks_row_invalid_amount():
    merged = {}
    row = {'Item': 'Book Fair', 'Categories': 'Fundraiser',
           'No. of Transactions': '1', 'Total': 'N/A'}
    _process_givebacks_row(row, 'test.csv', merged)
    assert merged['Book Fair']['total'] == 0.0

from pta_treasurer.parsers import parse_quickbooks_detail

FIXTURES = Path(__file__).parent / 'fixtures'

def test_parse_quickbooks_detail_expenses():
    result = parse_quickbooks_detail(FIXTURES, 'July', '2025')
    assert result['expense_total'] == 1439.34
    assert 'Accounting Expense (Quickbooks)' in result['expenses']
    assert result['expenses']['Accounting Expense (Quickbooks)'] == 1257.76

def test_parse_quickbooks_detail_no_income():
    result = parse_quickbooks_detail(FIXTURES, 'July', '2025')
    assert result['income_total'] == 0.0
    assert result['income'] == {}

def test_parse_quickbooks_detail_net_income():
    result = parse_quickbooks_detail(FIXTURES, 'July', '2025')
    assert result['net_income'] == -1439.34

def test_parse_quickbooks_detail_period():
    result = parse_quickbooks_detail(FIXTURES, 'July', '2025')
    assert 'July' in result['period']

def test_parse_quickbooks_detail_transactions_count():
    result = parse_quickbooks_detail(FIXTURES, 'July', '2025')
    assert len(result['transactions']) == 2

def test_parse_quickbooks_detail_transaction_types():
    result = parse_quickbooks_detail(FIXTURES, 'July', '2025')
    types = [t['type'] for t in result['transactions']]
    assert 'Expense' in types
    assert 'Check' in types

def test_parse_quickbooks_wrong_month_raises():
    import pytest
    with pytest.raises(FileNotFoundError):
        parse_quickbooks_detail(FIXTURES, 'August', '2025')


def test_parse_chase_pdf_balances():
    result = parse_chase_pdf(FIXTURES / 'Chase_july_2025.pdf')
    assert result['beginning_balance'] == 32630.10
    assert result['ending_balance']    == 31190.76

def test_parse_chase_pdf_period():
    result = parse_chase_pdf(FIXTURES / 'Chase_july_2025.pdf')
    assert 'July' in result['period']

def test_parse_chase_pdf_checks():
    result = parse_chase_pdf(FIXTURES / 'Chase_july_2025.pdf')
    assert result['total_checks'] == 181.58
    assert len(result['checks'])  == 1
    assert result['checks'][0]['check_no'] == '1077'

def test_parse_chase_pdf_withdrawals():
    result = parse_chase_pdf(FIXTURES / 'Chase_july_2025.pdf')
    assert result['total_withdrawals'] == 1257.76

def test_parse_chase_pdf_no_deposits():
    result = parse_chase_pdf(FIXTURES / 'Chase_july_2025.pdf')
    assert result['total_deposits'] == 0.0
    assert result['deposits']       == []

def test_parse_chase_pdf_reconciliation():
    result = parse_chase_pdf(FIXTURES / 'Chase_july_2025.pdf')
    calc = (result['beginning_balance']
            + result['total_deposits']
            - result['total_checks']
            - result['total_withdrawals']
            - result['total_fees'])
    assert abs(calc - result['ending_balance']) < 0.01

def test_parse_chase_pdf_source_file():
    result = parse_chase_pdf(FIXTURES / 'Chase_july_2025.pdf')
    assert result['source_file'] == 'Chase_july_2025.pdf'

def test_parse_chase_pdf_file_not_found():
    import pytest
    with pytest.raises(Exception):
        parse_chase_pdf(Path('nonexistent.pdf'))

def test_parse_givebacks_files_basic():
    gb_file = FIXTURES / 'givebacks_july_po_1Rql024TnYF4pDk8eQ4Q3ZZ8.csv'
    file_info = [(gb_file, 'July 2025', 0)]
    result = parse_givebacks_files(file_info)
    assert len(result) > 0

def test_parse_givebacks_files_total():
    gb_file = FIXTURES / 'givebacks_july_po_1Rql024TnYF4pDk8eQ4Q3ZZ8.csv'
    file_info = [(gb_file, 'July 2025', 0)]
    result = parse_givebacks_files(file_info)
    total = sum(r['total'] for r in result)
    assert total == 110.14  # $95.14 + $15.00

def test_parse_givebacks_files_item_names():
    gb_file = FIXTURES / 'givebacks_july_po_1Rql024TnYF4pDk8eQ4Q3ZZ8.csv'
    file_info = [(gb_file, 'July 2025', 0)]
    result = parse_givebacks_files(file_info)
    items = [r['item'] for r in result]
    assert 'Shop to Give Donation' in items
    assert 'Teacher/Staff' in items

def test_parse_givebacks_files_source_file():
    gb_file = FIXTURES / 'givebacks_july_po_1Rql024TnYF4pDk8eQ4Q3ZZ8.csv'
    file_info = [(gb_file, 'July 2025', 0)]
    result = parse_givebacks_files(file_info)
    assert all('givebacks_july_po_1Rql024TnYF4pDk8eQ4Q3ZZ8.csv' in r['source_file'] for r in result)

def test_parse_givebacks_empty_file_list():
     with pytest.raises(ValueError, match='No Givebacks files provided'):
        parse_givebacks_files([])


# ── Regression tests ported from the private notebook's later bug fixes ────
# (see ~/Desktop/pta/codes/pta_treasurer commit a443a09 for the originals)

def _mock_chase_pdf(text):
    fake_page = MagicMock()
    fake_page.extract_text.return_value = text
    fake_pdf = MagicMock()
    fake_pdf.__enter__.return_value = fake_pdf
    fake_pdf.__exit__.return_value = False
    fake_pdf.pages = [fake_page]
    return patch('pta_treasurer.parsers.pdfplumber.open', return_value=fake_pdf)


def test_parse_chase_pdf_check_with_gap_marker():
    """Chase marks non-sequential check numbers with a leading '*' before
    the '^' (e.g. '1257 *^ 05/12 4,280.30') - must not be silently dropped."""
    text = (
        "CHECKS PAID\n"
        "1243 ^ 05/12 1,255.00\n"
        "1257 *^ 05/12 4,280.30\n"
        "ELECTRONIC WITHDRAWALS\n"
    )
    with _mock_chase_pdf(text):
        result = parse_chase_pdf(Path('fake.pdf'))
    check_numbers = [c['check_no'] for c in result['checks']]
    assert '1243' in check_numbers
    assert '1257' in check_numbers
    gap_check = next(c for c in result['checks'] if c['check_no'] == '1257')
    assert gap_check['amount'] == 4280.30


def test_parse_chase_pdf_withdrawals_skips_summary_line_duplicate():
    """'Electronic Withdrawals' appears twice: once as a title-case one-line
    summary near the top, once as the real ALL CAPS itemized section header
    further down. Must parse from the real header, not the summary line."""
    text = (
        "Electronic Withdrawals   1   -37.00\n"  # summary line (title case)
        "CHECKS PAID\n"
        "1243 ^ 05/12 1,255.00\n"
        "ELECTRONIC WITHDRAWALS\n"                # real section header (caps)
        "05/26 Orig Co Name Some Vendor Desc $37.00\n"
        "DAILY ENDING BALANCE\n"
    )
    with _mock_chase_pdf(text):
        result = parse_chase_pdf(Path('fake.pdf'))
    assert len(result['withdrawals']) == 1
    assert result['withdrawals'][0]['amount'] == 37.00


# ── Tests for pass-through fund handling ────────────────────────────────────
# (generalized from the private notebook's "READTHON" feature — commit c348462)

def test_parse_quickbooks_no_pass_through_defaults_to_zero():
    result = parse_quickbooks_detail(FIXTURES, 'July', '2025')
    assert result['pass_through_income_total']  == 0.0
    assert result['pass_through_expense_total'] == 0.0
    assert result['pass_through_net']            == 0.0


def test_parse_quickbooks_pass_through_excluded_from_income_expenses(sample_qb_csv_with_pass_through):
    result = parse_quickbooks_detail(
        sample_qb_csv_with_pass_through.parent, 'July', '2025',
        pass_through_categories={'READTHON'}, pass_through_fund_name='READTHON')
    assert 'Readthon' not in result['income']
    assert 'Readthon' not in result['expenses']
    # Unaffected by the pass-through rows - only the "Accounting Expense
    # (Quickbooks)" category section counts; "Checking (4346)" is always
    # skipped by design (same as with no pass-through fund configured).
    assert result['income_total']  == 0.0
    assert result['expense_total'] == 1257.76


def test_parse_quickbooks_pass_through_totals(sample_qb_csv_with_pass_through):
    result = parse_quickbooks_detail(
        sample_qb_csv_with_pass_through.parent, 'July', '2025',
        pass_through_categories={'READTHON'}, pass_through_fund_name='READTHON')
    assert result['pass_through_income_total']  == 500.0
    assert result['pass_through_expense_total'] == 300.0
    assert result['pass_through_net']            == 200.0


def test_parse_quickbooks_pass_through_transactions_flagged(sample_qb_csv_with_pass_through):
    result = parse_quickbooks_detail(
        sample_qb_csv_with_pass_through.parent, 'July', '2025',
        pass_through_categories={'READTHON'}, pass_through_fund_name='READTHON')
    assert all('is_pass_through' in t for t in result['transactions'])
    pt_txns = [t for t in result['transactions'] if t['is_pass_through']]
    other_txns = [t for t in result['transactions'] if not t['is_pass_through']]
    assert len(pt_txns) == 2
    # Only the "Accounting Expense (Quickbooks)" row - the "Checking (4346)"
    # section's own listing is always skipped by design
    assert len(other_txns) == 1


def test_parse_quickbooks_pass_through_category_variant(tmp_path):
    """QuickBooks splits a fund's deposits/payouts into separately named
    sections in real exports - a second configured category name (e.g.
    'READTHON-EXPENSES') must be caught too, not just the primary one."""
    content = '''Setauket School PTA
Transaction Detail by Account
May 1-31, 2026
,,,,,,,,
,,Transaction date,Transaction type,Num,Name,Description,Split,Amount
Readthon,,,,,,,,
,,05/05/2026,Deposit,,Family B,Readathon pledge,Checking,400.00
Total for Readthon,,,,,,,,
Readthon-Expenses,,,,,,,,
,,05/15/2026,Check,3001,Middle School,Readathon payout,Checking,-250.00
Total for Readthon-Expenses,,,,,,,,
'''
    f = tmp_path / 'quickbooks_may_2026.csv'
    f.write_text(content, encoding='utf-8')
    result = parse_quickbooks_detail(
        tmp_path, 'May', '2026',
        pass_through_categories={'READTHON', 'READTHON-EXPENSES'},
        pass_through_fund_name='READTHON')
    assert 'Readthon' not in result['income']
    assert 'Readthon-Expenses' not in result['expenses']
    assert result['income_total']  == 0.0
    assert result['expense_total'] == 0.0
    assert result['pass_through_income_total']  == 400.0
    assert result['pass_through_expense_total'] == 250.0
    assert result['pass_through_net']            == 150.0
    assert all(t['is_pass_through'] for t in result['transactions'])


# ── Tests for find_bank_statement_month / match_credits_to_bank_statement ───
# (ported from the private notebook's Debits & Credits ledger work — commit a443a09)

def test_find_bank_statement_month_same_month_credit():
    txn = {'date': '07/30/2025', 'amount': 110.14, 'is_income': True}
    bank_by_month = {
        'July 2025': {'deposits': [{'date': '07/30', 'amount': 110.14}],
                      'checks': [], 'withdrawals': []},
    }
    result = find_bank_statement_month(txn, 'July 2025', bank_by_month, FISCAL_ORDER)
    assert result == 'July 2025'


def test_find_bank_statement_month_next_month_lag_credit():
    txn = {'date': '07/30/2025', 'amount': 110.14, 'is_income': True}
    bank_by_month = {
        'July 2025':   {'deposits': [], 'checks': [], 'withdrawals': []},
        'August 2025': {'deposits': [{'date': '08/01', 'amount': 110.14}],
                        'checks': [], 'withdrawals': []},
    }
    result = find_bank_statement_month(txn, 'July 2025', bank_by_month, FISCAL_ORDER)
    assert result == 'August 2025'


def test_find_bank_statement_month_unmatched():
    txn = {'date': '07/30/2025', 'amount': 999.99, 'is_income': True}
    bank_by_month = {
        'July 2025':   {'deposits': [{'date': '07/30', 'amount': 110.14}],
                        'checks': [], 'withdrawals': []},
        'August 2025': {'deposits': [{'date': '08/01', 'amount': 50.00}],
                        'checks': [], 'withdrawals': []},
    }
    result = find_bank_statement_month(txn, 'July 2025', bank_by_month, FISCAL_ORDER)
    assert result is None


def test_find_bank_statement_month_same_month_debit_check():
    txn = {'date': '07/01/2025', 'amount': 181.58, 'is_income': False}
    bank_by_month = {
        'July 2025': {'deposits': [], 'checks': [{'date': '07/01', 'amount': 181.58}],
                      'withdrawals': []},
    }
    result = find_bank_statement_month(txn, 'July 2025', bank_by_month, FISCAL_ORDER)
    assert result == 'July 2025'


def test_find_bank_statement_month_debit_matches_withdrawals():
    txn = {'date': '07/17/2025', 'amount': 1257.76, 'is_income': False}
    bank_by_month = {
        'July 2025': {'deposits': [], 'checks': [],
                      'withdrawals': [{'date': '07/17', 'amount': 1257.76}]},
    }
    result = find_bank_statement_month(txn, 'July 2025', bank_by_month, FISCAL_ORDER)
    assert result == 'July 2025'


def test_find_bank_statement_month_no_next_month_in_order():
    # Last month in fiscal_order - no "next month" to fall back to
    txn = {'date': '09/15/2025', 'amount': 42.00, 'is_income': True}
    bank_by_month = {'September 2025': {'deposits': [], 'checks': [], 'withdrawals': []}}
    result = find_bank_statement_month(txn, 'September 2025', bank_by_month, FISCAL_ORDER)
    assert result is None


def test_match_credits_to_bank_statement_same_day_group_sums_to_deposit():
    credits = [
        {'date': '07/10/2025', 'amount': 95.14, 'is_income': True},
        {'date': '07/10/2025', 'amount': 15.00, 'is_income': True},
    ]
    bank_by_month = {
        'July 2025': {'deposits': [{'date': '07/10', 'amount': 110.14}]},
    }
    match_credits_to_bank_statement(credits, 'July 2025', bank_by_month, FISCAL_ORDER)
    assert all(t['bank_statement_month'] == 'July 2025' for t in credits)


def test_match_credits_to_bank_statement_next_month_lag_group():
    credits = [
        {'date': '07/30/2025', 'amount': 60.0, 'is_income': True},
        {'date': '07/30/2025', 'amount': 40.0, 'is_income': True},
    ]
    bank_by_month = {
        'July 2025':   {'deposits': []},
        'August 2025': {'deposits': [{'date': '08/01', 'amount': 100.0}]},
    }
    match_credits_to_bank_statement(credits, 'July 2025', bank_by_month, FISCAL_ORDER)
    assert all(t['bank_statement_month'] == 'August 2025' for t in credits)


def test_match_credits_to_bank_statement_different_dates_grouped_separately():
    credits = [
        {'date': '07/01/2025', 'amount': 50.0, 'is_income': True},
        {'date': '07/02/2025', 'amount': 50.0, 'is_income': True},
    ]
    bank_by_month = {
        'July 2025': {'deposits': [
            {'date': '07/01', 'amount': 50.0},
            {'date': '07/02', 'amount': 50.0},
        ]},
    }
    match_credits_to_bank_statement(credits, 'July 2025', bank_by_month, FISCAL_ORDER)
    # each date is its own group, matched independently (not summed together)
    assert all(t['bank_statement_month'] == 'July 2025' for t in credits)


def test_match_credits_to_bank_statement_unmatched_group():
    credits = [{'date': '07/10/2025', 'amount': 999.99, 'is_income': True}]
    bank_by_month = {'July 2025': {'deposits': [{'date': '07/10', 'amount': 110.14}]}}
    match_credits_to_bank_statement(credits, 'July 2025', bank_by_month, FISCAL_ORDER)
    assert credits[0]['bank_statement_month'] is None


def test_match_credits_to_bank_statement_single_transaction_group():
    # Degenerate case: a "group" of one still works the same as the plain
    # per-transaction match.
    credits = [{'date': '08/05/2025', 'amount': 50.0, 'is_income': True}]
    bank_by_month = {'August 2025': {'deposits': [{'date': '08/05', 'amount': 50.0}]}}
    match_credits_to_bank_statement(credits, 'August 2025', bank_by_month, FISCAL_ORDER)
    assert credits[0]['bank_statement_month'] == 'August 2025'


# ── Tests for consolidate_givebacks_payouts ─────────────────────────────────
# (ported from the private notebook's Debits & Credits ledger work — commits
# a443a09/bc94cce/793c380)

def _gb_txn(date, amount, description='MemberHub/Givebacks Deposit', category='Membership'):
    return {'date': date, 'amount': amount, 'is_income': True,
            'description': description, 'category': category,
            'bank_statement_month': 'July 2025'}


def test_consolidate_givebacks_payouts_matches_and_consolidates():
    credits = [
        _gb_txn('07/10/2025', 95.14),
        _gb_txn('07/10/2025', 15.00),
    ]
    payouts = [{'total': 110.14, 'items': [{'item': 'Donations', 'total': 95.14},
                                            {'item': 'Teachers', 'total': 15.00}],
                'source_file': 'p1.csv'}]
    result, payout_dates = consolidate_givebacks_payouts(credits, payouts)
    assert len(result) == 1
    assert result[0]['amount'] == 110.14
    assert result[0]['category'] == 'MemberHub/Givebacks Deposit'
    assert result[0]['date'] == '07/10/2025'
    assert payout_dates == {0: '07/10/2025'}


def test_consolidate_givebacks_payouts_unmatched_payout_keeps_none():
    credits = [_gb_txn('07/10/2025', 50.0)]  # doesn't sum to any payout below
    payouts = [{'total': 999.99, 'items': [], 'source_file': 'p1.csv'}]
    result, payout_dates = consolidate_givebacks_payouts(credits, payouts)
    assert payout_dates == {0: None}
    # the unmatched Givebacks-tagged transaction passes through unconsolidated
    assert len(result) == 1
    assert result[0]['amount'] == 50.0


def test_consolidate_givebacks_payouts_matches_regardless_of_description():
    # Real payouts split across QuickBooks category lines that DON'T carry
    # the 'MemberHub/Givebacks Deposit' description tag (e.g. membership
    # dues categories) - matching is by date-group total, not description,
    # so these must still consolidate.
    credits = [
        {'date': '07/05/2025', 'amount': 150.0, 'is_income': True,
         'description': 'Family', 'category': 'Family',
         'bank_statement_month': 'July 2025'},
        {'date': '07/05/2025', 'amount': 50.0, 'is_income': True,
         'description': 'MemberHub/Givebacks Deposit', 'category': 'Membership',
         'bank_statement_month': 'July 2025'},
    ]
    payouts = [{'total': 200.0, 'items': [], 'source_file': 'p1.csv'}]
    result, payout_dates = consolidate_givebacks_payouts(credits, payouts)
    assert payout_dates == {0: '07/05/2025'}
    assert len(result) == 1
    assert result[0]['amount'] == 200.0
    assert result[0]['category'] == 'MemberHub/Givebacks Deposit'


def test_consolidate_givebacks_payouts_no_matching_payout_untouched():
    credits = [
        {'date': '07/05/2025', 'amount': 200.0, 'is_income': True,
         'description': 'Some Other Deposit', 'category': 'Sponsors',
         'bank_statement_month': 'July 2025'},
    ]
    payouts = [{'total': 999.99, 'items': [], 'source_file': 'p1.csv'}]
    result, payout_dates = consolidate_givebacks_payouts(credits, payouts)
    # doesn't match this payout's total, so it's left exactly as-is
    assert payout_dates == {0: None}
    assert len(result) == 1
    assert result[0]['description'] == 'Some Other Deposit'
    assert result[0]['category'] == 'Sponsors'


def test_consolidate_givebacks_payouts_multiple_payouts_matched_independently():
    credits = [
        _gb_txn('08/01/2025', 40.0), _gb_txn('08/01/2025', 10.0),   # group A: 50.0
        _gb_txn('08/05/2025', 30.0), _gb_txn('08/05/2025', 30.0),   # group B: 60.0
    ]
    payouts = [
        {'total': 60.0, 'items': [], 'source_file': 'p_b.csv'},
        {'total': 50.0, 'items': [], 'source_file': 'p_a.csv'},
    ]
    result, payout_dates = consolidate_givebacks_payouts(credits, payouts)
    assert payout_dates[0] == '08/05/2025'  # matched the 60.0 group
    assert payout_dates[1] == '08/01/2025'  # matched the 50.0 group
    amounts = sorted(r['amount'] for r in result)
    assert amounts == [50.0, 60.0]


def test_consolidate_givebacks_payouts_cross_date_combination_match():
    # Payout total (110.14) matches no single date, only 08/01 + 08/08 combined
    credits = [
        _gb_txn('08/01/2025', 95.14),
        _gb_txn('08/08/2025', 15.00),
    ]
    payouts = [{'total': 110.14, 'items': [], 'source_file': 'p1.csv'}]
    result, payout_dates = consolidate_givebacks_payouts(credits, payouts)
    assert payout_dates[0] == '08/01/2025'  # earliest of the combined dates
    assert len(result) == 1
    assert result[0]['amount'] == 110.14


def test_consolidate_givebacks_payouts_ambiguous_combination_stays_unmatched():
    # Two different date-groups both sum to 100.0 in different ways - genuinely
    # ambiguous, must not guess.
    credits = [
        _gb_txn('08/01/2025', 40.0),
        _gb_txn('08/08/2025', 60.0),
        _gb_txn('08/15/2025', 30.0),
        _gb_txn('08/22/2025', 70.0),
    ]
    payouts = [{'total': 100.0, 'items': [], 'source_file': 'p1.csv'}]
    result, payout_dates = consolidate_givebacks_payouts(credits, payouts)
    assert payout_dates[0] is None
    # nothing consolidated - all 4 original transactions pass through unchanged
    assert len(result) == 4


def test_consolidate_givebacks_payouts_within_date_subset_match():
    # Real scenario (October 2025): a payout's own entries share a date with
    # OTHER, unrelated income also entered that day, so the whole-date total
    # doesn't match - only a subset of that date's entries does.
    credits = [
        _gb_txn('10/03/2025', 100.0, description='Graduating Class Dues', category='Graduating Class Dues'),
        _gb_txn('10/03/2025', 60.0,  description='Standard', category='Standard'),
        _gb_txn('10/03/2025', 5.0,   description='Student', category='Student'),
        _gb_txn('10/03/2025', 60.0,  description='Teacher/Staff', category='Teachers'),
        _gb_txn('10/03/2025', 195.0, description='Birthday Books-Income', category='Birthday Books-Income'),
        _gb_txn('10/03/2025', 170.0, description='Grades 2-3', category='Fast Athletics'),   # unrelated
        _gb_txn('10/03/2025', 170.0, description='Trunk or Treat', category='Trunk or Treat'),  # unrelated
    ]
    payouts = [{'total': 420.0, 'items': [], 'source_file': 'p1.csv'}]
    result, payout_dates = consolidate_givebacks_payouts(credits, payouts)
    assert payout_dates[0] == '10/03/2025'
    consolidated = [r for r in result if r.get('category') == 'MemberHub/Givebacks Deposit']
    assert len(consolidated) == 1
    assert consolidated[0]['amount'] == 420.0
    # the two unrelated entries pass through untouched, not swept in
    leftover = sorted(r['amount'] for r in result if r is not consolidated[0])
    assert leftover == [170.0, 170.0]


def test_consolidate_givebacks_payouts_within_date_subset_ambiguous_stays_unmatched():
    # Two different subsets of the same date both sum to 50.0 - genuinely
    # ambiguous, must not guess.
    credits = [
        _gb_txn('10/03/2025', 10.0),
        _gb_txn('10/03/2025', 40.0),
        _gb_txn('10/03/2025', 20.0),
        _gb_txn('10/03/2025', 30.0),
        _gb_txn('10/03/2025', 900.0),  # padding so the whole-date total doesn't also match
    ]
    payouts = [{'total': 50.0, 'items': [], 'source_file': 'p1.csv'}]
    result, payout_dates = consolidate_givebacks_payouts(credits, payouts)
    assert payout_dates[0] is None
    assert len(result) == 5


def test_consolidate_givebacks_payouts_no_qb_entry_at_all_stays_unmatched():
    # Real scenario (October 2025): a Givebacks payout/reversal with no
    # corresponding QuickBooks entry anywhere that month - matching logic
    # can't invent a transaction that was never entered.
    credits = [_gb_txn('10/14/2025', 45.0), _gb_txn('10/14/2025', 15.0)]
    payouts = [{'total': -255.0, 'items': [], 'source_file': 'reversal.csv'}]
    result, payout_dates = consolidate_givebacks_payouts(credits, payouts)
    assert payout_dates[0] is None
    assert len(result) == 2


def test_extract_payout_id_basic():
    assert extract_payout_id('givebacks_august_po_1Rql024TnYF4pDk8eQ4Q3ZZ8.csv') == '1Rql024TnYF4pDk8eQ4Q3ZZ8'


def test_extract_payout_id_different_month_same_id():
    # the exact scenario that causes a cross-month duplicate: same id, two
    # different month prefixes
    july_id = extract_payout_id('givebacks_july_po_1Rql024TnYF4pDk8eQ4Q3ZZ8.csv')
    august_id = extract_payout_id('givebacks_august_po_1Rql024TnYF4pDk8eQ4Q3ZZ8.csv')
    assert july_id == august_id == '1Rql024TnYF4pDk8eQ4Q3ZZ8'


def test_extract_payout_id_no_match_returns_none():
    assert extract_payout_id('some_other_file.csv') is None
    assert extract_payout_id('givebacks_july.csv') is None


def test_parse_chase_pdf_withdrawal_amount_with_dollar_sign():
    """A '$' sitting directly against the digits (no space of its own)
    must not prevent the amount from matching."""
    text = (
        "CHECKS PAID\n"
        "ELECTRONIC WITHDRAWALS\n"
        "05/26 Orig Co Name Some Vendor Desc $37.00\n"
        "05/27 Orig Co Name Other Vendor Desc 6,990.00\n"
        "DAILY ENDING BALANCE\n"
    )
    with _mock_chase_pdf(text):
        result = parse_chase_pdf(Path('fake.pdf'))
    amounts = sorted(w['amount'] for w in result['withdrawals'])
    assert amounts == [37.00, 6990.00]
