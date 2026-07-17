"""
Tests for parsers.py helper functions
"""
import pytest
from pathlib import Path

from pta_treasurer.parsers import _parse_amount, _money_str
from pta_treasurer.parsers import parse_chase_pdf
from pta_treasurer.parsers import parse_givebacks_files


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
