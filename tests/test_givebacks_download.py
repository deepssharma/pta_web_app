"""
Tests for the pure (no-browser, no-network) helpers in
givebacks_download.py: payout filtering and CSV row construction. The
Playwright-driving parts (login, DOM scraping, session persistence)
aren't covered here -- they need a real or heavily-mocked browser and
are exercised manually per packaging/README.md-style hand-testing, not
unit tests.
"""
from pta_treasurer.givebacks_download import (
    parse_summary_table_rows,
    select_paid_payout_ids,
    write_payout_csv,
)


def test_select_paid_payout_ids_filters_by_month_and_status():
    payouts = [
        {'uuid': 'a', 'arrival_date': '1999-07-05T00:00:00Z', 'status': 'paid'},
        {'uuid': 'b', 'arrival_date': '1999-08-01T00:00:00Z', 'status': 'paid'},
        {'uuid': 'c', 'arrival_date': '1999-07-15T00:00:00Z', 'status': 'pending'},
        {'uuid': 'd', 'arrival_date': '1999-07-31T00:00:00Z', 'status': 'paid'},
    ]
    result = select_paid_payout_ids(payouts, target_year=1999, target_month=7)
    assert result == ['a', 'd']


def test_select_paid_payout_ids_dedupes_preserving_order():
    payouts = [
        {'uuid': 'a', 'arrival_date': '1999-07-05', 'status': 'paid'},
        {'uuid': 'a', 'arrival_date': '1999-07-05', 'status': 'paid'},
        {'uuid': 'b', 'arrival_date': '1999-07-06', 'status': 'paid'},
    ]
    assert select_paid_payout_ids(payouts, 1999, 7) == ['a', 'b']


def test_select_paid_payout_ids_skips_missing_uuid():
    payouts = [{'arrival_date': '1999-07-05', 'status': 'paid'}]
    assert select_paid_payout_ids(payouts, 1999, 7) == []


def test_parse_summary_table_rows_builds_expected_dicts():
    raw_rows = [
        ['Item', 'Categories', 'No. of Transactions', 'Total'],  # header, skipped
        ['Book Fair', 'Fundraising', '10', '$110.14'],
        ['Total', '', '10', '$110.14'],  # total row, skipped
    ]
    result = parse_summary_table_rows(raw_rows)
    assert result == [
        {'Item': 'Book Fair', 'Categories': 'Fundraising',
         'No. of Transactions': 10, 'Total': '$110.14'},
    ]


def test_parse_summary_table_rows_skips_blank_item():
    raw_rows = [['', 'Fundraising', '10', '$110.14']]
    assert parse_summary_table_rows(raw_rows) == []


def test_parse_summary_table_rows_handles_bad_numbers_gracefully():
    raw_rows = [['Book Fair', 'Fundraising', 'n/a', 'n/a']]
    result = parse_summary_table_rows(raw_rows)
    assert result == [
        {'Item': 'Book Fair', 'Categories': 'Fundraising',
         'No. of Transactions': 0, 'Total': '$0.00'},
    ]


def test_parse_summary_table_rows_skips_rows_with_too_few_cells():
    raw_rows = [['Book Fair', 'Fundraising']]
    assert parse_summary_table_rows(raw_rows) == []


def test_write_payout_csv_writes_expected_content(tmp_path):
    rows_data = [
        {'Item': 'Book Fair', 'Categories': 'Fundraising',
         'No. of Transactions': 10, 'Total': '$110.14'},
    ]
    dest = tmp_path / 'givebacks'
    path = write_payout_csv(rows_data, dest, month_short='july', payout_id='abc123')

    assert path == dest / 'givebacks_july_abc123.csv'
    content = path.read_text()
    assert 'Item,Categories,No. of Transactions,Total' in content
    assert 'Book Fair,Fundraising,10,$110.14' in content


def test_write_payout_csv_creates_dest_folder(tmp_path):
    dest = tmp_path / 'nested' / 'givebacks'
    assert not dest.exists()
    write_payout_csv([{'Item': 'x', 'Categories': 'y',
                        'No. of Transactions': 1, 'Total': '$1.00'}],
                      dest, 'july', 'id1')
    assert dest.exists()
