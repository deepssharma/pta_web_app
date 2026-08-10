"""
Tests for budget_io.py — the Excel-template budget config.
"""
import openpyxl
import pytest

from pta_treasurer.budget_io import (
    SHEET_NAMES, apply_dynamic_last_year, generate_template, load_budget,
    map_actuals_to_budget_items, merge_actuals_into_budget,
)


# ── generate_template ─────────────────────────────────────────────────────

def test_generate_template_creates_both_sheets(tmp_path):
    path = tmp_path / 'budget.xlsx'
    generate_template(path)
    wb = openpyxl.load_workbook(path)
    assert wb.sheetnames == list(SHEET_NAMES)


def test_generate_template_has_column_headers(tmp_path):
    path = tmp_path / 'budget.xlsx'
    generate_template(path)
    wb = openpyxl.load_workbook(path)
    header_row = [c.value for c in wb['Income Budget'][2]]
    assert header_row == [
        'Section', 'Item', 'QuickBooks Category Name(s)',
        'Last Year Actual', 'This Year Budget',
    ]


# ── load_budget ───────────────────────────────────────────────────────────

def _fill_budget(path, income_rows, expense_rows):
    generate_template(path)
    wb = openpyxl.load_workbook(path)
    for sheet_name, rows in [('Income Budget', income_rows), ('Expense Budget', expense_rows)]:
        ws = wb[sheet_name]
        for i, row in enumerate(rows, start=3):
            for ci, val in enumerate(row, start=1):
                ws.cell(row=i, column=ci, value=val)
    wb.save(path)


def test_load_budget_parses_rows(tmp_path):
    path = tmp_path / 'budget.xlsx'
    _fill_budget(
        path,
        income_rows=[('Fundraising', 'Book Fair', '', 9118.36, 500.0)],
        expense_rows=[('Admin/General', 'Bank Services', '', 234.94, 200.0)],
    )
    income_budget, expense_budget, qb_map = load_budget(path)
    assert income_budget == {'Fundraising': {'Book Fair': (9118.36, 500.0)}}
    assert expense_budget == {'Admin/General': {'Bank Services': (234.94, 200.0)}}


def test_load_budget_builds_qb_map_with_multiple_names(tmp_path):
    path = tmp_path / 'budget.xlsx'
    _fill_budget(
        path,
        income_rows=[],
        expense_rows=[
            ('Admin/General', 'Accounting', 'Accounting Expense (Quickbooks), QB Fees', 0, 1300.0),
        ],
    )
    _, _, qb_map = load_budget(path)
    assert qb_map == {
        'Accounting Expense (Quickbooks)': 'Accounting',
        'QB Fees': 'Accounting',
    }


def test_load_budget_skips_blank_rows(tmp_path):
    path = tmp_path / 'budget.xlsx'
    _fill_budget(
        path,
        income_rows=[('Fundraising', 'Book Fair', '', 100.0, 200.0), (None, None, None, None, None)],
        expense_rows=[],
    )
    income_budget, _, _ = load_budget(path)
    assert list(income_budget['Fundraising'].keys()) == ['Book Fair']


# ── merge_actuals_into_budget ─────────────────────────────────────────────

def test_merge_direct_name_match():
    budget = {'Program Expense': {'Picnic': (0.0, 200.0)}}
    actuals = {'Picnic': [181.58] + [0.0] * 11}
    merged = merge_actuals_into_budget(budget, actuals, qb_to_budget_map={})
    assert merged == {'Program Expense': {'Picnic': (0.0, 200.0, [181.58] + [0.0] * 11)}}


def test_merge_maps_qb_category_to_budget_item():
    budget = {'Admin/General': {'Accounting': (0.0, 1300.0)}}
    actuals = {'Accounting Expense (Quickbooks)': [1257.76] + [0.0] * 11}
    qb_map = {'Accounting Expense (Quickbooks)': 'Accounting'}
    merged = merge_actuals_into_budget(budget, actuals, qb_map)
    assert merged['Admin/General']['Accounting'][2][0] == 1257.76


def test_merge_sums_multiple_qb_categories_into_one_item():
    budget = {'Admin/General': {'Accounting': (0.0, 1300.0)}}
    actuals = {
        'Accounting Expense (Quickbooks)': [1000.0] + [0.0] * 11,
        'QB Fees': [50.0] + [0.0] * 11,
    }
    qb_map = {'Accounting Expense (Quickbooks)': 'Accounting', 'QB Fees': 'Accounting'}
    merged = merge_actuals_into_budget(budget, actuals, qb_map)
    assert merged['Admin/General']['Accounting'][2][0] == 1050.0


def test_merge_unmatched_qb_category_goes_to_other_section():
    budget = {'Admin/General': {'Accounting': (0.0, 1300.0)}}
    actuals = {'Mystery Category': [42.0] + [0.0] * 11}
    merged = merge_actuals_into_budget(budget, actuals, qb_to_budget_map={})
    assert merged['Other (from QuickBooks)']['Mystery Category'] == (0.0, 0.0, [42.0] + [0.0] * 11)


def test_merge_item_with_no_actuals_gets_zeros():
    budget = {'Fundraising': {'Book Fair': (9118.36, 500.0)}}
    merged = merge_actuals_into_budget(budget, actuals_dict={}, qb_to_budget_map={})
    assert merged['Fundraising']['Book Fair'] == (9118.36, 500.0, [0.0] * 12)


# ── map_actuals_to_budget_items ──────────────────────────────────────────

def test_map_actuals_translates_qb_name_to_budget_item():
    actuals = {'Accounting Expense (Quickbooks)': [1257.76] + [0.0] * 11}
    qb_map = {'Accounting Expense (Quickbooks)': 'Accounting'}
    mapped = map_actuals_to_budget_items(actuals, qb_map)
    assert mapped == {'Accounting': [1257.76] + [0.0] * 11}


def test_map_actuals_sums_multiple_qb_categories_into_one_item():
    actuals = {
        'Accounting Expense (Quickbooks)': [1000.0] + [0.0] * 11,
        'QB Fees': [50.0] + [0.0] * 11,
    }
    qb_map = {'Accounting Expense (Quickbooks)': 'Accounting', 'QB Fees': 'Accounting'}
    mapped = map_actuals_to_budget_items(actuals, qb_map)
    assert mapped['Accounting'][0] == 1050.0


def test_map_actuals_unmapped_category_keeps_its_own_name():
    actuals = {'Mystery Category': [42.0] + [0.0] * 11}
    mapped = map_actuals_to_budget_items(actuals, qb_to_budget_map={})
    assert mapped == {'Mystery Category': [42.0] + [0.0] * 11}


# ── apply_dynamic_last_year ──────────────────────────────────────────────

def test_apply_dynamic_last_year_no_prior_history_keeps_static_value():
    budget = {'Program Expense': {'Picnic': (200.0, 250.0)}}
    result = apply_dynamic_last_year(budget, prior_actuals_mapped={})
    assert result == budget


def test_apply_dynamic_last_year_replaces_with_prior_actual():
    budget = {'Program Expense': {'Picnic': (200.0, 250.0)}}
    prior = {'Picnic': [181.58] + [0.0] * 11}
    result = apply_dynamic_last_year(budget, prior)
    assert result == {'Program Expense': {'Picnic': (181.58, 250.0)}}


def test_apply_dynamic_last_year_sums_all_12_months():
    budget = {'Program Expense': {'Picnic': (200.0, 250.0)}}
    prior = {'Picnic': [100.0, 50.0] + [0.0] * 10}
    result = apply_dynamic_last_year(budget, prior)
    assert result['Program Expense']['Picnic'][0] == 150.0


def test_apply_dynamic_last_year_item_with_no_prior_actual_becomes_zero():
    # Some prior-year history exists (for a DIFFERENT item), but this
    # specific item has no matching prior-year actual at all -- that's a
    # genuinely zero prior year for it, not a reason to keep the stale
    # static value.
    budget = {'Program Expense': {'Picnic': (200.0, 250.0)}}
    prior = {'Some Other Item': [500.0] + [0.0] * 11}
    result = apply_dynamic_last_year(budget, prior)
    assert result == {'Program Expense': {'Picnic': (0.0, 250.0)}}
