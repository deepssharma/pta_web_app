"""
Tests for builders.py Excel sheet builder functions.
Uses mock data — no real files needed.
"""
import pytest
import openpyxl
from pta_treasurer.builders import (build_treasurer, build_budget, build_givebacks,
                      build_manifest, FISCAL_MONTHS)


# ── Mock data ─────────────────────────────────────────────────────────────────

MOCK_QRB = {
    'period':        'July 1-31, 2025',
    'income':        {},
    'income_total':  0.0,
    'expenses':      {'Accounting Quickbooks': 1257.76, '6th Grade Events': 181.58},
    'expense_total': 1439.34,
    'net_income':    -1439.34,
    'transactions':  [],
}

MOCK_BANK = {
    'period':            'July 01, 2025 through July 31, 2025',
    'account':           '4346',
    'beginning_balance': 32630.10,
    'ending_balance':    31190.76,
    'total_deposits':    0.0,
    'total_checks':      181.58,
    'total_withdrawals': 1257.76,
    'total_fees':        0.0,
    'deposits':          [],
    'checks':            [{'check_no': '1077', 'date': '07/01', 'amount': 181.58}],
    'withdrawals':       [],
    'fees':              [],
    'daily_balances':    {'07/31': 31190.76},
    'source_file':       'Chase_july_2025.pdf',
}

MOCK_GIVEBACKS = [
    {'item': 'Shop to Give Donation', 'category': '',
     'count': 1, 'total': 95.14, 'source_file': 'givebacks_july.csv'},
    {'item': 'Teacher/Staff', 'category': 'Memberships',
     'count': 1, 'total': 15.0,  'source_file': 'givebacks_july.csv'},
]


# ── Treasurer Report tests ────────────────────────────────────────────────────

def test_build_treasurer_creates_sheet():
    wb = openpyxl.Workbook()
    ws = wb.active
    build_treasurer(ws, MOCK_QRB, MOCK_BANK, 'July 2025', 'Setauket School PTA')
    assert ws['A1'].value == 'SETAUKET SCHOOL PTA'

def test_build_treasurer_has_org_name():
    wb = openpyxl.Workbook()
    ws = wb.active
    build_treasurer(ws, MOCK_QRB, MOCK_BANK, 'July 2025', 'Test PTA')
    assert ws['A1'].value == 'TEST PTA'

def test_build_treasurer_bank_reconciliation():
    wb = openpyxl.Workbook()
    ws = wb.active
    build_treasurer(ws, MOCK_QRB, MOCK_BANK, 'July 2025', 'Setauket School PTA')
    # Find beginning balance cell
    found = False
    for row in ws.iter_rows(values_only=True):
        if row[0] == 'Beginning Balance':
            assert row[1] == 32630.10
            found = True
            break
    assert found, 'Beginning Balance row not found'

def test_build_treasurer_difference_is_zero():
    wb = openpyxl.Workbook()
    ws = wb.active
    build_treasurer(ws, MOCK_QRB, MOCK_BANK, 'July 2025', 'Setauket School PTA')
    found = False
    for row in ws.iter_rows(values_only=True):
        if row[0] == 'Difference (should be $0.00)':
            assert abs(row[1]) < 0.01
            found = True
            break
    assert found, 'Difference row not found'


# ── Budget sheet tests ────────────────────────────────────────────────────────

def test_build_budget_income(sample_merged_income):
    wb = openpyxl.Workbook()
    ws = wb.active
    build_budget(ws, 'Budget vs Actuals - Income',
                 sample_merged_income, 'Setauket School PTA',
                 FISCAL_MONTHS, 0, show_pl=False)
    assert 'SETAUKET SCHOOL PTA' in ws['A1'].value

def test_build_budget_active_month_highlighted(sample_merged_income):
    wb = openpyxl.Workbook()
    ws = wb.active
    build_budget(ws, 'Test', sample_merged_income,
                 'Test PTA', FISCAL_MONTHS, 0, show_pl=False)
    # Column D (index 4) = JULY = fiscal index 0 — should be gold
    header_cell = ws.cell(row=3, column=4)
    # openpyxl stores colors as 8-char ARGB (alpha + RGB)
    # so FFD966 becomes 00FFD966
    assert header_cell.fill.fgColor.rgb.endswith('FFD966')  # ← use endswith

def test_build_budget_expense_with_pl(sample_merged_expense, sample_merged_income):
    wb = openpyxl.Workbook()
    ws = wb.active
    build_budget(ws, 'Budget vs Actuals - Expenses',
                 sample_merged_expense, 'Test PTA',
                 FISCAL_MONTHS, 0, show_pl=True,
                 income_merged=sample_merged_income)
    # Should have Profit/Loss header
    found = False
    for row in ws.iter_rows(min_row=3, max_row=3, values_only=True):
        if 'Profit/Loss' in [v for v in row if v]:
            found = True
    assert found, 'Profit/Loss column not found'


# ── Givebacks reconciliation tests ───────────────────────────────────────────

def test_build_givebacks_total():
    wb = openpyxl.Workbook()
    ws = wb.active
    build_givebacks(ws, MOCK_GIVEBACKS, MOCK_BANK, 'Test PTA')
    # Check total row exists
    found = False
    for row in ws.iter_rows(values_only=True):
        if row[0] == 'TOTAL':
            found = True
            break
    assert found, 'TOTAL row not found'

def test_build_givebacks_item_count():
    wb = openpyxl.Workbook()
    ws = wb.active
    build_givebacks(ws, MOCK_GIVEBACKS, MOCK_BANK, 'Test PTA')
    items_found = 0
    for row in ws.iter_rows(values_only=True):
        if row[0] in ('Shop to Give Donation', 'Teacher/Staff'):
            items_found += 1
    assert items_found == 2


# ── Fiscal months constants test ──────────────────────────────────────────────

def test_fiscal_months_count():
    assert len(FISCAL_MONTHS) == 12

def test_fiscal_months_order():
    assert FISCAL_MONTHS[0]  == 'JULY'
    assert FISCAL_MONTHS[5]  == 'DEC'
    assert FISCAL_MONTHS[6]  == 'JAN'
    assert FISCAL_MONTHS[11] == 'JUNE'
