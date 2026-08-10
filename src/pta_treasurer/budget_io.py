"""
budget_io.py
Reads and writes the Excel-template budget config that replaces the old
notebook's hardcoded INCOME_BUDGET / EXPENSE_BUDGET / QB_TO_BUDGET_MAP dicts.
"""

from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

from pta_treasurer.builders import (
    NAVY_FILL, SUBHDR_FONT, TEAL_FILL, THIN_BORDER, WHITE,
)
from openpyxl.styles import Alignment, Font

TEMPLATE_HEADERS = [
    'Section', 'Item', 'QuickBooks Category Name(s)',
    'Last Year Actual', 'This Year Budget',
]

SHEET_NAMES = ('Income Budget', 'Expense Budget')


def _write_sheet(ws, title: str):
    ws.sheet_view.showGridLines = False
    for col, w in zip(['A', 'B', 'C', 'D', 'E'], [22, 28, 40, 16, 16]):
        ws.column_dimensions[col].width = w

    ws.merge_cells('A1:E1')
    c = ws['A1']
    c.value = title
    c.font = Font(name='Arial', bold=True, size=13, color=WHITE)
    c.fill = NAVY_FILL
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 24

    for ci, hdr in enumerate(TEMPLATE_HEADERS, 1):
        cell = ws.cell(row=2, column=ci, value=hdr)
        cell.font = SUBHDR_FONT
        cell.fill = TEAL_FILL
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
    ws.row_dimensions[2].height = 30


def generate_template(path: Path) -> None:
    """
    Writes a blank budget workbook with 'Income Budget' / 'Expense Budget'
    sheets for the treasurer to fill in (or edit each fiscal year).
    """
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = SHEET_NAMES[0]
    _write_sheet(ws1, 'INCOME BUDGET')

    ws2 = wb.create_sheet(SHEET_NAMES[1])
    _write_sheet(ws2, 'EXPENSE BUDGET')

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def _read_budget_sheet(ws):
    """
    Returns (budget, qb_to_budget_map) for one sheet:
      budget: {section: {item: (last_year_actual, this_year_budget)}}
      qb_to_budget_map: {qb_category_name: item_name}
    """
    budget = {}
    qb_to_budget_map = {}

    for row in ws.iter_rows(min_row=3, values_only=True):
        section, item = row[0], row[1]
        if not section or not item:
            continue
        section = str(section).strip()
        item = str(item).strip()
        qb_names = str(row[2]).strip() if row[2] else ''
        last_year = float(row[3]) if row[3] not in (None, '') else 0.0
        this_year = float(row[4]) if row[4] not in (None, '') else 0.0

        budget.setdefault(section, {})[item] = (last_year, this_year)

        for qb_name in qb_names.split(','):
            qb_name = qb_name.strip()
            if qb_name:
                qb_to_budget_map[qb_name] = item

    return budget, qb_to_budget_map


def load_budget(path: Path):
    """
    Reads a budget workbook produced by generate_template (or hand-edited by
    the treasurer). Returns (income_budget, expense_budget, qb_to_budget_map)
    with the two budgets in the {section: {item: (last_yr, budget)}} shape
    that merge_actuals_into_budget expects, and qb_to_budget_map merged
    across both sheets.
    """
    wb = openpyxl.load_workbook(path, data_only=True)

    income_budget, income_map = _read_budget_sheet(wb[SHEET_NAMES[0]])
    expense_budget, expense_map = _read_budget_sheet(wb[SHEET_NAMES[1]])

    qb_to_budget_map = {**income_map, **expense_map}
    return income_budget, expense_budget, qb_to_budget_map


def map_actuals_to_budget_items(actuals_dict: dict, qb_to_budget_map: dict) -> dict:
    """
    Translates QuickBooks category names in actuals_dict to budget item
    names via qb_to_budget_map, summing entries that map to the same item.
    A QB category with no mapping keeps its own name unchanged (picked up
    later as an 'Other (from QuickBooks)' item by merge_actuals_into_budget,
    or simply ignored by callers -- like apply_dynamic_last_year -- that
    only care about items already in the budget).

    Args:
        actuals_dict:      {qb_category_name: [12 monthly actuals]}
        qb_to_budget_map: {qb_category_name: budget_item_name}

    Returns:
        {budget_item_name: [12 monthly actuals]}
    """
    mapped_actuals = {}
    for qb_name, vals in actuals_dict.items():
        budget_name = qb_to_budget_map.get(qb_name, qb_name)
        if budget_name in mapped_actuals:
            mapped_actuals[budget_name] = [
                mapped_actuals[budget_name][i] + vals[i] for i in range(12)
            ]
        else:
            mapped_actuals[budget_name] = list(vals)
    return mapped_actuals


def apply_dynamic_last_year(budget_dict: dict, prior_actuals_mapped: dict) -> dict:
    """
    Replaces each item's static 'last_year_actual' (originally hand-entered
    into budget.xlsx) with the real prior fiscal year's actual, computed
    from that year's own recorded history.

    Args:
        budget_dict:            {section: {item: (last_year_actual, this_year_budget)}}
        prior_actuals_mapped:   {budget_item_name: [12 monthly actuals]} --
                                 already translated via map_actuals_to_budget_items,
                                 for the PRIOR fiscal year specifically.

    Returns:
        budget_dict unchanged if prior_actuals_mapped is empty (no digital
        history exists at all for that prior fiscal year -- e.g. the very
        first fiscal year this app is used for -- so the static fallback
        value is kept). Otherwise, a new {section: {item: (dynamic_last_year,
        this_year_budget)}} -- an item with no matching prior-year actual
        becomes 0.0 (a genuinely zero prior year for that item), not the
        stale static value.
    """
    if not prior_actuals_mapped:
        return budget_dict

    return {
        section: {
            item: (round(sum(prior_actuals_mapped.get(item, [0.0] * 12)), 2), this_year)
            for item, (_, this_year) in items.items()
        }
        for section, items in budget_dict.items()
    }


def merge_actuals_into_budget(budget_dict: dict, actuals_dict: dict,
                               qb_to_budget_map: dict) -> dict:
    """
    Merges parsed QuickBooks actuals into a budget structure.

    Args:
        budget_dict:      {section: {item: (last_year_actual, this_year_budget)}}
        actuals_dict:      {qb_category_name: [12 monthly actuals]}
        qb_to_budget_map: {qb_category_name: budget_item_name}

    Returns:
        {section: {item: (last_year_actual, this_year_budget, [12 monthly actuals])}}
        plus an 'Other (from QuickBooks)' section for any QB category that
        doesn't map to a budget item.
    """
    mapped_actuals = map_actuals_to_budget_items(actuals_dict, qb_to_budget_map)

    result = {}
    all_budget_items = {item for section in budget_dict.values() for item in section}
    for section, items in budget_dict.items():
        result[section] = {}
        for item, (last_yr, budget) in items.items():
            result[section][item] = (last_yr, budget, mapped_actuals.get(item, [0.0] * 12))

    unmatched = [k for k in mapped_actuals if k not in all_budget_items]
    if unmatched:
        result['Other (from QuickBooks)'] = {
            item: (0.0, 0.0, mapped_actuals[item]) for item in unmatched
        }

    return result
