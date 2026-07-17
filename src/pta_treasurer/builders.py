"""
builders.py
All Excel sheet builder functions for the PTA Treasurer Report Generator.
Requires openpyxl.
"""

from datetime import datetime
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Colours ───────────────────────────────────────────────────────────────────
NAVY   = '1F3864'; TEAL   = '2E75B6'; LTBLUE = 'BDD7EE'
GOLD   = 'FFD966'; WHITE  = 'FFFFFF'; LGREY  = 'F2F2F2'
GREEN  = 'E2EFDA'; RED_BG = 'FCE4D6'

SUBHDR_FONT = Font(name='Arial', bold=True, color=WHITE, size=10)
BODY_FONT   = Font(name='Arial', size=10)
BOLD_FONT   = Font(name='Arial', bold=True, size=10)
TOTAL_FONT  = Font(name='Arial', bold=True, size=10, color=NAVY)

NAVY_FILL   = PatternFill('solid', fgColor=NAVY)
TEAL_FILL   = PatternFill('solid', fgColor=TEAL)
LTBLUE_FILL = PatternFill('solid', fgColor=LTBLUE)
GOLD_FILL   = PatternFill('solid', fgColor=GOLD)
LGREY_FILL  = PatternFill('solid', fgColor=LGREY)
GREEN_FILL  = PatternFill('solid', fgColor=GREEN)
RED_FILL    = PatternFill('solid', fgColor=RED_BG)

THIN        = Side(style='thin',   color='AAAAAA')
MED         = Side(style='medium', color=NAVY)
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MED_BORDER  = Border(left=MED,  right=MED,  top=MED,  bottom=MED)
MONEY_FMT   = '$#,##0.00_);($#,##0.00)'

FISCAL_MONTHS = ['JULY','AUG','SEPT','OCT','NOV','DEC',
                 'JAN','FEB','MAR','APR','MAY','JUNE']


# ── Shared helpers ────────────────────────────────────────────────────────────

def _sec_hdr(ws, label, row, n=4):
    ws.merge_cells(f'A{row}:{get_column_letter(n)}{row}')
    c = ws[f'A{row}']; c.value = label
    c.font = Font(name='Arial', bold=True, size=11, color=WHITE)
    c.fill = NAVY_FILL
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.row_dimensions[row].height = 20
    return row + 1


def _col_hdrs(ws, row, labels):
    for i, lbl in enumerate(labels):
        c = ws.cell(row=row, column=i+1, value=lbl)
        c.font = SUBHDR_FONT; c.fill = TEAL_FILL; c.border = THIN_BORDER
        c.alignment = Alignment(
            horizontal='center' if i > 0 else 'left',
            vertical='center', indent=1 if i == 0 else 0)
    ws.row_dimensions[row].height = 18
    return row + 1


def _data_row(ws, row, label, amount, shade=False):
    fill = LGREY_FILL if shade else PatternFill()
    ws[f'A{row}'].value = label; ws[f'A{row}'].font = BODY_FONT
    ws[f'A{row}'].fill = fill;   ws[f'A{row}'].border = THIN_BORDER
    ws[f'A{row}'].alignment = Alignment(indent=2)
    ws[f'B{row}'].value = amount; ws[f'B{row}'].font = BODY_FONT
    ws[f'B{row}'].fill = fill;    ws[f'B{row}'].border = THIN_BORDER
    ws[f'B{row}'].number_format = MONEY_FMT
    ws[f'B{row}'].alignment = Alignment(horizontal='right')
    for col in ['C', 'D']:
        ws[f'{col}{row}'].fill = fill
        ws[f'{col}{row}'].border = THIN_BORDER
    ws.row_dimensions[row].height = 16


def _total_row(ws, row, label, val):
    for col in ['A', 'B', 'C', 'D']:
        ws[f'{col}{row}'].fill = LTBLUE_FILL
        ws[f'{col}{row}'].border = MED_BORDER
    ws[f'A{row}'].value = label; ws[f'A{row}'].font = TOTAL_FONT
    ws[f'A{row}'].alignment = Alignment(indent=1)
    ws[f'B{row}'].value = val;   ws[f'B{row}'].font = TOTAL_FONT
    ws[f'B{row}'].number_format = MONEY_FMT
    ws[f'B{row}'].alignment = Alignment(horizontal='right')
    ws.row_dimensions[row].height = 18


# ── 1. TREASURER REPORT ───────────────────────────────────────────────────────

def build_treasurer(ws, qb, bank, month_label, org_name):
    ws.sheet_view.showGridLines = False
    for col, w in zip(['A','B','C','D'], [32,18,18,22]):
        ws.column_dimensions[col].width = w

    ws.merge_cells('A1:D1'); c = ws['A1']; c.value = org_name.upper()
    c.font = Font(name='Arial', bold=True, size=16, color=NAVY)
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 28

    ws.merge_cells('A2:D2'); c = ws['A2']
    c.value = f'Monthly Treasurer Report - {month_label}'
    c.font = Font(name='Arial', bold=True, size=12, color=TEAL)
    c.alignment = Alignment(horizontal='center')

    ws.merge_cells('A3:D3'); c = ws['A3']
    c.value = f'Generated: {datetime.today().strftime("%B %d, %Y")}'
    c.font = Font(name='Arial', italic=True, size=9, color='888888')
    c.alignment = Alignment(horizontal='center')

    row = 5
    row = _sec_hdr(ws, 'INCOME', row)
    row = _col_hdrs(ws, row, ['Category', 'Amount', '', ''])
    inc_s = row
    for i, (k, v) in enumerate(qb['income'].items()):
        _data_row(ws, row, k, v, shade=(i % 2 == 1)); row += 1
    _total_row(ws, row, 'Total Income', f'=SUM(B{inc_s}:B{row-1})')
    it = row; row += 2

    row = _sec_hdr(ws, 'EXPENSES', row)
    row = _col_hdrs(ws, row, ['Category', 'Amount', '', ''])
    exp_s = row
    for i, (k, v) in enumerate(qb['expenses'].items()):
        _data_row(ws, row, k, v, shade=(i % 2 == 1)); row += 1
    _total_row(ws, row, 'Total Expenses', f'=SUM(B{exp_s}:B{row-1})')
    et = row; row += 2

    ws.merge_cells(f'A{row}:D{row}')
    ws[f'A{row}'].value = 'NET INCOME / (LOSS)'
    ws[f'A{row}'].font = Font(name='Arial', bold=True, size=11, color=WHITE)
    ws[f'A{row}'].fill = NAVY_FILL
    ws[f'A{row}'].alignment = Alignment(horizontal='left', indent=1)
    row += 1
    for col in ['A','B','C','D']:
        ws[f'{col}{row}'].fill = GOLD_FILL
        ws[f'{col}{row}'].border = MED_BORDER
    ws[f'A{row}'].value = 'Net Income (Loss)'
    ws[f'A{row}'].font = TOTAL_FONT
    ws[f'A{row}'].alignment = Alignment(indent=1)
    ws[f'B{row}'].value = f'=B{it}-B{et}'
    ws[f'B{row}'].font = TOTAL_FONT
    ws[f'B{row}'].number_format = MONEY_FMT
    ws[f'B{row}'].alignment = Alignment(horizontal='right')
    row += 2

    row = _sec_hdr(ws, 'BANK RECONCILIATION - Chase Business Checking', row)
    row = _col_hdrs(ws, row, ['Description', 'Amount', '', ''])
    for i, (lbl, amt) in enumerate([
        ('Beginning Balance',            bank['beginning_balance']),
        ('(+) Deposits & Additions',     bank['total_deposits']),
        ('(-) Checks Paid',             -bank['total_checks']),
        ('(-) Electronic Withdrawals',  -bank.get('total_withdrawals', 0.0)),
        ('(-) Fees',                    -bank['total_fees']),
    ]):
        _data_row(ws, row, lbl, amt, shade=(i % 2 == 1)); row += 1
    _total_row(ws, row, 'Ending Balance (per statement)',
               bank['ending_balance']); row += 1

    calc = (bank['beginning_balance'] + bank['total_deposits']
            - bank['total_checks']
            - bank.get('total_withdrawals', 0.0)
            - bank['total_fees'])
    diff = calc - bank['ending_balance']

    for lbl, val, fill in [
        ('Calculated Ending Balance', calc, PatternFill()),
        ('Difference (should be $0.00)', diff,
         GREEN_FILL if abs(diff) < 0.01 else RED_FILL),
    ]:
        for col in ['A','B','C','D']:
            ws[f'{col}{row}'].fill = fill
            ws[f'{col}{row}'].border = THIN_BORDER
        ws[f'A{row}'].value = lbl; ws[f'A{row}'].font = BOLD_FONT
        ws[f'A{row}'].alignment = Alignment(indent=2)
        ws[f'B{row}'].value = val; ws[f'B{row}'].font = BOLD_FONT
        ws[f'B{row}'].number_format = MONEY_FMT
        ws[f'B{row}'].alignment = Alignment(horizontal='right')
        ws.row_dimensions[row].height = 16; row += 1
    row += 1

    if bank['daily_balances']:
        row = _sec_hdr(ws, 'DAILY ENDING BALANCES', row)
        for col, hdr in zip(['A','B'], ['Date', 'Balance']):
            ws[f'{col}{row}'].value = hdr
            ws[f'{col}{row}'].font = SUBHDR_FONT
            ws[f'{col}{row}'].fill = TEAL_FILL
            ws[f'{col}{row}'].border = THIN_BORDER
            ws[f'{col}{row}'].alignment = Alignment(horizontal='center')
        row += 1
        for i, (dt, bal) in enumerate(sorted(bank['daily_balances'].items())):
            fill = LGREY_FILL if i % 2 else PatternFill()
            ws[f'A{row}'].value = dt;  ws[f'A{row}'].font = BODY_FONT
            ws[f'A{row}'].fill = fill; ws[f'A{row}'].border = THIN_BORDER
            ws[f'A{row}'].alignment = Alignment(horizontal='center')
            ws[f'B{row}'].value = bal; ws[f'B{row}'].font = BODY_FONT
            ws[f'B{row}'].fill = fill; ws[f'B{row}'].border = THIN_BORDER
            ws[f'B{row}'].number_format = MONEY_FMT
            ws[f'B{row}'].alignment = Alignment(horizontal='right')
            row += 1


# ── 2. BUDGET VS ACTUALS ──────────────────────────────────────────────────────

def build_budget(ws, title, merged_data, org_name, fiscal_months,
                 current_idx, show_pl=False, income_merged=None):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 28
    ws.column_dimensions['B'].width = 13
    ws.column_dimensions['C'].width = 13
    for i in range(12):
        ws.column_dimensions[get_column_letter(4+i)].width = 9
    ws.column_dimensions[get_column_letter(16)].width = 11
    ws.column_dimensions[get_column_letter(17)].width = 11

    ws.merge_cells(f'A1:{get_column_letter(17)}1'); c = ws['A1']
    c.value = f'{org_name.upper()}  -  {title}'
    c.font = Font(name='Arial', bold=True, size=13, color=WHITE)
    c.fill = NAVY_FILL
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 26

    ws.merge_cells(f'A2:{get_column_letter(17)}2'); c = ws['A2']
    c.value = (f'As of {datetime.today().strftime("%B %d, %Y")}  |  '
               f'Fiscal Year July 2025 - June 2026  |  '
               f'Active month: {fiscal_months[current_idx]}')
    c.font = Font(name='Arial', italic=True, size=9, color='666666')
    c.alignment = Alignment(horizontal='center')

    pl_label = 'Profit/Loss' if show_pl else ''
    headers = (['Category', 'Last Year', 'Budget (Annual)']
               + fiscal_months + ['Total', pl_label])
    for ci, hdr in enumerate(headers, 1):
        c = ws.cell(row=3, column=ci, value=hdr)
        is_active = (4 <= ci <= 15 and (ci-4) == current_idx)
        c.font = Font(name='Arial', bold=True, size=9,
                      color=NAVY if is_active else WHITE)
        c.fill = GOLD_FILL if is_active else TEAL_FILL
        c.alignment = Alignment(
            horizontal='center' if ci > 1 else 'left',
            vertical='center', wrap_text=True)
        c.border = THIN_BORDER
    ws.row_dimensions[3].height = 30

    # Build income lookup for P&L column
    inc_lookup = {}
    if show_pl and income_merged:
        inc_lookup = {
            item.lower(): sum(monthly[:current_idx+1])
            for section in income_merged.values()
            for item, (_, _, monthly) in section.items()
        }

    dr = 4
    for section, items in merged_data.items():
        ws.merge_cells(f'A{dr}:{get_column_letter(17)}{dr}')
        c = ws[f'A{dr}']; c.value = section
        c.font = Font(name='Arial', bold=True, size=10, color=WHITE)
        c.fill = NAVY_FILL
        c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws.row_dimensions[dr].height = 18
        ss = dr + 1; dr += 1

        for ir, (item, (last_yr, budget, monthly)) in enumerate(items.items()):
            fill = LGREY_FILL if ir % 2 == 1 else PatternFill()
            c = ws.cell(row=dr, column=1, value=item)
            c.font = BODY_FONT; c.fill = fill
            c.alignment = Alignment(indent=2); c.border = THIN_BORDER

            for ci, v in [(2, last_yr), (3, budget)]:
                c = ws.cell(row=dr, column=ci, value=v if v else None)
                c.font = BODY_FONT; c.fill = fill
                c.number_format = MONEY_FMT
                c.alignment = Alignment(horizontal='right')
                c.border = THIN_BORDER

            for mi, v in enumerate(monthly):
                is_active = (mi == current_idx)
                c = ws.cell(row=dr, column=4+mi, value=v if v else None)
                c.font = Font(name='Arial', bold=is_active, size=10)
                c.fill = (GOLD_FILL if (is_active and v)
                          else PatternFill('solid', fgColor='FFF9E6')
                          if is_active else fill)
                c.number_format = MONEY_FMT
                c.alignment = Alignment(horizontal='right')
                c.border = THIN_BORDER

            c = ws.cell(row=dr, column=16, value=f'=SUM(D{dr}:O{dr})')
            c.font = BODY_FONT; c.fill = fill; c.number_format = MONEY_FMT
            c.alignment = Alignment(horizontal='right'); c.border = THIN_BORDER

            # Profit/Loss column
            if show_pl:
                act_exp = sum(monthly[:current_idx+1])
                act_inc = inc_lookup.get(item.lower(), 0.0)
                pl_val  = (act_inc - act_exp
                           if (act_inc or act_exp) else None)
            else:
                pl_val = None
            c = ws.cell(row=dr, column=17, value=pl_val)
            c.font = BODY_FONT; c.fill = fill; c.number_format = MONEY_FMT
            c.alignment = Alignment(horizontal='right'); c.border = THIN_BORDER

            ws.row_dimensions[dr].height = 15; dr += 1

        se = dr - 1
        c = ws.cell(row=dr, column=1, value=f'Total {section}')
        c.font = TOTAL_FONT; c.fill = LTBLUE_FILL
        c.alignment = Alignment(indent=1); c.border = MED_BORDER
        for ci in range(2, 18):
            cl = get_column_letter(ci)
            c = ws.cell(row=dr, column=ci,
                        value=f'=SUM({cl}{ss}:{cl}{se})')
            c.font = TOTAL_FONT; c.fill = LTBLUE_FILL
            c.number_format = MONEY_FMT
            c.alignment = Alignment(horizontal='right'); c.border = MED_BORDER
        ws.row_dimensions[dr].height = 18; dr += 2

    c = ws.cell(row=dr, column=1, value='GRAND TOTAL')
    c.font = Font(name='Arial', bold=True, size=11, color=WHITE)
    c.fill = NAVY_FILL; c.alignment = Alignment(indent=1); c.border = MED_BORDER
    for ci in range(2, 18):
        cl = get_column_letter(ci)
        c = ws.cell(row=dr, column=ci,
                    value=f'=SUMIF(A4:A{dr-1},"Total*",{cl}4:{cl}{dr-1})')
        c.font = Font(name='Arial', bold=True, size=10, color=WHITE)
        c.fill = NAVY_FILL; c.number_format = MONEY_FMT
        c.alignment = Alignment(horizontal='right'); c.border = MED_BORDER
    ws.row_dimensions[dr].height = 20


# ── 3. GIVEBACK RECONCILIATION ────────────────────────────────────────────────

def build_givebacks(ws, givebacks, bank, org_name):
    ws.sheet_view.showGridLines = False
    for col, w in zip(['A','B','C','D','E','F'], [30,18,12,16,11,25]):
        ws.column_dimensions[col].width = w

    ws.merge_cells('A1:F1'); c = ws['A1']
    c.value = f'{org_name.upper()}  -  Giveback Reconciliation'
    c.font = Font(name='Arial', bold=True, size=13, color=WHITE)
    c.fill = NAVY_FILL
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 26

    ws.merge_cells('A2:F2'); c = ws['A2']
    c.value = f'Generated: {datetime.today().strftime("%B %d, %Y")}'
    c.font = Font(name='Arial', italic=True, size=9, color='666666')
    c.alignment = Alignment(horizontal='center')

    row = 4
    for col, hdr in zip(['A','B','C','D','E','F'],
                         ['Item','Category','Transactions',
                          'Amount','% of Total','Source File']):
        c = ws[f'{col}{row}']; c.value = hdr; c.font = SUBHDR_FONT
        c.fill = TEAL_FILL
        c.alignment = Alignment(
            horizontal='center' if col != 'A' else 'left')
        c.border = THIN_BORDER
    ws.row_dimensions[row].height = 18; row += 1

    ds = row; trn = ds + len(givebacks)
    for i, g in enumerate(givebacks):
        fill = LGREY_FILL if i % 2 == 1 else PatternFill()
        for col in ['A','B','C','D','E','F']:
            ws[f'{col}{row}'].fill = fill
            ws[f'{col}{row}'].border = THIN_BORDER
        ws[f'A{row}'].value = g['item']
        ws[f'A{row}'].font = BODY_FONT
        ws[f'A{row}'].alignment = Alignment(indent=1)
        ws[f'B{row}'].value = g['category']
        ws[f'B{row}'].font = BODY_FONT
        ws[f'B{row}'].alignment = Alignment(horizontal='center')
        ws[f'C{row}'].value = g['count']
        ws[f'C{row}'].font = BODY_FONT
        ws[f'C{row}'].alignment = Alignment(horizontal='center')
        ws[f'D{row}'].value = g['total']
        ws[f'D{row}'].font = BODY_FONT
        ws[f'D{row}'].number_format = MONEY_FMT
        ws[f'D{row}'].alignment = Alignment(horizontal='right')
        ws[f'E{row}'].value = f'=D{row}/D{trn}'
        ws[f'E{row}'].font = BODY_FONT
        ws[f'E{row}'].number_format = '0.0%'
        ws[f'E{row}'].alignment = Alignment(horizontal='center')
        ws[f'F{row}'].value = g.get('source_file', '')
        ws[f'F{row}'].font = BODY_FONT
        ws[f'F{row}'].alignment = Alignment(indent=1)
        ws.row_dimensions[row].height = 15; row += 1

    for col in ['A','B','C','D','E','F']:
        ws[f'{col}{row}'].fill = LTBLUE_FILL
        ws[f'{col}{row}'].border = MED_BORDER
    ws[f'A{row}'].value = 'TOTAL'; ws[f'A{row}'].font = TOTAL_FONT
    ws[f'A{row}'].alignment = Alignment(indent=1)
    ws[f'C{row}'].value = f'=SUM(C{ds}:C{row-1})'
    ws[f'C{row}'].font = TOTAL_FONT
    ws[f'C{row}'].alignment = Alignment(horizontal='center')
    ws[f'D{row}'].value = f'=SUM(D{ds}:D{row-1})'
    ws[f'D{row}'].font = TOTAL_FONT
    ws[f'D{row}'].number_format = MONEY_FMT
    ws[f'D{row}'].alignment = Alignment(horizontal='right')
    ws[f'E{row}'].value = '100.0%'; ws[f'E{row}'].font = TOTAL_FONT
    ws[f'E{row}'].alignment = Alignment(horizontal='center')
    ws.row_dimensions[row].height = 18; row += 2

    _sec_hdr(ws, 'GIVEBACK <-> BANK RECONCILIATION', row, n=6); row += 1
    gb_total = sum(g['total'] for g in givebacks)
    bank_gb  = next(
        (d['amount'] for d in bank['deposits']
         if 'gb payout' in d.get('description','').lower()
         or 'givebacks' in d.get('description','').lower()), 0.0)

    for lbl, val in [
        ('Givebacks Platform Total',            gb_total),
        ('Givebacks Deposit in Bank Statement',  bank_gb),
        ('Difference',                           gb_total - bank_gb),
    ]:
        is_d = lbl == 'Difference'
        fill = (GREEN_FILL if abs(gb_total-bank_gb) < 0.01
                else RED_FILL) if is_d else PatternFill()
        for col in ['A','B','C','D','E','F']:
            ws[f'{col}{row}'].fill = fill
            ws[f'{col}{row}'].border = THIN_BORDER
        ws[f'A{row}'].value = lbl; ws[f'A{row}'].font = BOLD_FONT
        ws[f'A{row}'].alignment = Alignment(indent=2)
        ws[f'D{row}'].value = val; ws[f'D{row}'].font = BOLD_FONT
        ws[f'D{row}'].number_format = MONEY_FMT
        ws[f'D{row}'].alignment = Alignment(horizontal='right')
        ws.row_dimensions[row].height = 16; row += 1


# ── 4. FILE MANIFEST ──────────────────────────────────────────────────────────

def build_manifest(ws, month_folder, org_name, month_label,
                   fiscal_idx, fiscal_months):
    ws.sheet_view.showGridLines = False
    for col, w in zip(['A','B','C','D'], [15,35,22,14]):
        ws.column_dimensions[col].width = w

    ws.merge_cells('A1:D1'); c = ws['A1']
    c.value = f'{org_name.upper()}  -  File Manifest'
    c.font = Font(name='Arial', bold=True, size=13, color=WHITE)
    c.fill = NAVY_FILL
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 26

    ws.merge_cells('A2:D2'); c = ws['A2']
    c.value = (f'Report: {month_label}  |  '
               f'Column: {fiscal_months[fiscal_idx]}  |  '
               f'Generated: {datetime.today().strftime("%B %d, %Y at %I:%M %p")}')
    c.font = Font(name='Arial', italic=True, size=9, color='666666')
    c.alignment = Alignment(horizontal='center')

    row = 4
    for col, hdr in zip(['A','B','C','D'],
                         ['Type','Filename','Last Modified','Size']):
        c = ws[f'{col}{row}']; c.value = hdr; c.font = SUBHDR_FONT
        c.fill = TEAL_FILL
        c.alignment = Alignment(
            horizontal='left' if col == 'B' else 'center')
        c.border = THIN_BORDER
    ws.row_dimensions[row].height = 18; row += 1

    month_folder = Path(month_folder)
    all_files = (
        [('QuickBooks', f) for f in sorted(month_folder.glob('*.csv'))
         if 'givebacks' not in f.name.lower()] +
        [('Bank Statement', f) for f in sorted(month_folder.glob('*.pdf'))] +
        [('Givebacks', f) for f in sorted(
            (month_folder / 'givebacks').glob('*.csv'))
         if (month_folder / 'givebacks').exists()]
    )

    for i, (ft, fp) in enumerate(all_files):
        fill = LGREY_FILL if i % 2 == 1 else PatternFill()
        stat = fp.stat()
        mod  = datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M')
        size = f'{stat.st_size/1024:.1f} KB'
        for col in ['A','B','C','D']:
            ws[f'{col}{row}'].fill = fill
            ws[f'{col}{row}'].border = THIN_BORDER
        ws[f'A{row}'].value = ft; ws[f'A{row}'].font = BODY_FONT
        ws[f'A{row}'].alignment = Alignment(horizontal='center')
        ws[f'B{row}'].value = fp.name; ws[f'B{row}'].font = BODY_FONT
        ws[f'B{row}'].alignment = Alignment(indent=1)
        ws[f'C{row}'].value = mod; ws[f'C{row}'].font = BODY_FONT
        ws[f'C{row}'].alignment = Alignment(horizontal='center')
        ws[f'D{row}'].value = size; ws[f'D{row}'].font = BODY_FONT
        ws[f'D{row}'].alignment = Alignment(horizontal='right', indent=1)
        ws.row_dimensions[row].height = 16; row += 1

    ws.merge_cells(f'A{row}:D{row}')
    ws[f'A{row}'].value = f'Total files processed: {len(all_files)}'
    ws[f'A{row}'].font = BOLD_FONT
    ws[f'A{row}'].alignment = Alignment(indent=1)


# ── 5. YTD SUMMARY ───────────────────────────────────────────────────────────
def build_ytd_summary(ws, income_merged, expense_merged, org_name,
                      month_label, fiscal_idx, fiscal_months,
                      bank, balance_forward=0.0):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 28
    for col in ['B','C','D','E','F','G','H','I']:
        ws.column_dimensions[col].width = 14

    # Title
    ws.merge_cells('A1:I1'); c = ws['A1']
    c.value = f"{org_name}  —  Treasurer's Report"
    c.font = Font(name='Arial', bold=True, size=14, color=NAVY)
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 28

    ws.merge_cells('A2:I2'); c = ws['A2']
    c.value = f'7/1/25 - {datetime.today().strftime("%-m/%-d/%Y")}'
    c.font = Font(name='Arial', italic=True, size=10, color='666666')
    c.alignment = Alignment(horizontal='center')
    ws.row_dimensions[2].height = 16

    # Balance forward and current balance
    ws['A3'].value = 'Balance Forward:'
    ws['A3'].font = BOLD_FONT
    ws['B3'].value = balance_forward
    ws['B3'].font = BOLD_FONT
    ws['B3'].number_format = MONEY_FMT
    ws['C3'].value = 'as of 6/30/25'
    ws['C3'].font = Font(name='Arial', italic=True, size=9, color='666666')

    ws['E3'].value = 'Current Balance:'
    ws['E3'].font = BOLD_FONT
    ws['F3'].value = bank['ending_balance']
    ws['F3'].font = BOLD_FONT
    ws['F3'].number_format = MONEY_FMT
    period_end = (bank.get('period','').split('through')[-1].strip()
                  if bank.get('period') else '')
    ws['G3'].value = f'as of {period_end}'
    ws['G3'].font = Font(name='Arial', italic=True, size=9, color='666666')
    ws.row_dimensions[3].height = 18

    # YTD totals
    ytd_inc = sum(
        sum(monthly[:fiscal_idx+1])
        for section in income_merged.values()
        for _, (_, _, monthly) in section.items()
    )
    ytd_exp = sum(
        sum(monthly[:fiscal_idx+1])
        for section in expense_merged.values()
        for _, (_, _, monthly) in section.items()
    )
    ws['A4'].value = f'Total Income: ${ytd_inc:,.2f}'
    ws['A4'].font = BOLD_FONT
    ws['E4'].value = f'Total Expenses: ${ytd_exp:,.2f}'
    ws['E4'].font = BOLD_FONT
    ws.row_dimensions[4].height = 18

    # Build lookups
    income_lookup = {}
    for section, items in income_merged.items():
        for item, (last_yr, budget, monthly) in items.items():
            income_lookup[item.lower()] = (last_yr, budget, monthly)

    expense_lookup = {}
    for section, items in expense_merged.items():
        for item, (last_yr, budget, monthly) in items.items():
            expense_lookup[item.lower()] = (last_yr, budget, monthly)

    # Build section structure from expense sections
    section_items = {}
    for section, items in expense_merged.items():
        section_items[section] = {}
        for item, (last_yr_exp, exp_budget, exp_monthly) in items.items():
            inc_data    = income_lookup.get(item.lower())
            last_yr_inc = inc_data[0] if inc_data else 0.0
            inc_budget  = inc_data[1] if inc_data else 0.0
            inc_monthly = inc_data[2] if inc_data else [0.0]*12
            section_items[section][item] = {
                'last_yr_inc':  last_yr_inc,
                'last_yr_exp':  last_yr_exp,
                'exp_budget':   exp_budget,
                'inc_budget':   inc_budget,
                'act_income':   sum(inc_monthly[:fiscal_idx+1]),
                'act_expenses': sum(exp_monthly[:fiscal_idx+1]),
            }

    # Add income-only items
    # Special merged rows
    MERGE_PAIRS = {
        'membership expenses': {
            'label':        'Membership Expenses/Income',
            'income_keys':  ['single', 'family', 'donations', 'teachers', 'student'],
            'expense_keys': ['membership expenses', 'council dues'],
        },
        'entertainment': {
            'label':        'Basket Dinner',
            'income_keys':  ['ticket & raffle sales', 'sponsors'],
            'expense_keys': ['entertainment', 'raffles', 'venue'],
        },
    }
    
    all_merged_income_keys  = set()
    all_merged_expense_keys = set()
    for merge in MERGE_PAIRS.values():
        all_merged_income_keys.update(merge['income_keys'])
        all_merged_expense_keys.update(merge.get('expense_keys', []))
    
    other_income = {}
    
    for section, items in income_merged.items():
        for item, (last_yr_inc, inc_budget, inc_monthly) in items.items():
            if item.lower() not in expense_lookup and item.lower() not in all_merged_income_keys:
                other_income[item] = {
                    'last_yr_inc':  last_yr_inc,
                    'last_yr_exp':  0.0,
                    'exp_budget':   0.0,
                    'inc_budget':   inc_budget,
                    'act_income':   sum(inc_monthly[:fiscal_idx+1]),
                    'act_expenses': 0.0,
                }
    if other_income:
        section_items['Other Income'] = other_income

    def money_or_none(val):
        return val if val and abs(val) > 0.001 else None

    def money_or_dash(val):       # ← add this
        if val is None or abs(val) < 0.001:
            return '--'
        return val

    def sec_hdr_ytd(label, r):
        ws.merge_cells(f'A{r}:I{r}')
        c = ws[f'A{r}']; c.value = label
        c.font = Font(name='Arial', bold=True, size=11, color=WHITE)
        c.fill = NAVY_FILL
        c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws.row_dimensions[r].height = 20
        return r + 1

    def col_hdrs_ytd(r):
        for ci, hdr in enumerate([
            'Category',
            "Last Year's Income", "Last Year's Expenses",
            'Expected Income',    'Expected Expenses',
            'Actual Income',      'Actual Expenses',
            'Net Profit',         "Last Year's Profit",
        ], 1):
            c = ws.cell(row=r, column=ci, value=hdr)
            c.font = Font(name='Arial', bold=True, size=8, color=WHITE)
            c.fill = TEAL_FILL; c.border = THIN_BORDER
            c.alignment = Alignment(
                horizontal='center' if ci > 1 else 'left',
                vertical='center', wrap_text=True)
        ws.row_dimensions[r].height = 28
        return r + 1

    row = 6
    row = col_hdrs_ytd(row)
    
    for section_name, items in section_items.items():
        if not items:
            continue

        row = sec_hdr_ytd(section_name, row)
        #row = col_hdrs_ytd(row)

        sec = defaultdict(float)

        rendered_income_keys = set()  # track which income items were merged
        rendered_expense_keys = set()
        
        for i, (item, vals) in enumerate(items.items()):
            item_lower = item.lower()

            # Skip items already merged into another row
            if (item_lower in all_merged_expense_keys and
                    item_lower not in MERGE_PAIRS):
                continue

            if item_lower in MERGE_PAIRS:
                merge = MERGE_PAIRS[item_lower]
                
                # Sum income
                merged_inc_actual  = sum(
                    sum(income_lookup[k][2][:fiscal_idx+1])
                    for k in merge['income_keys'] if k in income_lookup)
                merged_inc_last_yr = sum(
                    income_lookup[k][0]
                    for k in merge['income_keys'] if k in income_lookup)
                merged_inc_budget  = sum(
                    income_lookup[k][1]
                    for k in merge['income_keys'] if k in income_lookup)

                # Sum expenses
                merged_exp_actual  = sum(
                    sum(expense_lookup[k][2][:fiscal_idx+1])
                    for k in merge.get('expense_keys', []) if k in expense_lookup)
                merged_exp_last_yr = sum(
                    expense_lookup[k][0]
                    for k in merge.get('expense_keys', []) if k in expense_lookup)
                merged_exp_budget  = sum(
                    expense_lookup[k][1]
                    for k in merge.get('expense_keys', []) if k in expense_lookup)

                rendered_income_keys.update(merge['income_keys'])
                rendered_expense_keys.update(merge.get('expense_keys', []))

                display_vals = {
                    'last_yr_inc':  merged_inc_last_yr,
                    'last_yr_exp':  merged_exp_last_yr,
                    'inc_budget':   merged_inc_budget,
                    'exp_budget':   merged_exp_budget,
                    'act_income':   merged_inc_actual,
                    'act_expenses': merged_exp_actual,
                }
                display_label = merge['label']
            else:
                display_label = item
                display_vals  = vals

            has_data = (
                abs(display_vals['last_yr_inc'])  > 0.001 or
                abs(display_vals['last_yr_exp'])  > 0.001 or
                abs(display_vals['inc_budget'])   > 0.001 or
                abs(display_vals['exp_budget'])   > 0.001 or
                abs(display_vals['act_income'])   > 0.001 or
                abs(display_vals['act_expenses']) > 0.001
            )
            if not has_data:
                continue

            fill = LGREY_FILL if i % 2 == 1 else PatternFill()
            net      = display_vals['act_income'] - display_vals['act_expenses']
            last_net = display_vals['last_yr_inc'] - display_vals['last_yr_exp']

            row_vals = [
                display_label,
                money_or_dash(display_vals['last_yr_inc']),
                money_or_dash(display_vals['last_yr_exp']),
                money_or_dash(display_vals['inc_budget']),
                money_or_dash(display_vals['exp_budget']),
                money_or_dash(display_vals['act_income']),
                money_or_dash(display_vals['act_expenses']),
                money_or_dash(net),
                money_or_dash(last_net),
            ]

            # ... rest of cell rendering stays the same
            for ci, val in enumerate(row_vals, 1):
                c = ws.cell(row=row, column=ci, value=val)
                c.font = BODY_FONT; c.fill = fill; c.border = THIN_BORDER
                if ci == 1:
                    c.alignment = Alignment(indent=2)
                else:
                    c.number_format = MONEY_FMT
                    c.alignment = Alignment(horizontal='right')

            # Highlight net profit
            if vals['act_income'] > 0 or vals['act_expenses'] > 0:
                net_cell = ws.cell(row=row, column=8)
                net_cell.fill = GREEN_FILL if net >= 0 else RED_FILL

            ws.row_dimensions[row].height = 15

            for k in ['last_yr_inc','last_yr_exp','inc_budget',
                      'exp_budget','act_income','act_expenses']:
                sec[k] += display_vals[k]
            row += 1

        # Section total
        sec_net      = sec['act_income'] - sec['act_expenses']
        sec_last_net = sec['last_yr_inc'] - sec['last_yr_exp']

        for ci in range(1, 10):
            ws.cell(row=row, column=ci).fill   = LTBLUE_FILL
            ws.cell(row=row, column=ci).border = MED_BORDER

        ws.cell(row=row, column=1, value='TOTAL:').font = TOTAL_FONT
        ws.cell(row=row, column=1).alignment = Alignment(indent=1)

        for ci, val in enumerate([
            money_or_dash(sec['last_yr_inc']),
            money_or_dash(sec['last_yr_exp']),
            money_or_dash(sec['inc_budget']),
            money_or_dash(sec['exp_budget']),
            money_or_dash(sec['act_income']),
            money_or_dash(sec['act_expenses']),
            money_or_dash(sec_net),
            money_or_dash(sec_last_net),
        ], 2):
            c = ws.cell(row=row, column=ci, value=val)
            c.font = TOTAL_FONT; c.number_format = MONEY_FMT
            c.alignment = Alignment(horizontal='right')

        ws.row_dimensions[row].height = 18
        row += 2
