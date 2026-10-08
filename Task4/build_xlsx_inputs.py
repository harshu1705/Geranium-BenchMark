# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime
import os

BASE = os.path.join(os.path.dirname(__file__), "PIVOT2_inputs")
os.makedirs(BASE, exist_ok=True)

HEADER_FILL = PatternFill(fill_type='solid', fgColor='1F4E78')
HEADER_FONT = Font(bold=True, color='FFFFFF')

def style_header(ws, ncols, row=1):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.freeze_panes = ws.cell(row=row + 1, column=1).coordinate

def autosize(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

# ============================================================
# 1. Contractor Repair Estimate
# ============================================================
wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'Repair Estimate'
headers = ['Line Item', 'Category', 'Amount ($)', 'Notes']
ws.append(headers)
rows = [
    ['Structural repair - production floor & roof framing', 'Like-Kind-and-Quality Repair', 620000, 'Fire and heat damage to steel framing and roof deck'],
    ['Electrical system replacement (like-kind)', 'Like-Kind-and-Quality Repair', 310000, 'Replacement of damaged panels, conduit, and wiring to prior specification'],
    ['Roofing and exterior wall restoration', 'Like-Kind-and-Quality Repair', 220000, 'Replacement of damaged roofing membrane and exterior wall panels'],
    ['Interior finishes and fixtures', 'Like-Kind-and-Quality Repair', 0, 'No separate claim; included in structural line above'],
    ['Fire sprinkler system upgrade to current code', 'Code-Required Upgrade (Ordinance or Law)', 120000, 'Current fire code requires full sprinkler coverage; prior system was only partial and grandfathered'],
    ['Electrical service upgrade to current code', 'Code-Required Upgrade (Ordinance or Law)', 60000, 'Current electrical code requires upgraded panel capacity and arc-fault protection beyond prior installation'],
    ['Debris removal and disposal', 'Debris Removal', 45000, 'Removal of fire debris, damaged materials, and hazardous waste disposal'],
]
for r in rows:
    ws.append(r)
lkq_total = sum(r[2] for r in rows if r[1] == 'Like-Kind-and-Quality Repair')
ol_total = sum(r[2] for r in rows if 'Ordinance' in r[1])
debris_total = sum(r[2] for r in rows if r[1] == 'Debris Removal')
ws.append(['Like-Kind-and-Quality Repair Subtotal', '', lkq_total, ''])
ws.append(['Code-Required Upgrade (Ordinance or Law) Subtotal', '', ol_total, ''])
ws.append(['Debris Removal Subtotal', '', debris_total, ''])
ws.append(['Estimated Repair Completion Timeline', '', '', '7 months from date of loss (by approximately May 14, 2026)'])
style_header(ws, len(headers))
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=3, max_col=3):
    for cell in row:
        if isinstance(cell.value, (int, float)):
            cell.number_format = '#,##0'
autosize(ws, [42, 32, 14, 55])
wb.save(os.path.join(BASE, 'Contractor_Repair_Estimate.xlsx'))
print('LKQ total', lkq_total, 'O&L total', ol_total, 'Debris total', debris_total)

# ============================================================
# 2. Business Financial Statements
# ============================================================
wb = openpyxl.Workbook()
ws1 = wb.active
ws1.title = 'Pre-Loss Monthly Trend'
headers = ['Month', 'Net Revenue ($)', 'Total Operating Expenses ($)', 'Net Income ($)']
ws1.append(headers)
# Oct 2024 - Sep 2025, monthly net income averaging toward a ~592,000 total for Oct24-Apr25 (7 months)
pre_loss = [
    (datetime(2024,10,1), 640000, 555000),
    (datetime(2024,11,1), 655000, 565000),
    (datetime(2024,12,1), 700000, 598000),
    (datetime(2025,1,1),  610000, 528000),
    (datetime(2025,2,1),  615000, 533000),
    (datetime(2025,3,1),  680000, 580000),
    (datetime(2025,4,1),  660000, 568000),
    (datetime(2025,5,1),  670000, 575000),
    (datetime(2025,6,1),  690000, 588000),
    (datetime(2025,7,1),  700000, 596000),
    (datetime(2025,8,1),  685000, 584000),
    (datetime(2025,9,1),  695000, 591000),
]
for dt, rev, exp in pre_loss:
    ws1.append([dt, rev, exp, rev - exp])
style_header(ws1, len(headers))
for row in ws1.iter_rows(min_row=2, max_row=ws1.max_row, min_col=2, max_col=4):
    for cell in row:
        cell.number_format = '#,##0'
for row in ws1.iter_rows(min_row=2, max_row=ws1.max_row, min_col=1, max_col=1):
    for cell in row:
        cell.number_format = 'mmm-yyyy'
autosize(ws1, [14, 18, 26, 16])

oct24_apr25_ni = sum(rev-exp for dt,rev,exp in pre_loss if dt <= datetime(2025,4,1))
print('Oct24-Apr25 actual Net Income (prior-year comparable for period of restoration):', oct24_apr25_ni)

ws1.append([])
ws1.append(['Note: Same 7-month window one year prior to the loss (October 2024 through April 2025) is the prior-year comparable period for the Business Income calculation below. Company revenue has grown at approximately 3% year-over-year for the trailing eight quarters per management review.'])

ws2 = wb.create_sheet('Post-Loss Partial Operations')
headers2 = ['Month', 'Net Revenue at Temporary Facility ($)', 'Operating Expenses ($)', 'Net Income ($)', 'Extra Expense Incurred ($)']
ws2.append(headers2)
post_loss = [
    (datetime(2025,10,1), 0, 0, 0, 18000),
    (datetime(2025,11,1), 0, 0, 0, 22000),
    (datetime(2025,12,1), 140000, 122000, 18000, 15000),
    (datetime(2026,1,1),  190000, 163000, 27000, 12000),
    (datetime(2026,2,1),  205000, 175000, 30000, 10000),
    (datetime(2026,3,1),  215000, 182000, 33000, 9000),
    (datetime(2026,4,1),  225000, 188000, 37000, 9000),
]
for dt, rev, exp, ni, ee in post_loss:
    ws2.append([dt, rev, exp, ni, ee])
style_header(ws2, len(headers2))
for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row, min_col=2, max_col=5):
    for cell in row:
        cell.number_format = '#,##0'
for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row, min_col=1, max_col=1):
    for cell in row:
        cell.number_format = 'mmm-yyyy'
autosize(ws2, [14, 30, 20, 16, 20])
actual_ni_total = sum(ni for dt,rev,exp,ni,ee in post_loss)
extra_expense_total = sum(ee for dt,rev,exp,ni,ee in post_loss)
print('Actual Net Income earned during period of restoration:', actual_ni_total)
print('Total Extra Expense incurred:', extra_expense_total)

ws3 = wb.create_sheet('BPP Loss Detail')
headers3 = ['Equipment / Inventory Item', 'Replacement Cost Value ($)', 'Accumulated Depreciation (%)', 'Actual Cash Value ($)']
ws3.append(headers3)
bpp_items = [
    ['CNC machining center (2019)', 280000, 0.40, None],
    ['Precision grinding equipment (2021)', 150000, 0.25, None],
    ['Raw material and WIP inventory', 110000, 0.10, None],
    ['Tooling, fixtures, and small equipment', 80000, 0.50, None],
]
for item in bpp_items:
    item[3] = round(item[1] * (1 - item[2]))
    ws3.append(item)
rcv_total = sum(i[1] for i in bpp_items)
acv_total = sum(i[3] for i in bpp_items)
ws3.append(['Total', rcv_total, '', acv_total])
style_header(ws3, len(headers3))
for row in ws3.iter_rows(min_row=2, max_row=ws3.max_row, min_col=2, max_col=2):
    for cell in row:
        cell.number_format = '#,##0'
for row in ws3.iter_rows(min_row=2, max_row=ws3.max_row, min_col=3, max_col=3):
    for cell in row:
        if isinstance(cell.value, float):
            cell.number_format = '0%'
for row in ws3.iter_rows(min_row=2, max_row=ws3.max_row, min_col=4, max_col=4):
    for cell in row:
        cell.number_format = '#,##0'
autosize(ws3, [40, 24, 24, 22])
print('BPP RCV total', rcv_total, 'ACV total', acv_total, 'depreciation', rcv_total-acv_total)

wb.save(os.path.join(BASE, 'Business_Financial_Statements.xlsx'))

# ============================================================
# 3. Statement of Values / Appraisal
# ============================================================
wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'Statement of Values'
headers = ['Coverage', 'Appraised Value at Time of Loss ($)', 'Amount of Insurance Carried ($)', 'Coinsurance Percentage', 'Appraisal Date', 'Appraiser']
ws.append(headers)
ws.append(['Building (Coverage A)', 2700000, 2000000, 0.80, datetime(2025,11,10), 'Overstreet Valuation Group'])
ws.append(['Business Personal Property (Coverage B)', 780000, 800000, 0.80, datetime(2025,11,10), 'Overstreet Valuation Group'])
style_header(ws, len(headers))
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=2, max_col=3):
    for cell in row:
        cell.number_format = '#,##0'
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=4, max_col=4):
    for cell in row:
        cell.number_format = '0%'
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=5, max_col=5):
    for cell in row:
        cell.number_format = 'mm/dd/yyyy'
autosize(ws, [38, 30, 28, 20, 16, 26])
ws.append([])
ws.append(['Note: Appraised values reflect current replacement cost construction pricing as of the appraisal date, obtained at the Company\'s request following notice of loss. The Named Insured had not updated declared values since policy inception in January 2025.'])
wb.save(os.path.join(BASE, 'Statement_of_Values_Appraisal.xlsx'))

print('DONE')
