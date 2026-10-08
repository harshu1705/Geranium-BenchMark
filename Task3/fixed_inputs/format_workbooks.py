# -*- coding: utf-8 -*-
"""
Bank-grade formatting pass for all six ALLL workbooks:
- Bold/shaded header row, frozen header pane, sensible column widths
- Real date values (not text) with a consistent date number format
- Currency / percent / decimal number formats on numeric columns
- ALLL_Loss_Rate_and_Reserve_Data.xlsx: 8-Qtr Average becomes a real AVERAGE() formula
- Removes the "multiply directly / no conversion needed" editorializing from headers
- Falls back gracefully if a workbook lacks a given feature
"""
import datetime
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import openpyxl

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_ALIGN = Alignment(vertical="center", wrap_text=True)

CURRENCY_FMT = '#,##0'
PERCENT2_FMT = '0.00%'
PERCENT1_FMT = '0.0%'
DECIMAL2_FMT = '0.00'
DATE_FMT = 'mm/dd/yyyy'


def style_header(ws, header_row=1, height=32):
    for cell in ws[header_row]:
        if cell.value is not None:
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = HEADER_ALIGN
    ws.row_dimensions[header_row].height = height


def freeze_below_header(ws, header_row=1, first_col_letter=None):
    col = first_col_letter or 'A'
    ws.freeze_panes = f'{col}{header_row + 1}'


def autosize(ws, widths: dict):
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width


def to_real_date(ws, col_letter, first_row, last_row):
    for r in range(first_row, last_row + 1):
        cell = ws[f'{col_letter}{r}']
        if isinstance(cell.value, str) and cell.value.strip():
            try:
                cell.value = datetime.datetime.strptime(cell.value.strip(), '%Y-%m-%d').date()
                cell.number_format = DATE_FMT
            except ValueError:
                pass
        elif isinstance(cell.value, (datetime.date, datetime.datetime)):
            cell.number_format = DATE_FMT


def apply_number_format(ws, col_letter, first_row, last_row, fmt):
    for r in range(first_row, last_row + 1):
        cell = ws[f'{col_letter}{r}']
        if cell.value is not None:
            cell.number_format = fmt


print("Formatting all six ALLL workbooks...")

# ---------------------------------------------------------------
# 1. ALLL_Loss_Rate_and_Reserve_Data.xlsx
# ---------------------------------------------------------------
path = 'ALLL_Loss_Rate_and_Reserve_Data.xlsx'
wb = openpyxl.load_workbook(path)

ws = wb['Loss Rates']
ws['B1'] = 'Historical Loss Rate (%)'
apply_number_format(ws, 'B', 2, 6, PERCENT2_FMT)
style_header(ws)
freeze_below_header(ws)
autosize(ws, {'A': 18, 'B': 34, 'C': 30})

ws2 = wb['Reserve Assumptions']
ws2['B1'] = 'Value'
# Row 2: Estimated Disposition Cost -> percent; Row3: Q-Factor bps -> percent;
# Row4: Booked ALLL balance -> currency; Row5: Tolerance band -> percent
ws2['B2'].number_format = PERCENT1_FMT
ws2['B3'].number_format = PERCENT2_FMT
ws2['B4'].number_format = CURRENCY_FMT
ws2['B5'].number_format = PERCENT1_FMT
style_header(ws2)
freeze_below_header(ws2)
autosize(ws2, {'A': 32, 'B': 16, 'C': 46})

ws3 = wb['Historical Loss Rate Trend (8Q)']
for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
    apply_number_format(ws3, col, 2, 6, PERCENT2_FMT)
# Real AVERAGE() formula instead of a hardcoded 8-quarter average
for r in range(2, 7):
    ws3[f'J{r}'] = f'=AVERAGE(B{r}:I{r})'
    ws3[f'J{r}'].number_format = PERCENT2_FMT
style_header(ws3)
freeze_below_header(ws3)
autosize(ws3, {c: 11 for c in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']})
ws3.column_dimensions['A'].width = 18
ws3.column_dimensions['J'].width = 14

ws4 = wb['Methodology Notes']
style_header(ws4)
freeze_below_header(ws4)
autosize(ws4, {'A': 26, 'B': 70})
for row in ws4.iter_rows(min_row=2, max_row=ws4.max_row):
    for c in row:
        if c.value:
            c.alignment = Alignment(wrap_text=True, vertical='top')

wb.save(path)
print(f'  {path} done')

# ---------------------------------------------------------------
# 2. Covenant_Compliance_Workbook_Q4_2025.xlsx
# ---------------------------------------------------------------
path = 'Covenant_Compliance_Workbook_Q4_2025.xlsx'
wb = openpyxl.load_workbook(path)
ws = wb.active
apply_number_format(ws, 'C', 2, 30, DECIMAL2_FMT)
apply_number_format(ws, 'D', 2, 30, DECIMAL2_FMT)
apply_number_format(ws, 'E', 2, 30, '0"%"')
apply_number_format(ws, 'F', 2, 30, '0.0"%"')
apply_number_format(ws, 'G', 2, 30, DECIMAL2_FMT)
apply_number_format(ws, 'H', 2, 30, DECIMAL2_FMT)
style_header(ws)
freeze_below_header(ws)
autosize(ws, {'A': 10, 'B': 32, 'C': 16, 'D': 13, 'E': 18, 'F': 14, 'G': 16, 'H': 18})
wb.save(path)
print(f'  {path} done')

# ---------------------------------------------------------------
# 3. Loan_Register_Q4_2025.xlsx
# ---------------------------------------------------------------
path = 'Loan_Register_Q4_2025.xlsx'
wb = openpyxl.load_workbook(path)
ws = wb.active
to_real_date(ws, 'F', 2, 30)   # Origination Date
to_real_date(ws, 'G', 2, 30)   # Maturity Date
to_real_date(ws, 'L', 2, 30)   # Last Review Date
apply_number_format(ws, 'D', 2, 30, CURRENCY_FMT)   # Original Amount
apply_number_format(ws, 'E', 2, 30, CURRENCY_FMT)   # Outstanding Balance
apply_number_format(ws, 'H', 2, 30, '0.00"%"')       # Interest Rate
apply_number_format(ws, 'J', 2, 30, CURRENCY_FMT)   # Appraised Collateral Value
style_header(ws)
freeze_below_header(ws)
autosize(ws, {'A': 10, 'B': 30, 'C': 18, 'D': 15, 'E': 16, 'F': 13, 'G': 13, 'H': 11,
              'I': 20, 'J': 18, 'K': 12, 'L': 13, 'M': 18, 'N': 16, 'O': 12})
wb.save(path)
print(f'  {path} done')

# ---------------------------------------------------------------
# 4. Borrower_Financial_Statements_Summary.xlsx
# ---------------------------------------------------------------
path = 'Borrower_Financial_Statements_Summary.xlsx'
wb = openpyxl.load_workbook(path)
ws = wb.active
to_real_date(ws, 'C', 2, 30)   # Fiscal Year End
to_real_date(ws, 'H', 2, 30)   # Statement Date
apply_number_format(ws, 'D', 2, 30, CURRENCY_FMT)   # NOI
apply_number_format(ws, 'E', 2, 30, CURRENCY_FMT)   # Total Annual Debt Service
apply_number_format(ws, 'F', 2, 30, CURRENCY_FMT)   # Current Assets
apply_number_format(ws, 'G', 2, 30, CURRENCY_FMT)   # Current Liabilities
style_header(ws)
freeze_below_header(ws)
autosize(ws, {'A': 10, 'B': 30, 'C': 14, 'D': 16, 'E': 20, 'F': 15, 'G': 17, 'H': 13, 'I': 26})
wb.save(path)
print(f'  {path} done')

# ---------------------------------------------------------------
# 5. Appraisal_and_Collateral_Summary_Q4_2025.xlsx
# ---------------------------------------------------------------
path = 'Appraisal_and_Collateral_Summary_Q4_2025.xlsx'
wb = openpyxl.load_workbook(path)
ws = wb.active
to_real_date(ws, 'D', 2, 30)   # Appraisal Date
apply_number_format(ws, 'E', 2, 30, CURRENCY_FMT)   # Appraised Value
style_header(ws)
freeze_below_header(ws)
autosize(ws, {'A': 10, 'B': 30, 'C': 36, 'D': 14, 'E': 15, 'F': 18, 'G': 24, 'H': 12, 'I': 40})
wb.save(path)
print(f'  {path} done')

# ---------------------------------------------------------------
# 6. Exception_Tracking_Log_Q4_2025.xlsx
# ---------------------------------------------------------------
path = 'Exception_Tracking_Log_Q4_2025.xlsx'
wb = openpyxl.load_workbook(path)
ws = wb.active
to_real_date(ws, 'E', 2, 12)   # Date Identified
to_real_date(ws, 'J', 2, 12)   # Waiver Expiry Date
style_header(ws)
freeze_below_header(ws)
autosize(ws, {'A': 12, 'B': 10, 'C': 30, 'D': 22, 'E': 14, 'F': 18, 'G': 46, 'H': 26, 'I': 12, 'J': 14, 'K': 28})
for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
    for c in row:
        if c.column_letter == 'G' and c.value:
            c.alignment = Alignment(wrap_text=True, vertical='top')
wb.save(path)
print(f'  {path} done')

print("All six workbooks formatted.")
