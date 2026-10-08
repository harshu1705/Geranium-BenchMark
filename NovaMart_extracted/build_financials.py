import openpyxl
from openpyxl.styles import Font
import zipfile
import shutil
import os

path = 'Inputs/Financial_Statements_Summary.xlsx'
wb = openpyxl.load_workbook(path)
ws = wb.active

cols = ['B', 'C', 'D', 'E']  # FY2023, FY2024, FY2025, TTM

# 1. Gross Profit (row 4) = Revenue (row2) - COGS (row3)
for c in cols:
    ws[f'{c}4'] = f'={c}2-{c}3'

# 2. EBITDA (row 6) = Gross Profit (row4) - SG&A (row5)
for c in cols:
    ws[f'{c}6'] = f'={c}4-{c}5'

# 3. Pre-Tax Income (row 9) = EBITDA (row6) - D&A (row7) - Interest Expense (row8)
for c in cols:
    ws[f'{c}9'] = f'={c}6-{c}7-{c}8'

# 4. Insert an "Income Tax Expense" row before Net Income (row 10), pushing everything down by 1
ws.insert_rows(10)
ws['A10'] = 'Income Tax Expense'
tax_values = {'B10': 3000000, 'C10': 4000000, 'D10': 2500000, 'E10': 2300000}
for coord, val in tax_values.items():
    ws[coord] = val
    ws[coord].number_format = '#,##0'

# 5. Net Income is now row 11 = Pre-Tax Income (row9) - Income Tax Expense (row10)
for c in cols:
    ws[f'{c}11'] = f'={c}9-{c}10'

# After the insert, subsequent original rows shift down by 1:
# row11(orig10)->Net Income (row11), row12(dash)->13, Current Assets(orig12)->13,
# Current Liabilities(orig13)->14, Current Ratio(orig14)->15, dash(orig15)->16,
# Total Assets(orig16)->17, Total Debt(orig17)->18, Total Equity(orig18)->19

# 6. Current Ratio (now row 15) = Current Assets(row13) / Current Liabilities(row14)
for c in cols:
    ws[f'{c}15'] = f'={c}13/{c}14'
    ws[f'{c}15'].number_format = '0.00'

# Verify row labels align as expected before continuing
labels = {ws[f'A{r}'].value: r for r in range(1, 20)}
print(labels)

# 7. Append a Ratios & Supplemental Metrics section starting at row 21
start = 21
ws[f'A{start}'] = 'Ratios & Supplemental Metrics'
ws[f'A{start}'].font = Font(bold=True)
r = start + 1

ws[f'A{r}'] = 'Working Capital'
for c in cols:
    ws[f'{c}{r}'] = f'={c}13-{c}14'
    ws[f'{c}{r}'].number_format = '#,##0'
r += 1

ws[f'A{r}'] = 'EBITDA Margin %'
for c in cols:
    ws[f'{c}{r}'] = f'={c}6/{c}2'
    ws[f'{c}{r}'].number_format = '0.0%'
r += 1

ws[f'A{r}'] = 'Net Income Margin %'
for c in cols:
    ws[f'{c}{r}'] = f'={c}11/{c}2'
    ws[f'{c}{r}'].number_format = '0.0%'
r += 1

ws[f'A{r}'] = 'Revenue YoY Growth %'
ws[f'C{r}'] = '=(C2-B2)/B2'
ws[f'D{r}'] = '=(D2-C2)/C2'
ws[f'E{r}'] = '=(E2-D2)/D2'
for c in ['C', 'D', 'E']:
    ws[f'{c}{r}'].number_format = '0.0%'
r += 1

ws[f'A{r}'] = 'Debt-to-Equity Ratio'
for c in cols:
    ws[f'{c}{r}'] = f'={c}18/{c}19'
    ws[f'{c}{r}'].number_format = '0.00'
r += 1

ws[f'A{r}'] = 'Return on Equity (ROE) %'
for c in cols:
    ws[f'{c}{r}'] = f'={c}11/{c}19'
    ws[f'{c}{r}'].number_format = '0.0%'
r += 1

wb.save(path)
print('saved')

# --- Post-process: strip the "Openpyxl" generator string from docProps/app.xml ---
def clean_app_xml(xlsx_path):
    tmp_path = xlsx_path + '.tmp'
    with zipfile.ZipFile(xlsx_path, 'r') as zin:
        names = zin.namelist()
        with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'docProps/app.xml':
                    text = data.decode('utf-8')
                    text = text.replace(
                        'Microsoft Excel Compatible / Openpyxl 3.1.5',
                        'Microsoft Excel'
                    )
                    text = text.replace('<AppVersion>3.1</AppVersion>', '<AppVersion>16.0300</AppVersion>')
                    data = text.encode('utf-8')
                zout.writestr(item, data)
    shutil.move(tmp_path, xlsx_path)

clean_app_xml(path)
print('cleaned app.xml')

# verify
with zipfile.ZipFile(path) as z:
    print(z.read('docProps/app.xml').decode('utf-8'))
