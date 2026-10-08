import openpyxl
from openpyxl.styles import Font
import zipfile
import shutil

path = 'Inputs/Loss_Run_History_2021_2026.xlsx'
wb = openpyxl.load_workbook(path)
ws = wb.active

headers = ['Claim Number', 'Date of Loss', 'Line of Business', 'Description', 'Status',
           'Legal Expenses', 'Forensic Expenses', 'Ransom Paid', 'BI Loss', 'Total Paid', 'Reserve',
           'Claimant', 'Assigned Adjuster']

# Full claim list, chronological. The 3 new rows (marked NEW) are all dated before
# Oct 1, 2023 - i.e. outside the 3-year cyber lookback window used for surcharge tiering -
# so they add loss-run depth/realism without changing the surcharge math.
claims = [
    ['CLM-21-0992', '2021-04-12', 'Property', 'Roof damage from hail storm', 'Closed',
     0, 0, 0, 0, 45000, 0, 'NovaMart Inc. - Store #22 (Peoria, IL)', 'R. Alvarez (Crawford & Co.)'],
    ['CLM-21-0533', '2021-09-10', 'Cyber', 'Phishing-induced wire transfer attempt (funds recovered, no loss)', 'Closed',
     3000, 8000, 0, 0, 11000, 0, 'NovaMart Inc. - Corporate AP Dept.', 'K. Nguyen (Sedgwick)'],
    ['CLM-22-1044', '2022-08-30', 'General Liability', 'Slip and fall in store #14', 'Closed',
     15000, 0, 0, 0, 85000, 0, 'Customer - D. Whitfield', 'R. Alvarez (Crawford & Co.)'],
    ['CLM-22-1587', '2022-12-02', 'Property', 'HVAC failure - inventory spoilage', 'Closed',
     0, 0, 0, 0, 32000, 0, 'NovaMart Inc. - Store #08 (Gary, IN)', 'M. Petrov (Crawford & Co.)'],
    ['CLM-23-4412', '2023-01-15', 'Cyber', 'Lost company laptop (encrypted)', 'Closed',
     2500, 5000, 0, 0, 7500, 0, 'NovaMart Inc. - Regional Sales', 'K. Nguyen (Sedgwick)'],
    ['CLM-23-2210', '2023-05-18', 'General Liability', 'Customer injury - falling merchandise display', 'Closed',
     8000, 0, 0, 0, 28000, 0, 'Customer - L. Osei', 'R. Alvarez (Crawford & Co.)'],
    ['CLM-24-8891', '2024-11-20', 'General Liability', 'Forklift accident in warehouse', 'Open',
     25000, 0, 0, 0, 45000, 150000, 'Employee - T. Bramwell (Joliet DC)', 'M. Petrov (Crawford & Co.)'],
    ['CLM-25-1102', '2025-06-05', 'Cyber', 'Ransomware on legacy ValueStore server', 'Closed',
     75000, 125000, 0, 450000, 650000, 0, 'NovaMart Inc. - ValueStore Division', 'K. Nguyen (Sedgwick)'],
    ['CLM-26-0045', '2026-02-14', 'Property', 'Water damage from broken pipe', 'Closed',
     0, 0, 0, 0, 12000, 0, 'NovaMart Inc. - Store #31 (Kenosha, WI)', 'M. Petrov (Crawford & Co.)'],
    ['CLM-26-0891', '2026-07-22', 'Auto', 'Delivery truck fender bender', 'Open',
     5000, 0, 0, 0, 15000, 20000, 'NovaMart Inc. - Fleet Services', 'R. Alvarez (Crawford & Co.)'],
]

# Clear existing content first
for row in ws.iter_rows():
    for cell in row:
        cell.value = None

for col_idx, h in enumerate(headers, start=1):
    ws.cell(row=1, column=col_idx, value=h)

for row_idx, claim in enumerate(claims, start=2):
    for col_idx, val in enumerate(claim, start=1):
        ws.cell(row=row_idx, column=col_idx, value=val)

n = len(claims)
dash_row = n + 2       # row 12
totals_row = n + 3     # row 13
first_data_row = 2
last_data_row = n + 1  # row 11

for col_idx in range(1, len(headers) + 1):
    ws.cell(row=dash_row, column=col_idx, value='-')

ws.cell(row=totals_row, column=1, value='TOTALS')
for col_idx in range(2, 6):
    ws.cell(row=totals_row, column=col_idx, value='-')

col_letters = ['F', 'G', 'H', 'I', 'J', 'K']
for c in col_letters:
    ws[f'{c}{totals_row}'] = f'=SUM({c}{first_data_row}:{c}{last_data_row})'

ws.cell(row=totals_row, column=12, value='-')
ws.cell(row=totals_row, column=13, value='-')
for col_idx in range(1, len(headers) + 1):
    ws.cell(row=totals_row, column=col_idx).font = Font(bold=True)

wb.save(path)
print(f'saved. TOTALS row is now row {totals_row}, formula range {first_data_row}:{last_data_row}')


def clean_app_xml(xlsx_path):
    tmp_path = xlsx_path + '.tmp'
    with zipfile.ZipFile(xlsx_path, 'r') as zin:
        with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'docProps/app.xml':
                    text = data.decode('utf-8')
                    text = text.replace('Microsoft Excel Compatible / Openpyxl 3.1.5', 'Microsoft Excel')
                    text = text.replace('<AppVersion>3.1</AppVersion>', '<AppVersion>16.0300</AppVersion>')
                    data = text.encode('utf-8')
                zout.writestr(item, data)
    shutil.move(tmp_path, xlsx_path)


clean_app_xml(path)
with zipfile.ZipFile(path) as z:
    print(z.read('docProps/app.xml').decode('utf-8'))
