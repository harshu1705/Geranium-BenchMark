import sys
sys.stdout.reconfigure(encoding='utf-8')
import openpyxl
import os

BASE = r'c:\Users\user\OneDrive\Desktop\geranium project\Task3'
SRC = os.path.join(BASE, 'fixed_inputs')
os.makedirs(SRC, exist_ok=True)

NAME_MAP = {
    'Beta Healthcare Group': 'Ridgeview Healthcare Group',
    'Delta Auto Group': 'Coastal Auto Group',
    'Gamma Apartments LLC': 'Fairview Apartments LLC',
    'Epsilon Restaurant Group': 'Harbor Grille Restaurant Group',
    'Zeta Office Park LLC': 'Meridian Office Park LLC',
    'Eta Print & Packaging Inc.': 'Palmetto Print & Packaging Inc.',
    'Theta Logistics LLC': 'Redline Logistics LLC',
    'Iota Storage Partners': 'Keystone Storage Partners',
    'Kappa Medical Supplies': 'Anderson Medical Supplies',
    'Lambda Homebuilders Inc.': 'Desert Ridge Homebuilders Inc.',
    'Mu Foods Distribution': 'Blue Ridge Foods Distribution',
    'Nu Marina LLC': 'Chesapeake Marina LLC',
    'Xi Fitness Holdings': 'Ironclad Fitness Holdings',
    'Omicron Hotel Group Inc.': 'Lowcountry Hotel Group Inc.',
    'Pi Cold Storage LLC': 'Arctic Cold Storage LLC',
    'Rho Law Group PC': 'Whitfield Law Group PC',
    'Sigma Automotive LLC': 'Interstate Automotive LLC',
    'Tau Senior Living': 'Magnolia Senior Living',
    'Upsilon Manufacturing': 'Buckeye Manufacturing',
    'Phi Data Centers Inc.': 'Lonestar Data Centers Inc.',
    'Chi Pharmacy Group': 'Riverbend Pharmacy Group',
    'Psi Freight LLC': 'Summit Freight LLC',
    'Omega Vineyards': 'Silverado Vineyards',
    'Alpha Property Management': 'Pioneer Property Management',
    'B1 Energy Services': 'Permian Energy Services',
    'C2 Tech Campus': 'Cascade Tech Campus',
    'D3 Staffing Solutions': 'Horizon Staffing Solutions',
}

# Fiscal-year de-mechanization: (loan id -> (new FYE, new Statement Date, new Source))
# Only touches loans NOT involved in any covenant breach narrative (avoids LN-001,002,004,006,008,010)
FYE_FIXES = {
    'LN-005': ('2025-12-31', '2025-11-05', 'CPA-Audited Financial Statements'),   # source varied, FYE kept
    'LN-007': ('2025-07-31', '2025-09-10', 'CPA-Audited Financial Statements'),   # FYE shifted off Jun30
    'LN-011': ('2025-06-30', '2025-05-31', 'Tax Returns'),                        # source varied, FYE kept
    'LN-015': ('2025-08-31', '2025-10-05', 'Internal Financial Statements'),      # FYE shifted off Jun30
    'LN-021': ('2026-01-31', '2025-12-20', 'CPA-Reviewed Financial Statements'),  # FYE shifted off Dec31
    'LN-025': ('2025-10-31', '2025-09-20', 'CPA-Reviewed Financial Statements'),  # ag business, harvest-season FYE
}

def rename_sheet(path, name_col_idx):
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    changed = 0
    for row in ws.iter_rows(min_row=2):
        cell = row[name_col_idx]
        if cell.value in NAME_MAP:
            cell.value = NAME_MAP[cell.value]
            changed += 1
    out_path = os.path.join(SRC, os.path.basename(path))
    wb.save(out_path)
    print(f"{os.path.basename(path)}: renamed {changed} borrower cells -> {out_path}")

# Loan_Register_Q4_2025.xlsx: Loan ID=0, Borrower Name=1
rename_sheet(os.path.join(BASE, 'extracted_inputs', 'Loan_Register_Q4_2025.xlsx'), 1)
# Appraisal_and_Collateral_Summary_Q4_2025.xlsx: Borrower Name=1
rename_sheet(os.path.join(BASE, 'extracted_inputs', 'Appraisal_and_Collateral_Summary_Q4_2025.xlsx'), 1)
# Covenant_Compliance_Workbook_Q4_2025.xlsx: Borrower Name=1
rename_sheet(os.path.join(BASE, 'extracted_inputs', 'Covenant_Compliance_Workbook_Q4_2025.xlsx'), 1)
# Exception_Tracking_Log_Q4_2025.xlsx: Borrower Name=2
rename_sheet(os.path.join(BASE, 'extracted_inputs', 'Exception_Tracking_Log_Q4_2025.xlsx'), 2)

# Borrower_Financial_Statements_Summary.xlsx: Borrower Name=1, plus FYE/StatementDate/Source de-mechanization
path = os.path.join(BASE, 'extracted_inputs', 'Borrower_Financial_Statements_Summary.xlsx')
wb = openpyxl.load_workbook(path)
ws = wb.active
changed = 0
fye_changed = 0
for row in ws.iter_rows(min_row=2):
    loan_id = row[0].value
    name_cell = row[1].value
    if name_cell in NAME_MAP:
        row[1].value = NAME_MAP[name_cell]
        changed += 1
    if loan_id in FYE_FIXES:
        new_fye, new_stmt_date, new_source = FYE_FIXES[loan_id]
        row[2].value = new_fye       # Fiscal Year End
        row[7].value = new_stmt_date # Statement Date
        row[8].value = new_source    # Statement Source
        fye_changed += 1
out_path = os.path.join(SRC, os.path.basename(path))
wb.save(out_path)
print(f"Borrower_Financial_Statements_Summary.xlsx: renamed {changed} borrower cells, de-mechanized {fye_changed} FYE rows -> {out_path}")

# Copy the policy doc unchanged (no borrower names in it)
import shutil
shutil.copy(os.path.join(BASE, 'extracted_inputs', 'Loan_Portfolio_Review_Policy.docx'), SRC)
print("Copied Loan_Portfolio_Review_Policy.docx unchanged")

print("\nDONE. NAME_MAP has", len(NAME_MAP), "entries.")
