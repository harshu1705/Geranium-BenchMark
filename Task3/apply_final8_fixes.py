import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
import shutil, os

BASE = os.path.join(os.path.dirname(__file__), "FINAL8_inputs")

RENAMES = {
    "Horizon Staffing Solutions": "Bramwell Staffing Group",
    "Summit Freight LLC": "Ridgeline Freight LLC",
    "Pioneer Property Management": "Cascade Property Management",
    "Innovance Business Park LLC": "Meridian Business Park LLC",
}

def rename_in_workbook(path):
    wb = openpyxl.load_workbook(path)
    changed = False
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    for old, new in RENAMES.items():
                        if old in cell.value:
                            cell.value = cell.value.replace(old, new)
                            changed = True
    if changed:
        wb.save(path)
    print(f"renamed in {os.path.basename(path)}: {changed}")

# 1. Rename LLM-flagged borrower names across all data files
for fname in [
    "Loan_Register_Q4_2025.xlsx",
    "Borrower_Financial_Statements_Summary.xlsx",
    "Appraisal_and_Collateral_Summary_Q4_2025.xlsx",
    "Covenant_Compliance_Workbook_Q4_2025.xlsx",
]:
    rename_in_workbook(os.path.join(BASE, fname))

# 2. Appraisal_and_Collateral_Summary_Q4_2025.xlsx - update LN-019 to a genuine
#    Q4 2025 (pre-close) revaluation that is LOWER than the stale figure baked
#    into the Covenant Compliance Workbook's Actual LTV, so independent
#    recomputation flips LN-019 into an LTV breach that the workbook does not show
#    and that has no corresponding Exception Tracking Log entry at all.
path = os.path.join(BASE, "Appraisal_and_Collateral_Summary_Q4_2025.xlsx")
wb = openpyxl.load_workbook(path)
ws = wb.active
for row in ws.iter_rows(min_row=2):
    if row[0].value == "LN-019":
        row[2].value = "Vehicle Inventory – Floor Plan Audit Dec 2025"  # Collateral Description
        row[3].value = datetime(2025, 12, 5)  # Appraisal Date
        row[4].value = 4800000  # Appraised Value ($)
        row[8].value = ("Value declined from $5.8M at the January 2025 audit; December 2025 floor-plan "
                         "audit reflects softened used-vehicle wholesale pricing.")  # Notes
        print("Updated LN-019 appraisal row:", [c.value for c in row])
wb.save(path)

# 3. Loan_Register_Q4_2025.xlsx - keep the Appraised Collateral Value column in sync
#    with the true current appraisal (this is the "source of truth" value used for
#    recomputation; the Covenant Compliance Workbook's Actual LTV is deliberately
#    NOT updated, mirroring the LN-002 stale-workbook pattern).
path = os.path.join(BASE, "Loan_Register_Q4_2025.xlsx")
wb = openpyxl.load_workbook(path)
ws = wb.active
header = [c.value for c in ws[1]]
col_appraised = header.index("Appraised Collateral Value ($)")
for row in ws.iter_rows(min_row=2):
    if row[0].value == "LN-019":
        row[col_appraised].value = 4800000
        print("Updated LN-019 loan register appraised value:", row[col_appraised].value)
wb.save(path)

# 4. Exception_Tracking_Log_Q4_2025.xlsx - add an explicit Waiver Approved Date
#    column so LN-001's waiver can be confirmed as approved and in force before
#    the 12/31/2025 quarter-end review date (previously only an expiry date and a
#    remediation-plan submission date were present, leaving approval timing unclear).
path = os.path.join(BASE, "Exception_Tracking_Log_Q4_2025.xlsx")
wb = openpyxl.load_workbook(path)
ws = wb.active
header_font = ws["A1"].font
header_fill = ws["A1"].fill
new_col = ws.max_column + 1
header_cell = ws.cell(row=1, column=new_col, value="Waiver Approved Date")
header_cell.font = Font(bold=True, color=header_font.color)
header_cell.fill = PatternFill(fill_type=header_fill.fill_type, fgColor=header_fill.fgColor)
header_cell.alignment = Alignment(horizontal="center")
date_fmt = ws["J2"].number_format  # matches Waiver Expiry Date format
for row in ws.iter_rows(min_row=2):
    loan_row_id = row[0].value  # Exception ID
    cell = ws.cell(row=row[0].row, column=new_col)
    if loan_row_id == "EXC-001":
        cell.value = datetime(2025, 11, 25)
        cell.number_format = date_fmt
    else:
        cell.value = None
from openpyxl.utils import get_column_letter
ws.column_dimensions[get_column_letter(new_col)].width = 20
wb.save(path)
print("Added Waiver Approved Date column to Exception Tracking Log")

# 5. ALLL_Loss_Rate_and_Reserve_Data.xlsx - wrap the 8-quarter average formulas in
#    ROUND(...,4) so the cached/computed values don't carry IEEE-754 floating point
#    noise (e.g. 0.0009750000000000001) when the workbook is opened.
path = os.path.join(BASE, "ALLL_Loss_Rate_and_Reserve_Data.xlsx")
wb = openpyxl.load_workbook(path)
ws = wb["Historical Loss Rate Trend (8Q)"]
for r in range(2, 7):
    cell = ws.cell(row=r, column=10)
    formula = cell.value
    if isinstance(formula, str) and formula.startswith("=AVERAGE("):
        inner = formula[len("=AVERAGE("):-1]
        cell.value = f"=ROUND(AVERAGE({inner}),4)"
        print("Updated formula row", r, "->", cell.value)
wb.save(path)

print("DONE")
