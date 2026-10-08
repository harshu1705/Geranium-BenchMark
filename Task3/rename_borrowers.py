import openpyxl

RENAME = {
    'ABC Industrial LLC': 'Talbot Steel Fabricators LLC',
    'XYZ Retail Partners': 'Corrigan Retail Partners',
    'Meridian Office Park LLC': 'Fenwick Office Park LLC',
    'Keystone Storage Partners': 'Wexford Storage Partners',
    'Lonestar Data Centers Inc.': 'Prairie Fiber Data Centers Inc.',
    'Cascade Tech Campus': 'Wrenfield Tech Campus',
}

files = [
    'pivot_inputs/Loan_Register_Q4_2025.xlsx',
    'pivot_inputs/Borrower_Financial_Statements_Summary.xlsx',
    'pivot_inputs/Appraisal_and_Collateral_Summary_Q4_2025.xlsx',
    'pivot_inputs/Covenant_Compliance_Workbook_Q4_2025.xlsx',
    'pivot_inputs/Exception_Tracking_Log_Q4_2025.xlsx',
]

for f in files:
    wb = openpyxl.load_workbook(f)
    changed = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value in RENAME:
                    c.value = RENAME[c.value]
                    changed += 1
    wb.save(f)
    print(f, 'changed', changed)
