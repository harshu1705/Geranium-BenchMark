import pandas as pd
import json
import os

input_dir = r"c:/Users/user/OneDrive/Desktop/geranium project/Task3/extracted_inputs"
files = {
    "Loan_Register": "Loan_Register_Q4_2025.xlsx",
    "Financials": "Borrower_Financial_Statements_Summary.xlsx",
    "Appraisals": "Appraisal_and_Collateral_Summary_Q4_2025.xlsx",
    "Covenants": "Covenant_Compliance_Workbook_Q4_2025.xlsx",
    "Exceptions": "Exception_Tracking_Log_Q4_2025.xlsx"
}

data = {}
for key, file in files.items():
    df = pd.read_excel(os.path.join(input_dir, file))
    print(f"--- {key} ---")
    print(df.head())
    data[key] = df.to_dict(orient='records')
    
with open("data_summary.json", "w") as f:
    json.dump(data, f, indent=4)
