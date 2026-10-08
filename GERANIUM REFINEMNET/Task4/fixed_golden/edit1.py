import openpyxl

wb = openpyxl.load_workbook('QSR_Comps_LTM_1Q26.xlsx', data_only=False)

ltm = wb['LTM Build']
# Headers
ltm['B3'] = "McDonald's (MCD)  \u2014  FY-base: FY2025"
ltm['B4'] = 'FY2025'

# MCD FY-base annual figures -> FY2025 (year ended Dec 31, 2025)
ltm['B5'] = 26885      # Total revenues
ltm['B6'] = 12393      # GAAP operating income
ltm['B7'] = 229        # Impairment and other charges (gains), net - pretax one-time add-back
ltm['B9'] = 2199       # D&A per cash flow statement
ltm['B11'] = 8563      # GAAP net income
ltm['B12'] = 178       # (Gains)/Charges, net of tax - non-GAAP net income reconciliation
ltm['B14'] = 716.4     # FY2025 weighted-average diluted shares (informational; LTM EPS uses Q1'26 count in C14)

# DPZ one-time items, after-tax: tax-effect the pre-tax refranchising gain / aircraft-sale gain
# using each period's own effective tax rate from the DPZ filings, instead of leaving at 0.
# FY2025: provision for income taxes 168.634 / income before taxes 770.338 (10-K p.56)
# Q1'26: provision for income taxes 40.831 / income before taxes 180.642 (10-Q p.4)
# Q1'25: provision for income taxes 42.831 / income before taxes 192.482 (10-Q p.4) - no one-time item, tax-effect is moot but kept for consistency
ltm['L12'] = '=L7*(1-168.634/770.338)'
ltm['M12'] = '=M7*(1-40.831/180.642)'
ltm['N12'] = '=N7*(1-42.831/192.482)'

# Update source note (row 21)
ltm['A21'] = ("Source: MCD 10-K FY2025 (p.16, 18, 25, 53, 56) + MCD 10-Q Q1'26 (p.3, 4, 6, 21).  "
              "YUM 10-K FY2025 (p.38, 59, 61, 75) + YUM 10-Q Q1'26 (p.4, 6, 7, 10, 12, 18, 27, 28).  "
              "DPZ 10-K FY2025 (p.56, 59, 68) + DPZ 10-Q Q1'26 (p.3, 4, 6, 9, 13).  "
              "D&A from the cash flow statement.  One-time items from each company's non-GAAP reconciliation, or from the income statement where the company publishes none (DPZ).  "
              "DPZ one-time items are tax-effected using DPZ's own effective tax rate for each period (FY2025 10-K p.56; Q1'26 and Q1'25 10-Q p.4), since DPZ publishes no adjusted net-income figure.  "
              "DPZ figures converted from $000s to $mm.  LTM EPS uses the most recent (Q1'26) diluted share count.")

wb.save('QSR_Comps_LTM_1Q26.xlsx')
print("LTM Build updated")
