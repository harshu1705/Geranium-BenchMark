import openpyxl

wb = openpyxl.load_workbook('QSR_Comps_LTM_1Q26.xlsx', data_only=False)

summ = wb['Summary']
summ['A2'] = ("$ in millions except per-share data and multiples  |  Share prices: May 28, 2026  |  "
              "Multiples use comparable (adj) EBITDA, EBIT, and EPS  |  EV on financial-debt basis  |  "
              "MCD FY-base = FY2025 (10-K fiscal year ended Dec 31, 2025)")

evb = wb['EV Bridge']
evb['A24'] = ("Financial debt = long-term debt + short-term borrowings / current portion.  "
              "YUM finance lease obligations of $146mm are a component of the Note 11 long-term debt total (10-Q p.18), and DPZ finance lease principal of $127.259mm sits inside the current and long-term debt captions (10-Q p.13), so neither is added again.  "
              "MCD reports lease liabilities in separate captions from debt and does not split operating from finance at Q1'26 (10-Q p.3).  "
              "Memo leases: MCD $14,776mm total lease liability (current 707 + long-term 14,069), YUM $1,190mm non-current operating lease liabilities (10-Q p.12, the only lease amount disclosed), DPZ $252.240mm operating lease liabilities (current 49.490 + long-term 202.750).  "
              "MCD's finance lease liability at Dec 31, 2025 was $2,352mm (current 23 + long-term 2,329, 10-K p.71).  "
              "Headline multiples use the financial-debt EV.")

wb.save('QSR_Comps_LTM_1Q26.xlsx')
print("Summary + EV Bridge notes updated")
