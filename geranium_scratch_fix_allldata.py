import openpyxl
from openpyxl import Workbook

src = "C:/Users/user/OneDrive/Desktop/geranium project/Task3/pivot_inputs/ALLL_Loss_Rate_and_Reserve_Data.xlsx"

wb = openpyxl.load_workbook(src)

# Sheet 3: Historical Loss Rate Trend (supports the current-quarter rates used in Sheet 1)
ws3 = wb.create_sheet("Historical Loss Rate Trend (8Q)")
ws3.append(["Internal Risk Rating", "2024-Q1", "2024-Q2", "2024-Q3", "2024-Q4",
            "2025-Q1", "2025-Q2", "2025-Q3", "2025-Q4 (Current)", "8-Qtr Average"])
trend = {
    1: [0.0009, 0.0011, 0.0010, 0.0008, 0.0010, 0.0009, 0.0011, 0.0010],
    2: [0.0033, 0.0038, 0.0036, 0.0032, 0.0037, 0.0034, 0.0035, 0.0035],
    3: [0.0081, 0.0095, 0.0088, 0.0079, 0.0092, 0.0086, 0.0090, 0.0090],
    4: [0.0298, 0.0341, 0.0312, 0.0289, 0.0335, 0.0308, 0.0327, 0.0325],
    5: [0.0872, 0.0985, 0.0918, 0.0846, 0.0961, 0.0903, 0.0947, 0.0950],
}
for rating, vals in trend.items():
    avg = round(sum(vals) / len(vals), 4)
    ws3.append([rating] + vals + [avg])
ws3.append([
    "Note: Rating 6 is excluded from pooled trending because Rating-6 loans are always individually "
    "evaluated under Policy Section 4 and never carry a pooled historical loss rate.",
    None, None, None, None, None, None, None, None, None
])
ws3.append([
    "Current-quarter rates in the 'Loss Rates' tab reflect the 2025-Q4 column above and are refreshed "
    "quarterly by the Finance Department credit risk modeling team.",
    None, None, None, None, None, None, None, None, None
])

# Sheet 4: Methodology and data-source notes
ws4 = wb.create_sheet("Methodology Notes")
ws4.append(["Item", "Detail"])
ws4.append(["Loss rate derivation", "12-quarter rolling historical charge-off rate by Internal Risk Rating, calculated by Finance Department credit modeling, refreshed quarterly."])
ws4.append(["Pooling basis", "Pooled by Internal Risk Rating per Policy Section 3; ratings share similar risk characteristics (collateral type, payment history, financial trend) validated annually by the Model Risk Committee."])
ws4.append(["Disposition cost basis", "10% estimate derived from FY2025 collateral liquidation cost study covering auction, legal, and holding costs on foreclosed commercial collateral."])
ws4.append(["Q-Factor basis", "25bps add-on reflects emerging risk in a pool from an active, uncured covenant breach not yet captured in historical loss experience; approved per Policy Section 5."])
ws4.append(["Booked ALLL balance source", "General Ledger account 1720, quarter-end trial balance as of 12/31/2025."])
ws4.append(["Tolerance band basis", "Board-approved +/-10% band per Policy Section 6, consistent with regulatory guidance on CECL reserve estimation uncertainty."])
ws4.append(["Model validation", "Loss rate model last independently validated by internal audit Q3 2025; next validation scheduled Q3 2026."])
ws4.append(["Data refresh cycle", "Loss rates, disposition cost, and booked balance are refreshed each quarter-end prior to the ALLL Reserve Adequacy Memorandum being drafted."])

wb.save(src)
print("saved")

# verify count
wb2 = openpyxl.load_workbook(src, data_only=True)
total = 0
for ws in wb2.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None:
                total += 1
print("total populated cells:", total)
