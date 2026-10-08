import openpyxl

wb = openpyxl.load_workbook('QSR_Comps_LTM_1Q26.xlsx', data_only=False)
rd = wb['Raw Data']

# --- Annual base period (rows 3-18) ---
rd['B4'] = 'Dec 31, 2025'
rd['E4'] = 'MCD 10-K is FY2025 (year ended Dec 31, 2025); YUM FY2025 (Dec 31, 2025); DPZ FY2025 (Dec 28, 2025)'
rd['B5'] = 'FY2025'

rd['B7'] = 26885
rd['E7'] = 'MCD: 10-K p.53; YUM: 10-K p.59; DPZ: 10-K p.56 total revenues 4,939,994 ($000s converted to $mm)'

rd['B8'] = 12393
rd['E8'] = 'MCD: 10-K p.53; YUM: 10-K p.59; DPZ: 10-K p.56 income from operations 953,974 ($000s converted to $mm)'

rd['B9'] = 229
rd['E9'] = ("MCD: impairment and other charges (gains), net 229, primarily restructuring related to Accelerating the Organization, 10-K p.25; "
            "YUM: Special Items - Operating Profit 122, 10-K p.38; DPZ: refranchising gain (4,028), 10-K p.56 - sign: positive = expense")

rd['B11'] = 2199
rd['E11'] = 'MCD: cash flow statement, 10-K p.56; YUM: cash flow statement, 10-K p.61; DPZ: cash flow statement, 10-K p.59 (88,827 converted to $mm)'

rd['B13'] = 8563
rd['E13'] = 'MCD: 10-K p.53; YUM: 10-K p.59; DPZ: 10-K p.56 (601,704 converted to $mm)'

rd['B14'] = 178
rd['D14'] = -3.146
rd['E14'] = ("MCD: net income reconciliation 8,741 less 8,563, 10-K p.18; YUM: net income excluding Special Items 1,700 less 1,559, 10-K p.38; "
             "DPZ: no adjusted net-income figure published - refranchising gain of (4,028) tax-effected at DPZ's FY2025 effective tax rate of 21.9% (168,634 / 770,338, 10-K p.56)")

rd['B16'] = 11.95
rd['E16'] = 'MCD: 10-K p.53; YUM: 10-K p.59; DPZ: 10-K p.56'

rd['B17'] = 12.20
rd['D17'] = 17.48
rd['E17'] = ("MCD: non-GAAP EPS reconciliation, 10-K p.18; YUM: diluted EPS excluding Special Items, 10-K p.38; "
             "DPZ: no adjusted EPS published - GAAP net income plus the tax-effected refranchising gain, divided by FY2025 weighted-average diluted shares of 34.238mm")

rd['B18'] = 716.4
rd['E18'] = 'MCD: weighted-average diluted, 10-K p.53; YUM: EPS note, 10-K p.75; DPZ: EPS note 34,237,646, 10-K p.68'

# --- Q1 2026 section (rows 20-27) ---
rd['E20'] = "MCD: 10-Q p.4; YUM: 10-Q p.4; DPZ: 10-Q p.4 total revenues 1,150,594 ($000s converted to $mm)"
rd['E21'] = "MCD: 10-Q p.4; YUM: 10-Q p.4; DPZ: 10-Q p.4 income from operations 230,357 ($000s converted to $mm)"
rd['E23'] = "MCD: cash flow statement, 10-Q p.6; YUM: cash flow statement, 10-Q p.6; DPZ: cash flow statement, 10-Q p.6 (20,434 converted to $mm)"
rd['D25'] = -6.021
rd['E25'] = ("MCD: (Gains)/Charges net of tax 36, 10-Q p.21; YUM: net income excluding Special Items 418 less 432, 10-Q p.27; "
             "DPZ: no adjusted figure published - $7.8mm pre-tax gain on sale of assets tax-effected at DPZ's Q1'26 effective tax rate of 22.6% (40,831 / 180,642, 10-Q p.4)")
rd['D26'] = 3.94
rd['E26'] = ("MCD: non-GAAP EPS 2.83, 10-Q p.21; YUM: diluted EPS excluding Special Items 1.50, 10-Q p.28; "
             "DPZ: GAAP net income plus the tax-effected gain on sale of assets, divided by Q1'26 diluted shares of 33.815mm")

# --- Q1 2025 section (rows 28-36) ---
rd['E29'] = "MCD: prior-year column, 10-Q p.4; YUM: prior-year column, 10-Q p.4; DPZ: prior-year column, 10-Q p.4 (1,112,056 converted to $mm)"
rd['E30'] = "MCD: 10-Q p.4; YUM: 10-Q p.4; DPZ: 10-Q p.4 (210,095 converted to $mm)"
rd['E32'] = "MCD: cash flow statement, 10-Q p.6; YUM: cash flow statement, 10-Q p.6; DPZ: cash flow statement, 10-Q p.6 (20,362 converted to $mm)"
rd['E34'] = ("MCD: (Gains)/Charges net of tax 51, 10-Q p.21; YUM: net income excluding Special Items 366 less 253, 10-Q p.27; "
             "DPZ: no one-time item in the prior-year quarter, so no tax-effect applies - 0")
rd['E35'] = "MCD: non-GAAP EPS 2.67, 10-Q p.21; YUM: diluted EPS excluding Special Items 1.30, 10-Q p.28; DPZ: 10-Q p.4 - GAAP EPS, no one-time item in the prior-year quarter"

wb.save('QSR_Comps_LTM_1Q26.xlsx')
print("Raw Data updated")
