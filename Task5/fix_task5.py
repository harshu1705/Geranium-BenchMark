# -*- coding: utf-8 -*-
"""
One-shot fix script for Task 5 (Pruitt Retirement Readiness Assessment).
Applies, in place, on the live files under PIVOT3_inputs/ and PIVOT3_golden/:

  1. Age alignment: Renee is changed from 58 to 60 (same age as Daniel) so both
     spouses reach every Social Security / Medicare milestone age in the same
     year. This removes the cross-eligibility timing defects flagged in the
     reference-soundness and golden-source-fidelity reviews (Renee's own-record
     benefit / Medicare eligibility no longer lags Daniel's by ~2 years).
  2. RMD age fixed from 73 to 75 (SECURE 2.0 applies age 75 to anyone born in
     1960 or later; both spouses are born in 1966).
  3. Self-consistency: the two "X / 0.04 = Y" sentences now show the unrounded
     pre-rounding figure inline so the stated arithmetic reproduces the stated
     result exactly.
  4. Estate-document dates un-conflated: wills (2011) vs. beneficiary
     designations (2019) vs. POA/healthcare directives (never on file) are now
     stated separately instead of being bundled as "since 2019 at the latest".
  5. Added the missing flag that Renee's own group term life ends at
     retirement, leaving her with no coverage at all.
  6. Expanded input file content depth (new supporting sheets / sections) with
     no change to any figure used by the rubric.
"""
import os
import copy
import docx
from docx.shared import Pt
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(__file__)
INPUTS = os.path.join(BASE, "PIVOT3_inputs")
GOLDEN = os.path.join(BASE, "PIVOT3_golden")

HEADER_FILL = PatternFill(fill_type='solid', fgColor='1F4E78')
HEADER_FONT = Font(bold=True, color='FFFFFF')


def style_header(ws, ncols, row=1):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.freeze_panes = ws.cell(row=row + 1, column=1).coordinate


def autosize(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def set_para_text(para, new_text):
    """Replace a paragraph's visible text while keeping its first run's formatting."""
    if not para.runs:
        para.add_run(new_text)
        return
    para.runs[0].text = new_text
    for extra in para.runs[1:]:
        extra.text = ''


def find_para(doc, needle):
    for p in doc.paragraphs:
        if needle in p.text:
            return p
    raise ValueError(f"paragraph not found: {needle[:60]!r}")


# ============================================================
# 1. Client_Intake_and_Goals_Summary.docx -- age fix only
# ============================================================
path = os.path.join(INPUTS, "Client_Intake_and_Goals_Summary.docx")
d = docx.Document(path)
para = find_para(d, "Renee Pruitt, age 58")
set_para_text(para, "Renee Pruitt, age 60, Office Administrator, plans to retire at the same time as Daniel.")
d.save(path)
print("fixed Client_Intake_and_Goals_Summary.docx")

# ============================================================
# 2. Social_Security_Benefit_Estimates.docx -- DOB fix + expanded content
# ============================================================
path = os.path.join(INPUTS, "Social_Security_Benefit_Estimates.docx")
d = docx.Document(path)
para = find_para(d, "Renee Pruitt (DOB:")
set_para_text(para, "Renee Pruitt (DOB: October 1966)")

d.add_heading('How These Estimates Are Calculated', level=1)
d.add_paragraph(
    'Each estimate is based on the worker\'s Average Indexed Monthly Earnings (AIME), calculated from the '
    'highest 35 years of wage-indexed earnings on record, which is then run through the Social Security '
    'Primary Insurance Amount (PIA) formula to arrive at the benefit payable at Full Retirement Age. The '
    'age-62 and age-67 figures above are the age-62 (early) and age-67 (Full Retirement Age) versions of that '
    'same PIA, adjusted for early or delayed claiming as described below. These are estimates only; actual '
    'benefits depend on continued earnings between now and the claiming date and are not guaranteed.'
)

d.add_heading('Full Retirement Age and the Effect of Early or Delayed Claiming', level=1)
d.add_paragraph(
    'For anyone born in 1960 or later, including both Daniel and Renee, Full Retirement Age (FRA) under '
    'current law is 67. Claiming at 62, the earliest eligibility age, results in a permanent reduction of '
    'approximately 30% from the FRA benefit. Claiming at 65 results in a smaller permanent reduction of '
    'approximately 13.3% from the FRA benefit. Waiting past FRA accrues delayed retirement credits of '
    'approximately 8% per year up to age 70, after which no further credits accrue. The figures shown above '
    'for ages 62, 65, and 67 already reflect these respective reductions and credits.'
)

d.add_heading('Spousal and Survivor Benefits', level=1)
d.add_paragraph(
    'The estimates above reflect only each individual\'s own earnings record. Spousal benefits (up to 50% of '
    'a spouse\'s FRA benefit) and survivor benefits (up to 100% of a deceased spouse\'s benefit) may be '
    'available under separate rules and are not included in the figures above. Because both Daniel\'s and '
    'Renee\'s own-record benefits at every claiming age shown exceed any spousal benefit either would be '
    'entitled to on the other\'s record, spousal benefits are not expected to change the claiming analysis '
    'for this household, but eligibility should be confirmed with the Social Security Administration directly '
    'at the time of filing.'
)

d.add_heading('Cost-of-Living Adjustments (COLA)', level=1)
d.add_paragraph(
    'Once benefits begin, they are adjusted annually for inflation via the Cost-of-Living Adjustment, based on '
    'changes in the Consumer Price Index for Urban Wage Earners and Clerical Workers (CPI-W). The estimates '
    'above are stated in today\'s dollars and do not attempt to project future COLA amounts, which vary from '
    'year to year and cannot be known in advance.'
)

d.add_heading('Taxation of Social Security Benefits', level=1)
d.add_paragraph(
    'Depending on total household income in retirement, up to 85% of Social Security benefits may be subject '
    'to federal income tax under the combined-income thresholds set by the Internal Revenue Service. State '
    'taxation of benefits varies by state of residence. This statement does not constitute tax advice; the '
    'household should confirm the expected tax treatment of these benefits with a qualified tax professional '
    'as part of finalizing a claiming strategy.'
)

d.add_heading('Earnings Test for Benefits Claimed Before Full Retirement Age', level=1)
d.add_paragraph(
    'If a benefit is claimed before Full Retirement Age and the recipient continues to work, the annual '
    'earnings test may temporarily withhold a portion of the benefit once earned income exceeds an annually '
    'set threshold; withheld amounts are later credited back into the benefit calculation after FRA is '
    'reached. This provision only applies to a recipient who is both claiming early and still working; it is '
    'not relevant to a household that is fully retired and not earning wages at the time benefits begin.'
)

d.save(path)
print("fixed + expanded Social_Security_Benefit_Estimates.docx")

# ============================================================
# 3. Account_Statements_Summary.xlsx -- add supporting sheets
# ============================================================
path = os.path.join(INPUTS, "Account_Statements_Summary.xlsx")
wb = openpyxl.load_workbook(path)

ws3 = wb.create_sheet('Balance History (Trailing 4Q)')
ws3.append(['Account', 'Q1 2026 Balance ($)', 'Q2 2026 Balance ($)', 'Q3 2026 Balance ($)', 'Q4 2026 / Current Balance ($)'])
hist_rows = [
    ['Daniel 401(k)', 640000, 655000, 668000, 680000],
    ['Renee 403(b)', 295000, 300000, 305000, 310000],
    ['Joint Taxable Brokerage', 205000, 208000, 211000, 215000],
    ['Renee Roth IRA', 88000, 90000, 92500, 95000],
    ['Joint Savings', 40000, 42000, 43500, 45000],
]
for r in hist_rows:
    ws3.append(r)
style_header(ws3, 5)
for row in ws3.iter_rows(min_row=2, max_row=ws3.max_row, min_col=2, max_col=5):
    for cell in row:
        cell.number_format = '#,##0'
autosize(ws3, [26, 20, 20, 20, 26])

ws4 = wb.create_sheet('Recent Contribution Detail')
ws4.append(['Month', "Daniel 401(k) Contribution ($)", 'Renee 403(b) Contribution ($)', 'Brokerage Contribution ($)', 'Total ($)'])
contrib_rows = [
    ['May 2026', 2000, 800, 300, 3100],
    ['June 2026', 2000, 800, 300, 3100],
    ['July 2026', 2000, 800, 300, 3100],
    ['August 2026', 2000, 800, 300, 3100],
]
for r in contrib_rows:
    ws4.append(r)
ws4.append(['4-Month Total', 8000, 3200, 1200, 12400])
style_header(ws4, 5)
for row in ws4.iter_rows(min_row=2, max_row=ws4.max_row, min_col=2, max_col=5):
    for cell in row:
        cell.number_format = '#,##0'
autosize(ws4, [18, 26, 24, 24, 14])

wb.save(path)
print("expanded Account_Statements_Summary.xlsx")

# ============================================================
# 4. Household_Budget_and_Planning_Assumptions.xlsx -- add supporting sheets
# ============================================================
path = os.path.join(INPUTS, "Household_Budget_and_Planning_Assumptions.xlsx")
wb = openpyxl.load_workbook(path)

# ---- Annualized Budget Detail (pure derivation of Sheet 1, no new figures) ----
ws3 = wb.create_sheet('Annualized Budget Detail')
ws3.append(['Expense Category', 'Current Monthly ($)', 'Current Annual ($)', 'Retirement Monthly ($)', 'Retirement Annual ($)'])
cat_rows = [
    ('Housing (mortgage, taxes, insurance, maintenance)', 2400, 2600),
    ('Healthcare (premiums, out-of-pocket)', 650, 1200),
    ('Food and household', 900, 850),
    ('Transportation', 700, 500),
    ('Travel and leisure', 500, 1300),
    ('Insurance (life, auto, umbrella)', 380, 380),
    ('Retirement account contributions', 2800, 0),
    ('Brokerage contributions', 300, 0),
    ('Miscellaneous / discretionary', 570, 1670),
]
for name, cur_m, ret_m in cat_rows:
    ws3.append([name, cur_m, cur_m * 12, ret_m, ret_m * 12])
cur_m_total = sum(r[1] for r in cat_rows)
ret_m_total = sum(r[2] for r in cat_rows)
ws3.append(['Total', cur_m_total, cur_m_total * 12, ret_m_total, ret_m_total * 12])
style_header(ws3, 5)
for row in ws3.iter_rows(min_row=2, max_row=ws3.max_row, min_col=2, max_col=5):
    for cell in row:
        cell.number_format = '#,##0'
autosize(ws3, [42, 18, 18, 20, 20])

# ---- Pre-Retirement Contribution Schedule (ties to the $74,400 used in the projection) ----
ws4 = wb.create_sheet('Pre-Retirement Contributions')
ws4.append(['Quarter', 'Combined Monthly Contribution ($)', 'Contribution This Quarter ($)', 'Cumulative Contribution ($)'])
cumulative = 0
for i in range(1, 9):
    q_amt = 3100 * 3
    cumulative += q_amt
    ws4.append([f'Quarter {i} (of 8 to retirement)', 3100, q_amt, cumulative])
ws4.append(['Total (24 months)', '', 3100 * 24, cumulative])
style_header(ws4, 4)
for row in ws4.iter_rows(min_row=2, max_row=ws4.max_row, min_col=2, max_col=4):
    for cell in row:
        cell.number_format = '#,##0'
autosize(ws4, [28, 26, 26, 24])

# ---- Discretionary Spending Detail (breakdown of two existing retirement lines) ----
ws5 = wb.create_sheet('Discretionary Spending Detail')
ws5.append(['Retirement Line Item Detail', 'Monthly Amount ($)'])
detail_rows = [
    ('Travel and leisure -- Domestic trips', 650),
    ('Travel and leisure -- International trip fund', 400),
    ('Travel and leisure -- Local entertainment/dining', 250),
    ('Travel and leisure -- Subtotal', 1300),
    ('Miscellaneous/discretionary -- Gifts and family support', 300),
    ('Miscellaneous/discretionary -- Hobbies and recreation', 400),
    ('Miscellaneous/discretionary -- Home projects', 500),
    ('Miscellaneous/discretionary -- Personal care and subscriptions', 470),
    ('Miscellaneous/discretionary -- Subtotal', 1670),
]
for r in detail_rows:
    ws5.append(r)
style_header(ws5, 2)
for row in ws5.iter_rows(min_row=2, max_row=ws5.max_row, min_col=2, max_col=2):
    for cell in row:
        cell.number_format = '#,##0'
autosize(ws5, [52, 20])

wb.save(path)
print("expanded Household_Budget_and_Planning_Assumptions.xlsx")

# ============================================================
# 5. Golden: Pruitt_Retirement_Readiness_Assessment.docx
# ============================================================
path = os.path.join(GOLDEN, "Pruitt_Retirement_Readiness_Assessment.docx")
d = docx.Document(path)

# 5a. header age fix
para = find_para(d, "Daniel Pruitt (age 60) and Renee Pruitt (age 58)")
set_para_text(para, "Daniel Pruitt (age 60) and Renee Pruitt (age 60)  |  Target Retirement: Daniel age 62, approximately 2 years from today")

# 5b. Section 4 intro -- clarify same-age milestone timing
para = find_para(d, "Daniel and Renee have indicated a preference to claim Social Security at 62.")
set_para_text(
    para,
    "Daniel and Renee have indicated a preference to claim Social Security at 62. Daniel and Renee are the "
    "same age and reach each Social Security milestone age (62, 65, 67) within the same calendar year, so the "
    "combined household benefit at a given milestone is available in full from the start of that scenario in "
    "both cases below. Per the Social Security Benefit Estimates, claiming at 62 produces a combined $42,600 "
    "annually ($2,400 and $1,150 monthly, respectively). Delaying to age 65 raises the combined benefit to "
    "$54,600 annually ($3,100 and $1,450 monthly)."
)

# 5c. Scenario A -- inline unrounded figure for self-consistent arithmetic
para = find_para(d, "Subtracting the $42,600 of Social Security")
set_para_text(
    para,
    "Subtracting the $42,600 of Social Security from the $129,811.80 first-year need (before rounding to "
    "$129,812) leaves a portfolio withdrawal requirement of approximately $87,212 in year one (precisely "
    "$87,211.80). Applying the firm's 4% initial withdrawal rate sustainability screen, sustaining that "
    "withdrawal requires a portfolio of approximately $2,180,295 ($87,211.80 / 0.04) at retirement. Against "
    "the $1,585,642 projected portfolio, this scenario is short by approximately $594,653, a substantial gap "
    "under a standard 30-year sustainability test."
)

# 5d. Scenario B -- inline unrounded figure + same-age Medicare clarification
para = find_para(d, "From age 65 onward, Medicare replaces the private health bridge")
set_para_text(
    para,
    "From age 65 onward, Medicare replaces the private health bridge for both spouses (Daniel and Renee reach "
    "65 in the same year) and the $54,600 age-65 Social Security benefit begins. The base need, inflated "
    "through that fifth year, is approximately $118,246 (precisely $118,245.96), leaving an ongoing withdrawal "
    "requirement of approximately $63,646 per year (precisely $63,645.96). At the 4% rule, sustaining that "
    "requires a portfolio of approximately $1,591,149 ($63,645.96 / 0.04). Against the $1,186,370 available at "
    "65, this scenario is short by approximately $404,779."
)

# 5e. RMD age fix (73 -> 75, SECURE 2.0)
para = find_para(d, "Required Minimum Distributions beginning at age 73")
set_para_text(
    para,
    "For withdrawal sequencing once retired, recommend drawing from the joint taxable brokerage first (gains "
    "taxed at capital gains rates, and only on the $65,000 of built-in gain), then the 401(k) and 403(b) "
    "(taxed as ordinary income, and subject to Required Minimum Distributions beginning at age 75, the "
    "applicable RMD age under SECURE 2.0 for individuals born in 1960 or later), and preserving the Roth IRA "
    "for last, since it grows tax-free, carries no RMD requirement, and is the most tax-efficient asset to "
    "leave toward the household's stated legacy goal for their children."
)

# 5f. Estate documents bullet -- un-conflate the dates
para = find_para(d, "Update the estate documents.")
set_para_text(
    para,
    "Update the estate documents. The wills were executed in 2011 and have not been reviewed since; "
    "beneficiary designations on the retirement accounts were last confirmed in 2019; and no powers of "
    "attorney or healthcare directives are on file at all. Recommend a full review of all four items given "
    "the approaching retirement."
)

# 5g. Add the missing Renee life-insurance flag as a new bullet at the end
d.add_paragraph(
    "Separately, Renee's $72,000 employer group term life policy ends at retirement/separation from her "
    "employer, leaving her with no life insurance coverage at all from day one of retirement. Evaluate "
    "whether any individual coverage is warranted for Renee given the household's continued reliance on the "
    "portfolio during the drawdown years, particularly under Scenario A."
)

d.save(path)
print("fixed Pruitt_Retirement_Readiness_Assessment.docx")

print("ALL FIXES APPLIED")
