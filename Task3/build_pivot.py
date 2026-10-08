import sys
sys.stdout.reconfigure(encoding='utf-8')
from docx import Document
from docx.shared import Pt
import openpyxl
import os

BASE = r'c:\Users\user\OneDrive\Desktop\geranium project\Task3'
NEWDIR = os.path.join(BASE, 'pivot_inputs')
os.makedirs(NEWDIR, exist_ok=True)

# ---------------------------------------------------------------------------
# 1. ALLL_CECL_Reserve_Policy.docx
# ---------------------------------------------------------------------------
d = Document()
d.add_heading('COMMERCIAL LOAN ALLOWANCE FOR LOAN AND LEASE LOSSES (ALLL) POLICY', level=0)
d.add_paragraph('Tri-County Community Bank | Credit Risk Division')
d.add_paragraph('Policy Effective Date: January 1, 2025 | Review Cycle: Quarterly (CECL)')

d.add_heading('1. Purpose and Scope', level=1)
d.add_paragraph(
    "This policy establishes the Bank's methodology for estimating and testing the adequacy of the "
    "Allowance for Loan and Lease Losses (ALLL) under the Current Expected Credit Loss (CECL) standard "
    "for the commercial loan portfolio. The methodology applies to all commercial loan relationships "
    "exceeding $250,000 in outstanding balance. The Credit Risk Division must produce a quarterly ALLL "
    "Reserve Adequacy Memorandum testing the sufficiency of the booked reserve against the reserve "
    "required under this methodology."
)

d.add_heading('2. Definitions', level=1)
p = d.add_paragraph()
p.add_run('Allowance for Loan and Lease Losses (ALLL): ').bold = True
p.add_run('A valuation reserve, carried on the Bank\'s books, established to absorb estimated credit losses inherent in the loan portfolio.')
p = d.add_paragraph()
p.add_run('Pooled Reserve: ').bold = True
p.add_run('A reserve computed by applying a historical loss rate to the aggregate outstanding balance of all loans sharing a given Internal Risk Rating, excluding loans that are individually evaluated under Section 4.')
p = d.add_paragraph()
p.add_run('Individually Evaluated (Collateral-Dependent) Loan: ').bold = True
p.add_run('A loan that must be removed from pooled treatment and evaluated on a standalone basis because repayment is expected to be provided substantially through the operation or sale of collateral.')
p = d.add_paragraph()
p.add_run('Qualitative Factor (Q-Factor) Adjustment: ').bold = True
p.add_run('A basis-point add-on to a pool\'s historical loss rate, applied to capture emerging risk in that pool that is not yet reflected in historical loss experience.')
p = d.add_paragraph()
p.add_run('Reserve Shortfall: ').bold = True
p.add_run('The amount by which the reserve required under this methodology exceeds the currently booked ALLL balance. A shortfall must be recognized as additional provision expense in the current quarter.')

d.add_heading('3. Historical Loss Rate Pools', level=1)
d.add_paragraph(
    'All commercial loans not individually evaluated under Section 4 must be pooled strictly by Internal '
    'Risk Rating (per the 6-point scale defined in the Loan Portfolio Review Policy) and assigned the '
    'following historical loss rate, applied to the aggregate outstanding balance of the pool. Historical '
    'loss rates are sourced from the ALLL_Loss_Rate_and_Reserve_Data.xlsx workbook maintained by the '
    'Finance Department and must not be altered or estimated independently.'
)
table = d.add_table(rows=1, cols=2)
table.style = 'Light Grid Accent 1'
hdr = table.rows[0].cells
hdr[0].text = 'Internal Risk Rating'
hdr[1].text = 'Treatment'
rows = [
    ('1 — Minimal Risk', 'Pooled at the Rating 1 historical loss rate.'),
    ('2 — Acceptable Risk', 'Pooled at the Rating 2 historical loss rate.'),
    ('3 — Moderate Risk', 'Pooled at the Rating 3 historical loss rate.'),
    ('4 — Watch', 'Pooled at the Rating 4 historical loss rate.'),
    ('5 — Special Mention', 'Pooled at the Rating 5 historical loss rate.'),
    ('6 — Substandard / Loss', 'Not pooled. Individually evaluated per Section 4.'),
]
for a,b in rows:
    r = table.add_row().cells
    r[0].text = a
    r[1].text = b

d.add_heading('4. Individually Evaluated (Collateral-Dependent) Loans', level=1)
d.add_paragraph(
    'Any loan carrying an Internal Risk Rating of 6, or classified as non-performing (90 or more days past '
    'due), must be removed from pooled treatment and specifically evaluated for a standalone reserve. The '
    'specific reserve is calculated as:'
)
p = d.add_paragraph()
p.add_run('Specific Reserve = MAX(0, Outstanding Balance \u2212 (Appraised Collateral Value \u00d7 (1 \u2212 Estimated Disposition Cost %)))').italic = True
d.add_paragraph(
    'The Estimated Disposition Cost % is sourced from ALLL_Loss_Rate_and_Reserve_Data.xlsx and reflects '
    'the estimated cost of liquidating the collateral (auction fees, legal fees, holding costs). The most '
    'recent appraised value on file must be used regardless of loan origination date.'
)

d.add_heading('5. Qualitative Factor (Q-Factor) Adjustments', level=1)
d.add_paragraph(
    'Any pool (Ratings 1 through 5) that contains one or more loans carrying an active covenant breach as '
    'of the review date \u2014 a breach not yet cured or offset by a documented, approved waiver \u2014 must receive '
    'a Q-Factor add-on of 25 basis points (0.25%) to that pool\'s historical loss rate, applied to the full '
    'outstanding balance of the pool. This adjustment reflects emerging credit deterioration in the pool '
    'that has not yet worked its way into historical loss experience. The add-on applies at the pool level, '
    'not to the individual breaching loan, and stacks with the base historical loss rate for that pool. '
    'Covenant breach status must be independently verified against the Covenant Compliance Workbook, '
    'Borrower Financial Statements, and Appraisal and Collateral Summary before a Q-Factor is applied \u2014 '
    'pre-calculated breach flags must not be accepted without recomputation.'
)

d.add_heading('6. Reserve Adequacy Determination', level=1)
d.add_paragraph(
    'Total Required Reserve is the sum of all pooled reserves (Ratings 1 through 5, inclusive of any '
    'applicable Q-Factor add-ons) plus all specific reserves for individually evaluated loans. This total '
    'must be compared against the currently booked ALLL balance as recorded in '
    'ALLL_Loss_Rate_and_Reserve_Data.xlsx:'
)
d.add_paragraph('\u2022 If Total Required Reserve exceeds the booked balance, the shortfall must be recognized as additional provision expense in the current quarter and reported to the Board Audit Committee.')
d.add_paragraph('\u2022 If the booked balance exceeds Total Required Reserve by more than 10%, the excess must be evaluated for release, subject to Board approval.')
d.add_paragraph('\u2022 If the booked balance is within 10% of Total Required Reserve in either direction, the reserve is considered adequate and no provision action is required.')

d.add_heading('7. Quarterly ALLL Reserve Adequacy Memorandum Requirements', level=1)
d.add_paragraph('At the end of each quarter, the Credit Risk Division must produce an ALLL Reserve Adequacy Memorandum. The memorandum must include:')
d.add_paragraph('\u2022 Executive Summary with the overall reserve adequacy conclusion and the dollar shortfall or excess.')
d.add_paragraph('\u2022 Portfolio Segmentation showing loan count and aggregate balance by Internal Risk Rating pool.')
d.add_paragraph('\u2022 Pooled Reserve Calculation showing, for each pool, the base historical loss rate, any Q-Factor add-on and its trigger, the adjusted rate, and the resulting pooled reserve.')
d.add_paragraph('\u2022 Individually Evaluated Loan Analysis showing, for each collateral-dependent loan, the outstanding balance, appraised collateral value, disposition cost assumption, and resulting specific reserve.')
d.add_paragraph('\u2022 Covenant Breach and Watch List Cross-Reference documenting which loans and pools triggered a Q-Factor and why, cross-checked against the underlying financial and appraisal data.')
d.add_paragraph('\u2022 Reserve Adequacy Conclusion comparing Total Required Reserve to the booked balance and stating the recommended provision action.')
d.add_paragraph('\u2022 Policy/Source Reconciliation documenting how pre-calculated fields were verified against underlying source data and how any conflicts were resolved.')

d.add_heading('8. Approval and Distribution', level=1)
d.add_paragraph(
    'The completed ALLL Reserve Adequacy Memorandum must be reviewed and approved by the Chief Credit '
    'Officer prior to submission to the Board Audit Committee. Distribution: Chief Credit Officer, Chief '
    'Financial Officer, Board Audit Committee, and the Bank\'s external auditor.'
)

policy_path = os.path.join(NEWDIR, 'ALLL_CECL_Reserve_Policy.docx')
d.save(policy_path)
print('Saved', policy_path)

# ---------------------------------------------------------------------------
# 2. ALLL_Loss_Rate_and_Reserve_Data.xlsx
# ---------------------------------------------------------------------------
wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'Loss Rates'
ws.append(['Internal Risk Rating', 'Historical Loss Rate (%)', 'Notes'])
data = [
    (1, 0.10, 'Minimal Risk pool'),
    (2, 0.35, 'Acceptable Risk pool'),
    (3, 0.90, 'Moderate Risk pool'),
    (4, 3.25, 'Watch pool'),
    (5, 9.50, 'Special Mention pool'),
    (6, None, 'Not pooled - individually evaluated per Policy Section 4'),
]
for row in data:
    ws.append(row)

ws2 = wb.create_sheet('Reserve Assumptions')
ws2.append(['Assumption', 'Value', 'Source'])
ws2.append(['Estimated Disposition Cost %', 0.10, 'Finance Department - collateral liquidation cost study, FY2025'])
ws2.append(['Q-Factor Add-On for Breaching Pools (bps)', 0.25, 'Policy Section 5'])
ws2.append(['Currently Booked ALLL Balance ($)', 1250000, 'General Ledger account 1720, balance as of 12/31/2025'])
ws2.append(['Reserve Adequacy Tolerance Band', 0.10, 'Policy Section 6 - +/-10% of required reserve'])

data_path = os.path.join(NEWDIR, 'ALLL_Loss_Rate_and_Reserve_Data.xlsx')
wb.save(data_path)
print('Saved', data_path)
