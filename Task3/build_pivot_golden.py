# -*- coding: utf-8 -*-
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os

OUT_DIR = os.path.join(os.path.dirname(__file__), "PIVOT_golden")
os.makedirs(OUT_DIR, exist_ok=True)
OUT = os.path.join(OUT_DIR, "Loan_Portfolio_Review_Report_Q4_2025.docx")

d = docx.Document()

# base style
style = d.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(11)

def h1(text):
    d.add_heading(text, level=1)

def h2(text):
    d.add_heading(text, level=2)

def p(text):
    d.add_paragraph(text)

def table(headers, rows, widths=None):
    t = d.add_table(rows=1, cols=len(headers))
    t.style = 'Light Grid Accent 1'
    hdr = t.rows[0].cells
    for i, htext in enumerate(headers):
        hdr[i].text = htext
        for para in hdr[i].paragraphs:
            for run in para.runs:
                run.bold = True
    for row in rows:
        cells = t.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = str(val)
    d.add_paragraph()
    return t

# Title
title = d.add_heading('Tri-County Commercial Loan Portfolio Review', level=0)
sub = d.add_paragraph('Q4 2025  |  Prepared by Credit Risk Division')
sub2 = d.add_paragraph('Review Date: January 12, 2026, based on portfolio status as of the December 31, 2025 quarter-end close.')

# 1. Executive Summary
h1('1. Executive Summary')
p(
    "This report presents the Q4 2025 quarterly review of the Tri-County commercial loan portfolio, conducted "
    "under the Loan Portfolio Review Policy. The portfolio comprises 29 loans totaling $134,656,509 in outstanding "
    "balance. Independent recomputation of DSCR, LTV, and Current Ratio from the underlying financial statements "
    "and appraisal data, rather than reliance on the Covenant Compliance Workbook's pre-filled figures, surfaced "
    "10 individual covenant breaches across 7 loans, including two loans, LN-002 and LN-019, whose breaches do not "
    "appear at all in the workbook's Actual LTV column and would have been missed under a face-value review."
)
p(
    "Three loans carry an Internal Risk Rating of 4 or higher and sit on the Watch List: LN-001, LN-002, and "
    "LN-008, together $12,119,961 (9.00% of the portfolio). One loan, LN-008, meets the policy's non-performing "
    "definition at 97 days past due and in Chapter 11. All three Board-approved concentration limits, single "
    "borrower, single loan type, and single loan officer, are currently breached. Overall portfolio quality is "
    "adequate but requires close management of the seven breaching credits and near-term attention to concentration "
    "exposure."
)

# 2. Covenant Compliance Summary
h1('2. Covenant Compliance Summary')
p(
    "Per Policy Section 3, all 29 loans were tested against their loan-specific DSCR, LTV, and Current Ratio "
    "thresholds. Rather than accepting the Actual DSCR, Actual LTV, and Actual Current Ratio columns already "
    "filled in on Covenant_Compliance_Workbook_Q4_2025.xlsx, each ratio was independently rebuilt: DSCR as Net "
    "Operating Income over Total Annual Debt Service from Borrower_Financial_Statements_Summary.xlsx, LTV as "
    "Outstanding Balance from Loan_Register_Q4_2025.xlsx over the appraised value in "
    "Appraisal_and_Collateral_Summary_Q4_2025.xlsx, and Current Ratio as Current Assets over Current Liabilities "
    "from the financial statements file."
)
h2('DSCR Breaches')
p("Three loans fall below their loan-specific DSCR minimum:")
table(
    ["Loan", "Borrower", "Recomputed DSCR", "Minimum"],
    [
        ["LN-001", "Talbot Steel Fabricators LLC", "1.07x", "1.20x"],
        ["LN-004", "Coastal Auto Group", "1.14x", "1.20x"],
        ["LN-008", "Palmetto Print & Packaging Inc.", "0.81x", "1.20x"],
    ],
)
h2('LTV Breaches')
p(
    "Four loans exceed their loan-specific LTV maximum. Two of these, LN-002 and LN-019, are not visible in the "
    "workbook at all: the workbook's Actual LTV for LN-002 reads 71.5% (calculated against the $9,600,000 "
    "origination-date value rather than the current $8,200,000 appraisal) and for LN-019 reads 72.8% (calculated "
    "against the prior $5,800,000 collateral figure rather than the $4,800,000 December 2025 floor-plan "
    "revaluation). Both figures are stale, and both loans were confirmed breaching only by recomputing LTV "
    "directly from the Loan Register balance and the current appraised value."
)
table(
    ["Loan", "Borrower", "Recomputed LTV", "Maximum", "Workbook Actual LTV"],
    [
        ["LN-002", "Corrigan Retail Partners", "88.0%", "75%", "71.5% (stale)"],
        ["LN-008", "Palmetto Print & Packaging Inc.", "131.0%", "80%", "130.9%"],
        ["LN-010", "U-Save Self Storage LLC", "85.2%", "75%", "85.2%"],
        ["LN-019", "Interstate Automotive LLC", "88.0%", "80%", "72.8% (stale)"],
    ],
)
h2('Current Ratio Breaches')
p("Three loans fall below their loan-specific Current Ratio minimum:")
table(
    ["Loan", "Borrower", "Recomputed Current Ratio", "Minimum"],
    [
        ["LN-002", "Corrigan Retail Partners", "0.94x", "1.00x"],
        ["LN-006", "Harbor Grille Restaurant Group", "1.03x", "1.20x"],
        ["LN-008", "Palmetto Print & Packaging Inc.", "0.72x", "1.10x"],
    ],
)
h2('Concurrent Violations')
p(
    "LN-008 (Palmetto Print & Packaging Inc.) carries concurrent DSCR, LTV, and Current Ratio breaches, the only "
    "loan in the portfolio in default on all three covenants simultaneously. LN-002 (Corrigan Retail Partners) "
    "carries two concurrent breaches (LTV and Current Ratio), both of which the Covenant Compliance Workbook's "
    "pre-filled figures fail to show. In total, 10 individual covenant violations span 7 distinct loans: LN-001, "
    "LN-002, LN-004, LN-006, LN-008, LN-010, and LN-019."
)

# 3. Watch List Analysis
h1('3. Watch List Analysis')
p(
    "Per Policy Section 4, any loan carrying an Internal Risk Rating of 4 or higher is automatically placed on "
    "the Watch List. Three loans qualify:"
)
table(
    ["Loan", "Borrower", "Risk Rating", "Balance", "Officer", "Status per Exception Log"],
    [
        ["LN-001", "Talbot Steel Fabricators LLC", "4", "$3,817,432", "Diana Flores",
         "Waiver approved 11/25/2025 (EXC-001), expires 3/31/2026; remediation plan on file"],
        ["LN-002", "Corrigan Retail Partners", "5", "$7,214,889", "Marcus Webb",
         "No waiver; open LTV and Current Ratio exceptions (EXC-002, EXC-003); escalated to Watch List; formal demand issued (EXC-004)"],
        ["LN-008", "Palmetto Print & Packaging Inc.", "6", "$1,087,640", "Marcus Webb",
         "Legal hold; Chapter 11 (EXC-008, EXC-009, EXC-010)"],
    ],
)
p(
    "Combined Watch List exposure is $12,119,961, or 9.00% of the portfolio. No other loan in the Loan Register "
    "carries a risk rating of 4 or above. LN-001 is the only one of the three with an active, unexpired waiver "
    "offsetting its breach; LN-002 and LN-008 have no offsetting waiver on file."
)

# 4. Concentration Risk Analysis
h1('4. Concentration Risk Analysis')
p(
    "Policy Section 6 sets three Board-approved concentration limits: no single borrower over 15% of total "
    "portfolio outstanding, no single loan type over 40%, and no single loan officer's book over 40%. All three "
    "are currently breached."
)
h2('Single Borrower Concentration')
p(
    "The largest single relationship is LN-028 (Meridian Business Park LLC) at $21,117,450, or 15.68% of the "
    "$134,656,509 portfolio, 0.68 points over the 15% limit. No other single borrower exceeds the limit; the "
    "next-largest, LN-022 (DataVault Systems LLC) at $17,224,110, is 12.79% of the portfolio."
)
h2('Loan Type Concentration')
table(
    ["Loan Type", "Loan Count", "Balance", "% of Portfolio", "vs. 40% Limit"],
    [
        ["Commercial Mortgage", "9", "$88,481,044", "65.71%", "Over by 25.71 pts"],
        ["Term Loan", "9", "$20,281,745", "15.06%", "Within limit"],
        ["Construction Loan", "2", "$10,550,680", "7.84%", "Within limit"],
        ["Floor Plan Loan", "2", "$7,194,920", "5.34%", "Within limit"],
        ["Equipment Loan", "3", "$3,051,850", "2.27%", "Within limit"],
        ["Line of Credit", "3", "$2,902,460", "2.16%", "Within limit"],
        ["Agricultural Term Loan", "1", "$2,193,810", "1.63%", "Within limit"],
    ],
)
h2('Loan Officer Concentration')
table(
    ["Officer", "Loan Count", "Balance", "% of Portfolio", "vs. 40% Limit"],
    [
        ["James Trotter", "9", "$82,101,025", "60.97%", "Over by 20.97 pts"],
        ["Diana Flores", "11", "$26,929,522", "20.00%", "Within limit"],
        ["Marcus Webb", "9", "$25,625,962", "19.03%", "Within limit"],
    ],
)
p(
    "The Commercial Mortgage concentration is the most severe breach in dollar and percentage terms, followed by "
    "James Trotter's loan-officer book. Both warrant presentation to the Senior Loan Committee per Policy "
    "Section 6."
)

# 5. Exception Tracking
h1('5. Exception Tracking')
p(
    "Eleven exceptions are currently open, transcribed from Exception_Tracking_Log_Q4_2025.xlsx:"
)
table(
    ["Exception", "Loan", "Type", "Officer", "Status"],
    [
        ["EXC-001", "LN-001", "DSCR Covenant Breach", "Diana Flores", "Open, Monitoring; waiver approved 11/25/2025, expires 3/31/2026"],
        ["EXC-002", "LN-002", "LTV Covenant Breach", "Marcus Webb", "Open, Escalated to Watch List"],
        ["EXC-003", "LN-002", "Current Ratio Breach", "Marcus Webb", "Open, Escalated to Watch List"],
        ["EXC-004", "LN-002", "Payment Delinquency", "Marcus Webb", "Open, Formal demand issued"],
        ["EXC-005", "LN-004", "DSCR Covenant Breach", "James Trotter", "Open, Monitoring; remediation plan submitted"],
        ["EXC-006", "LN-004", "Payment Delinquency", "James Trotter", "Open, Awaiting payment confirmation"],
        ["EXC-007", "LN-006", "Current Ratio Breach", "Marcus Webb", "Open, Follow-up scheduled 3/1/2026"],
        ["EXC-008", "LN-008", "DSCR Covenant Breach", "Marcus Webb", "Open, Legal Hold (Chapter 11)"],
        ["EXC-009", "LN-008", "LTV Covenant Breach", "Marcus Webb", "Open, Legal Hold (Chapter 11)"],
        ["EXC-010", "LN-008", "Non-Performing / 90+ DPD", "Marcus Webb", "Open, Loss Reserve Active"],
        ["EXC-011", "LN-010", "LTV Covenant Breach", "James Trotter", "Open, Appraisal Dispute"],
    ],
)
p(
    "Every breach identified in Section 2 above has a corresponding exception on file except one: LN-019's LTV "
    "breach has no exception record anywhere in the tracking log. That gap was only caught because DSCR, LTV, and "
    "Current Ratio were rebuilt from source data as this review requires rather than cross-checked solely against "
    "existing exceptions."
)

# 6. Non-Performing Loan Schedule
h1('6. Non-Performing Loan Schedule')
p(
    "One loan meets the policy's non-performing definition of 90 or more days past due:"
)
table(
    ["Loan", "Borrower", "Balance", "Days Past Due", "Collateral Value", "LTV", "Status"],
    [
        ["LN-008", "Palmetto Print & Packaging Inc.", "$1,087,640", "97", "$830,000", "131.0%", "Chapter 11, legal hold"],
    ],
)
p(
    "LN-008's manufacturing equipment collateral, appraised at $830,000 as of the December 11, 2025 inspection, "
    "covers approximately 76.3% of the outstanding balance. Two other loans carry delinquency exceptions but do "
    "not meet the non-performing threshold: LN-002 is 45 days past due (Watch List, not non-performing) and LN-004 "
    "is 15 days past due (delinquency exception only)."
)

# 7. Final Portfolio Risk Assessment
h1('7. Final Portfolio Risk Assessment')
p(
    "The Q4 2025 portfolio carries a moderately elevated risk profile driven by four interconnected concerns."
)
p(
    "First, covenant breaches are broader than a workbook-only review would show: 10 individual violations across "
    "7 loans, two of which, LN-002 and LN-019, only came to light because DSCR, LTV, and Current Ratio were "
    "rebuilt from source data rather than taken from the Covenant Compliance Workbook. LN-019 in particular has no "
    "exception on file at all, which is a gap in the exception-logging process worth flagging to the Credit Risk "
    "Division independent of the loan itself."
)
p(
    "Second, non-performing exposure is concentrated in a single credit: LN-008 is 97 days past due, in Chapter 11, "
    "and its collateral covers only about 76% of the outstanding balance."
)
p(
    "Third, Watch List exposure totals $12,119,961 across three loans; only LN-001 currently has an offsetting "
    "waiver, leaving LN-002 and LN-008 unaddressed by any formal accommodation."
)
p(
    "Fourth, all three Board-approved concentration limits are currently breached, most severely the 65.71% "
    "Commercial Mortgage concentration (versus a 40% limit) and James Trotter's 60.97% officer book (versus a 40% "
    "limit). The single-borrower breach on LN-028 is comparatively minor at 0.68 points over the limit."
)
p(
    "Recommended actions: present the loan-type and loan-officer concentration breaches to the Senior Loan "
    "Committee per Policy Section 6; obtain a current appraisal and covenant reassessment for LN-002 and require "
    "an exception record be opened for LN-019 given its recomputed breach has none; continue legal coordination "
    "on LN-008's bankruptcy proceeding. The remaining 22 of 29 loans are performing within covenant and payment "
    "terms and require no immediate action beyond standard quarterly monitoring."
)

d.save(OUT)
print("saved", OUT)
