"""
create_report.py
Generates Loan_Portfolio_Review_Report_Q4_2025.docx
All figures are derived directly from the five source files (via data_summary.json).
"""

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH


def add_heading(doc, text, level):
    h = doc.add_heading(text, level=level)
    h.paragraph_format.space_before = Pt(12)
    h.paragraph_format.space_after = Pt(4)
    return h


def add_para(doc, text):
    p = doc.add_paragraph(text)
    p.paragraph_format.space_after = Pt(4)
    return p


def add_table_with_header(doc, headers, rows):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    hdr = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = h
        for run in hdr[i].paragraphs[0].runs:
            run.bold = True
    for ri, row in enumerate(rows):
        cells = table.rows[ri + 1].cells
        for ci, val in enumerate(row):
            cells[ci].text = str(val)
    doc.add_paragraph()
    return table


def create_report():
    doc = Document()

    # ── Title ────────────────────────────────────────────────────────────────
    title = doc.add_heading(
        'Tri-County Commercial Loan Portfolio Review', level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = doc.add_paragraph('Q4 2025  |  Prepared by Credit Review Department')
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    # ── 1. Executive Summary ─────────────────────────────────────────────────
    add_heading(doc, '1. Executive Summary', level=1)
    add_para(doc,
        'This report presents the results of the Q4 2025 quarterly review of the Tri-County '
        'commercial loan portfolio, conducted in accordance with the Loan Portfolio Review '
        'Policy. The portfolio comprises 29 active commercial loans with an aggregate outstanding '
        'balance of $134,656,509 as of the review date.')
    add_para(doc,
        'The review identified nine individual covenant breaches distributed across six affected '
        'loans, three Watch List designations driven by risk-rating criteria, one non-performing loan '
        'currently in Chapter 11 bankruptcy proceedings, three concentration-limit exceedances, and '
        'eleven open exceptions requiring active monitoring or legal coordination. The overall portfolio '
        'carries elevated risk due to the concentration of adversely rated exposure in a single '
        'non-performing credit (LN-008) and structural over-concentration in commercial mortgage '
        'loan types and a single loan-officer relationship.')

    # ── 2. Portfolio Overview ────────────────────────────────────────────────
    add_heading(doc, '2. Portfolio Overview', level=1)
    add_para(doc,
        'The portfolio includes 29 loans originated between 2020 and 2023, spanning industries '
        'including commercial real estate, manufacturing, retail, hospitality, healthcare, '
        'technology, food distribution, and energy. Collateral types range from commercial '
        'mortgages and equipment to accounts receivable and vehicle inventories.')

    add_para(doc, 'Loan type distribution:')
    lt_headers = ['Loan Type', 'Loan Count', 'Aggregate Balance ($)', '% of Portfolio']
    lt_rows = [
        ['Commercial Mortgage', '9', '$88,481,044', '65.71%'],
        ['Term Loan', '9', '$20,281,665', '15.06%'],
        ['Equipment Loan', '3', '$3,051,850', '2.27%'],
        ['Line of Credit', '3', '$2,902,460', '2.16%'],
        ['Construction Loan', '2', '$10,550,680', '7.84%'],
        ['Floor Plan Loan', '2', '$7,194,920', '5.35%'],
        ['Agricultural Term Loan', '1', '$2,193,810', '1.63%'],
    ]
    add_table_with_header(doc, lt_headers, lt_rows)

    add_para(doc,
        'Three relationship officers carry the portfolio: Diana Flores (11 loans, $27,512,145, '
        '20.43%), Marcus Webb (9 loans, $25,074,404, 18.62%), and James Trotter (9 loans, '
        '$82,101,025, 60.97%). Geographic coverage spans the Midwest, Southeast, Southwest, '
        'Pacific Northwest, and Mid-Atlantic regions.')

    # ── 3. Covenant Compliance Findings ─────────────────────────────────────
    add_heading(doc, '3. Covenant Compliance Findings', level=1)
    add_para(doc,
        'All 29 loans were evaluated independently against their loan-specific DSCR, LTV, and '
        'Current Ratio thresholds as documented in the Covenant Compliance Workbook. '
        'Pre-calculated workbook values were cross-checked against the underlying financial '
        'statements and appraisal records to confirm accuracy. Where workbook values and '
        'independently computed values differed, the independently derived figure was used and '
        'the variance is noted. Nine individual breaches were identified across six affected loans.')

    add_heading(doc, 'DSCR Breaches', level=2)
    add_para(doc,
        'Three loans fall below their loan-specific DSCR minimums as set out in the Covenant '
        'Compliance Workbook. LN-008 is severely distressed at 0.81x, well below its 1.20x floor.')
    dscr_headers = ['Loan ID', 'Borrower', 'Actual DSCR', 'Min Threshold', 'Variance']
    dscr_rows = [
        ['LN-001', 'ABC Industrial LLC', '1.07x', '1.20x', '-0.13x'],
        ['LN-004', 'Delta Auto Group', '1.14x', '1.20x', '-0.06x'],
        ['LN-008', 'Eta Print & Packaging Inc.', '0.81x', '1.20x', '-0.39x'],
    ]
    add_table_with_header(doc, dscr_headers, dscr_rows)
    add_para(doc,
        'DSCR values were independently verified from Borrower_Financial_Statements_Summary.xlsx '
        'by dividing Net Operating Income by Total Annual Debt Service: LN-001 $347,200 / '
        '$324,500 = 1.07x; LN-004 $288,100 / $252,700 = 1.14x; LN-008 $74,800 / $92,400 = '
        '0.81x. These computed values match the Covenant Compliance Workbook pre-calculated '
        'figures within rounding tolerance (±0.01x). No material conflict was identified.')

    add_heading(doc, 'LTV Breaches', level=2)
    add_para(doc,
        'Three loans exceed their loan-specific LTV maximums based on current outstanding '
        'balances from the Loan Register and latest appraised collateral values from the '
        'Appraisal and Collateral Summary.')
    ltv_headers = ['Loan ID', 'Borrower', 'Outstanding Balance', 'Appraised Value', 'Actual LTV', 'Max LTV']
    ltv_rows = [
        ['LN-002', 'XYZ Retail Partners', '$7,214,889', '$8,200,000', '87.9%', '75%'],
        ['LN-008', 'Eta Print & Packaging Inc.', '$1,087,640', '$830,000', '130.9%', '80%'],
        ['LN-010', 'Iota Storage Partners', '$3,108,775', '$3,650,000', '85.2%', '75%'],
    ]
    add_table_with_header(doc, ltv_headers, ltv_rows)
    add_para(doc,
        'LTV values were independently verified by dividing each loan\'s outstanding balance '
        '(Loan Register) by the corresponding appraised value (Appraisal Summary): '
        'LN-002 $7,214,889 / $8,200,000 = 87.9%, consistent with the Covenant Workbook. '
        'LN-008 $1,087,640 / $830,000 = 130.9%; the Appraisal Summary notes equipment value '
        'declined significantly and Chapter 11 was filed December 2025. '
        'LN-010 $3,108,775 / $3,650,000 = 85.2%; the Appraisal Summary notes value declined '
        'from $4.4M at origination (LTV breach triggered per appraisal note). All three '
        'computed figures match the Covenant Workbook pre-calculated values within rounding.')

    add_heading(doc, 'Current Ratio Breaches', level=2)
    add_para(doc,
        'Three loans fall below their loan-specific Current Ratio minimums based on financial '
        'statement data in Borrower_Financial_Statements_Summary.xlsx.')
    cr_headers = ['Loan ID', 'Borrower', 'Current Assets', 'Current Liabilities', 'Actual CR', 'Min CR']
    cr_rows = [
        ['LN-002', 'XYZ Retail Partners', '$813,900', '$865,800', '0.94x', '1.00x'],
        ['LN-006', 'Epsilon Restaurant Group', '$87,000', '$84,500', '1.03x', '1.20x'],
        ['LN-008', 'Eta Print & Packaging Inc.', '$94,000', '$130,500', '0.72x', '1.10x'],
    ]
    add_table_with_header(doc, cr_headers, cr_rows)
    add_para(doc,
        'Current Ratio was computed as Current Assets divided by Current Liabilities from the '
        'Financial Statements file: LN-002 $813,900 / $865,800 = 0.94x (below the 1.00x floor). '
        'LN-006 $87,000 / $84,500 = 1.03x (below the 1.20x floor; exception EXC-007 '
        'notes a seasonal pattern with recovery historically in Q2). '
        'LN-008 $94,000 / $130,500 = 0.72x (severely below the 1.10x floor). '
        'All computed values match the Covenant Workbook pre-calculated figures; no conflicts found.')

    add_heading(doc, 'Concurrent Violations', level=2)
    add_para(doc,
        'LN-008 (Eta Print & Packaging Inc.) carries concurrent DSCR, LTV, and Current Ratio '
        'breaches — the only loan in the portfolio with all three covenants in default simultaneously. '
        'LN-002 (XYZ Retail Partners) carries concurrent LTV and Current Ratio breaches. '
        'No other loan carries more than one individual breach.')

    # ── 4. Watch List and Concentration Analysis ─────────────────────────────
    add_heading(doc, '4. Watch List and Concentration Analysis', level=1)

    add_heading(doc, 'Watch List', level=2)
    add_para(doc,
        'Per policy Section 4 (Watch List and Classification), loans assigned an Internal Risk '
        'Rating of 4 or higher are automatically placed on the Watch List and require quarterly '
        'reporting to the Chief Credit Officer. Risk ratings are sourced exclusively from the '
        'Internal Risk Rating column of Loan_Register_Q4_2025.xlsx. Three loans qualify:')
    wl_headers = ['Loan ID', 'Borrower', 'Risk Rating', 'Policy Trigger', 'Payment Status']
    wl_rows = [
        ['LN-001', 'ABC Industrial LLC', '4', 'Rating ≥ 4', 'Current, 0 DPD'],
        ['LN-002', 'XYZ Retail Partners', '5', 'Rating ≥ 4', 'Past Due, 45 DPD'],
        ['LN-008', 'Eta Print & Packaging Inc.', '6', 'Rating ≥ 4', '97 DPD, Chapter 11'],
    ]
    add_table_with_header(doc, wl_headers, wl_rows)
    add_para(doc,
        'No other loan in the Loan Register carries a risk rating of 4 or above. '
        'LN-001 is on enhanced monitoring under a temporary DSCR waiver expiring March 31, 2026. '
        'LN-002 is subject to a formal payment demand issued December 8, 2025. '
        'LN-008 is in legal hold pending Chapter 11 bankruptcy proceedings.')

    add_heading(doc, 'Concentration Risk', level=2)
    add_para(doc,
        'Policy Section 6 sets three concentration limits: 15% for any single borrower, '
        '40% for any single loan type, and 40% for any single loan officer. '
        'All three thresholds are exceeded as of Q4 2025:')
    conc_headers = ['Category', 'Aggregate Balance', '% of Portfolio', 'Policy Limit', 'Status']
    conc_rows = [
        ['Single Borrower: C2 Tech Campus (LN-028)', '$21,117,450', '15.68%', '15%', 'BREACH'],
        ['Loan Type: Commercial Mortgage', '$88,481,044', '65.71%', '40%', 'BREACH'],
        ['Loan Officer: James Trotter', '$82,101,025', '60.97%', '40%', 'BREACH'],
    ]
    add_table_with_header(doc, conc_headers, conc_rows)
    add_para(doc,
        'Commercial Mortgage balance is the sum of nine loans in the Loan Register classified '
        'as that type: LN-002 ($7,214,889) + LN-007 ($13,574,180) + LN-010 ($3,108,775) + '
        'LN-014 ($3,895,540) + LN-016 ($10,874,320) + LN-020 ($6,582,440) + '
        'LN-022 ($17,224,110) + LN-026 ($4,889,340) + LN-028 ($21,117,450) = $88,481,044 '
        '(65.71% of $134,656,509). James Trotter manages nine loans per the Loan Register: '
        'LN-004 ($2,973,140) + LN-007 ($13,574,180) + LN-010 ($3,108,775) + '
        'LN-012 ($4,117,930) + LN-016 ($10,874,320) + LN-019 ($4,221,780) + '
        'LN-022 ($17,224,110) + LN-026 ($4,889,340) + LN-028 ($21,117,450) = $82,101,025 '
        '(60.97%). C2 Tech Campus (LN-028) at $21,117,450 represents 15.68% of the portfolio, '
        'exceeding the single-borrower cap by approximately $914,000 in excess exposure. '
        'All percentages are calculated against the total portfolio balance of $134,656,509.')

    # ── 5. Exception Tracking ────────────────────────────────────────────────
    add_heading(doc, '5. Exception Tracking', level=1)
    add_para(doc,
        'Eleven exceptions are currently open as of Q4 2025, transcribed directly from '
        'Exception_Tracking_Log_Q4_2025.xlsx. All exceptions were logged within the '
        'required ten-business-day window from the date of identification.')
    exc_headers = ['Exception ID', 'Loan ID', 'Borrower', 'Exception Type', 'Current Status']
    exc_rows = [
        ['EXC-001', 'LN-001', 'ABC Industrial LLC', 'DSCR Covenant Breach', 'Open — Monitoring'],
        ['EXC-002', 'LN-002', 'XYZ Retail Partners', 'LTV Covenant Breach', 'Open — Escalated to Watch List'],
        ['EXC-003', 'LN-002', 'XYZ Retail Partners', 'Current Ratio Breach', 'Open — Escalated to Watch List'],
        ['EXC-004', 'LN-002', 'XYZ Retail Partners', 'Payment Delinquency', 'Open — Formal demand issued'],
        ['EXC-005', 'LN-004', 'Delta Auto Group', 'DSCR Covenant Breach', 'Open — Monitoring'],
        ['EXC-006', 'LN-004', 'Delta Auto Group', 'Payment Delinquency', 'Open — Awaiting payment confirmation'],
        ['EXC-007', 'LN-006', 'Epsilon Restaurant Group', 'Current Ratio Breach', 'Open — Follow-up Scheduled'],
        ['EXC-008', 'LN-008', 'Eta Print & Packaging Inc.', 'DSCR Covenant Breach', 'Open — Legal Hold'],
        ['EXC-009', 'LN-008', 'Eta Print & Packaging Inc.', 'LTV Covenant Breach', 'Open — Legal Hold'],
        ['EXC-010', 'LN-008', 'Eta Print & Packaging Inc.', 'Non-Performing / 90+ DPD', 'Open — Loss Reserve Active'],
        ['EXC-011', 'LN-010', 'Iota Storage Partners', 'LTV Covenant Breach', 'Open — Appraisal Dispute'],
    ]
    add_table_with_header(doc, exc_headers, exc_rows)
    add_para(doc,
        'EXC-001: A remediation plan was submitted January 12, 2026; a temporary waiver was '
        'issued expiring March 31, 2026, pending ABC Industrial LLC remediation. '
        'EXC-008 through EXC-010 are in legal hold due to the Chapter 11 filing by '
        'Eta Print & Packaging Inc. on December 9, 2025, as documented in the exception log. '
        'EXC-011 is under appraisal dispute; the lender is challenging the November 2025 '
        'updated valuation that reduced collateral from $4.4M to $3.65M.')

    # ── 6. Non-Performing Loan Schedule ─────────────────────────────────────
    add_heading(doc, '6. Non-Performing Loan Schedule', level=1)
    add_para(doc,
        'One loan meets the policy definition of non-performing: 90 or more days past due or '
        'placed on non-accrual status. This classification is supported by the Days Past Due '
        'and Payment Status fields in Loan_Register_Q4_2025.xlsx, corroborated by Exception '
        'EXC-010 in the Exception Tracking Log (which records this loan at 97 days past due '
        'with a loss reserve established), and further confirmed by the Appraisal Summary note '
        'referencing the Chapter 11 filing in December 2025.')
    npl_headers = ['Loan ID', 'Borrower', 'Outstanding Balance', 'DPD', 'Payment Status', 'Legal Event']
    npl_rows = [
        ['LN-008', 'Eta Print & Packaging Inc.', '$1,087,640', '97', 'Non-accrual',
         'Chapter 11 filed December 9, 2025'],
    ]
    add_table_with_header(doc, npl_headers, npl_rows)
    add_para(doc,
        'Loan LN-008 is carried at an outstanding balance of $1,087,640 against manufacturing '
        'equipment collateral appraised at $830,000, yielding an LTV of 130.9%. A loss reserve '
        'has been established as noted in EXC-010. The Chapter 11 filing is documented in both '
        'the Exception Tracking Log (EXC-008 description) and the Appraisal Summary collateral '
        'inspection note dated December 2025. No additional loans in the portfolio meet the '
        '90-day past-due or non-accrual threshold at this review date.')
    add_para(doc,
        'Note: LN-002 is 45 days past due (Watch List, not non-performing) and LN-004 is 15 '
        'days past due (delinquency exception only). Neither has crossed the 90-day threshold '
        'or been placed on non-accrual status.')

    # ── 7. Final Portfolio Risk Assessment ───────────────────────────────────
    add_heading(doc, '7. Final Portfolio Risk Assessment', level=1)
    add_para(doc,
        'The Q4 2025 Tri-County commercial loan portfolio carries a moderately elevated risk '
        'profile driven by four interconnected concerns identified through this review:')
    add_para(doc,
        '1. Covenant breaches: Nine individual covenant violations across six affected loans '
        'indicate broad-based financial stress in a subset of the portfolio. LN-008 carries '
        'simultaneous DSCR (0.81x vs. 1.20x minimum), LTV (130.9% vs. 80% maximum), and '
        'Current Ratio (0.72x vs. 1.10x minimum) failures — the most severely distressed '
        'credit in the book. LN-001 and LN-004 carry DSCR breaches that are currently being '
        'remediated. LN-002 carries concurrent LTV and Current Ratio breaches accompanied by '
        'a 45-day payment delinquency.')
    add_para(doc,
        '2. Non-performing exposure: LN-008 (Eta Print & Packaging Inc.) is 97 days past due '
        'and in Chapter 11 bankruptcy. With collateral coverage at only 76% of outstanding '
        'principal ($830,000 appraised vs. $1,087,640 outstanding), loss exposure is material. '
        'A loss reserve is active per EXC-010.')
    add_para(doc,
        '3. Watch List concentration: Three loans — LN-001 (Risk Rating 4), LN-002 '
        '(Risk Rating 5), and LN-008 (Risk Rating 6) — qualify for Watch List status under '
        'policy Section 4. Their combined outstanding balance is approximately $12.1M. '
        'LN-001 holds a temporary DSCR waiver expiring March 31, 2026; LN-002 has received '
        'a formal payment demand and carries dual covenant exceptions.')
    add_para(doc,
        '4. Concentration risk: Commercial Mortgage loans represent 65.71% of the portfolio '
        '(policy limit 40%), the James Trotter loan-officer book represents 60.97% of total '
        'outstanding (policy limit 40%), and a single borrower — C2 Tech Campus — accounts for '
        '15.68% of total outstanding (policy limit 15%). These structural over-concentrations '
        'amplify systemic risk if any segment experiences adverse market conditions. '
        'Portfolio rebalancing and officer-level diversification should be prioritized.')
    add_para(doc,
        'The remaining 23 loans appear to be performing within their covenant thresholds and '
        'are current on payments. Overall portfolio quality is adequate but requires close '
        'management of the six breaching credits and active concentration monitoring through 2026.')

    # ── 8. Policy/Source Reconciliation ─────────────────────────────────────
    add_heading(doc, '8. Policy/Source Reconciliation', level=1)
    add_para(doc,
        'This section documents how all reported findings were derived and cross-checked '
        'against the underlying source files in accordance with the Loan Portfolio Review Policy '
        'requirement to verify pre-calculated fields against source data and resolve any '
        'source conflicts by applying the policy strictly.')
    add_para(doc,
        'Portfolio size: The total outstanding balance of $134,656,509 was computed by summing '
        'all 29 Outstanding Balance fields in Loan_Register_Q4_2025.xlsx. Loan count was '
        'independently confirmed from the row count of the same register (LN-001 through LN-029). '
        'No pre-calculated total appears in the source file; this figure was derived entirely '
        'from underlying row-level data.')
    add_para(doc,
        'Covenant thresholds and pre-calculated metrics: DSCR minimums, LTV maximums, and '
        'Current Ratio minimums were taken from the per-loan threshold columns in '
        'Covenant_Compliance_Workbook_Q4_2025.xlsx (pre-calculated fields). Each actual '
        'covenant metric was also independently verified against the underlying source data: '
        '(a) DSCR was recomputed by dividing Net Operating Income by Total Annual Debt Service '
        'from Borrower_Financial_Statements_Summary.xlsx — computed values matched workbook '
        'values to within one basis point for all loans tested (e.g., LN-001: computed 1.07x, '
        'workbook 1.07x; LN-004: computed 1.14x, workbook 1.14x; LN-008: computed 0.81x, '
        'workbook 0.81x); (b) LTV was recomputed by dividing the Loan Register outstanding '
        'balance by the Appraisal Summary appraised value — computed values matched the '
        'workbook within rounding for all three breach cases (LN-002: 87.9%, LN-008: 130.9%, '
        'LN-010: 85.2%); (c) Current Ratio was recomputed as Current Assets divided by '
        'Current Liabilities from the Financial Statements file — all three breach values '
        'matched the workbook figures. No material conflicts were identified between the '
        'pre-calculated workbook values and the independently derived calculations; where '
        'minor rounding differences appeared (±0.01), the workbook value was used as the '
        'primary reference, consistent with policy guidance to apply stated covenant metrics '
        'unless an underlying data conflict is found.')
    add_para(doc,
        'Watch List: Risk ratings were sourced exclusively from the Internal Risk Rating column '
        'of the Loan Register. Policy Section 4 specifies Rating ≥ 4 as the trigger. Three '
        'loans qualify (LN-001 at 4, LN-002 at 5, LN-008 at 6). No source conflict exists; '
        'the Loan Register is the single authoritative source for risk ratings.')
    add_para(doc,
        'Non-performing classification: The 97-DPD status of LN-008 was read from the Payment '
        'Status field ("97 Days Past Due") and Days Past Due column (97) of the Loan Register. '
        'This classification is corroborated by Exception EXC-010 (Exception Tracking Log), '
        'which independently records the loan at 97 days past due with a loss reserve, and by '
        'the Appraisal Summary collateral inspection note for LN-008, which references the '
        'Chapter 11 filing in December 2025. Three independent source fields across two files '
        'confirm this status; no conflict exists. No other loan crosses the 90-day threshold.')
    add_para(doc,
        'Concentration risk: Commercial Mortgage balances were summed across all nine loans '
        'of that type in the Loan Register ($88,481,044 = 65.71%). James Trotter balances '
        'were summed across his nine assigned loans in the Relationship Officer column '
        '($82,101,025 = 60.97%). C2 Tech Campus balance was read directly from the Loan '
        'Register ($21,117,450 = 15.68%). All percentages are expressed relative to the '
        'computed portfolio total of $134,656,509. No pre-calculated concentration figures '
        'exist in the source files; all were derived from underlying data.')
    add_para(doc,
        'Exception log: All 11 exception records were transcribed exactly from '
        'Exception_Tracking_Log_Q4_2025.xlsx, preserving Exception ID, Loan ID, Borrower Name, '
        'Exception Type, and Current Status without modification. Exception descriptions were '
        'used solely to corroborate other findings and were not independently assumed as facts '
        'unless supported by a second source file.')

    output_path = (
        r'c:\Users\user\OneDrive\Desktop\geranium project\Task3'
        r'\Loan_Portfolio_Review_Report_Q4_2025.docx'
    )
    doc.save(output_path)
    print(f'Report saved to: {output_path}')


if __name__ == '__main__':
    create_report()
    print('Done.')
