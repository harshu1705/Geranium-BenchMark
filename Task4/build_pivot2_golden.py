# -*- coding: utf-8 -*-
import docx
from docx.shared import Pt
import os

OUT_DIR = os.path.join(os.path.dirname(__file__), "PIVOT2_golden")
os.makedirs(OUT_DIR, exist_ok=True)
OUT = os.path.join(OUT_DIR, "Claim_Settlement_Recommendation_CLM-2025-08841.docx")

d = docx.Document()
d.styles['Normal'].font.name = 'Calibri'
d.styles['Normal'].font.size = Pt(11)

def h1(t): d.add_heading(t, level=1)
def h2(t): d.add_heading(t, level=2)
def p(t): d.add_paragraph(t)

def table(headers, rows):
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

d.add_heading('Claim Settlement Recommendation', level=0)
p('Claim Number: CLM-2025-08841  |  Policy Number: CP-4471932')
p('Named Insured: Bluegrass Precision Tooling, LLC  |  Date of Loss: October 14, 2025')
p('Prepared by: Property Claims Examiner, Heartland Mutual Insurance Company  |  Report Date: December 22, 2025')

h1('1. Claim Summary')
p(
    "Bluegrass Precision Tooling, LLC submitted a claim following an October 14, 2025 fire at its Elizabethtown, KY "
    "production facility. The Named Insured's public adjuster, Ellsworth Claims Advocates LLC, submitted a Proof of "
    "Loss on December 8, 2025 claiming $2,596,990: $1,330,000 for building repair (including code-required upgrades), "
    "$45,000 for debris removal, $620,000 for Business Personal Property at replacement cost, and $601,990 for "
    "Business Income and Extra Expense. The claim requests payment without any coinsurance adjustment on the basis "
    "of the Named Insured's claims-free history, which is not a basis the policy provides for waiving the "
    "Coinsurance Clause in Section 5."
)
p(
    "This recommendation applies the policy's coverage, valuation, coinsurance, ordinance or law, and business "
    "income provisions to the claimed amounts and the independent contractor, appraisal, and financial documentation "
    "on file, and recommends an immediate settlement of $2,184,315, with an additional $200,500 of recoverable "
    "depreciation payable upon proof of Business Personal Property replacement."
)

h1('2. Coverage and Cause of Loss Analysis')
p(
    "Per the Origin and Cause Investigation Report, the fire originated at an electrical subpanel modified in March "
    "2023 by a third-party contractor using undersized conductors that did not meet the applicable code. That "
    "modification is faulty workmanship, excluded under Policy Section 3. However, Section 3's Ensuing Loss "
    "provision covers a resulting Covered Cause of Loss, and the investigator's conclusion that the fire spread by "
    "ordinary combustion once ignited establishes fire, not the underlying wiring defect, as the proximate cause of "
    "the building and property damage. The fire damage is therefore a covered loss under Coverage A and Coverage B, "
    "and the resulting suspension of operations is a covered loss under Coverage C. No exclusion for Flood, Earth "
    "Movement, or Wear and Tear is implicated, consistent with the investigator's findings."
)

h1('3. Property Valuation')
h2('Building (Coverage A)')
p(
    "The Contractor_Repair_Estimate.xlsx separates the $1,330,000 claimed building figure into two categories that "
    "the policy treats differently: $1,150,000 of like-kind-and-quality structural, electrical, and roofing repair, "
    "and $180,000 of code-required upgrades to the fire sprinkler and electrical service, driven by current code "
    "rather than restoring the building to its pre-loss condition. Only the $1,150,000 like-kind-and-quality figure "
    "is valued under Coverage A; the $180,000 code-required upgrade is addressed separately under Ordinance or Law "
    "coverage in Section 5 below. Building repairs are valued on a Replacement Cost basis per Policy Section 4 once "
    "repairs of like kind and quality are contracted, which the Named Insured has done."
)
h2('Business Personal Property (Coverage B)')
p(
    "Per the BPP Loss Detail in Business_Financial_Statements.xlsx, damaged equipment and inventory carry a "
    "Replacement Cost Value of $620,000 against accumulated depreciation ranging from 10% to 50% by item, producing "
    "an Actual Cash Value of $419,500. Per Policy Section 4, the $200,500 difference between Replacement Cost Value "
    "and Actual Cash Value is Recoverable Depreciation, payable only after the Named Insured replaces the damaged "
    "property and submits proof of replacement within 180 days of the date of loss. The public adjuster's claim of "
    "the full $620,000 Replacement Cost Value up front does not reflect this holdback requirement."
)

h1('4. Coinsurance Analysis')
p(
    "Per Statement_of_Values_Appraisal.xlsx, the Building's appraised value at the time of loss was $2,700,000 "
    "against an 80% Coinsurance Percentage, requiring $2,160,000 of insurance to avoid a penalty. The Named Insured "
    "carried only $2,000,000, a coinsurance ratio of 92.59% ($2,000,000 / $2,160,000). Per Policy Section 5, the "
    "recovery on the $1,150,000 like-kind-and-quality building loss is therefore $1,150,000 x 92.59% = $1,064,815, "
    "before the $25,000 per-occurrence Deductible, for a net Building recovery of $1,039,815. Business Personal "
    "Property was appraised at $780,000 against $800,000 carried, above the $624,000 required under the 80% test, so "
    "no coinsurance penalty applies to Coverage B."
)

h1('5. Ordinance or Law and Debris Removal')
p(
    "The $180,000 code-required fire sprinkler and electrical service upgrade is covered under the Ordinance or Law "
    "Coverage Endorsement, which carries a $250,000 sublimit and, per Policy Section 6, is not subject to the "
    "Coinsurance Clause. The $180,000 claimed falls within the $250,000 sublimit and is payable in full. Debris "
    "removal of $45,000 is payable in full as an Additional Coverage per Policy Section 7, which is likewise not "
    "subject to coinsurance or the Deductible."
)

h1('6. Business Income and Extra Expense Calculation')
p(
    "Per Policy Section 8, Business Income is valued on an Actual Loss Sustained basis: the Net Income the business "
    "would have earned had no loss occurred, less the Net Income actually earned during the Period of Restoration. "
    "The Pre-Loss Monthly Trend in Business_Financial_Statements.xlsx shows $633,000 of actual Net Income in the "
    "same seven-month window one year earlier (October 2024 through April 2025), the prior-year comparable period "
    "for the Period of Restoration here (October 2025 through April 2026, per the contractor's estimated 7-month "
    "completion timeline). Applying the approximately 3% year-over-year growth trend noted in that file produces a "
    "projected Net Income of $651,990 had no loss occurred."
)
p(
    "The Post-Loss Partial Operations schedule shows the Named Insured resumed limited production at a temporary "
    "facility starting in December 2025, earning $145,000 of actual Net Income through April 2026. The Business "
    "Income loss is therefore $651,990 minus $145,000, or $506,990."
)
p(
    "The Named Insured incurred $95,000 of Extra Expense to establish and operate the temporary facility. That "
    "expense enabled $145,000 of Net Income that otherwise would not have been earned, a net benefit to the loss "
    "otherwise payable, so it qualifies in full under Policy Section 8's requirement that Extra Expense reduce the "
    "loss. Combined Business Income and Extra Expense total $506,990 plus $95,000, or $601,990, against Coverage C's "
    "$500,000 combined Limit of Insurance. Per the Declarations, the combined limit caps payment at $500,000, "
    "$101,990 less than the amount claimed and less than the amount otherwise calculated under the Actual Loss "
    "Sustained method."
)

h1('7. Total Settlement Recommendation')
table(
    ["Coverage / Item", "Amount Recommended", "Basis"],
    [
        ["Building (Coverage A), net of coinsurance and deductible", "$1,039,815", "$1,150,000 LKQ x 92.59% coinsurance ratio, less $25,000 deductible"],
        ["Debris Removal", "$45,000", "Additional Coverage, not subject to coinsurance or deductible"],
        ["Ordinance or Law (code-required upgrades)", "$180,000", "Within $250,000 sublimit; not subject to coinsurance"],
        ["Business Personal Property (Coverage B), Actual Cash Value", "$419,500", "$620,000 RCV less $200,500 depreciation; no coinsurance penalty"],
        ["Business Income and Extra Expense (Coverage C)", "$500,000", "Capped at combined Limit of Insurance ($601,990 calculated)"],
        ["Total Immediate Settlement Recommended", "$2,184,315", ""],
    ],
)
p(
    "The claimed amount of $2,596,990 is reduced by $412,675 through application of the Coinsurance Clause to the "
    "Building, deferral of $200,500 of Business Personal Property depreciation pending replacement, and the "
    "$101,990 amount of Business Income and Extra Expense loss exceeding Coverage C's combined limit."
)

h1('8. Reserve and Payment Recommendation')
p(
    "Recommend issuing immediate payment of $2,184,315. Recommend holding an additional $200,500 in reserve as "
    "Recoverable Depreciation on Business Personal Property, payable upon the Named Insured's proof of replacement "
    "within 180 days of the October 14, 2025 date of loss, per Policy Section 4. No further reserve is warranted for "
    "Business Income and Extra Expense, as the combined Limit of Insurance has been fully exhausted and the "
    "$101,990 excess is not recoverable under this policy. Recommend advising the Named Insured in writing that the "
    "coinsurance penalty applies notwithstanding its claims history, since the policy provides no basis to waive "
    "Section 5 on that ground."
)

d.save(OUT)
print('saved', OUT)
