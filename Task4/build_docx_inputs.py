# -*- coding: utf-8 -*-
import docx
from docx.shared import Pt
import os

BASE = os.path.join(os.path.dirname(__file__), "PIVOT2_inputs")
os.makedirs(BASE, exist_ok=True)

def new_doc():
    d = docx.Document()
    d.styles['Normal'].font.name = 'Calibri'
    d.styles['Normal'].font.size = Pt(11)
    return d

# ============================================================
# 1. POLICY
# ============================================================
d = new_doc()
d.add_heading('COMMERCIAL PROPERTY AND BUSINESS INCOME POLICY', level=0)
d.add_paragraph('Heartland Mutual Insurance Company')
d.add_paragraph('Policy Number: CP-4471932  |  Policy Period: January 1, 2025 to January 1, 2026')
d.add_paragraph('Named Insured: Bluegrass Precision Tooling, LLC  |  Location: 4410 Ironworks Pike, Elizabethtown, KY')

d.add_heading('1. Declarations', level=1)
d.add_paragraph('Coverage A - Building: Limit of Insurance $2,000,000. Valuation: Replacement Cost. Coinsurance: 80%.')
d.add_paragraph('Coverage B - Business Personal Property: Limit of Insurance $800,000. Valuation: Replacement Cost. Coinsurance: 80%.')
d.add_paragraph('Coverage C - Business Income and Extra Expense: Combined Limit of Insurance $500,000. Valuation: Actual Loss Sustained. Coinsurance: Not Applicable to this Coverage.')
d.add_paragraph('Ordinance or Law Coverage Endorsement: Sublimit $250,000, applies to the increased cost of construction required to comply with building codes in effect at the time of repair. Not subject to the Coinsurance Clause in Section 5.')
d.add_paragraph('Deductible: $25,000 per occurrence, applied to Coverage A and B combined before any coinsurance adjustment is finalized on the remaining balance.')

d.add_heading('2. Definitions', level=1)
d.add_paragraph('Replacement Cost Value (RCV): The cost to repair or replace damaged property with new property of like kind and quality, without deduction for depreciation.')
d.add_paragraph('Actual Cash Value (ACV): Replacement Cost Value less physical depreciation.')
d.add_paragraph('Period of Restoration: The period beginning on the date of direct physical loss and ending on the date the property should be repaired, rebuilt, or replaced with reasonable speed and similar quality, not to exceed 12 consecutive months unless extended by endorsement.')
d.add_paragraph('Actual Loss Sustained: The Net Income the business would have earned had no loss occurred, less the Net Income actually earned during the Period of Restoration, plus continuing normal operating expenses.')

d.add_heading('3. Covered Causes of Loss and Exclusions', level=1)
d.add_paragraph('This policy covers direct physical loss to Covered Property unless the loss is excluded below. Covered Causes of Loss include fire, lightning, explosion, windstorm, and other perils not otherwise excluded.')
d.add_paragraph('Exclusions include, but are not limited to: Flood, Earth Movement, ordinary Wear and Tear, and faulty, inadequate, or defective workmanship, construction, or design (the "Faulty Workmanship Exclusion").')
d.add_paragraph('Ensuing Loss: If an excluded cause of loss, including the Faulty Workmanship Exclusion, results in a Covered Cause of Loss, such as fire, this policy covers the loss or damage caused by that resulting Covered Cause of Loss, even though the original excluded cause is not covered.')

d.add_heading('4. Valuation of Loss', level=1)
d.add_paragraph('Losses to Building property covered under Coverage A are valued and payable on a Replacement Cost basis once repairs of like kind and quality are contracted or completed, subject to the Coinsurance Clause in Section 5.')
d.add_paragraph('Losses to Business Personal Property covered under Coverage B are initially payable on an Actual Cash Value basis. The difference between Actual Cash Value and Replacement Cost Value ("Recoverable Depreciation") is payable only after the damaged or destroyed property is actually repaired or replaced and proof of replacement is submitted to the Company within 180 days of the date of loss, or such longer period as the Company agrees to in writing.')

d.add_heading('5. Coinsurance Clause', level=1)
d.add_paragraph(
    'This clause applies separately to Coverage A and Coverage B. We will not pay the full amount of any loss if, at '
    'the time of loss, the Limit of Insurance carried on the affected Coverage is less than the specified Coinsurance '
    'Percentage multiplied by the value of Covered Property at the time of loss. Instead, we will pay only the '
    'proportion of the loss that the Limit of Insurance carried bears to the amount of insurance required to meet the '
    'Coinsurance Percentage, computed as: Recovery = Loss x (Limit of Insurance Carried / (Coinsurance Percentage x '
    'Value at Time of Loss)), subject to the applicable Limit of Insurance and Deductible. This Coinsurance Clause '
    'does not apply to Coverage C or to the Ordinance or Law Coverage Endorsement.'
)

d.add_heading('6. Ordinance or Law Coverage', level=1)
d.add_paragraph(
    'When a building code or ordinance requires an increase in the cost of construction, repair, or replacement of '
    'damaged Covered Property beyond the cost of like kind and quality repair, this policy will pay the increased '
    'cost, up to the Ordinance or Law Coverage sublimit stated in the Declarations, in addition to the Coverage A '
    'Limit of Insurance. This coverage is not subject to the Coinsurance Clause in Section 5.'
)

d.add_heading('7. Debris Removal', level=1)
d.add_paragraph(
    'We will pay the reasonable cost to remove debris of Covered Property caused by a Covered Cause of Loss, in '
    'addition to the applicable Limit of Insurance, up to 25% of the amount we pay for the direct physical loss to '
    'that property. Debris Removal costs are not subject to the Coinsurance Clause in Section 5 or to the Deductible '
    'in Section 1.'
)

d.add_heading('8. Business Income and Extra Expense', level=1)
d.add_paragraph(
    'We will pay for the Actual Loss Sustained by the Named Insured due to the necessary suspension of operations '
    'during the Period of Restoration, computed as the Net Income the business would have earned had no loss '
    'occurred, less the Net Income actually earned during the Period of Restoration.'
)
d.add_paragraph(
    'We will also pay Extra Expense: necessary expenses the Named Insured incurs during the Period of Restoration '
    'that it would not have incurred had no loss occurred, to the extent such expenses reduce the amount of loss '
    'otherwise payable under this Coverage. Business Income and Extra Expense share a single combined Limit of '
    'Insurance as stated in the Declarations.'
)

d.add_heading('9. Claim Reporting and Proof of Loss', level=1)
d.add_paragraph(
    'The Named Insured must give prompt notice of loss and submit a signed, sworn Proof of Loss with supporting '
    'documentation, including repair estimates, financial records, and any independent investigation reports, within '
    '60 days of the Company\'s request, unless extended in writing.'
)

d.add_heading('10. Claim Settlement Recommendation Requirements', level=1)
d.add_paragraph('For every commercial property claim exceeding $100,000 in claimed value, the assigned Claims Examiner must prepare a Claim Settlement Recommendation. The recommendation must include:')
for item in [
    'Claim Summary describing the loss, the parties involved, and the amount claimed.',
    'Coverage and Cause of Loss Analysis determining which claimed items are covered, excluded, or subject to the Ensuing Loss provision, with reference to the applicable policy section.',
    'Property Valuation for Building and Business Personal Property, distinguishing Replacement Cost Value from Actual Cash Value and identifying any Recoverable Depreciation held pending replacement.',
    'Coinsurance Analysis showing whether each affected Coverage meets the Coinsurance Percentage and, if not, the resulting adjusted recovery.',
    'Business Income and Extra Expense Calculation showing the Actual Loss Sustained computation and application of the combined Limit of Insurance.',
    'Total Settlement Recommendation itemizing the amount payable under each Coverage and the aggregate recommended payment.',
    'Reserve and Payment Recommendation stating any amounts to be held as Recoverable Depreciation and the recommended next steps.',
]:
    d.add_paragraph(item, style='List Bullet')

d.save(os.path.join(BASE, 'Heartland_Mutual_Commercial_Property_Policy.docx'))
print('saved policy')

# ============================================================
# 2. PUBLIC ADJUSTER CLAIM SUBMISSION
# ============================================================
d = new_doc()
d.add_heading('Proof of Loss and Claim Submission', level=0)
d.add_paragraph('Claim Number: CLM-2025-08841  |  Policy Number: CP-4471932')
d.add_paragraph('Named Insured: Bluegrass Precision Tooling, LLC')
d.add_paragraph('Date of Loss: October 14, 2025  |  Date Submitted: December 8, 2025')
d.add_paragraph('Submitted by: Karen Ellsworth, Public Adjuster, Ellsworth Claims Advocates LLC, on behalf of the Named Insured')

d.add_heading('Description of Loss', level=1)
d.add_paragraph(
    'On the evening of October 14, 2025, a fire originated in the main production building at 4410 Ironworks Pike, '
    'Elizabethtown, KY, causing significant damage to the building structure, electrical systems, roof, and '
    'manufacturing equipment. The Named Insured has retained an independent fire investigator and a licensed general '
    'contractor to document the loss. Operations were fully suspended for approximately eight weeks and have since '
    'resumed on a limited basis at a temporary leased facility while repairs are completed.'
)

d.add_heading('Claimed Amounts', level=1)
d.add_paragraph('Based on the attached contractor estimate and the Named Insured\'s financial records, the following amounts are claimed:')
for line in [
    'Building repair and restoration (per contractor estimate, including code-required upgrades): $1,330,000',
    'Debris removal: $45,000',
    'Business Personal Property (manufacturing equipment and inventory), replacement cost: $620,000',
    'Business Income and Extra Expense loss during the period of suspended and limited operations: $560,000',
]:
    d.add_paragraph(line, style='List Bullet')
d.add_paragraph('Total Claimed: $2,555,000')
d.add_paragraph(
    'The Named Insured respectfully requests payment of the full claimed amount at Replacement Cost Value for all '
    'property items, without any coinsurance adjustment, as the Named Insured has maintained continuous coverage '
    'with this Company for over eight years and has not previously submitted a property claim.'
)

d.add_heading('Supporting Documentation Referenced', level=1)
for line in [
    'Contractor_Repair_Estimate.xlsx (licensed general contractor, itemized)',
    'Business_Financial_Statements.xlsx (pre-loss and post-loss operating results)',
    'Statement_of_Values_Appraisal.xlsx (independent appraisal at time of loss)',
    'Origin_and_Cause_Investigation_Report.docx (independent fire investigator)',
]:
    d.add_paragraph(line, style='List Bullet')

d.save(os.path.join(BASE, 'Public_Adjuster_Claim_Submission.docx'))
print('saved claim submission')

# ============================================================
# 3. ORIGIN AND CAUSE INVESTIGATION REPORT
# ============================================================
d = new_doc()
d.add_heading('Origin and Cause Investigation Report', level=0)
d.add_paragraph('Prepared for: Heartland Mutual Insurance Company  |  Claim Number: CLM-2025-08841')
d.add_paragraph('Investigator: Russell Kagawa, CFI, Bluegrass Forensic Investigations LLC')
d.add_paragraph('Date of Inspection: October 17, 2025  |  Report Date: November 3, 2025')

d.add_heading('Scope', level=1)
d.add_paragraph(
    'This investigation was conducted to determine the origin and cause of the fire at 4410 Ironworks Pike, '
    'Elizabethtown, KY, on October 14, 2025, including physical inspection of the structure, review of maintenance '
    'and electrical records, and witness interviews.'
)

d.add_heading('Findings', level=1)
d.add_paragraph(
    'Physical evidence and burn pattern analysis indicate the fire originated at an electrical subpanel located in '
    'the northeast corner of the main production floor. Maintenance records show this subpanel was modified in March '
    '2023 by a third-party electrical contractor to accommodate additional production equipment. The modification did '
    'not comply with the applicable electrical code in effect at the time, specifically undersized conductors were '
    'installed for the added load.'
)
d.add_paragraph(
    'Over approximately 30 months of operation, the undersized conductors experienced progressive thermal '
    'degradation of their insulation. On the date of loss, this degradation resulted in an electrical short that '
    'ignited surrounding insulation and combustible debris, and the resulting fire spread rapidly through the '
    'production floor and into the roof structure.'
)
d.add_paragraph(
    'No evidence of incendiary origin, flood, or earth movement was identified. No other contributing cause was '
    'found.'
)

d.add_heading('Conclusion', level=1)
d.add_paragraph(
    'The fire was caused by an electrical short resulting from a 2023 electrical modification that did not meet code. '
    'The undersized wiring installation itself was faulty workmanship; the fire that resulted from that faulty '
    'workmanship is, in this investigator\'s opinion, a distinct and independent physical event that spread by '
    'ordinary combustion once ignited.'
)

d.save(os.path.join(BASE, 'Origin_and_Cause_Investigation_Report.docx'))
print('saved origin and cause report')
