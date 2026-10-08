# -*- coding: utf-8 -*-
import docx
from docx.shared import Pt
import os

BASE = os.path.join(os.path.dirname(__file__), "PIVOT3_inputs")
os.makedirs(BASE, exist_ok=True)

def new_doc():
    d = docx.Document()
    d.styles['Normal'].font.name = 'Calibri'
    d.styles['Normal'].font.size = Pt(11)
    return d

# ============================================================
# 1. Social Security Benefit Estimates
# ============================================================
d = new_doc()
d.add_heading('Social Security Administration', level=0)
d.add_paragraph('Your Social Security Statement - Estimated Benefits Summary')
d.add_paragraph('Prepared for financial planning purposes based on current earnings record. Actual benefits at claiming may vary based on continued earnings and annual COLA adjustments.')

d.add_heading('Gerald Whitfield (DOB: November 1966)', level=1)
d.add_paragraph('Estimated monthly retirement benefit if you claim at:')
for age, amt in [('Age 62 (earliest eligibility)', '$2,400'), ('Age 65', '$3,100'), ('Age 67 (Full Retirement Age)', '$3,400')]:
    d.add_paragraph(f'{age}: {amt} per month', style='List Bullet')

d.add_heading('Linda Whitfield (DOB: March 1968)', level=1)
d.add_paragraph('Estimated monthly retirement benefit based on her own earnings record, if she claims at:')
for age, amt in [('Age 62 (earliest eligibility)', '$1,150'), ('Age 65', '$1,450'), ('Age 67 (Full Retirement Age)', '$1,600')]:
    d.add_paragraph(f'{age}: {amt} per month', style='List Bullet')

d.add_paragraph(
    'Note: These estimates assume each spouse claims on their own earnings record. Benefits claimed before Full '
    'Retirement Age are permanently reduced; benefits claimed after Full Retirement Age (up to age 70) increase '
    'further under delayed retirement credits, not shown above.'
)
d.save(os.path.join(BASE, 'Social_Security_Benefit_Estimates.docx'))
print('saved SS estimates')

# ============================================================
# 2. Insurance and Estate Summary
# ============================================================
d = new_doc()
d.add_heading('Insurance and Estate Planning Summary', level=0)
d.add_paragraph('Whitfield Household  |  Prepared from client-provided policy documents and intake interview')

d.add_heading('Life Insurance', level=1)
d.add_paragraph('Gerald Whitfield: 20-year level term life policy, $500,000 face amount, purchased 2008, term expires when Gerald turns 68 (2034). No conversion to permanent coverage has been elected.', style='List Bullet')
d.add_paragraph('Linda Whitfield: Employer-provided group term life, 1x salary ($72,000), coverage ends at retirement/separation from employer.', style='List Bullet')

d.add_heading('Long-Term Care Insurance', level=1)
d.add_paragraph('Neither Gerald nor Linda currently holds a long-term care insurance policy. Both have declined LTC coverage in the past due to premium cost.')

d.add_heading('Health Insurance', level=1)
d.add_paragraph(
    'Both are currently covered under Gerald\'s employer group health plan. This coverage ends upon Gerald\'s '
    'retirement. Neither spouse is eligible for Medicare until age 65. A current quote for ACA marketplace or COBRA '
    'continuation coverage for the household runs approximately $1,800 per month combined until Medicare eligibility.'
)

d.add_heading('Estate Documents', level=1)
d.add_paragraph('Wills: Both have basic wills executed in 2011, naming each other as primary beneficiary and their two adult children as contingent beneficiaries. Not reviewed or updated since.', style='List Bullet')
d.add_paragraph('Beneficiary designations on file with account custodians (401(k), 403(b), Roth IRA) are consistent with the wills and were last confirmed in 2019.', style='List Bullet')
d.add_paragraph('No trust structures are in place. No power of attorney or healthcare directive documents are on file with this office.', style='List Bullet')

d.save(os.path.join(BASE, 'Insurance_and_Estate_Summary.docx'))
print('saved insurance and estate summary')

# ============================================================
# 3. Client Intake and Goals Summary
# ============================================================
d = new_doc()
d.add_heading('Client Intake and Goals Summary', level=0)
d.add_paragraph('Whitfield Household  |  Intake Meeting Notes, prepared by the advisory team')

d.add_heading('Household', level=1)
d.add_paragraph('Gerald Whitfield, age 60, Operations Manager, plans to retire at age 62.')
d.add_paragraph('Linda Whitfield, age 58, Office Administrator, plans to retire at the same time as Gerald.')
d.add_paragraph('Two adult children, financially independent, not a current support obligation.')

d.add_heading('Stated Goals', level=1)
for line in [
    'Retire together when Gerald turns 62, approximately 2 years from today.',
    'Maintain a comfortable but not lavish retirement lifestyle; both are willing to trim discretionary spending somewhat but do not want a dramatic lifestyle reduction.',
    'Both currently lean toward claiming Social Security as soon as they are eligible at 62, "to get something back" after years of paying in, but are open to a recommendation if the numbers say otherwise.',
    'No strong legacy or inheritance goal beyond leaving whatever is left to their children; preserving their own retirement security comes first.',
    'Risk tolerance: moderate. Comfortable with market fluctuation but increasingly want to reduce equity exposure as retirement approaches.',
]:
    d.add_paragraph(line, style='List Bullet')

d.add_heading('Firm Reference: Model Portfolio Guidance', level=1)
d.add_paragraph(
    'For clients within five years of their target retirement date and a moderate risk tolerance, the firm\'s '
    'standard model portfolio target is approximately 55% equities, 40% fixed income, and 5% cash, adjusted for '
    'individual circumstances. This is a planning reference point, not a fixed requirement.'
)

d.add_heading('Planning Assumptions on File', level=1)
for line in [
    'Assumed long-term portfolio growth rate for projection purposes: 6% annually, compounded.',
    'Assumed general inflation rate for expense projections: 3% annually.',
    'Combined ongoing retirement account and brokerage contributions until retirement: $3,100 per month.',
    'Sustainable withdrawal planning uses an initial 4% withdrawal rate of the portfolio balance at retirement as the firm\'s standard quick-screen sustainability test for a 30-year retirement horizon, with the withdrawal amount then adjusted for inflation in later years.',
]:
    d.add_paragraph(line, style='List Bullet')

d.save(os.path.join(BASE, 'Client_Intake_and_Goals_Summary.docx'))
print('saved client intake summary')
