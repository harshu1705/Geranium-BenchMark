import os
from docx import Document
from docx.shared import Pt, Inches

def create_application():
    doc = Document()
    
    doc.add_heading('Commercial Trade Credit Application', 0)
    doc.add_paragraph('Date: August 15, 2026')
    doc.add_paragraph('Applicant: Hartsfield Polymer Technologies, LLC')
    doc.add_paragraph('Address: 450 Industrial Parkway, Akron, OH 44301')
    doc.add_paragraph('Requested Credit Limit: $1,500,000')
    doc.add_paragraph('Requested Terms: Net 60')
    
    doc.add_heading('Company Background & Business Description', level=1)
    doc.add_paragraph(
        'Hartsfield Polymer Technologies, LLC is a mid-sized, privately-held manufacturer of specialized polymer casings and high-durability plastic injection molded components, primarily serving the North American automotive sector. '
        'Established in 2008, the company has grown steadily by focusing on precision engineering and strict quality control standards required by major automotive OEMs and Tier 1 suppliers. '
        'Our core products include custom-molded dashboard under-structures, specialized engine compartment fluid reservoirs, and reinforced polymer brackets. '
        'We operate out of a single 120,000 square-foot manufacturing facility in Akron, Ohio, employing approximately 145 full-time personnel.'
    )
    
    doc.add_heading('Reason for Credit Limit Increase', level=1)
    doc.add_paragraph(
        'We are formally requesting an increase from our current $400,000 (Net 30) credit limit with Crestline Chemical Supply to $1,500,000, along with an extension of terms to Net 60. '
        'This request is driven directly by a newly awarded, three-year supply contract with Vanguard Automotive Systems, a major Tier 1 automotive supplier. '
        'Under this new contract, we are required to significantly increase our production volume of high-density polyethylene (HDPE) components. '
        'The contract guarantees an estimated annual volume of $3,800,000 in direct resin purchases. '
        'To support the required inventory ramp-up and meet the strict just-in-time delivery schedules mandated by Vanguard, we must hold significantly more raw material inventory on hand. '
        'The extended Net 60 terms are necessary to align with the payment terms we are required to offer Vanguard Automotive Systems under the new master service agreement. '
        'Additionally, this expansion will require us to bring on line three new injection molding machines, which are currently being installed and calibrated. '
        'We anticipate that the increased capacity will not only fulfill the Vanguard contract but also provide an additional 15% output capacity for spot market orders, further diversifying our revenue streams. '
        'We have secured a favorable long-term lease on an adjacent 40,000 square-foot warehouse to handle the additional finished goods inventory and raw material storage, ensuring smooth logistics. '
        'The operational scale-up is fully funded through our existing capital reserves, meaning the requested trade credit increase is strictly for raw material procurement and working capital alignment.'
    )
    
    doc.add_heading('Corporate Information', level=1)
    doc.add_paragraph('Legal Structure: Limited Liability Company (LLC)')
    doc.add_paragraph('State of Organization: Ohio')
    doc.add_paragraph('Year Established: 2008')
    doc.add_paragraph('DUNS Number: 04-729-8135')
    doc.add_paragraph('Federal Tax ID (EIN): 34-8192754')
    doc.add_paragraph('Primary NAICS Code: 326199 (All Other Plastics Product Manufacturing)')
    
    doc.add_heading('Management & Ownership', level=1)
    doc.add_paragraph('David Harmon – Chief Financial Officer (CFO)')
    doc.add_paragraph('Sarah Jenkins – Chief Executive Officer (CEO) and 60% Majority Owner')
    doc.add_paragraph('Michael Hartsfield – Founder and 40% Minority Owner')
    
    doc.add_heading('Banking Information', level=1)
    doc.add_paragraph('Primary Bank: First Midwest Commercial Bank')
    doc.add_paragraph('Account Type: Corporate Operating Account & Payroll Account')
    doc.add_paragraph('Relationship Manager: Thomas Weaver')
    doc.add_paragraph('Average Balance: Low 6-figures')
    
    doc.add_heading('Trade References', level=1)
    doc.add_paragraph(
        '1. Rust Belt Resins Inc.\n'
        '   Contact: Accounting Department (accounting@rustbeltresins.com)\n'
        '   High Credit: $650,000\n'
        '   Terms: Net 30\n'
        '   Status: Current/Active'
    )
    doc.add_paragraph(
        '2. Buckeye Industrial Packaging\n'
        '   Contact: Billing Support (billing@buckeyepack.com)\n'
        '   High Credit: $250,000\n'
        '   Terms: Net 30\n'
        '   Status: Current/Active'
    )
    doc.add_paragraph(
        '3. Ohio Valley Logistics & Freight\n'
        '   Contact: Sarah Miller (smiller@ovlogistics.com)\n'
        '   High Credit: $120,000\n'
        '   Terms: Net 15\n'
        '   Status: Current/Active'
    )
    
    doc.add_heading('Certification and Signatures', level=1)
    doc.add_paragraph(
        'The undersigned hereby certifies that the information provided in this credit application is true, correct, and complete to the best of their knowledge. '
        'The applicant authorizes Crestline Chemical Supply, Inc. to investigate the banking and trade references provided herein, and to obtain commercial credit reports as necessary to evaluate this request.'
    )
    doc.add_paragraph('\nSigned: ___________________________')
    doc.add_paragraph('David Harmon, CFO')
    doc.add_paragraph('Hartsfield Polymer Technologies, LLC')
    doc.add_paragraph('Date: August 15, 2026')
    
    doc.save('Customer_Credit_Application_Hartsfield.docx')

def create_policy():
    doc = Document()
    
    doc.add_heading('Crestline Chemical Supply, Inc. - Corporate Credit Policy', 0)
    doc.add_paragraph('Effective Date: January 1, 2026')
    doc.add_paragraph('Document Owner: VP of Credit Risk')
    
    doc.add_heading('Section 1: Purpose and Scope', level=1)
    doc.add_paragraph(
        'The purpose of this Corporate Credit Policy is to establish clear guidelines and standardized procedures for the extension of trade credit to customers of Crestline Chemical Supply, Inc. '
        'This policy applies to all commercial accounts seeking open/unsecured credit terms for the purchase of bulk chemicals, industrial resins, and related supplies. '
        'Our primary objective is to maximize profitable sales while maintaining an acceptable level of credit risk and minimizing bad debt write-offs.'
    )
    
    doc.add_heading('Section 2: General Credit Terms and Standard Approvals', level=1)
    doc.add_paragraph(
        'The standard payment term for all Crestline commercial accounts is Net 30 days from the invoice date, unless otherwise approved by the Corporate Credit Department. '
        'Credit limits under $500,000 may be approved by a Senior Credit Analyst based on satisfactory trade references and a clear commercial credit report showing no major derogatory items. '
        'Requests between $500,000 and $1,000,000 require review of the most recent year-end financial statements and approval by the Credit Manager. '
        'In evaluating these mid-tier accounts, the Credit Department must ensure that working capital is positive and that the customer has demonstrated a track record of stable or growing revenue over the past two fiscal years. '
        'Any mid-tier account exhibiting a sudden drop in sales exceeding 20% year-over-year must be escalated to the Director of Credit for secondary review.'
    )
    
    doc.add_heading('Section 3: Trade Credit Approvals over $1,000,000', level=1)
    doc.add_paragraph(
        'Requests for unsecured open credit lines exceeding $1,000,000, or requests for extended payment terms (Net 60 or greater), represent a significant concentration of risk. '
        'These requests must undergo a rigorous quantitative financial analysis by the Corporate Credit Department. '
        'To qualify for an UNSECURED credit limit of this magnitude, the customer must meet ALL of the following mandatory criteria without exception:'
    )
    
    doc.add_paragraph(
        '1. Altman Z-Score for Private Companies: The customer must demonstrate financial stability by maintaining a calculated Z-score of 1.80 or higher based on their most recent financial statements.\n'
        '   - Formula: Z = 1.2(A) + 1.4(B) + 3.3(C) + 0.6(D) + 1.0(E)\n'
        '   - Where A = Working Capital / Total Assets\n'
        '   - Where B = Retained Earnings / Total Assets\n'
        '   - Where C = EBIT / Total Assets\n'
        '   - Where D = Book Value of Equity / Total Liabilities\n'
        '   - Where E = Sales / Total Assets'
    )
    
    doc.add_paragraph(
        "2. Internal Payment History: For existing accounts requesting limit increases, payment performance must be satisfactory. "
        "Their average 'Days to Pay' over the most recent 3-month period on the Crestline Accounts Receivable (AR) Ledger must not exceed 45 days. "
        "Accounts exhibiting a trend of slowing payments beyond this 45-day threshold are strictly ineligible for limit increases on unsecured terms."
    )
    
    doc.add_paragraph(
        '3. Public Records: The customer must have a clean external credit profile. '
        'Specifically, the commercial credit report must show no active or unresolved bankruptcies, no active judgments, and no active or unresolved tax liens exceeding $10,000.'
    )
    
    doc.add_heading('Section 4: Conditional Approvals and Credit Risk Mitigation (CRM)', level=1)
    doc.add_paragraph(
        'If a customer requires a credit limit exceeding $1,000,000 to support their purchasing needs, but fails to meet one or more of the mandatory criteria outlined in Section 3, the Credit Analyst MUST issue a formal recommendation based strictly on the following deterministic rules:'
    )
    doc.add_paragraph(
        'a) Approve with Conditions (1 Failure): If the applicant fails EXACTLY ONE of the mandatory criteria in Section 3, AND has no active tax liens of any amount, the analyst MUST conditionally approve the requested limit. '
        'This conditional approval requires Credit Risk Mitigation (CRM) in the form of an Irrevocable Letter of Credit (ILOC) drawn on a Tier 1 domestic bank, covering the entire requested limit.'
    )
    doc.add_paragraph(
        'b) Decline the Request (2+ Failures or Tax Lien): If the applicant fails TWO OR MORE of the mandatory criteria in Section 3, OR if they have any active tax lien exceeding $10,000, the analyst MUST decline the limit increase and maintain the customer on their existing, lower credit limit. '
        'In this scenario, Conditional Approval via CRM is strictly prohibited.'
    )
    
    doc.add_heading('Section 5: Annual Reviews and Monitoring', level=1)
    doc.add_paragraph(
        'All accounts with credit limits exceeding $500,000 are subject to an annual review. '
        'Customers must submit updated year-end financial statements within 120 days of their fiscal year-end. '
        'Failure to provide updated financials may result in an automatic suspension or reduction of the credit line until the required documentation is received and analyzed. '
        'During the annual review process, the credit analyst is responsible for recalculating all relevant financial ratios, pulling a fresh commercial credit report, and checking the internal accounts receivable ledger for any signs of payment deterioration. '
        'If an account that previously qualified for a $1,000,000+ unsecured line falls out of compliance with the Section 3 criteria during an annual review, the analyst must initiate a 90-day cure period. '
        'If the account cannot be brought back into compliance or adequately mitigated within that timeframe, the limit must be reduced to the maximum level supported by the revised financial assessment.'
    )
    
    doc.save('Crestline_Corporate_Credit_Policy.docx')

def create_dnb():
    doc = Document()
    
    doc.add_heading('Dun & Bradstreet Business Credit Report', 0)
    doc.add_paragraph('Report Date: August 18, 2026')
    doc.add_paragraph('Company Name: Hartsfield Polymer Technologies, LLC')
    doc.add_paragraph('D-U-N-S Number: 04-729-8135')
    doc.add_paragraph('Address: 450 Industrial Parkway, Akron, OH 44301')
    doc.add_paragraph('Primary SIC: 3089 (Plastics Products, Not Elsewhere Classified)')
    doc.add_paragraph('Employees: 145')
    doc.add_paragraph('Estimated Annual Sales: $11.2 Million')
    
    doc.add_heading('Risk Summary & Predictive Scores', level=1)
    doc.add_paragraph(
        'D&B Paydex Score: 68 (Industry Median: 75). \n'
        'A Paydex of 68 indicates that payments to suppliers are typically made 15 to 22 days beyond standard terms. '
        'This represents a deterioration from the score of 74 reported in Q1 2026. '
        'The Paydex score is a unique, dollar-weighted indicator of a business\'s past payment performance based on trade experiences reported to Dun & Bradstreet by various vendors. '
        'A score of 80 denotes that payments are made exactly within terms, while scores below 70 generally flag an elevated risk of late payments.'
    )
    doc.add_paragraph(
        'Financial Stress Score: 1450. \n'
        'Percentile: 32 (where 1 is highest risk and 100 is lowest). '
        'This indicates a higher than average risk of experiencing severe financial stress over the next 12 months compared to similar businesses in the manufacturing sector. '
        'The Financial Stress Score utilizes statistical modeling to predict the likelihood of business failure, combining demographic data, financial ratios, and payment history.'
    )
    doc.add_paragraph(
        'Overall Risk Assessment: Moderate to High Risk. \n'
        'The company has shown a recent deterioration in payment promptness across several major suppliers, alongside a newly filed public record. '
        'Suppliers are advised to closely monitor this account, potentially restricting open terms or requiring partial cash-in-advance for large orders until the payment trend stabilizes or improves.'
    )
    
    doc.add_heading('Payment Habits Analysis', level=1)
    doc.add_paragraph(
        'Based on 14 trade experiences reported in the last 12 months, the company demonstrates the following payment trends:\n'
        '- Within Terms: 42%\n'
        '- 1 to 30 Days Slow: 38%\n'
        '- 31 to 60 Days Slow: 15%\n'
        '- 61+ Days Slow: 5%'
    )
    doc.add_paragraph(
        'This represents a significant shift from the prior 12-month period, where 75% of trade experiences were reported as paying within terms, and only 5% were reported as 31+ days slow. '
        'The recent increase in late payments suggests a possible tightening of operating cash flow or a deliberate strategy to stretch payables. '
        'In particular, several key suppliers in the raw materials sector have reported a marked slowdown in payments over the last two quarters.'
    )
    
    doc.add_heading('Select Trade Lines', level=1)
    doc.add_paragraph(
        '1. Raw Materials Supplier: High Credit $650,000. Terms: Net 30. Recent experience: Paying 15-30 days beyond terms.\n'
        '2. Packaging Supplier: High Credit $250,000. Terms: Net 30. Recent experience: Paying on time or up to 10 days slow.\n'
        '3. Logistics/Freight: High Credit $120,000. Terms: Net 15. Recent experience: Paying on time.\n'
        '4. Equipment Leasing: High Credit $45,000. Terms: Net 30. Recent experience: Paying 30+ days slow.'
    )
    
    doc.add_heading('Public Records & Legal Filings', level=1)
    doc.add_paragraph('Bankruptcies: 0 recorded.')
    doc.add_paragraph('Judgments: 0 recorded.')
    doc.add_paragraph('Liens: 1 active record found.')
    
    doc.add_paragraph(
        '- Filed Date: July 12, 2026\n'
        '- Filer: State of Ohio Department of Taxation\n'
        '- Type: State Tax Lien\n'
        '- Amount: $45,000\n'
        '- Status: Unresolved / Open.'
    )
    
    doc.add_paragraph(
        'D&B Analyst Note: D&B contacted the company regarding the active lien on August 10, 2026. '
        'Management stated this is a dispute over equipment depreciation rules and they expect it to be resolved by Q4, but it remains officially active on public record.'
    )
    
    doc.save('Dun_and_Bradstreet_Credit_Report.docx')

if __name__ == '__main__':
    create_application()
    create_policy()
    create_dnb()
    print("Word documents generated successfully.")
