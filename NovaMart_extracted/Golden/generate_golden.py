import os
from docx import Document

def generate_memo():
    doc = Document()
    doc.add_heading('Underwriting Memorandum: NovaMart Inc.', 0)

    doc.add_heading('1. Executive Summary', level=1)
    doc.add_paragraph('Applicant: NovaMart Inc.')
    doc.add_paragraph('Industry: Retail (NAICS 452112)')
    doc.add_paragraph('Annual Revenue: $502,400,000')
    doc.add_paragraph('Requested Limit: $15,000,000')
    doc.add_paragraph('Recommendation: DECLINE')
    doc.add_paragraph(
        'Summary Justification: This risk represents a hard decline. The applicant is missing Endpoint Detection & '
        'Response (EDR) on 2% of their endpoints (legacy barcode scanners) per the Ransomware Supplemental Questionnaire. '
        'Section 5 of the Underwriting Guidelines explicitly states that EDR is "Required for all endpoints. If missing, Decline." '
        'Furthermore, the perimeter scan identified an open Critical vulnerability (CVE-2023-4528) on the VPN appliance. '
        'Because the risk triggers multiple simultaneous decline conditions (missing EDR and an unresolved critical CVE on the perimeter), '
        'Section 8 mandates immediate file closure with no VP referral permitted. '
        'The pricing analysis and further underwriting review completed below are strictly for file documentation and auditing purposes.'
    )

    doc.add_heading('2. Premium Calculation and Surcharge Analysis', level=1)
    doc.add_paragraph('Base Premium Calculation: $25,120')
    doc.add_paragraph(
        'Rationale: The TTM revenue derived from the Financial Statements Summary is $502,400,000. '
        'Using the standard retail base rate of $0.05 per $1,000 in revenue, the base premium is calculated as follows: '
        '($502,400,000 / 1000) * 0.05 = $25,120. The base premium is established at $25,120.'
    )
    doc.add_paragraph('Loss Run Surcharge Multiplier: 1.20x')
    doc.add_paragraph(
        'Rationale: The Loss Run History reveals two cyber claims in the last 3 years: a 2025 cyber ransomware claim (CLM-25-1102) '
        'totaling $650,000 paid, and a 2023 cyber claim (CLM-23-4412) totaling $7,500. The cyber-only 3-year total is $657,500. '
        'Section 4 of the Underwriting Guidelines states that claims totaling $500,000 or more in the last 3 years require a 1.20x penalty multiplier. '
        'Because the cyber-only total ($657,500) exceeds the $500,000 threshold, it warrants the maximum 1.20x loss penalty.'
    )
    doc.add_paragraph('Security Surcharge Multiplier: 1.25x')
    doc.add_paragraph(
        'Rationale: The Ransomware Supplemental App confirms that Multi-Factor Authentication (MFA) is missing for Third-Party/Contractor Remote Access. '
        'Per Section 5 of the underwriting guidelines, missing MFA on remote access mandates a 1.25x security surcharge multiplier. '
        'This is a critical gap given that third-party vendors often serve as the initial infection vector in modern retail ransomware deployments.'
    )
    doc.add_paragraph('Final Calculated Premium: $37,680 (Not Quoted due to Decline)')
    doc.add_paragraph(
        'Rationale: Base Premium ($25,120) * Loss Multiplier (1.20) * Security Multiplier (1.25) = $37,680. '
        'The calculated premium is $37,680. This significantly misses the broker\'s target premium of $45,000. '
        'Even if the decline conditions were resolved, the account is uncompetitive and would face heavy pricing resistance from the insured.'
    )

    doc.add_heading('3. Limit Cap and Authority Requirements', level=1)
    doc.add_paragraph('Eligible Limit: $10,000,000 (Moot due to Decline)')
    doc.add_paragraph(
        'Rationale: The missing Contractor MFA caps the maximum eligible limit at $10,000,000 per Section 5 of the guidelines. '
        'The requested limit of $15,000,000 is therefore ineligible, even setting aside the EDR decline trigger.'
    )
    doc.add_paragraph('Authority Required: None (Immediate File Closure Required)')
    doc.add_paragraph(
        'Rationale: While a $10,000,000 limit normally requires Senior Underwriter authority, '
        'Section 8 of the Underwriting Guidelines states that because the risk triggers multiple simultaneous decline conditions '
        '(missing EDR and an unresolved critical CVE on the perimeter), immediate file closure is required with no VP referral permitted.'
    )

    doc.add_heading('4. Subjectivity Analysis', level=1)
    doc.add_paragraph('Mandatory Security Condition: N/A (Declined)')
    doc.add_paragraph(
        'Rationale: The Third-Party Security Scan identified an Open Critical vulnerability, CVE-2023-4528 (Fortinet FortiOS Heap Buffer Overflow), '
        'on the VPN appliance (vpn.novamart.com). This vulnerability allows unauthenticated remote code execution, which is an extremely severe risk. '
        'Normally, this finding requires a 15-day remediation subjectivity clause per Section 6 of the guidelines, or a decline. '
        'However, because this condition combines with the EDR gap, it creates multiple simultaneous decline conditions, '
        'which Section 8 dictates must result in immediate file closure with no VP referral.'
    )

    doc.add_heading('5. Risk Matrix Summary', level=1)
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Category'
    hdr_cells[1].text = 'Underwriting Finding'

    row_cells = table.add_row().cells
    row_cells[0].text = 'Loss History'
    row_cells[1].text = '$657,500 total cyber-only losses in 3 years triggers 1.20x penalty multiplier.'

    row_cells = table.add_row().cells
    row_cells[0].text = 'Ransomware Controls'
    row_cells[1].text = 'Missing EDR on 2% of endpoints triggers hard decline. Missing MFA on Contractor Remote Access triggers 1.25x penalty and caps limit at $10,000,000.'

    row_cells = table.add_row().cells
    row_cells[0].text = 'Perimeter Security'
    row_cells[1].text = 'Open Critical vulnerability CVE-2023-4528 on VPN. Combines with EDR gap to trigger Section 8 multiple simultaneous decline conditions.'

    row_cells = table.add_row().cells
    row_cells[0].text = 'Broker Target Premium'
    row_cells[1].text = 'Final calculated premium is $37,680, which is well below the broker target premium of $45,000.'

    # Save
    doc.save('Underwriting_Memo_NovaMart.docx')
    print('Generated Underwriting_Memo_NovaMart.docx successfully.')

if __name__ == '__main__':
    generate_memo()
