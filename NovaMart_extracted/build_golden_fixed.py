from docx import Document

path = 'Golden/Underwriting_Memo_NovaMart.docx'
d = Document(path)
paras = d.paragraphs

# --- Paragraph 7: Executive Summary Justification ---
# Remove the framing that treats the open Critical CVE as an equal, automatic
# "decline condition" alongside EDR. Per Section 6, a Critical CVE independently
# permits EITHER decline OR a 15-day remediation subjectivity - it is not itself
# an unconditional mandatory decline the way the EDR gap (Section 5) is.
paras[7].runs[0].text = (
    'Summary Justification: This risk represents a hard decline. The applicant is missing Endpoint '
    'Detection & Response (EDR) on 2% of their endpoints (legacy barcode scanners) per the Ransomware '
    'Supplemental Questionnaire. Section 5 of the Underwriting Guidelines explicitly states that EDR is '
    '"Required for all endpoints. If missing, Decline," with no discretion to substitute a subjectivity '
    'clause or grace period. This EDR gap is, on its own, a mandatory, non-discretionary decline trigger. '
    'Separately, the perimeter scan identified an open Critical vulnerability (CVE-2023-4528) on the VPN '
    'appliance; under Section 6 that finding would independently permit either a decline or a 15-day '
    'remediation subjectivity, but because the file is already declined on the EDR trigger, that choice is '
    'moot. Per Section 8, no VP referral is required to close this file. The pricing analysis and further '
    'underwriting review completed below are strictly for file documentation and auditing purposes.'
)

# --- Paragraph 12: Loss Run Surcharge rationale ---
# Fix to $650,000 (CLM-25-1102 only), explicitly excluding CLM-23-4412 as outside the
# 3-year lookback window measured from the Oct 1, 2026 effective date (window = Oct 1, 2023 - present).
paras[12].runs[0].text = (
    'Rationale: The Loss Run History reveals one Cyber-classified claim within the 3-year lookback window: '
    'a 2025 ransomware claim (CLM-25-1102) totaling $650,000 paid. The 3-year window is measured from the '
    'October 1, 2026 desired effective date, meaning it runs from October 1, 2023 to present. The other '
    'Cyber-classified claim on file, CLM-23-4412 (a lost encrypted laptop dated January 15, 2023, totaling '
    '$7,500), falls before October 1, 2023 and is therefore excluded from the 3-year total. Per the revised '
    'Section 4 of the Underwriting Guidelines, the tiering calculation uses Cyber line-of-business claims '
    'only; non-cyber claims (e.g., the General Liability, Property, and Auto claims also in the loss run) '
    'are excluded regardless of date. Section 4 states that Cyber claims totaling $500,000 or more in the '
    'last 3 years require a 1.20x penalty multiplier. Because the in-window cyber-only total ($650,000) '
    'exceeds the $500,000 threshold, it warrants the maximum 1.20x loss penalty.'
)

# --- Paragraph 19: Eligible Limit rationale (keep, minor wording unchanged) - no edit needed ---

# --- Paragraph 20 & 21: Authority Required ---
paras[20].runs[0].text = 'Authority Required: Senior Underwriter (Moot due to Decline)'
paras[21].runs[0].text = (
    'Rationale: Both the requested $15,000,000 limit and the $10,000,000 eligible cap exceed the '
    '$5,000,000 Underwriter I authority limit under Section 7, so either would require Senior Underwriter '
    'authority to bind. However, the file is being declined outright on the mandatory EDR trigger under '
    'Section 5, so this authority determination is moot - no bind is being executed. Per Section 8 of the '
    'Underwriting Guidelines, no VP referral is required to close a file on a mandatory decline condition.'
)

# --- Paragraph 24: Subjectivity Analysis - add CVE-2022-40684, fix "multiple simultaneous decline" framing ---
paras[24].runs[0].text = (
    'Rationale: The Third-Party Security Scan identified an Open Critical vulnerability, CVE-2023-4528 '
    '(Fortinet FortiOS Heap Buffer Overflow, CVSS 9.6), on the VPN appliance (vpn.novamart.com). This '
    'vulnerability allows unauthenticated remote code execution, which is an extremely severe risk. '
    'Normally, this finding requires either a decline or a 15-day remediation subjectivity clause per '
    'Section 6 of the guidelines. The scan history also shows two additional Critical vulnerabilities that '
    'have already been remediated and require no further underwriting action: CVE-2024-3094 (XZ Utils '
    'Backdoor on SSH/port 22, CVSS 10) on web1.novamart.com, and CVE-2022-40684 (FortiOS Authentication '
    'Bypass, CVSS 9.6), also on the VPN appliance (vpn.novamart.com, port 10443). Because the file is '
    'already being declined on the separate, mandatory EDR trigger under Section 5, the Section 6 choice '
    'between decline and subjectivity for the still-open CVE-2023-4528 is moot - the file is declined.'
)

# --- Risk Matrix table ---
table = d.tables[0]
rows = table.rows
# Row 1: Loss History
rows[1].cells[1].text = '$650,000 in-window Cyber-only losses (CLM-25-1102) in the 3-year lookback triggers a 1.20x penalty multiplier.'
# Row 3: Perimeter Security
rows[3].cells[1].text = ('Open Critical vulnerability CVE-2023-4528 on VPN (decline-or-15-day-subjectivity choice per Section 6, '
                          'moot due to separate EDR decline). CVE-2024-3094 and CVE-2022-40684 also Critical but Remediated.')

d.save(path)
print('saved golden')

# --- Fix metadata ---
d2 = Document(path)
cp = d2.core_properties
cp.author = 'Sarah Thompson'
cp.last_modified_by = 'Sarah Thompson'
cp.comments = ''
d2.save(path)
print('metadata fixed')

# --- Verify ---
d3 = Document(path)
for i, p in enumerate(d3.paragraphs):
    if p.text.strip():
        print(i, p.text)
        print()
print('--- TABLE ---')
for r in d3.tables[0].rows:
    for c in r.cells:
        print(repr(c.text))
    print('--')
print('author:', d3.core_properties.author, '| comments:', repr(d3.core_properties.comments))
