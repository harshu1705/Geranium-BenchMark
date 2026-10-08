import sys
sys.stdout.reconfigure(encoding='utf-8')
from docx import Document
import copy
import shutil
import os

BASE = r'c:\Users\user\OneDrive\Desktop\geranium project\Task3'
src = os.path.join(BASE, 'Loan_Portfolio_Review_Report_Q4_2025.docx')
dst_dir = os.path.join(BASE, 'fixed_golden')
os.makedirs(dst_dir, exist_ok=True)
dst = os.path.join(dst_dir, 'Loan_Portfolio_Review_Report_Q4_2025.docx')
shutil.copy(src, dst)

NAME_MAP = {
    'Delta Auto Group': 'Coastal Auto Group',
    'Epsilon Restaurant Group': 'Harbor Grille Restaurant Group',
    'Eta Print & Packaging Inc.': 'Palmetto Print & Packaging Inc.',
    'Iota Storage Partners': 'Keystone Storage Partners',
    'C2 Tech Campus': 'Cascade Tech Campus',
}

doc = Document(dst)

def replace_in_paragraph(p):
    text = p.text
    new_text = text
    for old, new in NAME_MAP.items():
        new_text = new_text.replace(old, new)
    if new_text != text and p.runs:
        p.runs[0].text = new_text
        for r in p.runs[1:]:
            r.text = ''

for p in doc.paragraphs:
    replace_in_paragraph(p)

for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                replace_in_paragraph(p)

# --- Rewrite Section 7 "Final Portfolio Risk Assessment" numbered list into flowing prose ---
replacements_prose = {
    '1. Covenant breaches: Nine individual covenant violations across six affected loans indicate broad-based financial stress in a subset of the portfolio. LN-008 carries simultaneous DSCR (0.81x vs. 1.20x minimum), LTV (130.9% vs. 80% maximum), and Current Ratio (0.72x vs. 1.10x minimum) failures — the most severely distressed credit in the book. LN-001 and LN-004 carry DSCR breaches that are currently being remediated. LN-002 carries concurrent LTV and Current Ratio breaches accompanied by a 45-day payment delinquency.':
        'Covenant breaches are the most immediate concern: nine individual covenant violations across six affected loans indicate broad-based financial stress in a subset of the portfolio. Palmetto Print & Packaging Inc. (LN-008) carries simultaneous DSCR (0.81x vs. 1.20x minimum), LTV (130.9% vs. 80% maximum), and Current Ratio (0.72x vs. 1.10x minimum) failures, making it the most severely distressed credit in the book. LN-001 and LN-004 carry DSCR breaches that are currently being remediated, while LN-002 carries concurrent LTV and Current Ratio breaches accompanied by a 45-day payment delinquency.',
    '2. Non-performing exposure: LN-008 (Eta Print & Packaging Inc.) is 97 days past due and in Chapter 11 bankruptcy. With collateral coverage at only 76% of outstanding principal ($830,000 appraised vs. $1,087,640 outstanding), loss exposure is material. A loss reserve is active per EXC-010.':
        'Non-performing exposure is concentrated entirely in LN-008 (Palmetto Print & Packaging Inc.), which is 97 days past due and in Chapter 11 bankruptcy. With collateral coverage at only 76% of outstanding principal ($830,000 appraised vs. $1,087,640 outstanding), loss exposure is material, and a loss reserve is already active per EXC-010.',
    '3. Watch List concentration: Three loans — LN-001 (Risk Rating 4), LN-002 (Risk Rating 5), and LN-008 (Risk Rating 6) — qualify for Watch List status under policy Section 4. Their combined outstanding balance is approximately $12.1M. LN-001 holds a temporary DSCR waiver expiring March 31, 2026; LN-002 has received a formal payment demand and carries dual covenant exceptions.':
        'Watch List concentration adds a third layer of risk: three loans — LN-001 (Risk Rating 4), LN-002 (Risk Rating 5), and LN-008 (Risk Rating 6) — qualify for Watch List status under policy Section 4, with a combined outstanding balance of approximately $12.1M. LN-001 holds a temporary DSCR waiver expiring March 31, 2026, while LN-002 has received a formal payment demand and carries dual covenant exceptions.',
    '4. Concentration risk: Commercial Mortgage loans represent 65.71% of the portfolio (policy limit 40%), the James Trotter loan-officer book represents 60.97% of total outstanding (policy limit 40%), and a single borrower — C2 Tech Campus — accounts for 15.68% of total outstanding (policy limit 15%). These structural over-concentrations amplify systemic risk if any segment experiences adverse market conditions. Portfolio rebalancing and officer-level diversification should be prioritized.':
        'Finally, structural concentration risk compounds the picture: Commercial Mortgage loans represent 65.71% of the portfolio (policy limit 40%), the James Trotter loan-officer book represents 60.97% of total outstanding (policy limit 40%), and a single borrower — Cascade Tech Campus — accounts for 15.68% of total outstanding (policy limit 15%). These structural over-concentrations amplify systemic risk if any segment experiences adverse market conditions, and portfolio rebalancing and officer-level diversification should be prioritized.',
}

changed = 0
for p in doc.paragraphs:
    t = p.text
    if t in replacements_prose:
        new_t = replacements_prose[t]
        if p.runs:
            p.runs[0].text = new_t
            for r in p.runs[1:]:
                r.text = ''
        changed += 1
print(f"Section 7 prose rewrites applied: {changed}")

doc.save(dst)
print("Saved intermediate golden docx with names fixed + Section 7 de-listed.")
