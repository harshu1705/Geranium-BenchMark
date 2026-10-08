"""
Fix script to address reviewer feedback on the Snorkel task submission.

Fixes:
1. Broker email - put dollar amounts back ($45,000 and $15,000,000)
2. Golden solution - flip the competitiveness sentence ($37,680 is UNDER $45,000 target = good)
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')

from docx import Document
import os
import copy

BASE = r'c:\Users\user\OneDrive\Desktop\geranium project\NovaMart_extracted'

###############################################################################
# FIX 1: Broker Email - Restore dollar amounts
###############################################################################
print("=" * 60)
print("FIX 1: Fixing Broker Email dollar amounts...")
print("=" * 60)

broker_path = os.path.join(BASE, 'Inputs', 'Broker_Email_Correspondence.docx')
doc = Document(broker_path)

for i, para in enumerate(doc.paragraphs):
    text = para.text
    changed = False

    # Fix: "requested a  limit" → "requested a $15,000,000 limit"
    if 'requested a  limit' in text or 'requested a\xa0 limit' in text:
        for run in para.runs:
            if 'requested a  limit' in run.text:
                run.text = run.text.replace('requested a  limit', 'requested a $15,000,000 limit')
                changed = True
            elif 'requested a' in run.text and run.text.strip().endswith('requested a'):
                run.text = run.text.replace('requested a', 'requested a $15,000,000')
                changed = True

    # Fix: "a ,000 quote" → "a $45,000 quote"
    if ',000 quote' in text:
        for run in para.runs:
            if ',000 quote' in run.text:
                run.text = run.text.replace(',000 quote', '$45,000 quote')
                changed = True
            elif 'a ,000' in run.text:
                run.text = run.text.replace('a ,000', 'a $45,000')
                changed = True

    # Fix: "that ,000 mark" → "that $45,000 mark"
    if ',000 mark' in text:
        for run in para.runs:
            if ',000 mark' in run.text:
                run.text = run.text.replace(',000 mark', '$45,000 mark')
                changed = True
            elif 'that ,000' in run.text:
                run.text = run.text.replace('that ,000', 'that $45,000')
                changed = True

    # Fix: "offer the full  limit" → "offer the full $15,000,000 limit"
    if 'the full  limit' in text or 'the full\xa0 limit' in text:
        for run in para.runs:
            if 'the full  limit' in run.text:
                run.text = run.text.replace('the full  limit', 'the full $15,000,000 limit')
                changed = True
            elif 'full ' in run.text and run.text.strip().endswith('full'):
                run.text = run.text + ' $15,000,000'
                changed = True

    if changed:
        print(f"  Fixed paragraph {i}")

# Alternative approach: if run-level replacement didn't work, 
# reconstruct the paragraphs with correct text
# Let's just do full text check and fix
for i, para in enumerate(doc.paragraphs):
    full_text = para.text
    if any(x in full_text for x in [',000 quote', ',000 mark', 'requested a  limit', 'the full  limit',
                                      'a ,000 quote', 'a ,000']):
        # The run-level fix didn't fully work, let's do a more aggressive approach
        # Merge all runs into one and replace
        if len(para.runs) > 0:
            merged = full_text
            merged = merged.replace('a ,000 quote', 'a $45,000 quote')
            merged = merged.replace(',000 quote', '$45,000 quote')
            merged = merged.replace('that ,000 mark', 'that $45,000 mark')
            merged = merged.replace(',000 mark', '$45,000 mark')
            merged = merged.replace('requested a  limit', 'requested a $15,000,000 limit')
            merged = merged.replace('the full  limit', 'the full $15,000,000 limit')
            
            if merged != full_text:
                # Clear all runs except first, put merged text in first run
                para.runs[0].text = merged
                for run in para.runs[1:]:
                    run.text = ''
                print(f"  Force-fixed paragraph {i}: {merged[:80]}...")

doc.save(broker_path)
print("  Saved Broker_Email_Correspondence.docx")

# Verify
doc2 = Document(broker_path)
print("\n  VERIFICATION - Broker Email paragraphs:")
for i, para in enumerate(doc2.paragraphs):
    if para.text.strip():
        print(f"  P{i}: {para.text[:120]}...")

###############################################################################
# FIX 2: Golden Solution - Flip competitiveness sentence
###############################################################################
print("\n" + "=" * 60)
print("FIX 2: Fixing Golden Solution competitiveness sentence...")
print("=" * 60)

golden_path = os.path.join(BASE, 'Golden', 'Underwriting_Memo_NovaMart.docx')
doc = Document(golden_path)

for i, para in enumerate(doc.paragraphs):
    text = para.text

    # Fix: "significantly misses the broker's target premium of $45,000" 
    # → "comes in well below the broker's target premium of $45,000"
    # Also fix: "the account is uncompetitive and would face heavy pricing resistance"
    # → "the premium is competitive on price, though the decline conditions prevent quoting"
    
    if 'significantly misses' in text or 'uncompetitive' in text:
        for run in para.runs:
            if 'significantly misses' in run.text:
                run.text = run.text.replace(
                    'This significantly misses the broker\'s target premium of $45,000.',
                    'This comes in well below the broker\'s target premium of $45,000, meaning the account is price-competitive.'
                )
                print(f"  Fixed 'significantly misses' in paragraph {i}")
            if 'uncompetitive' in run.text:
                run.text = run.text.replace(
                    'Even if the decline conditions were resolved, the account is uncompetitive and would face heavy pricing resistance from the insured.',
                    'If the decline conditions were resolved, the premium would actually beat the broker\'s target, making the account competitive on price.'
                )
                print(f"  Fixed 'uncompetitive' in paragraph {i}")
    
    # Also check the Risk Matrix table for the same issue
    
# Fix tables too
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            if 'exceeds the broker target' in cell.text or 'misses' in cell.text.lower():
                old = cell.text
                # The risk matrix row says premium exceeds target
                cell.text = cell.text.replace(
                    'which exceeds the broker target premium of $45,000.',
                    'which is well below the broker target premium of $45,000.'
                )
                if cell.text != old:
                    print(f"  Fixed table cell: {cell.text[:80]}...")

doc.save(golden_path)
print("  Saved Underwriting_Memo_NovaMart.docx")

# Verify golden
doc3 = Document(golden_path)
print("\n  VERIFICATION - Golden Solution paragraphs:")
for i, para in enumerate(doc3.paragraphs):
    if para.text.strip():
        print(f"  P{i}: {para.text[:150]}...")
print("\n  Tables:")
for table in doc3.tables:
    for row in table.rows:
        cells = [cell.text[:80] for cell in row.cells]
        print(f"  {cells}")

print("\n" + "=" * 60)
print("ALL FIXES APPLIED SUCCESSFULLY")
print("=" * 60)
