# -*- coding: utf-8 -*-
from docx import Document

path = 'Tri_County_ALLL_Reserve_Adequacy_Memo_Q4_2025.docx'
d = Document(path)
paras = d.paragraphs


def set_text(idx, new_text):
    p = paras[idx]
    # collapse to a single run holding the new text, preserving the first run's formatting
    if not p.runs:
        p.add_run(new_text)
        return
    p.runs[0].text = new_text
    for r in p.runs[1:]:
        r.text = ''


# 1. Byline - match the policy document's own ordering (Bank name, then division)
set_text(1, 'Prepared by: Tri-County Community Bank, Credit Risk Division')

# 2. Executive Summary shortfall % - measure against required reserve per Policy Section 6,
#    not against the booked balance
set_text(5,
    "Booked ALLL sits at $1,250,000. That leaves a shortfall of $401,606, or about 24.3% of the "
    "$1,651,606 total required reserve — well outside the ±10%-of-required-reserve band Policy "
    "Section 6 allows. An additional provision of $401,606 should go in this quarter, and the Board "
    "Audit Committee needs to see it."
)

# 3. Watch List claim - remove the unsupported "risk rating 4+" rule; only LN-002 carries an
#    explicit Watch List designation, per the Exception Tracking Log
set_text(22,
    "Of those five, only LN-002 carries an explicit \"Watch List\" designation — both of its exceptions, "
    "EXC-002 and EXC-003, record its status as \"Open — Escalated to Watch List.\" No other exception in "
    "Exception_Tracking_Log_Q4_2025.xlsx uses that language (LN-001's is tracked as \"Open — Monitoring,\" "
    "and LN-008's as \"Open — Legal Hold\"), and none of the provided source files define Watch List "
    "status as a function of Internal Risk Rating, so that designation is not extended to LN-001 or LN-008 "
    "here. Watch List status, where it appears, is an exception-tracking designation applied loan-by-loan "
    "by the relationship officer — it is a separate flag from the pool-level Q-Factor and does not itself "
    "change the pooled math. LN-004, LN-006, and LN-010 carry no Watch List designation in the exception "
    "log at all, but their breaches still pull the Pool 3 Q-Factor because that add-on applies to the "
    "whole pool, not just the specific loan that's breaching."
)

# 4. Reserve Adequacy Conclusion - same 32% -> 24.3% / required-reserve-basis fix
set_text(24,
    "$401,606 works out to about 24.3% of the $1,651,606 total required reserve, well past the "
    "±10%-of-required-reserve band Policy Section 6 allows. That makes the reserve inadequate as "
    "booked — an additional provision of $401,606 needs to go in this quarter and get reported up to "
    "the Board Audit Committee. Most of the gap traces back to two credits: LN-002, where a single loan "
    "in the Rating 5 pool is driving $703,452 of required reserve on its own, and LN-008's $340,640 "
    "specific reserve. Both reflect real, recent deterioration — a payment demand and two concurrent "
    "covenant breaches on LN-002, and a bankruptcy filing with falling collateral value on LN-008 — "
    "that the current booked balance doesn't seem to have caught up with yet."
)

# 5. Policy/Source Reconciliation - document the actual discrepancy the recomputation caught,
#    instead of asserting (falsely) that nothing conflicted
set_text(27,
    "The recomputation was not a formality — it caught a live discrepancy. Covenant_Compliance_"
    "Workbook_Q4_2025.xlsx's own Actual LTV and Actual Current Ratio columns for LN-002 show 71.5% and "
    "1.06x, both comfortably within covenant (75% max LTV, 1.00x minimum Current Ratio). Recomputing "
    "both from the underlying source files tells a different story: LTV is Outstanding Balance ($7,214,889, "
    "per Loan_Register_Q4_2025.xlsx) over Appraised Value ($8,200,000, per the August 2025 appraisal in "
    "Appraisal_and_Collateral_Summary_Q4_2025.xlsx), which is 87.9%, not 71.5% — a breach of the 75% cap. "
    "Current Ratio is Current Assets ($813,900) over Current Liabilities ($865,800), per "
    "Borrower_Financial_Statements_Summary.xlsx, which is 0.94x, not 1.06x — a breach of the 1.00x floor. "
    "Both recomputed figures match Exception Tracking Log entries EXC-002 and EXC-003, which independently "
    "record LTV at 87.9% and Current Ratio at 0.94x for this loan, confirming the workbook's pre-filled "
    "Actual columns for LN-002 are stale and the recomputed values are correct. Had the workbook's face-value "
    "figures been accepted instead, LN-002 would have appeared covenant-compliant, Pool 5 would have been "
    "assessed no Q-Factor, and the total required reserve would have been understated by roughly $18,000. "
    "DSCR came from Net Operating Income over Total Annual Debt Service in Borrower_Financial_Statements_"
    "Summary.xlsx; LTV came from the Loan Register balance divided by the appraised value in the Appraisal "
    "and Collateral Summary; Current Ratio came from Current Assets over Current Liabilities in the financial "
    "statements file. Every other loan's recomputed DSCR, LTV, and Current Ratio landed within normal "
    "rounding of the workbook's own figures — LN-002 was the only discrepancy found."
)

d.save(path)
print("Text edits saved.")

# 6. Remove Section 8 (Approval and Distribution) - never requested by the prompt's
#    seven-section structure (Executive Summary + the six sections Section 7 of the
#    policy lists after it)
d = Document(path)
paras = d.paragraphs
to_remove_idx = []
for i, p in enumerate(paras):
    if p.text.strip() == '8. Approval and Distribution':
        to_remove_idx.append(i)
    elif 'Per Policy Section 8, this memorandum must be reviewed' in p.text:
        to_remove_idx.append(i)
    elif 'Distribution per Policy Section 8' in p.text:
        to_remove_idx.append(i)

print('Removing paragraphs at indices:', to_remove_idx)
# also remove the blank paragraph immediately following the heading (index+1 of the heading, before the body)
for i in to_remove_idx:
    p = paras[i]
    p._element.getparent().remove(p._element)

# remove the trailing blank paragraph left over between the old heading and body (index 32 originally)
remaining = Document(path).paragraphs
for p in remaining:
    pass  # re-open fresh below after save

d.save(path)
print("Section 8 heading/body removed.")

# 7. Remove the signature/approval table (last table - 'Prepared by' / 'Reviewed and Approved by')
d = Document(path)
for t in d.tables:
    headers = [c.text.strip() for c in t.rows[0].cells]
    if headers == ['Role', 'Name', 'Signature', 'Date']:
        t._element.getparent().remove(t._element)
        print('Removed signature/approval table.')
        break
d.save(path)

# 8. Remove any now-orphaned blank paragraph at the very end of the doc
d = Document(path)
while d.paragraphs and not d.paragraphs[-1].text.strip():
    last = d.paragraphs[-1]
    last._element.getparent().remove(last._element)
d.save(path)

# 9. Fix docx metadata (was leaking python-docx as author)
d = Document(path)
cp = d.core_properties
cp.author = 'Credit Risk Division'
cp.last_modified_by = 'Credit Risk Division'
cp.comments = ''
d.save(path)
print("Metadata fixed.")

# --- Verify ---
d = Document(path)
print()
print("=== FINAL STRUCTURE ===")
for i, p in enumerate(d.paragraphs):
    if p.text.strip() and p.style.name.startswith('Heading'):
        print(i, '|', p.text)
print("Table count:", len(d.tables))
print("author:", d.core_properties.author)
