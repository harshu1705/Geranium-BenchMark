import os
from docx import Document

doc_path = r'c:\Users\user\OneDrive\Desktop\geranium project\NovaMart_extracted\Golden\Underwriting_Memo_NovaMart.docx'
doc = Document(doc_path)

full_text = []
for p in doc.paragraphs:
    full_text.append(p.text)

full_text_str = '\n'.join(full_text)

# We want to verify that the numbers are intact.
required_strings = [
    '$502,400,000',
    '$25,120',
    '$657,500',
    '$37,680',
    'CVE-2023-4528',
    'EDR',
    'Section 8',
    'immediate file closure'
]

print("--- REGRESSION TEST ---")
all_passed = True
for req in required_strings:
    if req in full_text_str:
        print(f"PASS: Found {req}")
    else:
        print(f"FAIL: Missing {req}")
        all_passed = False

if all_passed:
    print("All strings verified successfully.")
else:
    print("Regression test failed.")

