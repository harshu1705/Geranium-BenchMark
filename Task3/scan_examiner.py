from docx import Document
d = Document('check_golden/Loan_Portfolio_Review_Report_Q4_2025.docx')
for i,p in enumerate(d.paragraphs):
    t = p.text
    low = t.lower()
    if 'examin' in low or 'department of banking' in low or 'report of examination' in low:
        print(i, repr(t))
