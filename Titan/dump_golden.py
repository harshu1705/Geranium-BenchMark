import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

p = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'golden', 'Project_Titan_Due_Diligence_Report.docx')
d = Document(p)
ti = 0
pi = 0
for child in d.element.body.iterchildren():
    tag = child.tag.split('}')[-1]
    if tag == 'p':
        para = Paragraph(child, d)
        if para.text.strip():
            print(f"P{pi} [{para.style.name}] {para.text}")
        pi += 1
    elif tag == 'tbl':
        t = Table(child, d)
        print(f"\n=== TABLE {ti} ===")
        for ri, r in enumerate(t.rows):
            cells = [c.text.replace("\n", " / ") for c in r.cells]
            print(f"  R{ri}: " + " | ".join(cells))
        print(f"=== END TABLE {ti} ===\n")
        ti += 1
