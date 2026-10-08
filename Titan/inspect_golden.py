import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

p = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'golden', 'Project_Titan_Due_Diligence_Report.docx')
d = Document(p)
print('STYLES:', [s.name for s in d.styles if s.type is not None][:40])
pi = 0
for child in d.element.body.iterchildren():
    tag = child.tag.split('}')[-1]
    if tag == 'p':
        para = Paragraph(child, d)
        if para.text.strip():
            runs = [(r.text[:45], r.bold, r.italic, r.font.size, r.font.color.rgb if r.font.color and r.font.color.type is not None else None) for r in para.runs]
            print(f"P{pi} [{para.style.name}] nruns={len(para.runs)} {runs}")
        pi += 1

# table cell run formatting sample
ti = 0
for child in d.element.body.iterchildren():
    if child.tag.split('}')[-1] == 'tbl':
        t = Table(child, d)
        print(f"\nTABLE {ti}: style={t.style.name if t.style else None} rows={len(t.rows)} cols={len(t.columns)}")
        for ri in (0, 1):
            if ri < len(t.rows):
                for ci, c in enumerate(t.rows[ri].cells):
                    for para in c.paragraphs:
                        rr = [(r.text[:30], r.bold, r.font.size) for r in para.runs]
                        print(f"   R{ri}C{ci} [{para.style.name}] {rr}")
        ti += 1
