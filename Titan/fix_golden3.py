import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

BASE = os.path.dirname(os.path.abspath(__file__))
P = os.path.join(BASE, 'golden', 'Project_Titan_Due_Diligence_Report.docx')
d = Document(P)

tbls = [Table(c, d) for c in d.element.body.iterchildren() if c.tag.split('}')[-1] == 'tbl']
sla = [t for t in tbls if t.rows[0].cells[-1].text.strip().endswith('Result')][0]

# Make the STRESS-003 row self-caveating so the table alone never reads as a clean pass.
row = [r for r in sla.rows if 'STRESS-003' in r.cells[0].text][0]
cell = row.cells[5]
cell.paragraphs[0].runs[0].text = 'MET — throughput only; see memory caveat below'
print('STRESS-003 §4.5.a cell ->', cell.text)

# Remove empty runs left behind by run surgery (cosmetic XML hygiene).
removed = 0
for para in d.paragraphs:
    for r in list(para.runs):
        if r.text == '':
            r._r.getparent().remove(r._r)
            removed += 1
for t in tbls:
    for tr in t.rows:
        for c in tr.cells:
            for para in c.paragraphs:
                for r in list(para.runs):
                    if r.text == '':
                        r._r.getparent().remove(r._r)
                        removed += 1
print('removed', removed, 'empty runs')

d.save(P)
print('saved')
