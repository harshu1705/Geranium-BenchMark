import docx
from docx import Document
from docx.shared import Pt

d = Document('Dun_and_Bradstreet_Credit_Report.docx')

def find_para(doc, text_start):
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().startswith(text_start):
            return i
    return None

def style_header_row(table):
    for cell in table.rows[0].cells:
        for p in cell.paragraphs:
            for run in p.runs:
                run.bold = True

def add_table_after(doc, anchor_paragraph, headers, rows):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Light Grid Accent 1'
    for i, h in enumerate(headers):
        table.rows[0].cells[i].text = h
    for row in rows:
        cells = table.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = str(val)
    style_header_row(table)
    anchor_paragraph._p.addnext(table._tbl)
    return table

# Remove the old prose paragraphs we are replacing with tables, keep intro sentences.
paras = d.paragraphs

# 1) Risk Summary & Predictive Scores -> score table right after the section heading
heading_idx = find_para(d, 'Risk Summary & Predictive Scores')
heading_para = paras[heading_idx]
add_table_after(d, heading_para, ['Score', 'Current Value', 'Prior Period Value', 'Industry Benchmark'],
                 [
                     ['D&B Paydex Score', 68, '74 (Q1 2026)', '75 (Industry Median)'],
                     ['Financial Stress Score', 1450, '—', '—'],
                     ['Financial Stress Percentile (1=highest risk, 100=lowest risk)', 32, '—', '—'],
                 ])

# 2) Payment Habits Analysis -> current vs prior table
heading_idx2 = find_para(d, 'Payment Habits Analysis')
heading_para2 = paras[heading_idx2]
# the paragraph right after the heading is the intro sentence about 14 trade experiences; keep it,
# insert table after the paragraph that lists the bullet breakdown
intro_idx = None
for i in range(heading_idx2, len(paras)):
    if paras[i].text.strip().startswith('Based on 14 trade experiences'):
        intro_idx = i
        break
add_table_after(d, paras[intro_idx],
                 ['Payment Bucket', 'Trailing 12 Months (Current)', 'Prior 12-Month Period'],
                 [
                     ['Within Terms', '42%', '75%'],
                     ['1 to 30 Days Slow', '38%', '—'],
                     ['31 to 60 Days Slow', '15%', '—'],
                     ['61+ Days Slow', '5%', '5%'],
                 ])

# 3) Select Trade Lines -> table
heading_idx3 = find_para(d, 'Select Trade Lines')
heading_para3 = paras[heading_idx3]
add_table_after(d, heading_para3,
                 ['Trade Line', 'High Credit', 'Terms', 'Recent Payment Experience'],
                 [
                     ['Raw Materials Supplier', '$650,000', 'Net 30', 'Paying 15-30 days beyond terms'],
                     ['Packaging Supplier', '$250,000', 'Net 30', 'Paying on time or up to 10 days slow'],
                     ['Logistics/Freight', '$120,000', 'Net 15', 'Paying on time'],
                     ['Equipment Leasing', '$45,000', 'Net 30', 'Paying 30+ days slow'],
                 ])

# 4) Public Records & Legal Filings -> table
heading_idx4 = find_para(d, 'Public Records & Legal Filings')
heading_para4 = paras[heading_idx4]
add_table_after(d, heading_para4,
                 ['Record Type', 'Count / Status', 'Details'],
                 [
                     ['Bankruptcies', '0 recorded', '—'],
                     ['Judgments', '0 recorded', '—'],
                     ['Liens', '1 active record found',
                      'Filed 7/12/2026; Filer: State of Ohio Dept. of Taxation; Type: State Tax Lien; '
                      'Amount: $45,000; Status: Unresolved / Open'],
                 ])

d.save('Dun_and_Bradstreet_Credit_Report.docx')
print('saved')
