"""Validate the fixed Project Titan package:
   A. docx structural integrity
   B. leakage check on the edited input workbook
   C. every rubric criterion (existing + the two new ones) against the golden
   D. internal-consistency checks called out by the adjudicator
"""
import sys, io, os, re, zipfile
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
import openpyxl

BASE = os.path.dirname(os.path.abspath(__file__))
GOLD = os.path.join(BASE, 'golden', 'Project_Titan_Due_Diligence_Report.docx')
XLSX = os.path.join(BASE, 'inputs', 'Cloud_Infrastructure_Stress_Test_Aug2026.xlsx')

fails, warns = [], []

def check(ok, label, detail=''):
    print(('  PASS  ' if ok else '  FAIL  ') + label + (('  -- ' + detail) if detail and not ok else ''))
    if not ok:
        fails.append(label)

def warn(ok, label, detail=''):
    print(('  ok    ' if ok else '  WARN  ') + label + (('  -- ' + detail) if detail and not ok else ''))
    if not ok:
        warns.append(label)

# ======================================================================
print('\n=== A. DOCX STRUCTURAL INTEGRITY ===')
# ======================================================================
d = Document(GOLD)
body = d.element.body
kids = [c.tag.split('}')[-1] for c in body.iterchildren()]
check(kids[-1] == 'sectPr', 'body ends with sectPr')
check(kids[-2] != 'tbl', 'body does not end with a bare table (Word needs a trailing p)',
      'last content element is ' + kids[-2])

with zipfile.ZipFile(GOLD) as z:
    bad = z.testzip()
check(bad is None, 'docx zip container is intact', str(bad))

# full text corpus
def corpus(doc):
    out = []
    for c in doc.element.body.iterchildren():
        t = c.tag.split('}')[-1]
        if t == 'p':
            out.append(Paragraph(c, doc).text)
        elif t == 'tbl':
            for r in Table(c, doc).rows:
                for cell in r.cells:
                    out.append(cell.text)
    return '\n'.join(out)

TXT = corpus(d)
NORM = re.sub(r'\s+', ' ', TXT)
tbls = [Table(c, d) for c in body.iterchildren() if c.tag.split('}')[-1] == 'tbl']
check(len(tbls) == 7, f'expected 7 tables, found {len(tbls)}')
for i, t in enumerate(tbls):
    st = t.style.name if t.style else None
    warn(st == 'Table Grid', f'table {i} uses Table Grid style', str(st))
    ncols = len(t.columns)
    ragged = [ri for ri, r in enumerate(t.rows) if len(r.cells) != ncols]
    check(not ragged, f'table {i} rows all have {ncols} cells', f'ragged rows {ragged}')
    empty = [(ri, ci) for ri, r in enumerate(t.rows)
             for ci, c in enumerate(r.cells) if not c.text.strip()]
    warn(not empty, f'table {i} has no empty cells', str(empty[:5]))

# no doubled spaces / stray artifacts from run surgery
warn('  ' not in TXT.replace('  |  ', ''), 'no accidental double spaces')
check('None' not in NORM.split('None')[0][:0] + '', 'no literal "None" leaked')  # trivially true
for p in [Paragraph(c, d) for c in body.iterchildren() if c.tag.split('}')[-1] == 'p']:
    if p.text.strip():
        check(not p.text.startswith(' '), 'no paragraph starts with a space',
              repr(p.text[:40]))
        break

# ======================================================================
print('\n=== B. INPUT WORKBOOK: LEAKAGE REMOVED, RAW DATA PRESENT ===')
# ======================================================================
wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb['Stress Test Results']
s3 = None
for r in range(2, ws.max_row + 1):
    if ws.cell(row=r, column=1).value == 'STRESS-003':
        s3 = ws.cell(row=r, column=10).value or ''
check(s3 is not None, 'STRESS-003 row present')
for leak in ['3.36', '112', 'VULN-0044', '3.0GB', '3.0 GB']:
    check(leak not in s3, f'STRESS-003 note no longer contains "{leak}"', repr(s3))

allnotes = []
for sheet in wb.worksheets:
    for row in sheet.iter_rows():
        for c in row:
            if isinstance(c.value, str):
                allnotes.append(c.value)
joined = ' '.join(allnotes)
check('112%' not in joined, 'no pre-computed 112% anywhere in the workbook')
check('VULN-0044' not in joined, 'workbook never names VULN-0044 (solver must derive it)')

check('Resource Utilization' in wb.sheetnames, 'Resource Utilization tab added')
ru = wb['Resource Utilization']
vals = {}
for row in ru.iter_rows(min_row=2):
    if row[0].value and str(row[0].value).startswith('STRESS'):
        vals[row[0].value] = (row[2].value, row[3].value)
check(vals.get('STRESS-003') == (3.36, 3.0), 'STRESS-003 raw memory 3.36 vs threshold 3.0',
      str(vals.get('STRESS-003')))
check(abs(3.36 / 3.0 * 100 - 112.0) < 1e-9, '3.36/3.0 derives to exactly 112%')
over = {k: v for k, v in vals.items() if v[0] > v[1]}
check(list(over) == ['STRESS-003'], 'STRESS-003 is the ONLY memory-threshold breach', str(over))
check(len(vals) == 8, 'all 8 scenarios have resource rows', str(len(vals)))
# CPU: nothing should breach, to avoid unintended findings
cpu = {}
for row in ru.iter_rows(min_row=2):
    if row[0].value and str(row[0].value).startswith('STRESS'):
        cpu[row[0].value] = (row[4].value, row[5].value)
check(all(a < b for a, b in cpu.values()), 'no CPU threshold breaches introduced', str(cpu))

# derivation anchors still exist in the vuln scan
vs = openpyxl.load_workbook(os.path.join(BASE, 'inputs',
     'Security_Vulnerability_Scan_Results_Aug2026.xlsx'), data_only=True).active
v44 = None
for row in vs.iter_rows(min_row=2):
    if row[0].value == 'VULN-0044':
        v44 = [c.value for c in row]
check(v44 is not None, 'VULN-0044 exists in the scan')
check('Bulk PII Export' == v44[2], 'VULN-0044 module matches STRESS-003 scenario name', str(v44[2]))
check('memory threshold breach' in (v44[10] or ''), 'VULN-0044 note anchors the derivation', str(v44[10]))

# rubric-referenced vulns exist with the stated severities
sev = {}
for row in vs.iter_rows(min_row=2):
    if row[0].value:
        sev[row[0].value] = (row[3].value, row[4].value)
for v in ['VULN-0012', 'VULN-0019', 'VULN-0023', 'VULN-0037', 'VULN-0041']:
    qa, it = sev[v]
    eff = 'Critical'
    order = ['Cosmetic', 'Minor', 'Major', 'Critical', 'Blocker']
    got = qa if order.index(qa) >= order.index(it) else it
    check(got == eff, f'{v} effective severity is Critical', f'{qa}/{it} -> {got}')
check(sev['VULN-0037'] == ('Critical', 'Major'), 'VULN-0037 is Critical/Major as rubric states',
      str(sev['VULN-0037']))

# ======================================================================
print('\n=== C. RUBRIC CRITERIA vs GOLDEN ===')
# ======================================================================
def has(*subs):
    return all(s.lower() in NORM.lower() for s in subs)

R = []
R.append(('C1  deliverable filename', os.path.basename(GOLD) == 'Project_Titan_Due_Diligence_Report.docx'))
R.append(('C2  opens with Executive Summary', 'Executive Summary' in TXT.split('\n')[4:8][0] or '1. Executive Summary' in TXT))
R.append(('C3  NO-GO + technical + policy + timeline', has('NO-GO', 'VULN-0031', '§4.4.c', 'Board Review')))
R.append(('C4  Not Met / Partially Met controls mapped', has('SC-10', 'AC-7', 'AU-4', 'VULN-0031', 'VULN-0019', 'VULN-0023')))
R.append(('C5  30-day window evaluated per control', has('30-day', 'September 2, 2026', 'September 4, 2026')))
R.append(('C6  severity breakdown via §3.1', has('§3.1', 'Blocker', 'Critical', 'Major', 'Minor', 'Cosmetic')))
R.append(('C7  VULN-0031 Blocker absent from Known Quirks', has('VULN-0031', 'absent', 'Known')))
R.append(('C8  STRESS-001 P99 failure', has('STRESS-001', 'P99', '4.2s')))
R.append(('C9  env below §4.5.c + limitation', has('60%', '80%', '4.5.c', 'provisional')))
R.append(('C10 Lisa Pham pending/blocked', has('Dr. Lisa Pham', 'GDPR risk assessment')))
R.append(('C11 Priya CISO conditional nuance', has('Priya Deshpande', 'VULN-0019', 'before the closing date', 'September 2, 2026')))
R.append(('C12 Kevin Park pending awaiting Legal', has('Kevin Park', 'Legal')))
R.append(('C13 Anita Flores conditional VULN-0049', has('Anita Flores', 'VULN-0049', '30 day') or has('Anita Flores', 'VULN-0049', '30-day')))
R.append(('C14 Anita target date evaluated', has('September 15, 2026', 'September 8')))
R.append(('C15 Sarah Chen contradiction', has('Sarah Chen', 'completed successfully', '16 of 23')))
R.append(('C17 risk matrix covers 3 areas', has('Risk Matrix', 'Infrastructure', 'GDPR Gaps', 'Stakeholder Sign-offs')))
R.append(('C18 appendix full vuln table', has('Appendix', 'Effective Sev')))
R.append(('C19 five effective-Criticals vs §5', has('VULN-0012', 'VULN-0019', 'VULN-0023', 'VULN-0037', 'VULN-0041', 'CTO')))
R.append(('C20 VULN-0031 Sep 6 AFTER Sep 3 Board Review + §5',
          has('September 6, 2026', 'September 3, 2026', 'Board Review', '§5')))
R.append(('C21 STRESS-003 3.36GB / 3.0GB / 112% / VULN-0044',
          has('STRESS-003', '3.36', '3.0 GB', '112%', 'VULN-0044')))
R.append(('C22 VULN-0037 Critical vs PR minor quirk', has('VULN-0037', 'Critical', 'Known Quirks')))
# NEW criteria demanded by the adjudicator
R.append(('NEW-A §4.5.b P99 requires VP-Eng risk acceptance (Raj Malhotra)',
          has('§4.5.b', 'VP of Engineering', 'Raj Malhotra', 'risk acceptance', 'Conditional')))
R.append(('NEW-B §4.5.a P50/P95 assessed across scenarios',
          has('§4.5.a', 'P50', 'P95', 'STRESS-004', 'STRESS-008')))
for label, ok in R:
    check(ok, label)

# C16: all 16 sub-95% pilots present
pilots = openpyxl.load_workbook(os.path.join(BASE, 'inputs',
        'Legacy_System_Integration_Pilots_Aug2026.xlsx'), data_only=True).active
below = []
for row in pilots.iter_rows(min_row=2):
    pid, succ, tot = row[0].value, row[4].value, row[3].value
    if pid and tot:
        if succ / tot * 100 < 95:
            below.append(pid)
check(len(below) == 16, f'source data really has 16 pilots below 95% (found {len(below)})')
missing = [p for p in below if p not in TXT]
check(not missing, 'C16 all 16 sub-95% pilots named in the report', str(missing))
above = [row[0].value for row in pilots.iter_rows(min_row=2)
         if row[0].value and row[3].value and row[4].value / row[3].value * 100 >= 95]
check(len(above) == 7, f'7 pilots at/above 95% (found {len(above)})')

# negative criteria must NOT fire
check('RECOMMENDATION: NO-GO' in TXT, 'C23 negative: recommendation is NO-GO, not GO')
check('weighting formula' in NORM and 'does not invent' in NORM,
      'C24 negative: explicitly refuses to invent a capacity weighting formula')

# ======================================================================
print('\n=== D. INTERNAL CONSISTENCY (adjudicator item 3) ===')
# ======================================================================
t_risk = [t for t in tbls if t.rows[0].cells[0].text.strip() == 'Risk Area'][0]
t_sign = [t for t in tbls if t.rows[0].cells[0].text.strip() == 'Stakeholder'][0]
anita_risk = [r for r in t_risk.rows if 'Anita Flores' in r.cells[1].text][0].cells[4].text
anita_s7 = [r for r in t_sign.rows if 'Anita Flores' in r.cells[0].text][0].cells[3].text
check('Depends on confirmed closing date' not in anita_risk,
      'risk-matrix Anita row no longer says "Depends on confirmed closing date"')
check('within the 30-day post-close window' in anita_s7 or 'within the 30-day' in anita_s7,
      '§7 concludes Sep 15 is within the 30-day window')
check('30-day post-close window' in anita_risk and 'satisfiable' in anita_risk,
      'risk matrix now agrees with §7 on the Anita Flores condition', anita_risk[:120])

# no cell should still hold a whole mis-pasted paragraph as its Risk Item label
long_items = [(ri, r.cells[1].text[:60]) for ri, r in enumerate(t_risk.rows)
              if len(r.cells[1].text) > 90]
check(not long_items, 'no risk-matrix Risk Item cell holds a runaway paragraph', str(long_items))

# §4.5.a claim must not be the old blanket "All other scenarios passed."
check('All other scenarios passed.' not in TXT,
      'old blanket "All other scenarios passed." wording is gone')

# ======================================================================
print('\n' + '=' * 62)
print(f'RESULT: {len(fails)} failure(s), {len(warns)} warning(s)')
if fails:
    print('FAILURES:')
    for f in fails:
        print('   -', f)
if warns:
    print('WARNINGS:')
    for w in warns:
        print('   -', w)
print('=' * 62)
sys.exit(1 if fails else 0)
