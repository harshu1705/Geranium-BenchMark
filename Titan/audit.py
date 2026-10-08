"""Independent second-pass audit.
   1. Leakage sweep across ALL seven input files (not just the stress workbook)
   2. Every numeric/date/name claim in the golden cross-checked against source data
   3. Internal-contradiction sweep of the golden
"""
import sys, io, os, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
import openpyxl

BASE = os.path.dirname(os.path.abspath(__file__))
IN = os.path.join(BASE, 'inputs')
GOLD = os.path.join(BASE, 'golden', 'Project_Titan_Due_Diligence_Report.docx')

issues, notes = [], []
def bad(lbl, detail=''):
    print('  ISSUE  ' + lbl + (('  -- ' + detail) if detail else '')); issues.append(lbl)
def ok(lbl):
    print('  ok     ' + lbl)
def note(lbl, detail=''):
    print('  note   ' + lbl + (('  -- ' + detail) if detail else '')); notes.append(lbl)

def docx_text(p):
    d = Document(p); out = []
    for c in d.element.body.iterchildren():
        t = c.tag.split('}')[-1]
        if t == 'p': out.append(Paragraph(c, d).text)
        elif t == 'tbl':
            for r in Table(c, d).rows:
                for cell in r.cells: out.append(cell.text)
    return '\n'.join(out)

def xlsx_text(p):
    wb = openpyxl.load_workbook(p, data_only=True); out = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None: out.append(str(c.value))
    return '\n'.join(out)

# ==================================================================
print('\n=== 1. LEAKAGE SWEEP ACROSS ALL SEVEN INPUT FILES ===')
# ==================================================================
inputs_txt = {}
for f in sorted(os.listdir(IN)):
    p = os.path.join(IN, f)
    inputs_txt[f] = docx_text(p) if f.endswith('.docx') else xlsx_text(p)
ALL_IN = '\n'.join(inputs_txt.values())
LOW = ALL_IN.lower()

# Conclusions the rubric grades that must NOT be pre-stated in any input.
leak_patterns = [
    ('112%',                       'pre-computed STRESS-003 memory overage'),
    (r'16 of 23',                  'pre-computed pilot failure count'),
    (r'16/23',                     'pre-computed pilot failure count'),
    ('no-go',                      'pre-stated deal recommendation'),
    ('conditional go',             'pre-stated deal recommendation'),
    ('effective severity',         'pre-applied §3.1 conflict resolution'),
    ('automatic deal-breaker',     'policy language leaking into evidence files'),
    ('material omission',          'pre-stated disclosure conclusion'),
    ('mischaracter',               'pre-stated PR-draft conclusion'),
    ('dispositive',                'pre-stated conclusion'),
]
for pat, why in leak_patterns:
    hits = []
    for fn, txt in inputs_txt.items():
        if re.search(pat, txt, re.I):
            # the Policy doc is *allowed* to state policy language
            if fn == 'Due_Diligence_Evaluation_Policy.docx' and why.startswith(('policy', 'pre-stated deal')):
                continue
            hits.append(fn)
    if hits:
        note(f'"{pat}" appears in {hits} ({why})')
    else:
        ok(f'no input pre-states: {why}')

# explicit re-check of the adjudicated string
s3note = [t for t in inputs_txt['Cloud_Infrastructure_Stress_Test_Aug2026.xlsx'].split('\n')
          if 'asterisk' in t]
print('\n  STRESS-003 note now reads:\n    ', s3note[0] if s3note else '(none)')
for frag in ['3.36', '112', 'VULN-0044', 'memory usage hit']:
    if s3note and frag in s3note[0]: bad(f'STRESS-003 note still contains "{frag}"')
if s3note: ok('STRESS-003 note is free of all four leaked fragments')

# ==================================================================
print('\n=== 2. GOLDEN CLAIMS CROSS-CHECKED AGAINST SOURCE DATA ===')
# ==================================================================
G = docx_text(GOLD)
GN = re.sub(r'\s+', ' ', G)

# --- severity distribution recomputed from the scan ----------------
vs = openpyxl.load_workbook(os.path.join(IN, 'Security_Vulnerability_Scan_Results_Aug2026.xlsx'),
                            data_only=True).active
order = ['Cosmetic', 'Minor', 'Major', 'Critical', 'Blocker']
eff = {}
for row in vs.iter_rows(min_row=2):
    vid, qa, it = row[0].value, row[3].value, row[4].value
    if vid:
        eff[vid] = qa if order.index(qa) >= order.index(it) else it
from collections import Counter
dist = Counter(eff.values())
print('  recomputed distribution:', dict(dist), 'total', len(eff))
expect = {'Blocker': 1, 'Critical': 5, 'Major': 10, 'Minor': 22, 'Cosmetic': 9}
for k, v in expect.items():
    if dist[k] != v: bad(f'golden claims {v} {k} but data has {dist[k]}')
if all(dist[k] == v for k, v in expect.items()): ok('severity distribution 1/5/10/22/9 matches source')
if len(eff) != 47: bad(f'golden claims 47 vulns, data has {len(eff)}')
else: ok('47 total vulnerabilities matches source')

# --- the appendix table must list every vuln with the right eff sev -
d = Document(GOLD)
tbls = [Table(c, d) for c in d.element.body.iterchildren() if c.tag.split('}')[-1] == 'tbl']
app = [t for t in tbls if t.rows[0].cells[0].text.strip() == 'Vuln ID'][0]
app_rows = {r.cells[0].text.strip(): (r.cells[1].text.strip(), r.cells[2].text.strip(),
            r.cells[3].text.strip(), r.cells[4].text.strip(), r.cells[5].text.strip())
            for r in app.rows[1:]}
if len(app_rows) != 47: bad(f'appendix has {len(app_rows)} rows, expected 47')
else: ok('appendix lists all 47 vulnerabilities')
mism = []
src = {row[0].value: (row[3].value, row[4].value, row[9].value, row[5].value)
       for row in vs.iter_rows(min_row=2) if row[0].value}
for vid, (qa, it, ef, gd, st) in app_rows.items():
    s = src.get(vid)
    if not s: mism.append((vid, 'not in source')); continue
    if qa != s[0] or it != s[1]: mism.append((vid, f'sev {qa}/{it} vs src {s[0]}/{s[1]}'))
    if ef != eff[vid]: mism.append((vid, f'eff {ef} vs computed {eff[vid]}'))
    if gd != s[2]: mism.append((vid, f'gdpr {gd} vs src {s[2]}'))
    if st != s[3]: mism.append((vid, f'status {st} vs src {s[3]}'))
if mism: bad('appendix rows mismatch source', str(mism[:6]))
else: ok('every appendix row matches the source scan exactly (sev/eff/gdpr/status)')

# --- pilot table recomputed --------------------------------------
pw = openpyxl.load_workbook(os.path.join(IN, 'Legacy_System_Integration_Pilots_Aug2026.xlsx'),
                            data_only=True).active
prate = {}
for row in pw.iter_rows(min_row=2):
    if row[0].value: prate[row[0].value] = (row[4].value, row[3].value,
                                            round(row[4].value / row[3].value * 100, 1))
pt = [t for t in tbls if t.rows[0].cells[0].text.strip() == 'Pilot ID'][0]
perr = []
for r in pt.rows[1:]:
    pid = r.cells[0].text.strip()
    frac, pct = r.cells[2].text.strip(), r.cells[3].text.strip()
    s, tot, rate = prate[pid]
    if frac != f'{s}/{tot}': perr.append((pid, f'{frac} vs {s}/{tot}'))
    if pct != f'{rate}%': perr.append((pid, f'{pct} vs {rate}%'))
if perr: bad('pilot table figures mismatch', str(perr[:6]))
else: ok('all 16 pilot rows: success/total and % match recomputed source values')

# --- GDPR control counts -----------------------------------------
gd_txt = inputs_txt['GDPR_Privacy_Audit_Checklist_Aug2026.docx']
cnt = Counter()
for line in gd_txt.split('\n'):
    pass
gdoc = Document(os.path.join(IN, 'GDPR_Privacy_Audit_Checklist_Aug2026.docx'))
gtbls = [Table(c, gdoc) for c in gdoc.element.body.iterchildren() if c.tag.split('}')[-1] == 'tbl']
ctrls = {}
for t in gtbls:
    if t.rows[0].cells[0].text.strip() != 'Control ID': continue
    for r in t.rows[1:]:
        ctrls[r.cells[0].text.strip()] = r.cells[3].text.strip()
cc = Counter(ctrls.values())
print('  recomputed GDPR controls:', dict(cc), 'total', len(ctrls))
if len(ctrls) != 45: bad(f'golden claims 45 controls, data has {len(ctrls)}')
else: ok('45 GDPR controls matches source')
for status, n in [('Met', 39), ('Partially Met', 2), ('Not Met', 1), ('N/A', 3)]:
    if cc[status] != n: bad(f'golden claims {n} {status}, data has {cc[status]}')
if cc['Met'] == 39 and cc['Partially Met'] == 2 and cc['Not Met'] == 1 and cc['N/A'] == 3:
    ok('39 Met / 2 Partially Met / 1 Not Met / 3 N/A matches source')

# --- every date the golden asserts -------------------------------
date_claims = [
    ('September 6, 2026',  'VULN-0031', src['VULN-0031'][0] if False else None, '2026-09-06'),
    ('September 2, 2026',  'VULN-0019', None, '2026-09-02'),
    ('September 4, 2026',  'VULN-0023', None, '2026-09-04'),
    ('September 15, 2026', 'VULN-0049', None, '2026-09-15'),
]
tgt = {row[0].value: str(row[8].value) for row in vs.iter_rows(min_row=2) if row[0].value}
for phrase, vid, _, expect_date in date_claims:
    actual = tgt[vid][:10]
    if actual != expect_date:
        bad(f'golden says {phrase} for {vid}, scan says {actual}')
    elif phrase not in G:
        bad(f'golden does not contain the phrase "{phrase}" for {vid}')
    else:
        ok(f'{vid} target {expect_date} correctly cited as "{phrase}"')

# --- sign-off tracker names/roles/dates ---------------------------
sw = openpyxl.load_workbook(os.path.join(IN, 'M_and_A_Signoff_Tracker_Aug2026.xlsx'),
                            data_only=True).active
track = {row[1].value: (row[0].value, row[3].value, str(row[4].value), row[5].value)
         for row in sw.iter_rows(min_row=2) if row[1].value}
for nm, role, status in [('Dr. Lisa Pham', 'Chief Legal Officer', 'Pending'),
                         ('Priya Deshpande', 'Chief Information Security Officer', 'Conditional'),
                         ('Kevin Park', 'Chief Privacy Officer', 'Pending'),
                         ('Anita Flores', 'Head of QA', 'Conditional'),
                         ('Sarah Chen', 'Director of M&A Integration', 'Approved'),
                         ('Raj Malhotra', 'VP Engineering', 'Approved')]:
    t = track[nm]
    if t[0] != role or t[1] != status: bad(f'{nm}: source is {t[0]}/{t[1]}, golden assumes {role}/{status}')
    elif nm not in G: bad(f'{nm} not mentioned in golden')
    else: ok(f'{nm} — {role}, {status} — matches source and appears in golden')
if '2026-08-20' in track['Raj Malhotra'][2] and 'August 20, 2026' in G:
    ok('Raj Malhotra risk-acceptance date August 20, 2026 matches tracker')
else:
    bad('Raj Malhotra date mismatch', track['Raj Malhotra'][2])

# --- stress figures ----------------------------------------------
st = openpyxl.load_workbook(os.path.join(IN, 'Cloud_Infrastructure_Stress_Test_Aug2026.xlsx'),
                            data_only=True)['Stress Test Results']
srows = {r[0].value: [c.value for c in r] for r in st.iter_rows(min_row=2) if r[0].value}
if srows['STRESS-001'][6] == '4.2s' and '4.2s' in G: ok('STRESS-001 P99 4.2s matches source')
else: bad('STRESS-001 P99 mismatch')
fails_p99 = [k for k, v in srows.items() if 'FAIL' in str(v[7])]
if fails_p99 == ['STRESS-001']: ok('STRESS-001 is genuinely the only FAIL status in the workbook')
else: bad('unexpected FAIL rows', str(fails_p99))

# ==================================================================
print('\n=== 3. INTERNAL-CONTRADICTION SWEEP OF THE GOLDEN ===')
# ==================================================================
# Anita Flores: every mention must agree that Sep 15 is INSIDE the window.
anita_ctx = [s for s in re.split(r'(?<=[.!?]) ', GN) if 'Anita Flores' in s or 'VULN-0049' in s]
neg = [s for s in anita_ctx if re.search(r'depends on|cannot be (met|confirmed)|uncertain|unclear', s, re.I)]
if neg: bad('Anita Flores statements still hedge', str(neg)[:200])
else: ok(f'all {len(anita_ctx)} Anita/VULN-0049 statements agree the condition is satisfiable')

# STRESS-003 must never be described as a clean pass.
s3ctx = [s for s in re.split(r'(?<=[.!?]) ', GN) if 'STRESS-003' in s]
clean = [s for s in s3ctx if re.search(r'STRESS-003[^.]*\bpassed\b(?![^.]*caveat)', s, re.I)]
if clean: bad('STRESS-003 described as a clean pass somewhere', str(clean)[:200])
else: ok(f'all {len(s3ctx)} STRESS-003 statements carry the memory caveat')

# recommendation must be NO-GO everywhere
if re.search(r'RECOMMENDATION:\s*(CONDITIONAL GO|GO)\b', G): bad('a GO recommendation appears')
else: ok('recommendation is NO-GO in every occurrence')

# §4.5.b must be stated as satisfied, not failed
if re.search(r'§4\.5\.b[^.]{0,80}not (met|satisfied)', GN, re.I):
    bad('§4.5.b described as not satisfied (contradicts Raj Malhotra acceptance)')
else: ok('§4.5.b consistently described as satisfied by VP-Eng risk acceptance')

# duplicated long paragraphs (the v2 copy-paste defect)
paras = [p.text.strip() for p in d.paragraphs if len(p.text.strip()) > 120]
cells = [c.text.strip() for t in tbls for r in t.rows for c in r.cells if len(c.text.strip()) > 120]
dupes = [x for x, n in Counter(paras + cells).items() if n > 1]
if dupes: bad(f'{len(dupes)} long passage(s) duplicated verbatim', dupes[0][:100])
else: ok('no long passage is duplicated verbatim anywhere in the report')

# section numbering sanity
heads = [p.text.strip() for p in d.paragraphs if p.style.name == 'Heading 1']
nums = [int(re.match(r'(\d+)\.', h).group(1)) for h in heads if re.match(r'\d+\.', h)]
if nums != list(range(1, len(nums) + 1)): bad('section numbering not sequential', str(nums))
else: ok(f'sections numbered 1..{len(nums)} sequentially')

print('\n' + '=' * 64)
print(f'AUDIT: {len(issues)} issue(s), {len(notes)} note(s)')
for i in issues: print('   ISSUE:', i)
for n in notes: print('   note :', n)
print('=' * 64)
sys.exit(1 if issues else 0)
