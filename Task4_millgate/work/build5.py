import csv, datetime, json
from collections import defaultdict
from fractions import Fraction as F

EFFECTIVE = datetime.date(2027,1,1)

def months_between(d1, d2):
    return (d2.year-d1.year)*12 + (d2.month-d1.month)

def parse(d):
    y,m,dd = d.split('-')
    return datetime.date(int(y),int(m),int(dd))

def round_to(x, unit):
    x = F(x)
    q = x / unit
    fl = q.__floor__()
    frac = q - fl
    if frac >= F(1,2):
        return (fl+1)*unit
    return fl*unit

def round_hundred(x):
    return round_to(x, 100)

def round_up_hundred(x):
    x = F(x)
    q = x/100
    fl = q.__floor__()
    if q == fl:
        return fl*100
    return (fl+1)*100

pts = {}
with open('inputs_extract/Job_Evaluation_Point_Schedule.csv') as f:
    for row in csv.DictReader(f):
        pts[row['job_code']] = (row['job_title'], int(row['evaluation_points']))

items = sorted(pts.items(), key=lambda kv: kv[1][1])
lo = (items[0][1][1]//10)*10
bands=[]
start=lo
while start<=items[-1][1][1]:
    bands.append((start,start+49))
    start+=50
n_grades=len(bands)

def grade_of(p):
    for i,(a,b) in enumerate(bands, start=1):
        if a<=p<=b: return i
    raise ValueError(p)

grade = {code: grade_of(p) for code,(t,p) in items}
points_of = {code:p for code,(t,p) in items}
titles = {code:t for code,(t,p) in items}

rows=[]
with open('inputs_extract/market_survey_match_detail.csv') as f:
    for row in csv.DictReader(f):
        rows.append((row['job_code'], row['benchmark_title'], int(row['evaluation_points']),
                      row['source'], row['source_effective'], int(row['reported_50th_percentile']),
                      row['scope_match'], int(row['incumbents_in_cut']), row['match_note']))

WEIGHTS = {'Cranmere Carrier Compensation Survey':F(6,10), 'Delaney National Insurance Survey':F(4,10)}

excluded = {}
usable_by_job = defaultdict(list)
for (job_code, title, evalpts, source, srceff, pct50, scope, incumbents, note) in rows:
    srceff_d = parse(srceff)
    age_months = months_between(srceff_d, EFFECTIVE)
    reasons = []
    if age_months > 24: reasons.append('Source effective ' + str(srceff) + ' is more than 24 months before the 1 Jan 2027 structure effective date')
    if scope not in ('Exact','Good'): reasons.append('Scope match recorded as ' + str(scope) + ', not Exact or Good')
    if incumbents < 5: reasons.append('Cut carries only ' + str(incumbents) + ' incumbents, below the 5-incumbent minimum')
    if reasons:
        excluded.setdefault(job_code, []).append((source, srceff, pct50, scope, incumbents, reasons))
        continue
    aged = F(pct50) * (1 + F(32,1000)*F(age_months,12))
    usable_by_job[job_code].append((source, aged, WEIGHTS[source]))

blended = {}
for job_code, entries in usable_by_job.items():
    if len(entries)==1:
        blended[job_code] = entries[0][1]
    else:
        num = sum(a*w for s,a,w in entries)
        den = sum(w for s,a,w in entries)
        blended[job_code] = num/den

by_grade_usable = defaultdict(list)
for code, rate in blended.items():
    by_grade_usable[grade[code]].append((code, rate))

SPREAD = {}
for g in range(1,n_grades+1):
    if g<=3: SPREAD[g]=F(40,100)
    elif g<=7: SPREAD[g]=F(45,100)
    else: SPREAD[g]=F(50,100)

raw_mid = {}
mechanism = {}
usable_count = {}
for g in range(1, n_grades+1):
    lst = sorted(by_grade_usable[g], key=lambda x: x[1])
    usable_count[g] = len(lst)
    if len(lst)==0:
        raw_mid[g] = None
        mechanism[g]='Interpolated'
    else:
        n=len(lst)
        if n%2==1: med = lst[n//2][1]
        else: med = (lst[n//2 -1][1] + lst[n//2][1]) / 2
        raw_mid[g] = med
        mechanism[g]='Market Median'

def nearest_benchmarked(g, direction):
    i=g
    while True:
        i+=direction
        if i<1 or i>n_grades: return None
        if raw_mid.get(i) is not None and mechanism[i]=='Market Median':
            return i

interp_from = {}
for g in range(1,n_grades+1):
    if raw_mid[g] is None:
        below = nearest_benchmarked(g,-1); above = nearest_benchmarked(g,1)
        interp_from[g] = (below, above)
        raw_mid[g] = (round_hundred(raw_mid[below]) + round_hundred(raw_mid[above])) / 2
        mechanism[g] = 'Interpolated'

rounded_mid_premarket = {g: round_hundred(raw_mid[g]) for g in range(1,n_grades+1)}

final_mid = {}
final_mid[1] = rounded_mid_premarket[1]
step_repair_detail = {}
for g in range(2, n_grades+1):
    prev = final_mid[g-1]
    cur = rounded_mid_premarket[g]
    lo_bound = prev * F(108,100)
    hi_bound = prev * F(114,100)
    if cur < lo_bound:
        val = lo_bound
        r = round_hundred(val)
        if r < lo_bound: r += 100
        final_mid[g] = r
        mechanism[g]='Step Repair'
        step_repair_detail[g] = ('low', cur, prev)
    elif cur > hi_bound:
        val = hi_bound
        r = round_hundred(val)
        if r > hi_bound: r -= 100
        final_mid[g] = r
        mechanism[g]='Step Repair'
        step_repair_detail[g] = ('high', cur, prev)
    else:
        final_mid[g] = cur

appendix_min_force = {1:33100,2:40200,3:39800,4:41500,5:45200,6:50900,7:57300,8:62800,9:69100,10:76400,11:84900}

grade_min={}; grade_max={}; grade_min_pref={}
for g in range(1,n_grades+1):
    spread=SPREAD[g]
    mid = final_mid[g]
    raw_min = F(mid) / (1+spread/2)
    rmin = round_hundred(raw_min)
    grade_min_pref[g] = rmin
    if appendix_min_force[g] > rmin:
        rmin = F(appendix_min_force[g])
        mechanism[g] = 'Continuity Floor'
    grade_min[g]=rmin
    raw_max = F(rmin)*(1+spread)
    grade_max[g]=round_hundred(raw_max)

ZONES = {'Zone 1': (F(92,100), ['Ohio','Indiana','Missouri']),
         'Zone 2': (F(100,100), ['Colorado','Illinois','Minnesota']),
         'Zone 3': (F(112,100), ['California','Washington','New York'])}

state_wage = {}
with open('inputs_extract/State_Wage_Floor_Schedule_2027.csv') as f:
    for row in csv.DictReader(f):
        state_wage[row['state']] = row

def zone_floor(zone_name, grade_ceiling):
    factor, states = ZONES[zone_name]
    if grade_ceiling < 420:
        vals = [F(state_wage[s]['hourly_minimum_2027'])*2080 for s in states]
        return max(vals), 'hourly'
    else:
        vals = [F(state_wage[s]['exempt_salary_threshold_2027']) for s in states]
        return max(vals), 'salary'

zone_table = {}
for g in range(1, n_grades+1):
    ceiling = bands[g-1][1]
    for zname,(factor,states) in ZONES.items():
        zmin_geo = round_hundred(F(grade_min[g]) * factor)
        zmax = round_hundred(F(grade_max[g]) * factor)
        zmid = round_hundred(F(final_mid[g]) * factor)
        floor_val, floor_type = zone_floor(zname, ceiling)
        if floor_val > zmin_geo:
            zmin_final = round_up_hundred(floor_val)
            floor_applied = True
        else:
            zmin_final = zmin_geo
            floor_applied = False
        zone_table[(g,zname)] = dict(min=int(zmin_final), mid=int(zmid), max=int(zmax),
                                       floor_applied=floor_applied, floor_val=float(floor_val), floor_type=floor_type,
                                       geo_min=int(zmin_geo))

print("EXCLUDED MATCHES:")
for code, entries in excluded.items():
    for e in entries:
        print(" ", code, e)

print("\nGRADE TABLE (base, zone-neutral factor=1.00 basis):")
for g in range(1,n_grades+1):
    print("G" + str(g), bands[g-1], "usable=" + str(usable_count[g]), "mech=" + mechanism[g],
          "min=" + str(grade_min[g]), "mid=" + str(final_mid[g]), "max=" + str(grade_max[g]),
          "(pre-floor min=" + str(grade_min_pref[g]) + ")")

print("\nSTEP REPAIR DETAIL:", step_repair_detail)
print("\nINTERP FROM:", interp_from)

print("\nZONE TABLE:")
for g in range(1,n_grades+1):
    for zname in ZONES:
        d = zone_table[(g,zname)]
        print("G" + str(g), zname, "min=" + str(d['min']), "mid=" + str(d['mid']), "max=" + str(d['max']),
              "floor_applied=" + str(d['floor_applied']), "floor=" + str(d['floor_val']), "(" + d['floor_type'] + ")",
              "geo_min=" + str(d['geo_min']))

out = dict(
    bands=bands,
    grade_min={str(k):str(v) for k,v in grade_min.items()},
    grade_mid={str(k):str(v) for k,v in final_mid.items()},
    grade_max={str(k):str(v) for k,v in grade_max.items()},
    grade_min_pref={str(k):str(v) for k,v in grade_min_pref.items()},
    usable_count={str(k):v for k,v in usable_count.items()},
    mechanism=mechanism,
    zone_table={str(g)+"|"+z: v for (g,z),v in zone_table.items()},
    grade_of=grade, points_of=points_of, titles=titles,
    excluded={k: [list(e) for e in v] for k,v in excluded.items()},
    step_repair_detail={str(k):list(v) for k,v in step_repair_detail.items()},
    interp_from={str(k):list(v) for k,v in interp_from.items()},
)
with open('work/grades.json','w') as f:
    json.dump(out, f, indent=2)
