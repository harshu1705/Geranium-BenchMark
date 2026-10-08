import csv, json, datetime
from fractions import Fraction as F

def round_to(x, unit):
    x = F(x)
    q = x/unit
    fl = q.__floor__()
    frac = q-fl
    if frac >= F(1,2):
        return (fl+1)*unit
    return fl*unit

def round_hundred(x):
    return round_to(x,100)

with open('work/grades.json') as f:
    G = json.load(f)

grade_of_code = G['grade_of']
zone_table = G['zone_table']  # "g|zone" -> dict

ZONE_STATES = {'Zone 1': ['Ohio','Indiana','Missouri'],
               'Zone 2': ['Colorado','Illinois','Minnesota'],
               'Zone 3': ['California','Washington','New York']}
state_to_zone = {}
for z, states in ZONE_STATES.items():
    for s in states: state_to_zone[s]=z

state_wage = {}
with open('inputs_extract/State_Wage_Floor_Schedule_2027.csv') as f:
    for row in csv.DictReader(f):
        state_wage[row['state']] = row

bands = G['bands']  # list of [lo,hi]

def own_state_floor(state, grade_ceiling):
    row = state_wage[state]
    if grade_ceiling < 420:
        return F(row['hourly_minimum_2027'])*2080, 'hourly (own state)'
    else:
        return F(row['exempt_salary_threshold_2027']), 'exempt salary threshold (own state)'

CUTOFF = datetime.date(2027,2,1)

roster = []
with open('inputs_extract/Incumbent_Roster.csv') as f:
    for row in csv.DictReader(f):
        roster.append(row)

print("Total roster rows:", len(roster))

in_employment = []
excluded_before_cutoff = []
for r in roster:
    if r['employment_status']=='Terminated' and r['termination_date']:
        y,m,d = r['termination_date'].split('-')
        tdate = datetime.date(int(y),int(m),int(d))
        if tdate < CUTOFF:
            excluded_before_cutoff.append(r)
            continue
    in_employment.append(r)

print("Excluded (terminated before 2027-02-01):", len(excluded_before_cutoff))
for r in excluded_before_cutoff:
    print("  ", r['employee_id'], r['name'], r['job_code'], r['termination_date'])

print("In employment on 2027-02-01:", len(in_employment))

ungradable = []
gradable = []
for r in in_employment:
    if r['job_code'] not in grade_of_code:
        ungradable.append(r)
    else:
        gradable.append(r)

print("\nUngradable (job code not on point schedule):", len(ungradable))
for r in ungradable:
    print("  ", r['employee_id'], r['name'], r['job_code'], r['job_title'])

print("\nGradable population:", len(gradable))

fte_check = set(r['fte'] for r in roster)
print("Distinct FTE values in roster:", fte_check)

status_check = set(r['employment_status'] for r in roster)
print("Distinct employment_status values:", status_check)

affected = []
not_affected = []
total_shortfall = F(0)

for r in gradable:
    g = grade_of_code[r['job_code']]
    zone = state_to_zone[r['work_state']]
    zt = zone_table[f"{g}|{zone}"]
    zmin = F(zt['min']); zmax = F(zt['max'])
    salary = F(int(r['current_base_salary']))
    if salary < zmin:
        shortfall = zmin - salary
        affected.append((r, g, zone, salary, zmin, zmax, shortfall))
        total_shortfall += shortfall
    else:
        not_affected.append((r, g, zone, salary, zmin, zmax))

print("\nAffected incumbents (salary below zone minimum):", len(affected))
print("Total cost of moving every affected incumbent to new minimum: $%s" % total_shortfall)

BUDGET = F(355000)
print("Budget:", BUDGET, "Exceeds:", total_shortfall > BUDGET)

ceiling_of_grade = {i+1: bands[i][1] for i in range(len(bands))}

rows2027 = []
total_2027 = F(0)
total_2028 = F(0)
for (r,g,zone,salary,zmin,zmax,shortfall) in affected:
    ceiling = ceiling_of_grade[g]
    floor_val, floor_type = own_state_floor(r['work_state'], ceiling)
    floor_topup = max(F(0), floor_val - salary)
    half = shortfall / 2
    half_r = round_hundred(half)
    floor_topup_r = round_hundred(floor_topup) if floor_topup > 0 else F(0)
    amt2027 = max(half_r, floor_topup_r)
    if amt2027 > shortfall:
        amt2027 = shortfall
    amt2028 = shortfall - amt2027
    total_2027 += amt2027
    total_2028 += amt2028
    rows2027.append((r['employee_id'], r['name'], r['job_code'], r['work_state'], zone, g,
                      int(salary), int(zmin), int(shortfall), float(floor_val), floor_type,
                      int(floor_topup_r), int(half_r), int(amt2027), int(amt2028)))

print("\nTotal 2027 payment: $%s" % total_2027)
print("Total 2028 payment: $%s" % total_2028)
print("Check sum:", total_2027+total_2028, "== total shortfall", total_shortfall)

for row in rows2027:
    if row[0] in ('E-4130','E-4307'):
        print("\nCHECK:", row)

with open('work/costing.json','w') as f:
    json.dump({
        'excluded_before_cutoff':[r['employee_id'] for r in excluded_before_cutoff],
        'ungradable':[r['employee_id'] for r in ungradable],
        'affected_count': len(affected),
        'not_affected_count': len(not_affected),
        'total_shortfall': str(total_shortfall),
        'total_2027': str(total_2027),
        'total_2028': str(total_2028),
        'rows2027': rows2027,
    }, f, indent=2)
