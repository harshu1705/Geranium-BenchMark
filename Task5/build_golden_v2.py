# -*- coding: utf-8 -*-
"""Build Pruitt_Retirement_Readiness_Assessment.docx.

Every dollar figure in the narrative is pulled from the calculation block below,
so the prose cannot drift away from the arithmetic.
"""
import os
import docx
from docx.shared import Pt

# ============================================================
# Calculation block
# ============================================================
BAL = {
    "Daniel 401(k)":           (680000, 0.65, 0.30, 0.05),
    "Renee 403(b)":            (310000, 0.70, 0.25, 0.05),
    "Joint Taxable Brokerage": (215000, 0.80, 0.20, 0.00),
    "Renee Roth IRA":          (95000,  0.90, 0.10, 0.00),
    "Joint Savings":           (45000,  0.00, 0.00, 1.00),
}
G, I, SWR = 0.06, 0.03, 0.04
AGE0, RET = 60, 62
BASE_TODAY, BRIDGE = 102000.0, 21600.0
SS_M = {62: (2400, 1150), 65: (2947, 1387), 67: (3400, 1600)}

TOTAL = sum(v[0] for v in BAL.values())
EQ = sum(v[0] * v[1] for v in BAL.values())
FI = sum(v[0] * v[2] for v in BAL.values())
CA = sum(v[0] * v[3] for v in BAL.values())
EQP, FIP, CAP = EQ / TOTAL * 100, FI / TOTAL * 100, CA / TOTAL * 100

GROWN = TOTAL * (1 + G) ** 2
CONTRIB = 3100 * 24
PORT_RET = GROWN + CONTRIB

f = lambda age: (1 + I) ** (age - AGE0)
base = lambda age: BASE_TODAY * f(age)
gross = lambda age: base(age) + (BRIDGE if 62 <= age <= 64 else 0.0)
ssa = lambda age: (SS_M[age][0] + SS_M[age][1]) * 12


def scenario(claim, growth_first=True):
    port, rows = PORT_RET, []
    for age in range(RET, claim):
        net = gross(age)
        port = port * (1 + G) - net if growth_first else (port - net) * (1 + G)
        rows.append((age, net, port))
    need = gross(claim)
    ss = ssa(claim) * f(claim)
    wd = need - ss
    req = wd / SWR
    return dict(rows=rows, port=port, need=need, ss=ss, wd=wd, req=req, gap=port - req)


def carry(claim, test):
    port = PORT_RET
    for age in range(RET, test):
        net = gross(age) - (ssa(claim) * f(age) if age >= claim else 0.0)
        port = port * (1 + G) - net
    wd = gross(test) - (ssa(claim) * f(test) if test >= claim else 0.0)
    req = wd / SWR
    return dict(port=port, wd=wd, req=req, gap=port - req)


A, B, C = scenario(62), scenario(65), scenario(67)
Balt, Calt = scenario(65, False), scenario(67, False)
A65 = carry(62, 65)
A67, B67 = carry(62, 67), carry(65, 67)

# Scenario A year-by-year net draws for the like-for-like walk
A_draws = [(age, ssa(62) * f(age), gross(age) - ssa(62) * f(age)) for age in (62, 63, 64)]
A_port_walk, p_ = [], PORT_RET
for age, _ss, net in A_draws:
    p_ = p_ * (1 + G) - net
    A_port_walk.append(p_)

d0 = lambda x: "${:,.0f}".format(x)
d2 = lambda x: "${:,.2f}".format(x)
pct = lambda x: "{:.1f}%".format(x)
pct2 = lambda x: "{:.2f}%".format(x)

# Rebalancing trades implied by the 55/40/5 target
TGT_EQ, TGT_FI, TGT_CA = 0.55 * TOTAL, 0.40 * TOTAL, 0.05 * TOTAL
EQ_SHIFT = EQ - TGT_EQ
TD_EQUITY = 680000 * 0.65 + 310000 * 0.70
TD_TOTAL = 680000 + 310000

# Breakeven ages on cumulative benefits (today's dollars, undiscounted)
BE_65 = 65 + (ssa(62) * 3) / (ssa(65) - ssa(62))
BE_67 = 67 + (ssa(62) * 5) / (ssa(67) - ssa(62))

# ============================================================
# Document
# ============================================================
OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "PIVOT3_golden")
os.makedirs(OUT_DIR, exist_ok=True)
OUT = os.path.join(OUT_DIR, "Pruitt_Retirement_Readiness_Assessment.docx")

doc = docx.Document()
doc.styles['Normal'].font.name = 'Calibri'
doc.styles['Normal'].font.size = Pt(11)

h1 = lambda t: doc.add_heading(t, level=1)
h2 = lambda t: doc.add_heading(t, level=2)
p = lambda t: doc.add_paragraph(t)
b = lambda t: doc.add_paragraph(t, style='List Bullet')


def table(headers, rows):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Light Grid Accent 1'
    for i, ht in enumerate(headers):
        cell = t.rows[0].cells[i]
        cell.text = ht
        for para in cell.paragraphs:
            for run in para.runs:
                run.bold = True
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = str(v)
    doc.add_paragraph()


doc.add_heading('Pruitt Household: Retirement Readiness Assessment', level=0)
p('Prepared by: Advisory Team  |  Report Date: September 2026')
p('Daniel Pruitt (age 60) and Renee Pruitt (age 60)  |  Target Retirement: Daniel age 62, '
  'approximately 2 years from today')
p('Purpose: to answer the question the Pruitts brought to us, namely whether they can both stop working '
  'when Daniel turns 62, and to give them a defensible recommendation on when to claim Social Security.')

# ---------------------------------------------------------------- 1
h1('1. Current Financial Snapshot')
p("The household holds {} in investable assets across five accounts: Daniel's 401(k) ({}), Renee's 403(b) "
  "({}), a joint taxable brokerage ({}, with a $150,000 cost basis and $65,000 of unrealized gain), Renee's "
  "Roth IRA ({}), and joint savings ({}). The trailing four quarters show steady growth in every account "
  "with no withdrawals, and contributions have run at a consistent $3,100 per month. The primary residence "
  "is owned outright and carries no mortgage, and there are no other liabilities on file."
  .format(d0(TOTAL), d0(680000), d0(310000), d0(215000), d0(95000), d0(45000)))

table(['Account', 'Balance', 'Equity', 'Fixed Income', 'Cash'],
      [['Daniel 401(k)', d0(680000), '65%', '30%', '5%'],
       ['Renee 403(b)', d0(310000), '70%', '25%', '5%'],
       ['Joint Taxable Brokerage', d0(215000), '80%', '20%', '0%'],
       ['Renee Roth IRA', d0(95000), '90%', '10%', '0%'],
       ['Joint Savings', d0(45000), '0%', '0%', '100%'],
       ['Total / blended', d0(TOTAL), '{} ({})'.format(pct(EQP), d0(EQ)),
        '{} ({})'.format(pct(FIP), d0(FI)), '{} ({})'.format(pct(CAP), d0(CA))]])

p("Weighting each account's stated allocation by its balance produces a blended portfolio of approximately "
  "{} equity ({}), {} fixed income ({}), and {} cash ({}). Against the firm's standard near-retirement "
  "target of 55% equity / 40% fixed income / 5% cash, the household is overweight equities by roughly 13 "
  "percentage points and correspondingly underweight fixed income by about 15. That is a more "
  "growth-oriented posture than the firm's glide path contemplates for a household two years from drawing "
  "on this portfolio, and it sits awkwardly against their own stated wish to reduce equity exposure as "
  "retirement approaches. It matters most in the years immediately around the retirement date, when a "
  "market decline arriving at the same time as the first withdrawals does disproportionate long-term "
  "damage."
  .format(pct(EQP), d0(EQ), pct(FIP), d0(FI), pct(CAP), d0(CA)))

# ---------------------------------------------------------------- 2
h1('2. Retirement Income Needs Analysis')
p("The budget workbook shows current spending of $9,200 per month ($110,400 annually), dropping to an "
  "expected $8,500 per month ($102,000 annually, in today's dollars) once retirement account and brokerage "
  "contributions stop and the housing, healthcare, transportation and travel lines shift. That headline "
  "drop of $700 a month understates what is happening underneath it. Contributions ending account for "
  "$3,100 a month on their own, so stripping those out, actual living costs rise by $2,400 a month in "
  "retirement rather than falling: travel up $800, miscellaneous up $1,100, healthcare up $550 and housing "
  "up $200, against only $250 of reductions in food and transportation. This is not a plan that leans on "
  "spending less. Inflated at the plan's 3% assumption over the two years until retirement, the $102,000 "
  "becomes "
  "approximately {} in the first year of retirement ($102,000 x 1.03^2)."
  .format(d0(base(62))))

p("Separately, because neither spouse is eligible for Medicare until 65, the household will need to buy "
  "private health coverage through the ACA marketplace or COBRA for the three years from ages 62 through "
  "64, quoted at $1,800 per month, or {} annually. The workbook states that figure as the planning cost "
  "for each of those three bridge years, so it is carried at face value rather than inflated further. This "
  "cost is not inside the $8,500 monthly retirement budget and has to be added on top of it, but only for "
  "those three years."
  .format(d0(BRIDGE)))

p("Total funding need in the first year of retirement, before any guaranteed income, is therefore "
  "approximately {} plus {}, or {}. The two years that follow carry the same bridge cost against a base "
  "budget that keeps inflating."
  .format(d0(base(62)), d0(BRIDGE), d0(gross(62))))

table(['Retirement year', 'Base budget (inflated 3%)', 'Health bridge', 'Total need'],
      [['Age 62', d2(base(62)), d0(BRIDGE), d2(gross(62))],
       ['Age 63', d2(base(63)), d0(BRIDGE), d2(gross(63))],
       ['Age 64', d2(base(64)), d0(BRIDGE), d2(gross(64))],
       ['Age 65 (Medicare begins)', d2(base(65)), '$0', d2(gross(65))],
       ['Age 66', d2(base(66)), '$0', d2(gross(66))],
       ['Age 67', d2(base(67)), '$0', d2(gross(67))]])

p("The step down at 65 is worth pointing out to the Pruitts directly. Their spending need falls by roughly "
  "$18,000 in the year Medicare starts even though their lifestyle does not change at all, which is exactly "
  "why the first three years of this plan are the hard ones and why the claiming decision below turns on "
  "how those three years are funded.")

# ---------------------------------------------------------------- 3
h1('3. Portfolio Projection to Retirement')
p("Applying the plan's 6% assumed annual growth rate to the current {} balance over two years (1.06 "
  "squared, a 12.36% cumulative factor) produces approximately {}. Adding the $3,100 per month in combined "
  "contributions expected to continue until retirement ({} over 24 months, counted at face value without "
  "crediting growth on the deposits themselves, which is the more conservative convention) brings the "
  "projected balance at retirement to approximately {}. Crediting those contributions with growth as they "
  "are made would add roughly $4,300 and is immaterial to any conclusion below."
  .format(d0(TOTAL), d0(GROWN), d0(CONTRIB), d0(PORT_RET)))

# ---------------------------------------------------------------- 4
h1('4. Social Security Claiming Analysis')
p("Daniel and Renee have indicated a preference to claim at 62, in their words to \"get something back\" "
  "after years of paying in. Both were born in 1966, so Full Retirement Age is 67 for each of them, and "
  "they reach each milestone age within the same calendar year as one another. The combined household "
  "benefit at a given milestone is therefore available in full from the start of each scenario below, "
  "which keeps the comparison clean.")

table(['Claiming age', 'Daniel (monthly)', 'Renee (monthly)', 'Combined monthly',
       'Combined annual', 'Reduction from FRA'],
      [['62 (earliest)', d0(2400), d0(1150), d0(3550), d0(ssa(62)), 'approx. 30%'],
       ['65', d0(2947), d0(1387), d0(2947 + 1387), d0(ssa(65)), 'approx. 13.3%'],
       ['67 (Full Retirement Age)', d0(3400), d0(1600), d0(5000), d0(ssa(67)), 'none']])

p("The gap between those rows is permanent, not a matter of timing. Because Full Retirement Age is 67 for "
  "this cohort, the reduction schedule is steeper in the early years than clients usually assume: claiming "
  "at 62 costs approximately 30% of the Full Retirement Age benefit for life, and even claiming at 65, only "
  "24 months early, still costs approximately 13.3%. That age-65 figure is larger than a straight-line "
  "reading between the 62 and 67 amounts would suggest, because the reduction accrues per month at the "
  "steeper five-ninths-of-one-percent rate across the whole 24-month window. Claiming at 62 would lock this "
  "household into {} a year in today's dollars instead of {}, a permanent difference of {} a year, and it "
  "is the single most consequential decision in this plan that is still entirely within their control."
  .format(d0(ssa(62)), d0(ssa(67)), d0(ssa(67) - ssa(62))))

p("On cumulative benefits alone, and setting the portfolio aside for a moment, waiting from 62 to 65 means "
  "forgoing three years of the reduced benefit, about {}, in exchange for about {} more per year for life "
  "thereafter, which pays back at roughly age {:.1f}. Waiting all the way to 67 means forgoing about {} in "
  "exchange for about {} more per year, which pays back at roughly age {:.1f}. Both breakevens land in the "
  "late seventies, against two people in good health with no chronic conditions on file and no dependents, "
  "which argues for delay on its own terms even before the portfolio arithmetic is brought in."
  .format(d0(ssa(62) * 3), d0(ssa(65) - ssa(62)), BE_65,
          d0(ssa(62) * 5), d0(ssa(67) - ssa(62)), BE_67))

p("Two further features of the benefit estimates point the same way. Survivor benefits run up to 100% of a "
  "deceased spouse's benefit, and Daniel's benefit is roughly double Renee's at every claiming age shown, "
  "so an age-62 election would permanently reduce not only his own benefit but the amount the surviving "
  "spouse steps up to, over what could be a long period of survivorship. Separately, the benefit statement notes that one "
  "half of Daniel's Full Retirement Age benefit ($1,700) is modestly above Renee's own Full Retirement Age "
  "benefit ($1,600), so a spousal top-up of up to roughly $100 per month could be available to Renee once "
  "Daniel has filed. That top-up is excluded from every figure in this report, both because it is subject to "
  "its own early-claiming reduction and because it cannot be confirmed without a determination from the "
  "Social Security Administration. It is upside, not income we are planning around, and it should be "
  "confirmed at the time of filing.")

p("Three caveats attach to the figures used throughout. The estimates are stated in today's dollars and do "
  "not project future cost-of-living adjustments, so for every sustainability test below they are grown at "
  "the plan's 3% inflation assumption to the nominal dollars of the year in which they are actually "
  "received, matching the basis on which spending is projected. Depending on total household income in "
  "retirement, up to 85% of benefits may be federally taxable, so the gross figures somewhat overstate what "
  "is available to spend; the household should confirm the expected treatment with a tax professional "
  "before finalizing the claiming decision. The earnings test that can withhold benefits claimed before "
  "Full Retirement Age does not apply here, since neither spouse expects wage income once retired.")

p("Because the household's own question was framed as claiming at 62 versus delaying, and because the "
  "benefit statement supplies amounts at 62, 65 and 67, all three claiming ages are carried into the "
  "sustainability test rather than only the first two. For reference throughout the rest of this report "
  "they are labeled Scenario A (claim at 62), Scenario B (claim at 65) and Scenario C (claim at 67). "
  "Delaying beyond 67 to 70 would add delayed retirement credits of roughly 8% a year, but the statement "
  "does not supply post-Full-Retirement-Age amounts, so age 70 is not modeled here and would need a fresh "
  "estimate from the Social Security Administration before it could be.")

# ---------------------------------------------------------------- 5
h1('5. Portfolio Sustainability Test')
p("Each claiming strategy is screened against the firm's standard 4% initial withdrawal rate test: the "
  "portfolio required to support the first full year of ongoing withdrawals at a 4% initial rate is "
  "compared with the portfolio actually projected to be available at that point. Two conventions from the "
  "planning assumptions on file govern how that is applied here. First, Social Security estimates are "
  "inflated at 3% to the nominal dollars of the year they are received, so income and spending are compared "
  "on the same basis. Second, where a strategy involves a self-funded bridge period before Social Security "
  "begins, the screen is applied at the point ongoing steady-state withdrawals begin, with each bridge "
  "year's drawdown modeled explicitly beforehand, and the portfolio keeps earning its assumed 6% during "
  "the drawdown rather than being treated as static.")

p("That second convention is worth stating plainly, because it is the difference between a meaningful "
  "answer and a misleading one. In Scenarios B and C the household draws {} in the first bridge year "
  "against a {} portfolio, which is about 8.2% of the balance. Screening that figure as though it were the "
  "permanent withdrawal rate would fail the plan instantly and would be wrong, because the elevated draw "
  "lasts only until Social Security and Medicare arrive. What matters is whether the portfolio that "
  "survives the bridge can support the steady-state withdrawal that follows it, and the bridge drawdown is "
  "shown year by year below so nothing is hidden inside that judgment. Applying the 4% rate at age 65 or "
  "67 while the 30-year horizon still runs from age 62 is conservative in the other direction, since fewer "
  "years remain to fund at that point."
  .format(d0(gross(62)), d0(PORT_RET)))

h2('Scenario A: Claim at 62')
p("Subtracting Social Security from the first-year need leaves the portfolio withdrawal requirement. The {} "
  "today's-dollar age-62 benefit grows to approximately {} in age-62 nominal dollars ({} x 1.03^2), which "
  "nets against the {} first-year need to leave a portfolio withdrawal requirement of approximately {} in "
  "year one. Sustaining that at the firm's 4% initial withdrawal rate requires a portfolio of approximately "
  "{} ({} / 0.04). Against the {} projected to be available at retirement, Scenario A is short by "
  "approximately {}. Put the other way, the household would be drawing {} of the portfolio in year one, "
  "well above the 4% the firm treats as sustainable over a 30-year horizon. This scenario fails the screen, "
  "and not narrowly."
  .format(d0(ssa(62)), d2(A['ss']), d0(ssa(62)), d2(A['need']), d2(A['wd']),
          d2(A['req']), d2(A['wd']), d0(PORT_RET), d2(-A['gap']),
          pct2(A['wd'] / PORT_RET * 100)))

h2('Scenario B: Delay to 65, Self-Fund a Three-Year Bridge')
p("Delaying means the household draws entirely on the portfolio, with no Social Security offset, for the "
  "three years from 62 through 64, inclusive of the health insurance bridge cost in each of them. Applying "
  "the plan's 6% growth to the portfolio in each year before that year's withdrawal, the balance moves as "
  "follows.")

table(['Age', 'Portfolio after growth', 'Withdrawal', 'Portfolio at year end'],
      [[str(age), d2((prev) * (1 + G)), d2(net), d2(port)]
       for (age, net, port), prev in zip(B['rows'], [PORT_RET] + [r[2] for r in B['rows'][:-1]])])

p("The portfolio ends the bridge at approximately {}. That is a decline of only about {} from the {} it "
  "started with, despite {} of withdrawals over the three years, because growth absorbs most of the draw. "
  "From 65 onward Medicare replaces the private health bridge for both spouses and Social Security begins. "
  "The {} today's-dollar age-65 benefit grows to approximately {} in age-65 nominal dollars ({} x 1.03^5), "
  "and the base need inflated through that fifth year is approximately {}, leaving an ongoing withdrawal "
  "requirement of approximately {}. At the 4% rule that requires a portfolio of approximately {} ({} / "
  "0.04). Against the {} available, Scenario B shows a surplus of approximately {} and passes the screen, "
  "with an implied initial withdrawal rate of {}."
  .format(d2(B['port']), pct((PORT_RET - B['port']) / PORT_RET * 100), d0(PORT_RET),
          d2(sum(r[1] for r in B['rows'])), d0(ssa(65)), d2(B['ss']), d0(ssa(65)),
          d2(B['need']), d2(B['wd']), d2(B['req']), d2(B['wd']), d2(B['port']),
          d2(B['gap']), pct2(B['wd'] / B['port'] * 100)))

p("It passes, but only just. The surplus is about {} of the required portfolio, which is inside the margin "
  "of error of a 6% growth assumption sustained over five consecutive years. The sensitivity note at the "
  "end of this section quantifies how little it takes to erase it."
  .format(pct(B['gap'] / B['req'] * 100)))

h2('Scenario C: Delay to 67, Self-Fund a Five-Year Bridge')
p("Nothing in the household's goals restricts the delay to age 65, and the benefit statement supplies "
  "Full Retirement Age amounts, so the same test is run on a delay all the way to 67. The bridge extends to "
  "five years, but the last two of them are materially cheaper because Medicare has already replaced the "
  "{} health insurance cost from 65 onward."
  .format(d0(BRIDGE)))

table(['Age', 'Portfolio after growth', 'Withdrawal', 'Portfolio at year end'],
      [[str(age), d2((prev) * (1 + G)), d2(net), d2(port)]
       for (age, net, port), prev in zip(C['rows'], [PORT_RET] + [r[2] for r in C['rows'][:-1]])])

p("The portfolio ends the longer bridge at approximately {}, about {} below where it started, having "
  "funded {} of withdrawals. At 67 the unreduced {} combined benefit grows to approximately {} in "
  "age-67 nominal dollars ({} x 1.03^7) against a base need of approximately {}, leaving an ongoing "
  "withdrawal requirement of approximately {}. That requires a portfolio of approximately {} at the 4% "
  "rule, against {} available, a surplus of approximately {} and an implied initial withdrawal rate of {}. "
  "Scenario C passes with a materially wider margin than Scenario B, which is the opposite of what the "
  "Pruitts would intuitively expect from spending two more years of their own money before any benefit "
  "arrives."
  .format(d2(C['port']), pct((PORT_RET - C['port']) / PORT_RET * 100),
          d2(sum(r[1] for r in C['rows'])), d0(ssa(67)), d2(C['ss']), d0(ssa(67)),
          d2(C['need']), d2(C['wd']), d2(C['req']), d2(C['port']), d2(C['gap']),
          pct2(C['wd'] / C['port'] * 100)))

p("The reason is that the unreduced benefit does more work than the extra two years of drawdown cost. "
  "Delaying from 65 to 67 lifts the combined benefit from {} to {} in today's dollars, roughly {} a year "
  "for life, which lowers the required portfolio by about {}. The two additional bridge years reduce the "
  "available portfolio by only about {}, because Medicare has already removed the health insurance cost "
  "and the portfolio is still compounding."
  .format(d0(ssa(65)), d0(ssa(67)), d0(ssa(67) - ssa(65)), d2(B['req'] - C['req']),
          d2(B['port'] - C['port'])))

h2('The Three Scenarios Side by Side')
table(['Scenario', 'Screen date', 'Portfolio available', 'Ongoing withdrawal',
       'Required at 4%', 'Result'],
      [['A: claim at 62', 'Age 62', d2(PORT_RET), d2(A['wd']), d2(A['req']),
        'Shortfall {}'.format(d2(-A['gap']))],
       ['B: claim at 65', 'Age 65', d2(B['port']), d2(B['wd']), d2(B['req']),
        'Surplus {}'.format(d2(B['gap']))],
       ['C: claim at 67', 'Age 67', d2(C['port']), d2(C['wd']), d2(C['req']),
        'Surplus {}'.format(d2(C['gap']))]])

h2('Comparing Scenarios A and B on a Common Date')
p("The three results above are each measured on the date that scenario's ongoing withdrawals begin, which "
  "is the correct application of an initial-withdrawal-rate test but does mean they are dated differently. "
  "To put the household's two stated alternatives on a like-for-like footing, Scenario A is carried forward "
  "to age 65 on the same 6%-growth, inflated-benefit basis used for Scenario B. With the inflated age-62 "
  "benefit offsetting the need in each of the first three years, Scenario A's net withdrawals at 62, 63 and "
  "64 are approximately {}, {} and {}, leaving approximately {} at age 65. Its ongoing need at 65 is the "
  "same {} inflated base less the age-62 benefit grown to approximately {}, or approximately {}, which at a "
  "4% initial rate requires approximately {}. Measured at age 65, Scenario A is still short by "
  "approximately {}, against Scenario B's surplus of approximately {}: a swing of approximately {} from the "
  "claiming decision alone."
  .format(d2(A_draws[0][2]), d2(A_draws[1][2]), d2(A_draws[2][2]), d2(A65['port']),
          d2(base(65)), d2(ssa(62) * f(65)), d2(A65['wd']), d2(A65['req']),
          d2(-A65['gap']), d2(B['gap']), d2(B['gap'] - A65['gap'])))

p("Note that Scenario A looks considerably better on this basis than it did at age 62, short by about {} "
  "rather than {}. That is not a contradiction. Carrying a scenario past the date its ongoing withdrawals "
  "actually begin credits it with three more years of growth on a portfolio that has already been screened, "
  "so the later reading is a comparison aid rather than the governing test. It is shown because it is the "
  "conservative way to frame the recommendation: even on the reading most favorable to claiming early, "
  "claiming at 62 still fails."
  .format(d2(-A65['gap']), d2(-A['gap'])))

h2('Sensitivity and Cross-Checks')
p("Two checks were run on the results above, because the Scenario B margin is thin enough that the answer "
  "should not rest on a modeling convention.")
b("Withdrawal timing. The tables above credit a full year of growth before each year's withdrawal. If "
  "instead each year's spending is taken at the start of the year and only the remainder grows, Scenario B "
  "arrives at age 65 with approximately {} and slips to a shortfall of approximately {}, while Scenario C "
  "arrives at 67 with approximately {} and still shows a surplus of approximately {}. Scenario B therefore "
  "sits on the pass/fail line depending on an assumption about withdrawal timing; Scenario C passes either "
  "way."
  .format(d2(Balt['port']), d2(-Balt['gap']), d2(Calt['port']), d2(Calt['gap'])))
b("Common screen date. Re-screening all three strategies at a single age-67 date gives a shortfall of "
  "approximately {} for Scenario A, a shortfall of approximately {} for Scenario B, and a surplus of "
  "approximately {} for Scenario C. This is not the firm's test, since the 4% rate is defined on the "
  "initial withdrawal and re-dating a scenario penalizes whichever strategy started its withdrawals "
  "earliest. It is shown only to confirm that the ranking C ahead of B ahead of A holds however the "
  "comparison is dated."
  .format(d2(-A67['gap']), d2(-B67['gap']), d2(C['gap'])))
b("Both checks point the same way. The choice between claiming at 62 and delaying is not close. The choice "
  "between 65 and 67 is the one that determines whether the plan has a real cushion or merely clears the "
  "screen on paper.")

# ---------------------------------------------------------------- 6
h1('6. Asset Allocation and Withdrawal Sequencing Recommendation')
p("The household's current {} equity weighting carries more sequence-of-returns risk than is advisable "
  "here, and the sensitivity check above shows why that matters concretely: every passing scenario depends "
  "on realizing something close to the full 6% growth assumption through the bridge years, which are "
  "precisely the years the portfolio is being drawn on hardest. Recommend rebalancing over the next two "
  "years toward the firm's 55% equity / 40% fixed income / 5% cash guidance, and executing that shift "
  "primarily inside the 401(k) and 403(b), where selling appreciated positions has no tax consequence, "
  "rather than in the taxable brokerage, where the {} of unrealized gain would be realized at capital gains "
  "rates. Reaching the target from today's balances means moving roughly {} out of equities and about {} "
  "into fixed income, with cash coming down about {}. That is comfortably inside the {} of equity currently "
  "sitting in the 401(k) and 403(b), so the entire shift can be executed without a single taxable sale in "
  "the brokerage account."
  .format(pct(EQP), d0(65000), d0(EQ_SHIFT), d0(TGT_FI - FI), d0(CA - TGT_CA), d0(TD_EQUITY)))

p("For withdrawal sequencing once retired, the recommended order is joint savings, then the taxable "
  "brokerage, then the tax-deferred accounts, with the Roth IRA preserved for last.")
b("Joint savings ({}) first. That cash sits inside the investable asset base the 4% test measures, so it is "
  "spend-down capital rather than a reserve held outside the plan, and it should be spent first because it "
  "carries no tax cost and no market risk on withdrawal. Use it for the opening months of year-one "
  "withdrawals, then hold the account at roughly six to twelve months of planned spending, replenishing it "
  "from the brokerage at each scheduled rebalance so that no year's spending has to be raised out of a "
  "falling market.".format(d0(45000)))
b("Joint taxable brokerage ({}) next. Only the {} of built-in gain is taxable, and at capital gains rather "
  "than ordinary rates, so this is the cheapest large source of bridge-year spending. Drawing it down early "
  "also has a second benefit specific to this plan: it keeps reportable income low in the two years before "
  "Medicare begins, which is what determines whether the household pays the income-related monthly "
  "adjustment amount on Part B and Part D premiums from 65 onward.".format(d0(215000), d0(65000)))
b("Daniel's 401(k) and Renee's 403(b) ({} combined) after that. These are taxed as ordinary income on "
  "withdrawal and are subject to required minimum distributions later, so drawing them in the bridge years "
  "is less efficient than the brokerage but becomes the main source once the brokerage is exhausted. The "
  "bridge years are also the household's lowest-income years, which makes them the natural window for any "
  "partial Roth conversions the household later decides to run, though those should be sized against the "
  "Medicare premium thresholds noted above.".format(d0(TD_TOTAL)))
b("Renee's Roth IRA ({}) last. It grows tax-free, carries no required minimum distribution, and is the most "
  "tax-efficient asset to leave to the children under the household's stated preference that whatever "
  "remains passes to them.".format(d0(95000)))

# ---------------------------------------------------------------- 7
h1('7. Insurance and Estate Issues Bearing on This Plan')
p("Four items in the insurance and estate file are relevant to whether this plan holds up, as distinct from "
  "general housekeeping.")
b("No long-term care coverage on either spouse. This is the largest unhedged risk in the plan. A care "
  "event would land directly on the same portfolio that Scenario B clears by roughly {}, and the household "
  "has no separate assets to absorb it. Quotes were last obtained about five years ago and declined on "
  "cost; premiums will be materially higher at their current ages, and no hybrid life/LTC or asset-based "
  "products have been evaluated at all. Re-quote standalone coverage and price hybrid and asset-based "
  "alternatives before the retirement date, while both spouses are in good health and still insurable.".format(d2(B['gap'])))
b("Renee's life insurance ends on her last day of work. Her coverage is employer-provided group term at 1x "
  "salary ($72,000), it is not portable, and it carries no individual conversion option, so she goes from "
  "covered to uncovered on the day she retires. Daniel's 20-year level term policy ($500,000, purchased "
  "2014) expires when he turns 68, in 2034, with no conversion elected, which is during the drawdown years "
  "and before the plan has built any cushion.")
b("Estate documents are stale, and in one respect absent. The wills were executed in 2011 and have not been "
  "reviewed since; beneficiary designations with the account custodians were separately confirmed more "
  "recently, in 2019; and no powers of attorney or healthcare directives are on file with this office at "
  "all. The third of those is the urgent one. Without them, incapacity of either spouse could force a "
  "court-supervised guardianship or conservatorship at exactly the point the household can least absorb the "
  "cost and delay.")
b("Health coverage between 62 and 65 is quoted but not committed. The {} annual bridge figure is a current "
  "preliminary quote. COBRA is generally capped at 18 months and typically costs more than a marketplace "
  "plan once the employer subsidy ends, so it cannot cover the full three-year gap on its own. Both routes "
  "should be priced formally in the months before Daniel's actual retirement date, because this line is a "
  "meaningful share of the bridge-year withdrawals that drive the whole analysis.".format(d0(BRIDGE)))

# ---------------------------------------------------------------- 8
h1('8. Action Plan and Recommendations')
p("The direct answer to the question the Pruitts asked is yes, they can both retire at 62, but only if they "
  "stop planning to claim Social Security at 62. The retirement date is affordable; the claiming date as "
  "currently intended is not.")

b("Do not claim Social Security at 62. On a like-for-like comparison at age 65, claiming at 62 leaves the "
  "plan short by approximately {} while delaying to 65 leaves a surplus of approximately {}, a swing of "
  "approximately {} that comes from the claiming decision alone. Measured at its own retirement-date screen, "
  "claiming at 62 is short by approximately {}. There is no reading of these numbers under which claiming at "
  "62 passes the firm's sustainability test."
  .format(d2(-A65['gap']), d2(B['gap']), d2(B['gap'] - A65['gap']), d2(-A['gap'])))

b("Delay to 67 if they can, and treat 65 as the floor rather than the target. Scenario C clears the 4% "
  "screen by approximately {}, about {} of the required portfolio, and holds up under both sensitivity "
  "checks. Scenario B clears it by approximately {}, about {}, and flips to a small shortfall if "
  "withdrawals are assumed at the start of each year instead of the end. Delaying to 65 is enough to clear "
  "the screen on the plan's stated conventions; delaying to 67 clears it with room to spare. This should "
  "be revisited at the six-month pre-retirement check-in the "
  "household has already asked for, before the claiming decision is finalized."
  .format(d2(C['gap']), pct(C['gap'] / C['req'] * 100), d2(B['gap']),
          pct(B['gap'] / B['req'] * 100)))

b("Consider splitting the claiming ages between the two spouses at that review. Daniel's benefit is roughly "
  "double Renee's and drives the survivor benefit, so the case for delaying his to 67 is stronger than the "
  "case for delaying hers. Renee claiming earlier while Daniel waits would shorten the bridge's cost "
  "without giving up the survivor protection, and is worth modeling once actual portfolio performance "
  "through the next two years is known.")

b("Begin rebalancing toward 55% equity / 40% fixed income / 5% cash over the next two years, executed "
  "inside the tax-deferred accounts to avoid realizing the {} of brokerage gain. The current {} equity "
  "weighting is the plan's main controllable risk, and every passing scenario depends on the growth "
  "assumption holding through the bridge years."
  .format(d0(65000), pct(EQP)))

b("Adopt the withdrawal order in Section 6 and hold the joint savings balance at six to twelve months of "
  "planned spending once the opening withdrawals are funded, so no bridge year has to be raised out of a "
  "falling market.")

b("Price long-term care coverage, including hybrid and asset-based options, before the retirement date. "
  "This is the largest risk the plan does not currently hedge, and the plan's cushion under Scenario B "
  "would not absorb a care event.")

b("Put powers of attorney and healthcare directives in place now, and review the 2011 wills. Beneficiary "
  "designations were confirmed in 2019 and are consistent with the wills, so they are the least urgent of "
  "the three, but all should be refreshed as part of the retirement transition.")

b("Review both life insurance positions. Decide before Renee's last day whether any individual coverage is "
  "warranted to replace the group term that ends with her employment, and revisit Daniel's term policy "
  "before it expires at 68.")

b("Build a spending-reduction contingency and test it at each annual review. If actual portfolio "
  "performance runs meaningfully below 6% before or during the bridge years, the levers are the travel and "
  "miscellaneous lines, which together account for $2,970 a month of the retirement budget and represent "
  "the bulk of the increase over current discretionary spending, or a later retirement date. Identifying "
  "which lever they would pull, and at what trigger, is worth doing now rather than in the moment.")

p("One framing point for the meeting. The Pruitts' instinct to claim at 62 is about getting back what they "
  "paid in, and the arithmetic says that instinct costs them the plan. Delaying is not a matter of leaving "
  "money on the table; it is what converts a portfolio that cannot support this retirement into one that "
  "can. The three bridge years are funded out of assets they already hold, and the analysis above shows the "
  "portfolio finishing that bridge only about {} below where it started."
  .format(pct((PORT_RET - B['port']) / PORT_RET * 100)))

doc.save(OUT)
print("wrote", OUT)
print()
print("--- key figures embedded ---")
for k, v in [("portfolio at retirement", PORT_RET), ("A shortfall", -A['gap']),
             ("B portfolio at 65", B['port']), ("B surplus", B['gap']),
             ("C portfolio at 67", C['port']), ("C surplus", C['gap']),
             ("A at 65 shortfall", -A65['gap']), ("B alt shortfall", -Balt['gap']),
             ("C alt surplus", Calt['gap'])]:
    print("  {:26s} {}".format(k, d2(v)))
