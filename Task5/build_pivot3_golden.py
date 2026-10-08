# -*- coding: utf-8 -*-
import docx
from docx.shared import Pt
import os

OUT_DIR = os.path.join(os.path.dirname(__file__), "PIVOT3_golden")
os.makedirs(OUT_DIR, exist_ok=True)
OUT = os.path.join(OUT_DIR, "Whitfield_Retirement_Readiness_Assessment.docx")

d = docx.Document()
d.styles['Normal'].font.name = 'Calibri'
d.styles['Normal'].font.size = Pt(11)

def h1(t): d.add_heading(t, level=1)
def h2(t): d.add_heading(t, level=2)
def p(t): d.add_paragraph(t)

def table(headers, rows):
    t = d.add_table(rows=1, cols=len(headers))
    t.style = 'Light Grid Accent 1'
    hdr = t.rows[0].cells
    for i, htext in enumerate(headers):
        hdr[i].text = htext
        for para in hdr[i].paragraphs:
            for run in para.runs:
                run.bold = True
    for row in rows:
        cells = t.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = str(val)
    d.add_paragraph()

d.add_heading('Whitfield Household: Retirement Readiness Assessment', level=0)
p('Prepared by: Advisory Team  |  Report Date: September 2026')
p('Gerald Whitfield (age 60) and Linda Whitfield (age 58)  |  Target Retirement: Gerald age 62, approximately 2 years from today')

h1('1. Current Financial Snapshot')
p(
    "The household holds $1,345,000 in investable assets across five accounts: Gerald's 401(k) ($680,000), Linda's "
    "403(b) ($310,000), a joint taxable brokerage ($215,000, with a $150,000 cost basis and $65,000 of unrealized "
    "gain), Linda's Roth IRA ($95,000), and joint savings ($45,000)."
)
p(
    "Weighting each account's stated allocation by balance produces a blended portfolio of approximately 68.1% "
    "equity ($916,500), 24.8% fixed income ($334,000), and 7.0% cash ($94,500). Against the firm's standard "
    "near-retirement target of 55% equity / 40% fixed income / 5% cash, the household is overweight equities by "
    "roughly 13 percentage points and correspondingly underweight fixed income, more growth-oriented than typical "
    "for a household two years from drawing on this portfolio."
)

h1('2. Retirement Income Needs Analysis')
p(
    "The Household Budget worksheet shows current spending of $9,200 per month ($110,400 annually), dropping to an "
    "expected $8,500 per month ($102,000 annually, in today's dollars) once retirement account and brokerage "
    "contributions stop and housing, healthcare, and travel categories shift. Inflated at the plan's 3% assumption "
    "over the 2 years until retirement, this becomes approximately $108,212 in the first year of retirement "
    "($102,000 x 1.03^2)."
)
p(
    "Separately, because Medicare does not begin until age 65, the household will need private health coverage "
    "(ACA marketplace or COBRA) for the three years from ages 62 through 64, quoted at $1,800 per month, or "
    "$21,600 annually. This bridge cost is not part of the retirement budget line above and must be added on top of "
    "it for the first three years of retirement only."
)
p(
    "Total funding need in the first year of retirement (age 62), before any guaranteed income, is therefore "
    "approximately $108,212 plus $21,600, or $129,812."
)

h1('3. Portfolio Projection to Retirement')
p(
    "Applying the plan's 6% assumed annual growth rate to the current $1,345,000 balance over 2 years "
    "(1.06 squared, or a 12.36% cumulative factor) produces approximately $1,511,242. Adding the $3,100 per month "
    "in combined contributions expected to continue until retirement ($74,400 over 24 months) brings the projected "
    "balance at retirement to approximately $1,585,642."
)

h1('4. Social Security Claiming Analysis')
p(
    "Gerald and Linda have indicated a preference to claim Social Security at 62. Per the Social Security Benefit "
    "Estimates, that produces a combined $42,600 annually ($2,400 and $1,150 monthly, respectively). Delaying to "
    "age 65 raises the combined benefit to $54,600 annually ($3,100 and $1,450 monthly)."
)
h2('Scenario A: Claim at 62')
p(
    "Subtracting the $42,600 of Social Security from the $129,812 first-year need leaves a portfolio withdrawal "
    "requirement of approximately $87,212 in year one. Applying the firm's 4% initial withdrawal rate "
    "sustainability screen, sustaining an $87,212 withdrawal requires a portfolio of approximately $2,180,295 "
    "($87,212 / 0.04) at retirement. Against the $1,585,642 projected portfolio, this scenario is short by "
    "approximately $594,653, a substantial gap under a standard 30-year sustainability test."
)
h2('Scenario B: Delay Social Security to 65, Self-Fund the Bridge')
p(
    "Delaying claiming means the household draws entirely on the portfolio, with no Social Security offset, for "
    "the three years from age 62 through 64 (inclusive of the health insurance bridge cost each of those years). "
    "Inflating the base need through each of those three years and adding the $21,600 bridge cost produces "
    "approximately $129,812, $133,058, and $136,402 in years one through three, a cumulative draw of approximately "
    "$399,272. Conservatively assuming no net portfolio growth during that three-year drawdown period (a "
    "deliberately cautious assumption for this screening purpose), the portfolio remaining at age 65 is "
    "approximately $1,186,370 ($1,585,642 less $399,272)."
)
p(
    "From age 65 onward, Medicare replaces the private health bridge and the $54,600 age-65 Social Security benefit "
    "begins. The base need, inflated through that fifth year, is approximately $118,246, leaving an ongoing "
    "withdrawal requirement of approximately $63,646 per year. At the 4% rule, sustaining that requires a portfolio "
    "of approximately $1,591,149. Against the $1,186,370 available at 65, this scenario is short by approximately "
    "$404,779."
)
p(
    "Delaying Social Security to 65 narrows the shortfall by roughly $189,874 compared with claiming at 62, but "
    "does not close it. Neither claiming strategy alone is sufficient to make the current plan sustainable under "
    "this screening test."
)

h1('5. Asset Allocation Recommendation')
p(
    "Given the projected shortfall under both claiming scenarios and only two years until retirement, the "
    "household's current 68.1% equity weighting carries more sequence-of-returns risk than is advisable, "
    "particularly given the plan already depends on selling into the portfolio at a higher-than-planned rate in "
    "the early retirement years. Recommend rebalancing over the next two years toward the firm's 55% equity / 40% "
    "fixed income / 5% cash guidance, funding the shift primarily out of the 401(k) and 403(b), where rebalancing "
    "has no tax consequence, rather than the taxable brokerage account."
)
p(
    "For withdrawal sequencing once retired, recommend drawing from the joint taxable brokerage first (gains taxed "
    "at capital gains rates, and only on the $65,000 of built-in gain), then the 401(k) and 403(b) (taxed as "
    "ordinary income, and subject to Required Minimum Distributions beginning at age 73), and preserving the Roth "
    "IRA for last, since it grows tax-free, carries no RMD requirement, and is the most tax-efficient asset to "
    "leave toward the household's stated legacy goal for their children."
)

h1('6. Action Plan and Recommendations')
for line in [
    "Do not claim Social Security at 62 as currently planned. Delaying to at least 65 narrows the sustainability gap by approximately $189,874 and should be the default recommendation absent a health or longevity concern that argues otherwise.",
    "Even with delayed claiming, a projected shortfall of approximately $404,779 remains. Close this gap through some combination of: working part-time for 1 to 2 years post-retirement, further trimming the $8,500 monthly retirement budget (the travel and miscellaneous increases are the most discretionary lines), or pushing the retirement date back.",
    "Begin rebalancing the portfolio toward 55% equity / 40% fixed income / 5% cash over the next two years, executed inside the tax-deferred accounts to avoid triggering capital gains in the brokerage account.",
    "Address the long-term care insurance gap; neither spouse carries LTC coverage, and a care event during retirement would materially worsen an already tight plan. Recommend at minimum pricing a hybrid life/LTC policy or an asset-based LTC rider.",
    "Update the estate documents. The wills, beneficiary designations, and lack of powers of attorney or healthcare directives have not been reviewed since 2019 at the latest; recommend a full review given the approaching retirement.",
    "Revisit the Gerald term life policy, which expires at age 68, before that date, given continued income dependency during the drawdown years discussed above.",
]:
    p(line)

d.save(OUT)
print('saved', OUT)
