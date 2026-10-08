# Geranium Refinement Suite: Multi-Domain Benchmark Calibration & Task Catalog

## 📌 Overview of the Refinement Workspace
The `GERANIUM REFINEMNET` directory serves as the iterative engineering lab where evaluation tasks are calibrated, edge cases from benchmark model submissions are analyzed, and surgical repairs are applied to goldens, input data files, and rubrics.

When models take benchmark tests, their raw submissions (e.g. `53afd282-8c85-4643...`, `faa74642-26a6-4f01...`, `4649aef0-d039...`) are unpacked, compared against the golden solutions, and audited for:
1. **Ambiguities in Prompts**: Prompts that allow multiple valid interpretations.
2. **Formula Integrity in Excel**: Hardcoded numbers vs live dynamic formulas (`openpyxl` formula injection).
3. **Data Discrepancies**: Subtle mismatches between raw CSV sources and summary tables.
4. **Grading Rubric Sensitivity**: Ensuring rubrics accurately reward nuanced domain reasoning while strictly penalizing hallucinations and shortcuts.

---

## 📂 Catalog of Tasks in GERANIUM REFINEMENT

### 1. Task 5 & Task 7: Granite Ridge Specialty Renewal Underwriting Review
* **Domain**: Commercial Property & Casualty / Habitational Real Estate Underwriting
* **Subject Portfolio**: Granite Ridge Specialty Insurance – October Renewal Batch (36 accounts, 75 insured locations)
* **Core Deliverable**: Multi-tab live financial workbook (`renewal_underwriting_review.xlsx`)
* **Key Files**:
  - `granite_ridge_renewal_accounts.csv` (Account register with stated premiums, unit counts, and stale loss ratios)
  - `granite_ridge_location_schedule.csv` (Detailed physical location schedules with TIV and unit counts)
  - `granite_ridge_loss_run_36mo.csv` (36-month historical loss run containing 42 claim records)
  - `granite_ridge_routing_notes.txt` (Underwriting authority guide & account routing thresholds)

#### Critical Domain Problems & Fixes (`fix_golden_and_rubric.py`, `apply_live_formulas.py`):
1. **Stale vs Recomputed Loss Ratios**:
   - The account list reports stated loss ratios that are out of date for 30 of the 36 accounts.
   - The underwriter must recompute the true 36-month attritional loss ratio for each account directly from `granite_ridge_loss_run_36mo.csv` against annual premium.
2. **Catastrophe Loss Exclusion (Section 9)**:
   - Out of 42 loss rows, 41 are attritional claims with blank `cat_code`.
   - Exactly one row carries a populated catastrophe code (`CAT-2408`). Under Section 9, catastrophe-coded losses must be carved out from the attritional loss ratio calculation.
3. **Location Schedule & TIV Reconciliation (Section 4)**:
   - Identifies significant variances on account `GR-2113` between the account summary unit count (list vs schedule) and Total Insured Value (TIV), triggering mandatory underwriting referral.
4. **Account Routing Dispositions**:
   - Routes each account into one of four buckets: **Cleared to Quote**, **Refer to Senior Underwriter**, **Decline** (`GR-2195` Humboldt Yard, whose recomputed loss ratio hit 292.3%), or **Hold for Broker Clarification** (`GR-2033` Lincolnwood Place).
5. **Live Formula Injection (`apply_live_formulas.py`)**:
   - Replaced static text in the golden solution with dynamic Excel formulas (`="Section 4 per granite_ridge_routing_notes.txt: unit count variance " & TEXT(...)`), linking narrative rationales directly to computed numbers in cells `AD34`, `AD35`, `AD36`, `A41`, `D13`, `D14`, `C2`, and `C3`.

---

### 2. Task 4: QSR Comps Trading Valuation Model
* **Domain**: Investment Banking / Equity Research / M&A Financial Modeling
* **Subject Companies**: Quick Service Restaurant (QSR) peers:
  - McDonald's Corporation (MCD)
  - Yum! Brands, Inc. (YUM)
  - Domino's Pizza, Inc. (DPZ)
* **Source Documents**: Actual SEC Form 10-K (FY2025) and Form 10-Q (1Q26) PDF filings for each company.
* **Core Deliverable**: `QSR_Comps_LTM_1Q26.xlsx`

#### Core Valuation Challenges & Solutions:
1. **LTM Stub Period Reconciliation**:
   - Calculates Last Twelve Months (LTM) financials: $\text{LTM} = \text{FY2025} + \text{1Q26} - \text{1Q25}$.
2. **Normalized EBITDA & D&A Extraction**:
   - Pulls true Depreciation & Amortization directly from the Statement of Cash Flows (rather than trusting partial operating expense footnotes).
   - Normalizes EBITDA by stripping non-recurring items (restructuring, impairment, litigation gains/losses).
3. **Tax-Effecting One-Off Items (The Domino's Trap)**:
   - McDonald's and Yum! report adjusted non-GAAP net income figures in their investor disclosures. Domino's Pizza does not.
   - The analyst must calculate Domino's effective tax rate from GAAP earnings and tax expense, and tax-effect Domino's one-time items to build a harmonized Adjusted Net Income and Adjusted P/E multiple.
4. **Enterprise Value (EV) Bridges**:
   - Constructs exact EV bridges: $\text{Market Cap} + \text{Total Debt} + \text{Finance Leases} - \text{Cash} + \text{Minority Interest}$.
   - Calculates peer trading multiples: EV/Revenue, EV/EBITDA, EV/EBIT, P/E, and applies peer median multiples back to each company to determine implied equity value per share.

---

### 3. Task 2: Cottonwood Market Conduct Examination Report
* **Domain**: Insurance Regulatory Compliance / State Insurance Department Examination
* **Target Role**: State Market Conduct Examiner
* **Regulator**: Kansas Insurance Department (Market Regulation Division)
* **Governing Rule**: Bulletin MC-9 (effective July 1, 2024, superseding MC-8)
* **Examined Carrier**: Cottonwood Casualty Insurance Company
* **Core Deliverable**: `Market_Conduct_Examination_Report.docx`

#### Evaluative Highlights:
1. **Working Days vs Calendar Days**:
   - Standard 4.1 (10 working days to acknowledge claim notice) and Standard 4.3 (10 working days to pay after signed release) require calendar filtering that strictly excludes weekends and declared state holidays.
   - Standard 4.2 (written decision on proof of loss) runs on a 30 **calendar day** clock.
2. **Sample Scoping & Denominator Rules**:
   - Claims with notice dates outside the examination period are purged from both exceptions and the sample denominator.
   - Claims closed without payment are tested for notice, decision, and denial compliance, but are excluded from payment timeliness denominators.
3. **Zero-Tolerance vs 7% Tolerance Standards**:
   - General claim handling error tolerance is 7.0%.
   - Denial standard 4.4 (stating specific policy provision) and complaint standard 5.1 carry a **strict 0.0% tolerance**, where even a single violation constitutes an official examination finding.
4. **Mandatory Statutory Settlement Fees**:
   - Total loss auto settlements must include sales tax at 7.5% of ACV plus the statutory title/registration fee of $42.50.

---

### 4. TASK 3: Arizona Monsoon Appraisal Quality Review
* **Domain**: Property & Casualty Auto Physical Damage / Direct Repair Program (DRP) Audit
* **Subject**: Review of 32 closed-paid auto repair files totaling $232,699.69 during the 2026 Arizona monsoon storm season.
* **Core Deliverable**: Physical damage executive audit report.

#### Key Findings & Traps:
1. **Territory Labor Rate vs Targeted Facility Control**:
   - Tested whether market pressures warranted an across-the-board territory labor rate increase.
   - Proved that labor rate leakage appeared on only 28.1% of claims (failing the 30% / 4-facility threshold required to open a territory rate adjustment).
2. **Leakage Concentration**:
   - Identified that **98.1% ($4,414.75) of all monetary exposure was concentrated in a single facility: Old Pueblo Auto Body (F-211)**, driven by unauthorized rates, missing supplement documentation, and improper pre-inspection approvals.
3. **Special Handling Validations**:
   - Verified legitimate code applications: `ALUM-08` (aluminum repair), `REMOTE-06` (remote scanning), `EMERG-2` (towing/storage), and `FAST-8` (expedited handling) to prevent false positive leakage penalties.

---

### 5. Linn County Corn Revenue Protection Pricing Indication
* **Domain**: Agricultural Economics / Crop Insurance Actuarial Science / USDA Risk Management Agency (RMA)
* **Subject**: Linn County, Iowa Corn Revenue Protection (RP) crop insurance pricing indication model (`linn_county_corn_rp_pricing_indication.xlsx`).
* **Components**:
  - Actuarial historical yield distributions (APH - Actual Production History).
  - RMA projected price volatility factor calculation from CME December corn futures options.
  - Multi-peril crop insurance premium rate indications across coverage levels (65%, 70%, 75%, 80%, 85%).
  - County base rate, trend-adjusted yield endorsements, and federal premium subsidy schedules.
