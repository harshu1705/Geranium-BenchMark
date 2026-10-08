# Task 3: ALLL / CECL Reserve Adequacy Assessment (Tri-County Community Bank)

## 📌 Executive Summary & Task Overview
* **Domain**: Commercial Banking / Credit Risk / Regulatory Accounting
* **Target Role**: Senior Credit Risk Officer / ALLL Committee Chair
* **Subject Institution**: Tri-County Community Bank
* **As-Of Reporting Date**: Q4 2025 (December 31, 2025)
* **Regulatory Framework**: CECL (Current Expected Credit Losses) / ASC 326 (and legacy ASC 310-10-35 / ASC 450-20 alignment)
* **Core Deliverable**: `Tri_County_ALLL_Reserve_Adequacy_Memo_Q4_2025.docx`
* **Booked ALLL Balance**: $812,000 (GL Account 1720)
* **Board Tolerance Band**: +/- 10% ($730,800 to $893,200)

---

## 🎯 Purpose & Why This Task Was Designed
Commercial banks are required by federal regulators (FDIC, OCC, Federal Reserve) to maintain an Allowance for Loan and Lease Losses (ALLL) that accurately reflects expected credit losses. This task tests whether an AI agent can conduct an end-to-end reserve adequacy test across seven complex workbooks and documents without falling for misleading pre-calculated flags.

### Key Input Files
1. `ALLL_CECL_Reserve_Policy.docx`: The authoritative governance document detailing loan segmentation, pooling, Q-Factor triggers, collateral valuation haircuts, and reporting structure.
2. `ALLL_Loss_Rate_and_Reserve_Data.xlsx`: Contains the 12-quarter rolling historical loss rates, the 25 bps Q-factor rule, the 10% collateral liquidation haircut, and booked GL balances.
3. `Loan_Register_Q4_2025.xlsx`: Master commercial loan tape containing borrower details, loan balances, interest rates, collateral types, and internal risk ratings (1 to 6).
4. `Borrower_Financial_Statements_Summary.xlsx`: Underlying financial data (Operating Income, Debt Service, Current Assets, Current Liabilities, Net Worth).
5. `Appraisal_and_Collateral_Summary_Q4_2025.xlsx`: Certified commercial real estate appraisal values, appraisal dates, and collateral types.
6. `Covenant_Compliance_Workbook_Q4_2025.xlsx`: Pre-calculated covenant tracking sheets containing intentional stale flags.
7. `Exception_Tracking_Log_Q4_2025.xlsx`: Documentation of policy exceptions, technical defaults, and remediation statuses.

---

## 🧮 Quantitative Methodology & Built-in Challenges

### 1. The Pre-Calculation Trap (Covenant Audit)
The `Covenant_Compliance_Workbook` contains pre-calculated DSCR, LTV, and Current Ratio numbers. Section 5 of the policy expressly warns that compliance flags cannot be taken at face value. The analyst is required to recalculate:
* **Debt Service Coverage Ratio (DSCR)**: $\text{Annual NOI / Operating Cash Flow} \div \text{Annual Debt Service}$
* **Loan-to-Value (LTV)**: $\text{Current Principal Balance} \div \text{Appraised Collateral Value}$
* **Current Ratio**: $\text{Current Assets} \div \text{Current Liabilities}$

By independently recalculating, the analyst identifies true active breaches versus cured or technical exceptions, determining which risk rating pools receive the **25 basis point (0.25%) Q-Factor adjustment**.

### 2. Loan Portfolio Segmentation & Tier Reserves
Loans are segmented by Internal Risk Rating:
* **Tiers 1 to 5 (Homogeneous Pools)**: Each pool's reserve is calculated as:
  $$\text{Reserve} = \text{Pool Total Balance} \times (\text{Historical Loss Rate} + \text{Q-Factor Add-on})$$
* **Tier 6 (Individually Evaluated / Impaired Loan)**: The credit to **Summit Ridge Properties** ($1,170,000 balance, Rating 6, 90+ days past due) is carved out from pooled loss rates under ASC 310. It is evaluated via the **Collateral Dependent Method**:
  $$\text{Net Realizable Value (NRV)} = \text{Appraised Value} \times (1 - 0.10 \text{ Liquidation Cost Haircut})$$
  $$\text{Specific Reserve} = \text{Book Balance} - \text{NRV}$$

### 3. Reserve Adequacy Decision & Tolerance Band
* The total required reserve is summed across the 5 pooled tiers plus the specific reserve.
* Total required reserve is benchmarked against the booked balance ($812,000).
* If the required reserve falls outside the $730,800 – $893,200 band, an immediate quarterly ALLL provision adjustment is mandated and reported to the Board Audit Committee.

---

## 🛠️ Evolution & Bug Remediation History
1. **Borrower Name Consistency (`rename_borrowers.py`, `rename_borrowers2.py`)**:
   - In early test builds, subtle inconsistencies existed between borrower legal entity names in the Loan Register versus the Covenant Workbook (e.g., "Apex Industrial Logistics LLC" vs "Apex Logistics Corp"). Scripts were written to harmonize all 30+ entity names across 6 workbooks.
2. **Loss Rate History Depth (`geranium_scratch_fix_allldata.py`)**:
   - Reviewers noted that Sheet 1 of the Loss Rate workbook presented current rates without supporting trend data. The script generated an 8-quarter historical trend tab (`Historical Loss Rate Trend (8Q)`) and a dedicated `Methodology Notes` tab detailing the 12-quarter rolling charge-off derivation and Model Risk Committee validation dates.
3. **Report Generation & Word Formatting (`create_report.py`, `build_final8_golden.py`)**:
   - Programmatically constructed the 7-section Word report, adhering to strict typography, cell shading, and professional executive summary presentation.
4. **Iterative Rubric Calibration (`rubric.md` through `rubric_FINAL8.md`)**:
   - Rubric criteria were expanded to 35+ items in `rubric_FINAL8.md`, incorporating strict weights on exact intermediate mathematical values, exclusion of Rating 6 from pooled totals, and explicit negative scoring for relying on unverified covenant sheets.
