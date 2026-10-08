# Task 4: Commercial P&C Forensic Fire Claim Settlement (Bluegrass Precision Tooling)

## 📌 Executive Summary & Task Overview
* **Domain**: Property & Casualty (P&C) Commercial Insurance Claims / Forensic Claims Adjusting
* **Target Role**: Senior Commercial Property Claims Examiner
* **Insured**: Bluegrass Precision Tooling, Inc.
* **Carrier**: Heartland Mutual Insurance Company
* **Policy Form**: Commercial Property Coverage Form (Heartland Mutual CP-0010)
* **Claim Reference**: `CLM-2025-08841`
* **Event**: Structural fire originating in the primary CNC machining wing.
* **Core Deliverable**: `Claim_Settlement_Recommendation_CLM-2025-08841.docx`
* **Core Dispute**: Public Adjuster's Proof of Loss ($3.8M+ lump sum demand) vs Carrier's policy-governed settlement calculation.

---

## 🎯 Purpose & Why This Task Was Designed
This task tests an AI agent's ability to interpret nuanced commercial insurance contract law, apply mathematical coinsurance penalties, evaluate depreciation holdbacks, and calculate business interruption losses.

### Master Document Inventory
1. `Heartland_Mutual_Commercial_Property_Policy.docx`: Governs insuring agreements, exclusions, exceptions (ensuing fire), coinsurance clauses, Ordinance or Law endorsements, and valuation conditions.
2. `Public_Adjuster_Claim_Submission.docx`: Aggressive claim submission filed on behalf of the insured, asserting no coinsurance penalties, immediate RCV payout on machinery, and inflated business income projections.
3. `Origin_and_Cause_Investigation_Report.docx`: Certified fire investigator report establishing ignition from an unpermitted 2023 480V three-phase electrical splice.
4. `Contractor_Repair_Estimate.xlsx`: Line-item reconstruction estimate combining structural repairs, architectural upgrades, and NEC 2024 electrical code compliance additions.
5. `Business_Financial_Statements.xlsx`: 3-year historical monthly P&L and balance sheets, plus 6 months post-fire operating statements from a temporary warehouse.
6. `Statement_of_Values_Appraisal.xlsx`: Pre-loss certified replacement cost appraisal for the structure and CNC machinery assets.

---

## ⚖️ Key Legal, Contractual & Actuarial Traps

### 1. Ensuing Loss Doctrine (Faulty Workmanship Exception)
* **Investigation Finding**: The fire originated from faulty, unpermitted electrical work done in 2023.
* **Policy Analysis**: Section B.3 (Exclusions) excludes the cost to repair or replace faulty, inadequate, or defective workmanship. However, the policy contains a standard **Ensuing Loss Clause**: *"If loss or damage by a Covered Cause of Loss results, we will pay for that resulting loss or damage."*
* **Resolution**: The carrier does not pay to rewire the defective 2023 panel itself, but the entire resulting fire damage to the building, machinery, and inventory is covered.

### 2. Separation of Ordinance or Law (Code Upgrades)
* The contractor's estimate lumped $185,000 of mandatory 2024 energy and electrical code upgrades directly into structural framing line items.
* Under Section A.5(e), base Building coverage covers only **direct physical loss of like kind and quality**.
* Increased cost of construction required by enforcement of local building codes must be carved out and evaluated strictly under the **Ordinance or Law Endorsement**, which carries a distinct sub-limit and specific qualification triggers.

### 3. Independent Coinsurance Penalty Testing
The Public Adjuster argued that coinsurance should be waived due to the insured’s 10-year loss-free history. The policy contains no such provision.
* **Building Coinsurance**:
  $$\text{Required Insurance} = \text{Appraised Replacement Cost} \times 80\% \text{ Coinsurance Percentage}$$
  $$\text{Coinsurance Ratio} = \min\left(1.0, \frac{\text{Carried Building Limit}}{\text{Required Insurance}}\right)$$
  $$\text{Covered Loss} = (\text{Loss Amount} \times \text{Coinsurance Ratio}) - \text{Deductible}$$
* **Business Personal Property (BPP) Coinsurance**: Evaluated independently using appraised machinery and inventory values versus carried contents limits.

### 4. ACV vs RCV Valuation & Depreciation Holdback
* Under Section 7 (Valuation), Business Personal Property is initially settled on an **Actual Cash Value (ACV)** basis:
  $$\text{ACV} = \text{Replacement Cost Value (RCV)} - \text{Depreciation}$$
* The **depreciation holdback** is payable only upon submission of paid invoices proving actual repair or replacement within 180 days. Paying full RCV immediately violates the policy conditions.

### 5. Business Interruption: Actual Loss Sustained & Extra Expense
* Net Income + Continuing Normal Operating Expenses calculated by establishing the pre-loss revenue baseline and adjusting for seasonality.
* Non-continuing expenses (e.g., hourly labor not retained, raw aluminum billets not purchased) must be deducted.
* Expediting expenses incurred to lease temporary CNC equipment in Shelbyville are evaluated under **Extra Expense Coverage**, capped by the joint Business Income / Extra Expense limit.

---

## 🛠️ Build Scripts & Verification
* `build_docx_inputs.py`: Generates the complex multi-page policy, investigation report, and public adjuster submission.
* `build_xlsx_inputs.py`: Assembles the multi-tab contractor estimate, historical financial statements, and Statement of Values workbook.
* `build_pivot2_golden.py`: Solves all mathematical reconciliations, applies deductibles, executes coinsurance formulas, and formats the golden recommendation memo.
* `rubric_PIVOT2.md`: Enforces strict grading across all 5 coverage heads, ensuring negative marks if coinsurance is waived or code upgrades are blended.
