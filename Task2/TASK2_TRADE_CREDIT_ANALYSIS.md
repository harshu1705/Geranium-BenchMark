# Task 2: Trade Credit Underwriting & Risk Analysis (Hartsfield Polymer / Crestline)

## 📌 Executive Summary & Task Overview
* **Domain**: Commercial Finance / Corporate Credit Underwriting
* **Target Role**: Senior Credit Analyst
* **Subject Account**: Hartsfield Polymer Technologies, LLC
* **Creditor**: Crestline Chemicals & Coatings Corporation
* **As-Of Decision Date**: August 20, 2026
* **Request**: Increase unsecured trade credit limit from $400,000 (Net 30) to $1,500,000 (Net 60) to support an automotive contract expansion.
* **Core Deliverable**: `Trade_Credit_Decision_Memo.docx`
* **Final Underwriting Recommendation**: **DECLINE** (with strict prohibition against conditional approval under Section 4(b) of Crestline Credit Policy).

---

## 🎯 Purpose & Why This Task Was Designed
This task evaluates whether an AI reasoning agent can perform rigorous, audit-level credit risk analysis by cross-referencing five multi-modal, heterogeneous documents:
1. `Customer_Credit_Application_Hartsfield.docx`
2. `Hartsfield_Polymer_Unaudited_Financials.xlsx`
3. `Crestline_AR_Ledger_Hartsfield_History.xlsx`
4. `Dun_and_Bradstreet_Credit_Report.docx`
5. `Crestline_Corporate_Credit_Policy.docx`

### Key Evaluative Challenges & Built-in Traps
1. **Temporal Boundary Filtering**: The Crestline Credit Policy mandates that the internal payment history test uses a strict **3-month rolling window** (May 20, 2026 to August 20, 2026). Invoices outside this window (specifically invoice `INV-10611` dated May 15, 2026) must be excluded. Naive LLMs average all paid invoices across the entire ledger (~39.5 days) and falsely conclude that the payment history is acceptable. When correctly filtered, the true 3-month average Days to Pay is **45.25 days**, strictly breaching the 45.0-day maximum policy threshold.
2. **Altman Z-Score Quantitative Precision**: The agent must compute the private manufacturing Altman Z-score ($Z = 0.717A + 0.847B + 3.107C + 0.420D + 0.998E$). The calculated score is **1.77**, which falls below the mandatory 1.80 investment-grade threshold into the distress / caution zone.
3. **Public Record Disqualification**: The D&B report reveals an active **$45,000 State Tax Lien**. Under Section 3 of the policy, any unresolved tax lien exceeding $10,000 is an immediate disqualifier for unsecured credit lines over $1,000,000.
4. **Deterministic Policy Escalation**: Under Section 4(b), having two or more primary criterion failures or an unresolved tax lien strictly prohibits conditional approval or executive overrides—mandating an outright **DECLINE**.

---

## 📊 Core Financial & Credit Metrics

| Metric | Formula / Source | Calculated Value | Policy Requirement | Status |
| :--- | :--- | :--- | :--- | :--- |
| **Working Capital** | Current Assets ($4,115,200) - Current Liabilities ($3,184,500) | **$930,700** | Positive | Pass |
| **Current Ratio** | Current Assets / Current Liabilities | **1.29** | >= 1.25 | Pass |
| **Quick Ratio** | (Current Assets - Inventory $1,935,000) / Current Liabilities | **0.68** | >= 0.80 | **Fail** |
| **Debt-to-Equity** | Total Liabilities ($7,245,600) / Net Worth ($2,936,900) | **2.47** | <= 2.50 | Borderline Pass |
| **2025 Net Margin** | Net Income ($496,950) / Revenue ($10,435,200) | **4.76%** | > 3.0% | Pass |
| **H1 2026 EBIT Margin** | EBIT ($134,800) / Revenue ($5,432,000) | **2.48%** | Sharp drop from 7.72% | **Warning** |
| **H1 2026 Interest Coverage**| EBIT ($134,800) / Interest Expense ($124,500) | **1.08x** | Plunged from 7.29x (2025) | **Severe Risk** |
| **Altman Z-Score** | $0.717A + 0.847B + 3.107C + 0.420D + 0.998E$ | **1.77** | >= 1.80 | **Fail** |
| **A (WC / TA)** | $930,700 / $10,182,500 | 0.0914 | - | - |
| **B (RE / TA)** | $937,200 / $10,182,500 | 0.0920 | - | - |
| **C (EBIT / TA)** | $805,400 / $10,182,500 | 0.0791 | - | - |
| **D (Equity / TL)** | $2,936,900 / $7,245,600 | 0.4053 | - | - |
| **E (Sales / TA)** | $10,435,200 / $10,182,500 | 1.0248 | - | - |
| **3-Mo Avg Days to Pay** | Invoices May 20 – Aug 20, 2026 (Excludes INV-10611) | **45.25 Days** | <= 45.0 Days | **Fail** |
| **D&B Paydex Score** | Dun & Bradstreet Report | **68 (down from 74)** | >= 70 | **Fail** |
| **D&B Stress Percentile** | Financial Stress Score Percentile | **32nd Percentile** | Lower risk tier | High Risk |
| **Public Filings** | State of Ohio Department of Taxation Lien | **$45,000 Active** | Max $10,000 | **Disqualifying** |

---

## 🛠️ Technical Scripts & Artifacts in Folder
* `generate_docs.py` & `generate_excel.py`: Programmatic builders that create the ground-truth DOCX and XLSX source input files with exact styles, headers, and metadata.
* `rebuild_ar_ledger.py`: Generates the accounts receivable ledger containing both in-window and out-of-window invoices with specific due dates, payment dates, and payment turnaround times.
* `clean_dnb.py` & `rebuild_dnb.py`: Builds the D&B credit report, ensuring realistic commercial credit scores, trade payment experiences (falling from 75% to 42% on-time), and registered state tax liens.
* `fix_policy_window.py` & `fix_role_grounding.py`: Refines the credit policy text and ground-truth solution to enforce exact wording regarding the 3-month lookback window and the analyst's delegation of authority.
* `verify.py`: Automated testing script that verifies the arithmetic of the Altman Z-score, ratio calculations, and window boundaries across generated deliverables.
* `prompt_and_rubric_FINAL6.txt`: Definitive prompt and weighted 29-criterion rubric used for scoring agent performance.

---

## 🔍 Problems Encountered & What Was Solved
1. **Invoice Boundary Ambiguity**: Early versions had ambiguities around whether invoices dated exactly 90 days prior or paid 90 days prior were included. `fix_policy_window.py` cemented that the policy explicitly requires invoices *incurred and paid* within the exact May 20 – August 20 calendar window, requiring explicit exclusion of invoice `INV-10611` (May 15).
2. **Z-Score Formula Divergence**: Different corporate finance textbooks use varying coefficients for private versus public firms (Altman 1968 vs Altman 1983). The policy was aligned explicitly with the five-factor manufacturing formula, and acceptable tolerance bands (1.76 to 1.78) were incorporated into Criterion 7.
3. **Negative Rubric Scoring**: A negative criterion (Criterion 21, weight -4) was introduced to penalize agents that take the naive average of all paid invoices without scoping to the window.
