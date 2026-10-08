# Rubric changes — adjudication item 4 only

Items 1, 2 and 3 were fixed in the **files** (no rubric change needed).
Item 4 requires **three new criteria** added to the end of the rubric.
Do **not** edit or delete any existing criterion — all 24 still grade correctly
against the revised golden.

---

## New Criterion 25

**Criterion text (copy/paste):**

```
The report determines that STRESS-001's P99 SLA failure invokes Policy §4.5.b, which permits the deal to proceed only as a Conditional Go and only where formal written risk acceptance is obtained from the VP of Engineering, and identifies VP Engineering Raj Malhotra's sign-off entry in the M&A Sign-off Tracker as that written risk acceptance.
```

**Weight:** `3`

Grounding in the golden: §2 Table row "P99 SLA (§4.5.b)" and §6 paragraph
"STRESS-001 — P99 SLA Failure (Policy §4.5.b)".
Grounding in the inputs: Policy §4.5.b + M&A Sign-off Tracker row
"VP Engineering | Raj Malhotra | Approved | 2026-08-20 | ...accepting P99 risk given test env scale."

---

## New Criterion 26

**Criterion text (copy/paste):**

```
The report assesses Policy §4.5.a P50/P95 SLA compliance across the stress-test scenarios rather than STRESS-001 alone, concluding that every scenario reporting numeric percentiles met its declared SLA metric at both the P50 and the P95 percentile.
```

**Weight:** `3`

Grounding in the golden: §2 Table row "SLA Thresholds P50/P95 (§4.5.a)" and the
§6 "§4.5.a — P50 / P95 Compliance Across All Scenarios" paragraph and table.

---

## New Criterion 27

**Criterion text (copy/paste):**

```
The report notes that STRESS-004 recorded its results qualitatively as "stable" rather than as numeric percentile values, so its Policy §4.5.a P50/P95 compliance cannot be numerically verified from the source data.
```

**Weight:** `1`

Grounding in the inputs: Stress Test Results sheet, STRESS-004 row has
`stable | stable | stable` in the P50/P95/P99 columns.

---

## Resulting rubric composition

| | before | after |
|---|---|---|
| positive criteria | 22 | 25 |
| negative criteria | 2 | 2 |
| max positive weight | 76 | 83 |

Composition stays healthy (negatives unchanged, minimal style/format weight),
so the "Rubric Quality Check" should continue to pass.
