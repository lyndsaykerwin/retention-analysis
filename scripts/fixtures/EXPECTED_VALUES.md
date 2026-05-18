# Expected values — `sample_retention_data.xlsx`

This document lists every hand-computed expected value from the synthetic retention fixture, with the math shown. It is the source of truth for the self-test assertions in `survey.py` and `compute.py`.

The next agent rewriting the self-test asserts should encode the **Top-line asserts** and **Survey asserts** sections verbatim, substituting the real call sites for the human-readable variables (`config.n_customers` → `result["config"]["n_customers"]`, etc.).

---

## 1. What this fixture represents

10 synthetic SaaS customers × 18 months of MRR (`Jan-25` through `Jun-26`). The dataset covers every retention bucket exactly once:

- **Flat** (customers 1, 2, 3) — constant MRR for the full period.
- **Upsell** (customer 4 stair-steps up; customer 9 mid-life upsell).
- **Downsell** (customer 5).
- **Churn** (customer 6 — never returns).
- **New customer joining mid-stream** (customers 7 and 9).
- **Churn + reactivation** (customer 8 — leaves, comes back later at a different price).
- **Downsell + recovery** (customer 10 — drops, then climbs back).

All customers appear as rows in Raw Data even in months where their MRR is zero. The dataset population is **10**. No negative values appear anywhere.

The source values are **MRR**. The ARR factor is **12** (multiply MRR × 12 to get ARR).

---

## 2. Customer behavior table

Re-stated from the spec, this is the ground truth.

| Customer ID | Behavior pattern | MRR by month |
|---|---|---|
| 1, 2, 3 | Flat | $1,000 every month, m1–m18 |
| 4 | Upsells twice | $500 m1–m5; $750 m6–m12; $1,000 m13–m18 |
| 5 | One downsell | $1,500 m1–m8; $1,000 m9–m18 |
| 6 | Churns at m11 | $800 m1–m10; $0 m11–m18 |
| 7 | Joins at m4 | $0 m1–m3; $500 m4–m18 |
| 8 | Churn + reactivate | $300 m1–m7; $0 m8–m13; $500 m14–m18 |
| 9 | Joins, then upsells | $0 m1–m6; $200 m7–m14; $400 m15–m18 |
| 10 | Downsell + recovery | $1,200 m1–m4; $700 m5–m11; $1,200 m12–m18 |

Month index: `m1 = Jan-25`, `m2 = Feb-25`, …, `m11 = Nov-25`, `m12 = Dec-25`, `m13 = Jan-26`, …, `m18 = Jun-26`.

---

## 3. Month-by-month rollforward (ARR scale)

All figures in dollars at **ARR scale** (MRR × 12). The classification rule for each customer's month-over-month change:

- `prev == 0` AND `curr > 0` → **new** (counts customer 7's m4 entry and customer 8's m14 reactivation as new)
- `prev > 0` AND `curr == 0` → **churn**
- `curr > prev` → **upsell** (amount = curr − prev)
- `curr < prev` (and both > 0) → **downsell** (amount = prev − curr)

Convention for the first month: `beginning_m1 = ending_m1` = total run-rate at start of period 1. No movements in m1 by definition.

**Sign convention in `compute.py` output:** `downsell` and `churn` are stored as **negative numbers**. The identity is `beginning + new + upsell + downsell + churn == ending`. Below I show absolute values for readability, but the negative-storage convention applies to assertions against `monthly[i]["downsell"]` and `monthly[i]["churn"]`.

| m  | Month   | Beginning | + New | + Upsell | − Downsell | − Churn | = Ending | n_active |
|----|---------|----------:|------:|---------:|-----------:|--------:|---------:|---------:|
| 1  | Jan-25  | 87,600 | 0 | 0 | 0 | 0 | 87,600 | 8 |
| 2  | Feb-25  | 87,600 | 0 | 0 | 0 | 0 | 87,600 | 8 |
| 3  | Mar-25  | 87,600 | 0 | 0 | 0 | 0 | 87,600 | 8 |
| 4  | Apr-25  | 87,600 | 6,000 | 0 | 0 | 0 | 93,600 | 9 |
| 5  | May-25  | 93,600 | 0 | 0 | 6,000 | 0 | 87,600 | 9 |
| 6  | Jun-25  | 87,600 | 0 | 3,000 | 0 | 0 | 90,600 | 9 |
| 7  | Jul-25  | 90,600 | 2,400 | 0 | 0 | 0 | 93,000 | 10 |
| 8  | Aug-25  | 93,000 | 0 | 0 | 0 | 3,600 | 89,400 | 9 |
| 9  | Sep-25  | 89,400 | 0 | 0 | 6,000 | 0 | 83,400 | 9 |
| 10 | Oct-25  | 83,400 | 0 | 0 | 0 | 0 | 83,400 | 9 |
| 11 | Nov-25  | 83,400 | 0 | 0 | 0 | 9,600 | 73,800 | 8 |
| 12 | Dec-25  | 73,800 | 0 | 6,000 | 0 | 0 | 79,800 | 8 |
| 13 | Jan-26  | 79,800 | 0 | 3,000 | 0 | 0 | 82,800 | 8 |
| 14 | Feb-26  | 82,800 | 6,000 | 0 | 0 | 0 | 88,800 | 9 |
| 15 | Mar-26  | 88,800 | 0 | 2,400 | 0 | 0 | 91,200 | 9 |
| 16 | Apr-26  | 91,200 | 0 | 0 | 0 | 0 | 91,200 | 9 |
| 17 | May-26  | 91,200 | 0 | 0 | 0 | 0 | 91,200 | 9 |
| 18 | Jun-26  | 91,200 | 0 | 0 | 0 | 0 | 91,200 | 9 |

**Rollforward identity verified:** `beg + new + upsell − downsell − churn == ending` for all 18 months.

### Narrative — how each non-zero movement was derived

- **m1 (Jan-25):** 8 customers active (1,2,3,4,5,6,8,10). Total MRR = $1,000×3 + $500 + $1,500 + $800 + $300 + $1,200 = **$7,300**. ARR = $7,300 × 12 = **$87,600**. (Customers 7 and 9 are not yet on the platform — their MRR is $0.)
- **m4 (Apr-25):** Customer 7 joins at $500 MRR. → New = $500 × 12 = **$6,000 ARR**. n_active grows from 8 to 9.
- **m5 (May-25):** Customer 10 drops from $1,200 → $700 MRR. Both >0, so this is a downsell of $500 × 12 = **$6,000 ARR**.
- **m6 (Jun-25):** Customer 4 upsells from $500 → $750 MRR. Upsell = $250 × 12 = **$3,000 ARR**.
- **m7 (Jul-25):** Customer 9 joins at $200 MRR. New = $200 × 12 = **$2,400 ARR**. n_active hits its peak of 10.
- **m8 (Aug-25):** Customer 8 churns ($300 → $0). Churn = $300 × 12 = **$3,600 ARR**. n_active back to 9.
- **m9 (Sep-25):** Customer 5 downsells ($1,500 → $1,000). Downsell = $500 × 12 = **$6,000 ARR**.
- **m11 (Nov-25):** Customer 6 churns ($800 → $0). Churn = $800 × 12 = **$9,600 ARR**. n_active drops to 8.
- **m12 (Dec-25):** Customer 10 recovers ($700 → $1,200). Treated as **upsell** (prev > 0), not new. Upsell = $500 × 12 = **$6,000 ARR**.
- **m13 (Jan-26):** Customer 4 upsells ($750 → $1,000). Upsell = $250 × 12 = **$3,000 ARR**.
- **m14 (Feb-26):** Customer 8 reactivates ($0 → $500). Treated as **new** (prev == 0). New = $500 × 12 = **$6,000 ARR**. n_active back to 9.
- **m15 (Mar-26):** Customer 9 upsells ($200 → $400). Upsell = $200 × 12 = **$2,400 ARR**.
- **m18 (Jun-26):** Ending ARR = total MRR ($1,000 × 3 + $1,000 + $1,000 + $0 + $500 + $500 + $400 + $1,200) × 12 = $7,600 × 12 = **$91,200**. n_active = 9 (everyone active except customer 6).

---

## 4. Top-line `compute()` asserts

After running `compute(rows, arr_factor=12.0)` on the long-format projection of this fixture:

| Path | Expected value | Notes |
|---|---|---|
| `config["n_customers"]` | `10` | dataset population, includes zero-revenue rows |
| `config["n_months"]` | `18` | |
| `config["arr_factor"]` | `12.0` | MRR input scaled to ARR |
| `config["month_range"]` | `["2025-01", "2026-06"]` | |
| `len(monthly)` | `18` | |
| `len(metrics_monthly)` | `17` | first month has `None` retention because no prior period |
| `metrics_ltm` | not None; `len == 6` | months m13..m18 each have an LTM (12-month-prior) comparison |
| `monthly[0]["beginning"]` | `87_600.00` | Jan-25 total MRR $7,300 × 12 |
| `monthly[0]["ending"]` | `87_600.00` | same as beginning — no movements in m1 |
| `monthly[0]["n_active"]` | `8` | customers 1,2,3,4,5,6,8,10 |
| `monthly[-1]["ending"]` | `91_200.00` | Jun-26 total MRR $7,600 × 12 |
| `monthly[-1]["beginning"]` | `91_200.00` | no movements in m17→m18 |
| `monthly[-1]["n_active"]` | `9` | everyone except customer 6 |
| `verification` | residuals all zero | rollforward identity holds every month |

If the existing self-tests use dotted access (`monthly[0].n_active`), translate the bracket forms above to attribute access. The dict structure is the API surface that `compute.py` currently returns.

---

## 5. Survey asserts

After running `survey.inspect_file("sample_retention_data.xlsx")`:

### Workbook-level

| Path | Expected value |
|---|---|
| `sheets[*].name` (the **inspected** sheets) | `["Raw Data"]` — only Raw Data is deep-inspected |
| `skipped_sheets[*].name` | contains `"Notes"` and `"Corkscrew"` |
| `near_tie` | `False` |
| `overall_sufficiency["verdict"]` | `"pass"` |

### `Raw Data` sheet detection

| Path | Expected value |
|---|---|
| `hypothesis["role"]` | `"source data"` |
| `hypothesis["shape"]` | `"sectioned-wide"` |
| `hypothesis["confidence"]` | `"high"` |
| `candidate_customer_column["col_letter"]` | `"B"` |
| `candidate_customer_column["col_index"]` | `2` |
| `candidate_date_columns["first_col_letter"]` | `"C"` |
| `candidate_date_columns["last_col_letter"]` | `"T"` |
| `candidate_date_columns["first_col"]` | `3` |
| `candidate_date_columns["last_col"]` | `20` |
| `candidate_date_columns["count"]` | `18` |
| `candidate_date_columns["range"]` | `"Jan-25 to Jun-26"` |
| `candidate_date_columns["header_row"]` | `7` |
| `scale_signal["verdict"]` | `"MRR (hypothesis)"` — **note the parenthetical**; use `startswith("MRR")` if you want to accept both this and a future `"MRR"` verdict |
| `derived_blocks_detected` | length `== 5` |
| `derived_blocks_detected[*]["label"]` | (in order) `"New customer MRR"`, `"Upsell MRR"`, `"Downsell MRR"`, `"Churn MRR"`, `"Check"` |
| `derived_blocks_detected[*]["col_range"]` | (in order) `"V-Y"`, `"AA-AD"`, `"AF-AI"`, `"AK-AN"`, `"AP-AS"` |
| `negative_values` | `[]` |
| `customer_count_hypothesis` | `10` |
| `sufficiency["verdict"]` | `"pass"` |

### `Corkscrew` sheet — must be skipped at Pass 1

- The workbook must contain a sheet named `"Corkscrew"`.
- It must appear in `report["skipped_sheets"]`, **not** in `report["sheets"]`.
- It must **never** be deep-inspected — the existing self-test for Pass 1 / Pass 2 logic relies on this.

### `Notes` sheet — also skipped

- Sheet name `"Notes"` must appear in `skipped_sheets`.
- Reason will be `"No date-like header row."` (since Notes only has one text cell at A1).

---

## 6. Workbook structure summary (for the assert-rewriter's reference)

```
sample_retention_data.xlsx
├── Raw Data        (1st sheet)
│   ├── Row 1, B1: "Sample SaaS Retention Workbook — Synthetic Fixture"
│   ├── Row 2, B2: "Generated for retention-analysis skill tests"
│   ├── Row 6: derived-block labels at V6, AA6, AF6, AK6, AP6
│   ├── Row 7: header row
│   │   ├── B7: "Customer ID"
│   │   ├── C7:T7: 18 date headers (Jan-25 → Jun-26) as real datetime values
│   │   └── V7:AS7: 5 derived-block date-header runs (each 4 columns wide)
│   └── Rows 8–17: 10 customer rows
│       ├── B8:B17: customer IDs 1..10
│       └── C8:T17: MRR values per the behavior table (use 0, not blank)
├── Corkscrew       (2nd sheet — finished-analysis dummy, must be skipped)
│   ├── B1: "ARR Corkscrew — Rollforward"
│   ├── A5..A11: "Beginning ARR", "+ New customer ARR", "+ Upsell",
│   │           "− Downsell", "− Churn", "Ending ARR", "External Check"
│   └── B4..D4 has 3 date headers + B5..D11 zero-filled
└── Notes           (3rd sheet — single-cell text)
    └── A1: free-text note pointing readers at this EXPECTED_VALUES.md doc
```

---

## 7. End-to-end check

The fixture was verified end-to-end:

1. Built the xlsx with `openpyxl`.
2. Re-opened it and confirmed: 3 sheets in the right order, C7:T7 holds 18 real datetime objects, B8:B17 holds customer IDs 1..10, spot-checked customer 6 m11 = 0 and customer 7 m1..m3 = 0 / m4 = 500.
3. Ran `survey.inspect_file(...)` and confirmed every field in §5 above matches.
4. Ran `compute(rows, arr_factor=12.0)` against a long-format projection of the customer behavior table and confirmed every field in §4 above matches. The full 18-month rollforward agreed with the hand calc in §3, and the identity `beg + new + upsell + downsell + churn == ending` holds in every month (downsell and churn stored as negative numbers).
