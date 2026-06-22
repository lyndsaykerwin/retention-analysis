---
name: retention-analysis
description: Use when the user wants an investor-grade retention analysis of recurring or re-occurring revenue from customer-level data — building Gross / Net / Logo retention metrics, an MRR-or-ARR corkscrew rollforward, and a formula-driven Excel deliverable where every number traces back to the source. Triggers on "retention", "churn", "NRR", "GRR", "logo retention", "ARR corkscrew", "customer retention", or when a user uploads customer-level revenue (wide customer × month or long/tidy form). Not for forecasting, LTV/CAC, cohort-by-acquisition-month curves, or consumer/transactional churn.
---

# retention-analysis

## Overview

Turn customer-level revenue into an investor-grade Excel deliverable whose every number is a live formula tracing back to the source. Change one customer's revenue in Raw Data and the whole model flexes.

The deliverable is **two or three sheets**:

- **Corkscrew** — the rollforward (period-by-period walk: beginning → new → expansion → contraction → churn → ending) plus the derived retention metrics (GRR, NRR, Logo).
- **Raw Data with Analysis** *(helper — present for workbook sources, skipped for tidy-CSV sources; see below)*.
- **Raw Data** — a verbatim copy of the source.

The deterministic work runs in the bundled Python scripts (`scripts/`), not by hand. Data can arrive wide (customers as rows, periods as columns) or long/tidy (one row per customer-period), as `.xlsx` or `.csv`.

## When the helper sheet appears (workbook → 3 sheets, tidy CSV → 2 sheets)

The helper ("Raw Data with Analysis") sits between the Corkscrew and Raw Data and gives the Corkscrew a clean, uniform grid to reference. `deliver.py` chooses automatically:

- **Source is an Excel workbook → a helper is built** (three sheets). It does as much or as little as the source needs:
  - **Aggregating** — the source needs real transformation: consolidate duplicate/renamed customers, sum multiple rows per customer (one per product/plan), or filter out-of-scope revenue types. The helper does this with `SUMIFS` and carries its own self-validation rows.
  - **1:1 pass-through** — the source already has one clean row per customer. The helper mirrors it with live references back to Raw Data and an "Excluded?" flag, normalizing the source's (often irregular) layout so the Corkscrew formulas stay uniform. It transforms nothing, but it standardizes position.
- **Source is a tidy long CSV → no helper** (two sheets). With no workbook layout to normalize, the Corkscrew references Raw Data directly.

So a clean single-row-per-customer **workbook still produces three sheets** (the pass-through helper). Tell the user the sheet count on that basis. (If you want a genuine two-sheet workbook deliverable, that's a future code change — today only the tidy-CSV path skips the helper.)

## Engineering principles

Apply DRY, YAGNI, and check-driven thinking to every cell: name the bug each non-data row catches or delete it; two paths to the same number is one check, not two; decorate only the title row and section banners — data rows are uniform.

---

## Critical Rules

### 1. Confirm scope once, upfront, then build
Retention fails silently when an early interpretive choice is wrong (misidentified customer column, ARR treated as MRR, "Total" rows counted as customers). Catch it once at the top, then build end-to-end without re-confirming mid-stream — the reconciliation checks at the end are the next user-facing checkpoint.

Ask **one consolidated question** (use your harness's structured-input mechanism) covering:
- **Which sheet** holds the raw data (and which to ignore)
- **Customer column, date columns, MRR vs ARR** — state your interpretation with 2-3 sample data points so the user can verify without opening the file
- **Revenue-type filter** (Recurring only / Recurring + Re-occurring / all)
- **Customer-unit definition** — aggregate to one row per customer, or keep line-level detail. Use whatever identifier the source actually has (a name, account number, or ID column).
- **Comparison period** — default year-over-year; see Rule 1a
- **Negative values in customer rows** — list each (customer, period, value) and ask: zero / treat as churn / leave (ignore negatives that sit in section/total rows — survey flags those separately and they're expected)
- **Actuals cutoff** — survey flags any month column that isn't a complete actual: the **current in-progress month** (today's month is never complete — the last complete month is the prior one) *and* any later forecast columns. Confirm where booked actuals end so the in-progress month and projections aren't counted as retention, then pass survey's `actuals_through` straight to `deliver.py --actuals-through` (e.g. `--actuals-through May-26`) — it drops those columns for you.

### 1a. Comparison period — default is year-over-year (YoY)
**Default to comparing periods one year apart** (this March vs last March) — right for businesses on **annual contracts**, where the renewal decision is annual, so a 12-month-apart comparison isolates true retention from seasonal noise. Offer **month-over-month (MoM)** if the user asks or the business runs on monthly contracts. Surface the default in the upfront question; let the user override (MoM / QoQ / YoY / LTM, or combinations).

### 2. Corkscrew is live formulas — no hardcoded values
The rollforward and the retention rates must be Excel formulas. The user must be able to change a customer's revenue and watch the corkscrew flex. Permitted hardcodes: raw revenue (the input), date headers, the ARR factor, methodology label cells. Everything else is a formula. If you catch yourself computing a sum or rate in Python and writing the result — stop, write a formula.

### 3. Reconcile against RAW data, not against derived sums
`Beginning + New + Upsell + Downsell + Churn = Ending` is tautological — Ending is *defined* as that sum, so it catches nothing. The real check compares two **independent** computations: **Corkscrew Ending(t) = sum-from-Raw-Data of in-scope revenue for period t × ARR_factor.**
- **Raw sums directly** (no transformation) → write the check straight against Raw Data: `=<Ending_cell> - SUMIFS('Raw Data'!<month_col>, 'Raw Data'!<type_col>, "Recurring") * $ARR_factor`.
- **Raw needs transformation** → that transformation lives on the helper sheet with its own check rows tying back to Raw Data via an independent path; the Corkscrew references the helper's pre-validated totals.
- **Multiple in-scope types** → show each type as its own row × ARR factor above ONE variance row (Variance = Ending − Sum of components). That variance IS the external check — don't also write a separate check row. One check, components visible above it.

### 4. Investigate failures before surfacing
If a check row is non-zero after recalc, diagnose it yourself — column-letter off-by-one? type-filter mismatch? customer row-range cutoff? Fix, re-run, then deliver. Only escalate genuine blockers (e.g. the source doesn't tie to its own stated totals).

### 5. Negatives: flag, don't coerce
Customer revenue shouldn't be negative (usually a refund or sign-flip). Scan before computing, list every negative **that falls in a customer row** in the upfront question, and ask how to handle. Negatives in section/total rows (e.g. a "Churn" or "Contraction" line) are expected — `survey.py` tags these separately, and you don't ask the user about them. Never silently coerce a customer-row negative.

### 6. Raw Data sheet is preserved verbatim — no exceptions
The Raw Data tab is an exact copy of the source: zero edits, no reformatting, no color changes, no reordering or renaming. Values, number formats, fonts, fills, borders, merged ranges, widths, heights, comments all preserved. It exists for user trust ("nothing was edited"). Apply the blue/green/black color convention only to the Corkscrew and helper sheets.

### 7. Formulas must be auditable — simple primitives
A formula that takes more than five seconds to parse is functionally wrong. Prefer `COUNTIF`/`SUMIFS` over `SUMPRODUCT`; compute a metric once and pull it with `HLOOKUP` or a direct reference rather than recomputing; decompose a sum-of-parts into visible rows. Keep `SUMPRODUCT` only where it genuinely earns it: differential row-level math across two periods (customers active in both *t* and *t-12*). Exact formula syntax lives in `reference/formulas-and-layout.md`.

### 8. Cell comments on every hardcoded input
Format `Source: [sheet]![cells], [description]`, e.g. `Source: Raw Data!B7:Q7, Customer 1 monthly MRR, Jan-25–Mar-26` or `Source: User-confirmed, MRR → annualization factor 12`. Add them as cells are populated, not at the end.

---

## Workflow

### Phase 1 — Survey the file, then confirm scope (ONE question)

**Start by running `survey.py` — it surveys the whole workbook in one pass.** It scores every sheet by name + structure, deep-inspects the best candidate, and reports the customer column, date columns, MRR-vs-ARR signal, revenue-type column, summary rows to exclude, and any negatives — as hypotheses for you to confirm. One script call replaces opening and scanning the file sheet by sheet.

```
python3 scripts/survey.py <path>            # human-readable
python3 scripts/survey.py <path> --json     # structured
python3 scripts/survey.py <path> --emit-config survey-config.json   # writes the source-sheet params for deliver.py
```

Pass that file to `deliver.py --config survey-config.json` and survey's confirmed findings (customer column, data-row range, first date column, **header row**, `actuals_through`) flow straight in — no hand-copied flags. Any explicit `deliver.py` flag still overrides the config.

**Prefer dispatching a subagent** to run `survey.py` (and read a few sample rows if needed) and return just its structured findings — the main thread stays lean. Reserve manual, cell-by-cell inspection for the rare file `survey.py` can't parse.

Then present **one consolidated confirmation question** (Critical Rule 1) with 2-3 sample data points so the user can verify without opening the workbook. **Wait for confirmation.**

### Phase 2 — Build the helper sheet *(only if the raw data needs transformation — see "When the helper sheet is needed")*

Skip this phase entirely when the raw data sums directly into the Corkscrew. When a helper *is* needed it normalizes the source to one row per customer per period:
- **Pass-through** — source already one row per customer; each helper cell is a live reference back to Raw Data, with an "Excluded?" column flagging summary/test rows (forced to 0 but kept for lineage).
- **Aggregating** — multiple rows per customer; each helper cell aggregates with `SUMIFS`.

Helper customer rows include only customers with at least one non-zero in-scope cell (customers with only excluded revenue are dropped — report the count). Sort by customer identifier deterministically (parse embedded numbers so "Account 17" sorts before "Account 4443"). When you filter + aggregate, the helper carries a small self-validation block that ties back to Raw Data before the Corkscrew references it. Exact layout: `reference/formulas-and-layout.md`.

### Phase 3 — Build the Corkscrew

Order: layout → headers → labels → blanks/dividers → formulas → formatting → comments → recalc. Lock all row positions before writing any formula. Exact layout and formulas: `reference/formulas-and-layout.md`.

**Comparison-period indexing — be precise.** For `M` months of source data and a lookback of `N`: number of comparison periods = `M − N`; first at source index `N` (the (N+1)th month); last at index `M − 1`. Example: Jan-21–Mar-24 (M=39), YoY (N=12) → 27 periods, first = Jan-22 (vs Jan-21), last = Mar-24 (vs Mar-23). Verify the first date-row label equals the (N+1)th source month — that catches off-by-ones immediately.

### Phase 4 — Number check, recalc once, validate, deliver

1. **Run `compute.py` as a fast number check — pass `--lookback` matching `deliver.py`** (default 12 = year-over-year). On the long-format CSV it computes, in well under a second with no spreadsheet recalc: the **lookback rollforward + GRR/NRR/Logo that match the Corkscrew period-for-period** (the `rollforward` and `metrics` keys — verified equal to the recalculated corkscrew to the cent), plus a month-over-month view (`monthly`/`metrics_monthly`), LTM cohort retention (`metrics_ltm`), and a 7-layer audit. The `rollforward`/`metrics` at your chosen lookback are your **expected values**: what every Corkscrew Beginning/Ending/external-check and headline retention number *should* equal. (Set `--lookback` to the SAME value you pass `deliver.py`; older versions computed only MoM+LTM, so on a YoY deliverable only the data-integrity layers matched — that gap is now closed.) Run it as soon as the long CSV exists.

2. **Recalculate the workbook with LibreOffice headless — once, not in a loop** (cross-platform):
   `time soffice --headless --calc --convert-to xlsx --outdir <dir> <path>`
   Keep the `time` — recalc is usually the slowest step of the whole run, and timing it is how you confirm where the wall-clock goes. If LibreOffice isn't installed, fall back to Excel manually, or the macOS `osascript` chain only on Mac.

3. **Re-open with `openpyxl.load_workbook(path, data_only=True)`** and confirm the cached values match the number check: every helper check row = 0 (if present); Corkscrew external check = 0 every period; decomposed variance = 0 (if present); headline endings/retention equal `compute.py`'s `rollforward`/`metrics` at the same `--lookback`; no `#REF!`/`#VALUE!`/`#NAME?`/`#DIV/0!` anywhere. Because step 1 already gave you the right answers, this is a single confirmation pass — **not** a recalc-diagnose-repeat loop.

4. **If a check is non-zero, the bug is in the formula generation, not the math** (the number check already verified the math). Diagnose against `compute.py`'s per-customer numbers (column-letter off-by-one, type-filter mismatch, row-range cutoff), fix, recalc once more.

5. **Spot-check 2-3 real customers** (one stable, one churned, one expanded) using the number check's per-customer classification as the expected answer, and reading the deliverable's cached cells (`data_only=True`) — see "Customer Spot-Check Trace."

6. **Show a one-screen reconciliation summary**: helper self-checks = 0 (if present); external check ties every period; decomposed variance = 0 (if present); spot-checked customers with cell refs; negatives handling; dropped-customer count. Then deliver.

**Why this order is fast:** the number check (step 1) is instant and tells you every right answer up front, so the slow LibreOffice recalc (step 2) runs **once** to populate display values — instead of a recalc → reopen → diagnose → recalc loop. Each script prints its own elapsed time; if a run feels slow, those plus the `time` on `soffice` pinpoint the culprit.

---

## Layouts & formatting (summary — full spec in `reference/`)

`deliver.py` builds the exact layout, formulas, and formatting automatically. The complete spec is in **`reference/formulas-and-layout.md`** — read it only if hand-building or repairing the workbook. The essentials:

- **Sheet order (left to right):** Corkscrew, [Raw Data with Analysis if built], Raw Data.
- **Corkscrew blocks, top to bottom:** title/ARR-factor → rollforward (Beginning, New, Upsell, Downsell, Churn, Ending, External Check) → customer counts → retention rates (GRR, NRR, Logo) → per-customer metrics → decomposed reconciliation (only when multiple in-scope types).
- **Text color convention:** blue = hardcoded input, green = reference to another sheet, black = formula within this sheet.
- **Fills:** navy `#1F4E79` banners (white bold), light-blue `#D9E1F2` sub-headers, medium-blue `#BDD7EE` only on the Beginning/Ending bookend rows. Check rows green when passing, red when failing — never ship red.
- **Numbers:** `$` only on the top and bottom row of a dollar block; interior rows `#,##0`; percentages `0.0%`; dates `mmm-yy`.
- **Never merge cells** (use `centerContinuous`). First N periods have no prior → leave those retention cells truly empty (no 0, no "N/A").

---

## Customer Spot-Check Trace

A spot-check is "show your work" — pick 2-3 real customers from THIS dataset (one stable, one churned, one expanded) and trace each end-to-end so the user can verify against the deliverable's cells.

1. **Read values from the deliverable, not Python.** After recalc, re-open `data_only=True` and read the cached helper-sheet values.
2. **Cite specific cells in both Raw Data and the helper**, e.g. *"row 55: N55 (Jan-22 in-scope MRR) = 2,669; Z55 (Jan-23) = 10,659; the Recurring line is Raw Data row 244."*
3. **MRR and ARR are different units — show the multiplication.** Write `MRR change +$7,990 → ARR +$7,990 × 12 = +$95,880.` Never label an MRR delta as "ARR" or skip the `× ARR_factor` step.

Example trace (per customer):

| Row in Raw Data | Product | Type | MRR period A | MRR period B |
|---|---|---|---:|---:|
| 241 | MergeLogic | Non-recurring → **excluded** | ~~$0~~ | ~~$0~~ |
| 243 | FlowBuilder | Re-occurring | $764 | $764 |
| 244 | InsightDash | Recurring | $1,905 | $9,895 |
| | **In-scope total** | | **$2,669** | **$10,659** |
| | **Helper cell** | | **N55 = 2,669** ✓ | **Z55 = 10,659** ✓ |

Then: `ΔMRR = +$7,990; ΔARR = +$7,990 × 12 = +$95,880 → Upsell.` If the trace and the helper disagree, do not ship — the mismatch is the bug.

---

## Common Mistakes

- **Hardcoded numbers in the corkscrew** → makes the model a screenshot. Use formulas (Rule 2).
- **"Total" rows treated as customers** → roughly doubles the corkscrew. Filter `Total`/`Grand Total`/`Sum`/`Subtotal`/`ACV` case-insensitively; report dropped rows.
- **Assuming MRR vs ARR from the header** → "Revenue" could be either. Use the numerical scale ($2,500/mo → MRR; $30,000 → ARR).
- **Tautological checks** → never check `Ending` against its own definition; compare to an independent path (Rule 3).
- **First-period retention shown as 0% / "N/A"** → leave those cells truly empty.
- **Color sprawl** → blues + grey + white only; red only for (never-shipped) failing checks.

---

## Bundled Scripts

Three Python scripts in `scripts/`. Run with `python3 scripts/<name>.py <args>`.

- **`survey.py`** — Phase 1. Two-pass: cheap structural scan of every sheet by name + shape, then deep-inspect the top candidate (and any runner-up within 80%). Outputs interpretation hypotheses, including: the real customer-row block (`first_row`/`last_row`, excluding section/total rows by label and dataless label rows — a row counts as a customer only when its name/ID is accompanied by at least one monthly value, so the block starts at the first real customer; position-agnostic, so it works whether section rows sit above, below, or among the customers); excluded non-customer clusters separated from the data by blank rows (a top/bottom summary block — reported, never silently swallowed) and any embedded column-total row caught by summation; an actuals flag marking the current in-progress month and any later forecast columns as not-complete-actuals (`actuals_through` = last complete month, i.e. the month *before* today's); and negatives tagged customer-row (needs a user decision) vs section-row (expected). `python3 scripts/survey.py <path> [--json]`
- **`compute.py`** — OPTIONAL independent cross-check. Recomputes the metrics in Python and runs a 7-layer self-check. The deliverable does **not** depend on it — `deliver.py` builds the workbook independently from the CSV. **Pass `--lookback` matching `deliver.py`** (default 12 = YoY): its `rollforward`/`metrics` keys then equal the Corkscrew period-for-period (verified against a real recalc), while `monthly`/`metrics_ltm` give MoM and LTM-cohort views. Run it to validate the numbers a second way, not as a required pipeline stage. `python3 scripts/compute.py <csv> [--arr-factor 12] [--lookback 12] [--output result.json] [--self-test]`
- **`deliver.py`** — builds the workbook (two- or three-sheet, mode chosen automatically) per the layout/formatting in `reference/`. Takes the ARR factor directly; `--compute-json` is optional (if given, its factor overrides `--arr-factor`); `--lookback` sets the comparison basis (default 12 = YoY, 1 = MoM). `python3 scripts/deliver.py <long.csv> <out.xlsx> --arr-factor 12 [--lookback 12] [--source <src.xlsx> --source-sheet "Raw Data" --source-customer-col B --source-first-data-row 8 --source-last-data-row 17 --source-first-date-col C --actuals-through 2026-05 --source-type-col D --type-filter "Recurring,Re-occurring"]`

  **`--source` contract — what passthrough/aggregating mode DOES and does NOT handle (read before relying on it):**
  - **Customer block bounds:** scanned from `--source-first-data-row` to `--source-last-data-row` (and down to the first blank / sheet max if the latter is omitted). Set **`--source-first-data-row`** to the first customer (skips a summary block ABOVE) and **`--source-last-data-row`** to the last customer (excludes a summary/total block BELOW). Feed both straight from survey's `customer_row_range`.
  - **Per-row exclusions (within the block):** still CSV-driven — a customer is excluded only by being **omitted from the long CSV** (it's then written below the block, zeroed, with an "Excluded?" flag for lineage). So an *embedded* total or a mid-block exclusion is handled by leaving that row out of the CSV.
  - **Months / actuals cutoff:** pass **`--actuals-through YYYY-MM`** (or a label like `May-26`) to drop the in-progress current month and any forecast tail — feed survey's `actuals_through` directly. (Without it, passthrough takes the month set from the long CSV, so you could alternatively omit those months from the CSV.)
  - **Net:** for the common shapes — a top/bottom summary block and a forecast/in-progress-month tail — `--source-first-data-row` / `--source-last-data-row` / `--actuals-through` handle it directly from survey's findings; no hand-shaped CSV needed. Only an *embedded mid-block* total/exclusion still requires omitting that one row from the CSV. `--source` supplies the live cell references and the external reconciliation check either way.

Each script has a built-in self-test: `python3 scripts/<name>.py --self-test` (survey/compute) runs against the bundled fixture.

---

## Final Checklist

- [ ] Scope confirmed via one upfront consolidated question (Phase 1)
- [ ] Sheets: Corkscrew + Raw Data, plus the helper *only if* the raw data needed transformation
- [ ] **Raw Data sheet is a verbatim copy — no edits, no reformatting, no color changes**
- [ ] All Corkscrew formulas live (no hardcoded sums or rates)
- [ ] Corkscrew external check = 0 in every period (vs Raw Data or pre-validated helper)
- [ ] Helper self-validation rows = 0 (if a helper was built)
- [ ] Decomposed reconciliation variance = 0 (if multiple in-scope revenue types)
- [ ] Cell comments on every hardcoded input
- [ ] 2-3 real customers spot-checked against the deliverable's cells
- [ ] Negatives reported and resolved; first-period retention cells empty
- [ ] Color palette: blues + grey + white; blue = hardcode, green = cross-sheet, black = formula

**File naming:** `<Company>_Retention_<YYYY-MM>_to_<YYYY-MM>.xlsx`
