---
name: retention-analysis
description: Use when the user wants an investor-grade retention analysis of SaaS subscription revenue from customer-level data — building Gross / Net / Logo retention metrics, an MRR-or-ARR corkscrew rollforward, and a formula-driven Excel deliverable where every number traces back to the source. Triggers on "retention", "churn", "NRR", "GRR", "logo retention", "ARR corkscrew", "customer retention", or when a user uploads customer-level revenue (wide customer × month or long/tidy form). Not for forecasting, LTV/CAC, cohort-by-acquisition-month curves, or consumer/transactional churn.
---

# retention-analysis

## Overview

Turn customer-level revenue into an investor-grade Excel deliverable. The output is two or three sheets:

- **Corkscrew** — the rollforward (a period-by-period walk from beginning to ending revenue: new, expansion, contraction, churn) plus the derived retention metrics.
- **Raw Data with Analysis** *(optional helper)* — an intermediate sheet that filters and/or aggregates the raw data into one clean row per customer per period. **Only built when the raw data needs transformation.** When the raw data sums directly into the Corkscrew, this sheet is skipped and the deliverable is just two sheets.
- **Raw Data** — a verbatim copy of the source.

Every metric is a live formula referencing the helper (or, in the two-sheet case, Raw Data directly), which traces back to the raw data. Change one customer's revenue in Raw Data and the entire model flexes.

Deterministic math runs in Python (the bundled scripts in `scripts/`). The data can arrive in either wide form (customers as rows, periods as columns) or long/tidy form (one row per customer per period); it can be an Excel workbook or a CSV. Interpretation is done by reading sheet/column structure and a few sample rows — not by brittle heuristics. Python supports interpretation, not the other way around.

## Engineering principles

Apply DRY, YAGNI, and check-driven thinking to every cell: name the bug each non-data row catches or delete it; two paths to the same number is one check, not two; decorate only the title row and section banners — data rows are uniform.

## Critical Rules

### 1. ONE upfront confirmation, then build

Retention fails silently when an early interpretive choice is wrong (misidentified customer column, ARR treated as MRR, "Total" rows counted as customers). Catch it once at the top, then build.

Ask the user **one consolidated question** after reading the file — use whatever structured-input mechanism your agent harness provides (a single multi-part prompt). Confirm:

- **Which sheet** holds the raw data (and which to ignore)
- **Customer column, date columns, MRR vs ARR** — state your interpretation with 2-3 sample data points so the user can verify without opening the file
- **Revenue-type filter** (Recurring only / Recurring + Re-occurring / all)
- **Customer-unit definition** — aggregate to the customer level (one row per customer), or keep product-/line-level detail. Use whatever the source's customer identifier is (a name, an account number, an ID column); don't assume a specific column name.
- **Comparison period** — see Rule 1a below for the default
- **Negative values found** (list each one and ask: zero / treat as churn / leave)

One question, all sub-parts. Wait for confirmation, then proceed end-to-end. Do not re-confirm in the middle — the reconciliation checks at the end are the next user-facing checkpoint.

### 1a. Comparison period — default is year-over-year

**Default to comparing equivalent periods one year apart (year-over-year, "YoY")** — e.g. this March vs last March. This is the right comparison for businesses on **annual contracts**, where a customer's renewal decision happens once a year, so a 12-month-apart comparison isolates true retention from seasonal noise.

If the user asks for it — or the business runs on **monthly contracts** — offer **month-over-month ("MoM")** comparisons instead (this period vs the immediately prior period). Surface the default in the upfront question and let the user override (MoM / QoQ / YoY / LTM, or combinations).

### 2. Don't read every sheet exhaustively

When the workbook has multiple sheets, a descriptive sheet name *can* point you at the raw data — but don't rely on it. Some files have unhelpful, generic, or no meaningful sheet names. Read every sheet's **name** first; if a name clearly flags the raw data, open the title block + first ~10 rows of *that one sheet* to confirm. If the names give you nothing, fall back to a quick structural peek at each sheet (row counts, whether it has a customer column and date headers) and pick the candidate that looks like raw customer-level revenue. Either way, stop searching once the right sheet is found — do not recite the structure of every sheet just because it's there.

If two sheets are plausibly the source, briefly check both and surface the choice to the user. Don't deep-profile every sheet "to be safe."

### 3. Corkscrew is live formulas — no hardcoded values

The rollforward (Beginning, New, Upsell, Downsell, Churn, Ending) and the derived retention rates (GRR, NRR, Logo) must be Excel formulas referencing the helper sheet. The user must be able to change a customer's revenue in Raw Data and watch the corkscrew flex.

Permitted hardcodes: the raw customer-level revenue (the input), date headers, the ARR factor (e.g., 12 if MRR data), and methodology label cells. Everything else is a formula. If you catch yourself computing a sum or rate in Python and writing the result — stop, write a formula.

### 4. Reconcile against RAW data, not against derived sums

The identity `Beginning + New + Upsell + Downsell + Churn = Ending` is tautological — Ending is defined as that sum. It catches nothing.

The real reconciliation: **Corkscrew Ending(t) = sum-from-Raw-Data of in-scope revenue for period t, × ARR_factor.** This compares two independent computations.

- **If the raw data can be summed directly** (single in-scope revenue type, no per-customer aggregation needed) — write the check formula straight against Raw Data: `=<Ending_cell> - SUMIFS('Raw Data'!<month_col>, 'Raw Data'!<type_col>, "Recurring") * $ARR_factor`. No helper-sheet self-validation block needed.

- **If the raw data needs transformation first** (filter by type, aggregate multiple product rows per customer) — that transformation goes on the helper sheet as an explicit interim step with its own check rows. The helper's own check ties back to Raw Data via an independent path (typically a column sum). The Corkscrew then references the helper's pre-validated totals.

- **Decomposed display by revenue type** — when multiple revenue types are in scope (e.g. Recurring + Re-occurring), surface them as separate rows × ARR factor above ONE variance row at the bottom (Variance = Ending − Sum of components). That variance IS the primary external check — do not also write a separate "External Check" row above the rollforward. One check, with the components visible right above it.

### 5. Investigate failures before surfacing

If a reconciliation row is non-zero after recalc, do not just hand the file to the user with a red flag. Walk the diagnosis yourself: is it a column-letter off-by-one? A type-filter mismatch ("<>Non-recurring" vs "Recurring" only)? A customer row-range that cuts off rows? Fix it, re-run, then deliver. Only escalate genuine blockers (e.g., the source data itself doesn't tie to its own stated totals).

### 6. Negative values: flag, don't coerce

Customer revenue should not be negative. Negatives usually mean a refund, sign-flip, or special adjustment. Before computing, scan for negatives. List every one to the user (customer, period, value) in the upfront confirmation question and ask how to handle. Never silently coerce.

### 7. Raw Data sheet is preserved verbatim — no exceptions

The third sheet is an exact copy of the source workbook's chosen sheet. Zero edits. No reformatting. No color changes. No reordering columns. No renaming. Values, number formats, fonts, fills, borders, merged ranges, column widths, row heights, and cell comments all preserved.

This tab is for user trust ("nothing was edited") — the Corkscrew references the helper sheet, not this one directly. Apply your blue/green/black color convention only to the Corkscrew and helper sheets.

### 8. Formula auditability — simple primitives over clever SUMPRODUCT

The model is being read by a non-engineer auditing every cell. A formula that takes more than five seconds to parse is functionally wrong.

- **Counts:** `COUNTIF(range, ">0")`, not `SUMPRODUCT(--(range>0))`
- **Conditional sums:** `SUMIFS(...)`, not `SUMPRODUCT((cond)*values)`
- **Dynamic column-by-header lookup:** `SUMIFS(INDEX(wide_block, 0, MATCH(header_label, header_row, 0)), criteria_range, criterion)` — copies horizontally without hand-edits
- **Reference repeated metrics, don't recompute.** Compute once on the helper sheet, pull into the Corkscrew with `HLOOKUP` or a direct cell reference
- **Keep SUMPRODUCT only where it earns its keep:** differential row-level math across two ranges (e.g., customers active in both period *t* and period *t-12*). That's the one case SUMIFS can't handle.
- **Decompose multi-component math into separate rows.** If the final number is a sum of parts, surface the parts.

### 9. Cell comments on every hardcoded input

Format: `Source: [sheet name]![cell range], [description]`. Examples:
- `Source: Raw Data!B7:Q7, Customer 1 monthly MRR, Jan-25 through Mar-26`
- `Source: User-confirmed, MRR data → annualization factor 12`
- `Source: User-confirmed, comparison period = YoY (T vs T-12)`

Add comments as cells are populated, not at the end.

---

## Workflow

### Phase 1 — Identify the data and confirm scope

1. Open the file. Read every sheet's **name** (not contents).
2. Pick the obvious raw-data candidate by name. Read its title block (rows above the data) + first ~10 rows of the data itself.
3. Identify: customer column, date columns, MRR vs ARR (use the numerical scale — $2,500 cells are MRR; $30,000 for the same customer is ARR), any revenue-type column, any obvious total/summary rows to exclude.
4. Scan for negatives.
5. Present to the user — one consolidated question with all sub-parts (per Critical Rule 1), using whatever structured-input mechanism your harness provides. Include 2-3 sample data points: one customer's full row, the date range, and a small slice of the values, so the user can confirm without opening the workbook.

Wait for confirmation.

### Phase 2 — Build the helper sheet (only if the raw data needs transformation)

**The helper sheet ("Raw Data with Analysis") is optional.** Build it only when the raw data isn't already in a clean one-row-per-customer-per-period shape — i.e. when you need to **filter** out-of-scope revenue types and/or **aggregate** multiple source rows into a single per-customer figure. If the raw data sums directly into the Corkscrew with no transformation, skip this phase entirely; the deliverable is just Corkscrew + Raw Data, and the Corkscrew references Raw Data directly.

When a helper *is* needed, it normalizes the source into one row per customer per period. Two common cases:

- **Pass-through** — the source already has one row per customer. Each helper cell is a live reference back to Raw Data (e.g. `=IF('Raw Data'!C8="","",'Raw Data'!C8)`), with an "Excluded?" column flagging summary/test rows; flagged rows are forced to 0 but kept for lineage.

- **Aggregating** — the source has multiple rows per customer (for example one row per product, plan, or line item). Each helper cell aggregates with `SUMIFS`. For wide source blocks, the customer-row formula uses a **direct column reference** for recalc speed (many customer rows × many period columns); the summary-block formulas at the top of the helper use `SUMIFS(INDEX(...), MATCH(...))` so they copy horizontally without hand-editing.

**Helper-sheet customer rows include only customers with at least one non-zero in-scope revenue cell.** Customers that exist in Raw Data but have ONLY excluded revenue types are filtered out (they'd just be rows of zeros). Report the dropped count to the user in the final summary.

**Sort order:** by the customer identifier, in a deterministic order — never source-row order. If the identifiers embed a number (e.g. "Account 17", "Account 4443"), parse the number and sort numerically rather than as text.

**Self-validation block (only when raw needs transformation):** when you filter by type and aggregate per customer, the helper validates itself before the Corkscrew references it. Stack a few check rows at the top — see "Helper Sheet Layout" below. Skip this block when the raw data sums directly into the Corkscrew check without transformation.

### Phase 3 — Build the Corkscrew

Build in this order: layout → headers → labels → blank rows and dividers → formulas → formatting → comments → recalc.

**Comparison-period indexing — be precise.** For monthly source data with `M` months and a lookback of `N` periods:

- Number of comparison periods = `M − N`
- First comparison period at source index `N` (the (N+1)th month)
- Last comparison period at source index `M − 1`

Example: Jan-21 to Mar-24 (M=39 months), YoY (N=12) → **27 comparison periods**, first = Jan-22 (vs Jan-21), last = Mar-24 (vs Mar-23). Verify the first period label in the date row equals the (N+1)th month of source data — that catches off-by-ones immediately.

### Phase 4 — Recalc, validate, deliver

1. **Recalculate the workbook with LibreOffice headless** (cross-platform — works on Mac and Linux):

   ```
   soffice --headless --calc --convert-to xlsx --outdir <dir> <path>
   ```

   If LibreOffice is not installed, fall back to opening manually in Excel and saving, or use the Mac-specific `osascript` chain only if running on macOS. LibreOffice is the default because the skill should not silently fail on a non-Mac machine.

2. **Re-open with `openpyxl.load_workbook(path, data_only=True)`** and verify:
   - All check rows on the helper sheet (if present) = 0 across every month
   - Corkscrew external check row = 0 across every period
   - Decomposed reconciliation variance (if present) = 0 across every period
   - No `#REF!`, `#VALUE!`, `#NAME?`, `#DIV/0!` strings anywhere

3. **If any check is non-zero, diagnose and fix before delivery** (per Critical Rule 5). Common causes: column-letter off-by-one in the formula generation, scope-filter mismatch, customer-row range bound that cuts off rows, helper sheet sorted differently from raw.

4. **Spot-check 2-3 real customers from this dataset** (one stable, one churned, one expanded). Read the values from the deliverable (`data_only=True`), not from Python — the user's question is "does the deliverable say what you think it says." Trace each customer's MRR figure back to specific cells in both the helper and Raw Data. See "Customer Spot-Check Trace" below.

5. **Show the user a one-screen reconciliation summary:**
   - Raw Data with Analysis self-checks (if present): all = 0 ✓
   - Corkscrew external check: ties in every period ✓
   - Decomposed reconciliation (if present): variance = 0 ✓
   - Spot-checked customers: list with cell references
   - Negatives in source: [none / handled per user's choice]
   - Dropped customer count: [N of total were all-excluded-type and filtered]

Then deliver.

---

## Layouts

### Corkscrew Sheet

Sheets in this order from leftmost: **Corkscrew**, **Raw Data with Analysis** *(present only when a helper is built — see Phase 2)*, **Raw Data**. When no helper is needed, it's just Corkscrew and Raw Data.

Define all row positions before writing any formula — a formula written before the layout is locked points to the wrong row when a later header insertion shifts everything down.

```
Row 1     Title: "<Company> ARR Corkscrew — Retention Analysis"
          centerContinuous, navy fill (#1F4E79), white bold text
Row 2     "Generated:" | timestamp
Row 3     "ARR Factor (MRR × N):" | factor value (BLUE — hardcode)
Row 5     Date headers across columns (Jan-22, Feb-22, …)
Row 6     Optional prior-period reference label ("vs Jan-21", …) for YoY layouts

Rollforward block
Row 8     Beginning ARR              [formula → helper sheet × $ARR_factor]
Row 9       + New customer ARR       [formula]
Row 10      + Expansion (Upsell)     [formula]
Row 11      − Contraction (Downsell) [formula, stored negative]
Row 12      − Churn                  [formula, stored negative]
Row 13    Ending ARR                 [= rows 8+9+10+11+12]
Row 14    External Check (= 0)       [= row13 − sum-from-Raw-Data × $ARR_factor]
                                      OR if helper has self-validation:
                                      [= row13 − helper-sheet-validated-total × $ARR_factor]

Customer count block
Row 16    SECTION BANNER "CUSTOMER COUNTS"
Row 17    # Active (prior period)    [HLOOKUP into helper]
Row 18    # Active (current)         [HLOOKUP into helper]
Row 19    # Churned                  [prior active − retained]
Row 20    # New                      [current active − retained]

Retention metrics block
Row 22    SECTION BANNER "RETENTION RATES"
Row 23    Gross Dollar Retention (GRR)   [= (Beg + Downsell + Churn) / Beg]
Row 24    Net Dollar Retention (NRR)     [= (Beg + Upsell + Downsell + Churn) / Beg]
Row 25    Logo Retention                 [= (Active prior − Churned) / Active prior]

Per-customer metrics
Row 27    SECTION BANNER "PER-CUSTOMER METRICS"
Row 28    Avg ARR per Active Customer    [= Ending ARR / # Active current]
Row 29    Avg ARR per New Customer       [= New ARR / # New]

Decomposed reconciliation (ONLY when multiple in-scope revenue types)
Row 31    SECTION BANNER "RECONCILIATION CHECKS"
Row 32    Recurring ARR              [= helper row 6 × $ARR_factor]
Row 33    Re-occurring ARR           [= helper row 7 × $ARR_factor]
Row 35    Sum customer ARR           [= row 32 + row 33]
Row 36    Variance vs Ending ARR     [= row 35 − row 13]   must = 0
```

Optional LTM corkscrew block (rows 38–48) with the same shape, comparison T-12, when LTM is part of the methodology.

### Helper Sheet (Raw Data with Analysis)

Summary block on top (only when self-validation is needed per Critical Rule 4), customer data below.

```
Row 1   Month headers           "2021-M1" … "2024-M3". Col A label = "Customer ID"
Row 2   # Active customers      = COUNTIF(<col>$12:<col>$<last>, ">0")
Row 3   # Retained vs prior     For first N columns (N = lookback) the value is "n/a"
                                — no prior period yet. From col N+1 onward, array formula:
                                = SUMPRODUCT((<curr>$12:<curr>$<last>>0) *
                                             (<prior>$12:<prior>$<last>>0))
                                Only place SUMPRODUCT is needed.
Row 4   Check # Active vs Raw   Independent recount directly from Raw Data, must = 0
Row 5   blank divider
Row 6   Recurring MRR total     SUMIFS(INDEX(Raw Data block, 0, MATCH(col$1, header_row, 0)),
                                       type_col, "Recurring")
Row 7   Re-occurring MRR total  Same pattern, "Re-occurring"
Row 8   Non-recurring MRR total Same pattern, "Non-recurring"
                                (Keep even when out of scope — needed for full-coverage recon)
Row 9   Total MRR (all types)   = <col>6 + <col>7 + <col>8
Row 10  Check vs Raw Data       = <col>9 − SUM('Raw Data'!<month_col>)   must = 0
Row 11  Check (Rec + Re-occ)    = (<col>6 + <col>7) − SUM(customer rows)  must = 0
Row 12+ Customer data           Col A = Customer ID. Each month cell uses DIRECT column
                                reference (not INDEX/MATCH) — thousands of rows × dozens
                                of columns, recalc speed matters:
                                = SUMIFS('Raw Data'!$<month>$<first>:$<month>$<last>,
                                         'Raw Data'!$<cust>$<first>:$<cust>$<last>, $A<row>,
                                         'Raw Data'!$<type>$<first>:$<type>$<last>, "<filter>")
```

**Formula style summary:**
- Summary block rows 6-8: `INDEX/MATCH` dynamic column lookup → easier to extend, formula is the same in every column except `<col>$1`
- Customer rows 12+: direct column reference → recalc speed
- Check rows 4, 10, 11: direct column reference → one-off, not a copyable pattern

**Freeze panes** at `B12` so labels and summary stay visible while scrolling.

### Raw Data Sheet

A verbatim copy of the source workbook's chosen sheet. No edits, no reformatting, no color changes. Preserve values, number formats, fonts, fills, borders, merged ranges, column widths, row heights, and cell comments.

This tab is for user trust. The Corkscrew references the helper, not this sheet directly.

---

## Formula patterns

All Corkscrew formulas reference the helper sheet (Raw Data with Analysis). The helper has month headers in row 1, customer data starting at row 12. For each Corkscrew comparison-period column, you need the **current period** and the **prior period** helper columns.

**Column mapping — work this out before writing any formulas.** For YoY (12-month lookback) over a 39-month source dataset:

| Corkscrew col | Period label | Helper current col | Helper prior col |
|---|---|---|---|
| C | 2022-M1 (idx 12) | N (idx 12) | B (idx 0) |
| D | 2022-M2 (idx 13) | O (idx 13) | C (idx 1) |
| AC | 2024-M3 (idx 38) | AN (idx 38) | AB (idx 26) |

Rule: for Corkscrew column at offset `i` from the first comparison-period column, helper current is at month-index `lookback + i`, helper prior is at month-index `i`. The helper's first month column is B; corresponding helper column letter is `get_column_letter(2 + month_index)`.

**Rollforward formulas** (`<curr>` and `<prior>` are helper column letters from the mapping):

```
Beginning ARR  =SUMPRODUCT(('Raw Data with Analysis'!<prior>$12:<prior>$<last>>0)*
                          'Raw Data with Analysis'!<prior>$12:<prior>$<last>)*$C$3

New ARR        =SUMPRODUCT(('Raw Data with Analysis'!<prior>$12:<prior>$<last>=0)*
                          ('Raw Data with Analysis'!<curr>$12:<curr>$<last>>0)*
                          'Raw Data with Analysis'!<curr>$12:<curr>$<last>)*$C$3

Upsell         =SUMPRODUCT(('Raw Data with Analysis'!<prior>$12:<prior>$<last>>0)*
                          ('Raw Data with Analysis'!<curr>$12:<curr>$<last>>'Raw Data with Analysis'!<prior>$12:<prior>$<last>)*
                          ('Raw Data with Analysis'!<curr>$12:<curr>$<last>-'Raw Data with Analysis'!<prior>$12:<prior>$<last>))*$C$3

Downsell       =SUMPRODUCT(('Raw Data with Analysis'!<prior>$12:<prior>$<last>>0)*
                          ('Raw Data with Analysis'!<curr>$12:<curr>$<last>>0)*
                          ('Raw Data with Analysis'!<curr>$12:<curr>$<last><'Raw Data with Analysis'!<prior>$12:<prior>$<last>)*
                          ('Raw Data with Analysis'!<curr>$12:<curr>$<last>-'Raw Data with Analysis'!<prior>$12:<prior>$<last>))*$C$3

Churn          =SUMPRODUCT(('Raw Data with Analysis'!<prior>$12:<prior>$<last>>0)*
                          ('Raw Data with Analysis'!<curr>$12:<curr>$<last>=0)*
                          (-'Raw Data with Analysis'!<prior>$12:<prior>$<last>))*$C$3

Ending         =<col>8+<col>9+<col>10+<col>11+<col>12

External Check =<col>13 - (independent_sum_path × $C$3)
               // independent_sum_path = SUMIFS on Raw Data when no transformation,
               // OR SUM of helper rows 6+7 when type-filter & per-customer aggregation needed
```

**Customer-count formulas** (HLOOKUP — simple, deterministic, easy to audit):

```
# Active prior    =HLOOKUP(SUBSTITUTE(<col>$6,"vs ",""),'Raw Data with Analysis'!$B$1:$<last>$2, 2, FALSE)
# Active current  =HLOOKUP(<col>$5, 'Raw Data with Analysis'!$B$1:$<last>$2, 2, FALSE)
# Churned         =<col>17 - HLOOKUP(<col>$5, 'Raw Data with Analysis'!$B$1:$<last>$3, 3, FALSE)
# New             =<col>18 - HLOOKUP(<col>$5, 'Raw Data with Analysis'!$B$1:$<last>$3, 3, FALSE)
```

**Retention metrics** (all use `IFERROR` so empty-prior periods don't error):

```
GRR    =IFERROR((<col>8 + <col>11 + <col>12) / <col>8, 0)
NRR    =IFERROR((<col>8 + <col>10 + <col>11 + <col>12) / <col>8, 0)
Logo   =IFERROR((<col>17 - <col>19) / <col>17, 0)
```

---

## Formatting Standards

### Cell colors (standard finance convention)

- **Blue text (RGB 0,0,255)** — hardcoded inputs (ARR factor, raw data values when displayed, date headers, methodology label values)
- **Green text (RGB 0,128,0 or #006100)** — references to another sheet (cells that pull from Raw Data or the helper sheet)
- **Black text** — formulas computed within the current sheet

### Fill colors

- Section header banner rows — dark blue `#1F4E79` with white bold text
- Sub-headers / row labels — light blue `#D9E1F2` with black bold
- Rollforward anchor rows (Beginning ARR, Ending ARR) — medium blue `#BDD7EE` with black bold. Apply to all data cells in those two rows as visual bookends of the rollforward block. Do NOT apply to retention rates, per-customer metrics, or other rows — sibling metrics should be visually uniform.
- Check rows: green text on white when passing, red text when failing. Never ship a red check.

No greens / yellows / oranges in the model body. Reserve red for failed checks (which should never ship) and green text only for passing checks and cross-sheet references.

### Number formats — dollar signs on top and bottom of a block, not every cell

Standard finance convention: in a vertical block of dollar values, only the **top row** and the **bottom (total / output) row** show the `$` symbol. Middle rows show numbers without the symbol. Same on the helper sheet's customer-by-month data table — interior cells are `#,##0`; only the totals row at the bottom carries `$`.

- **Top of block (e.g., Beginning ARR) and bottom (Ending ARR):** `"$"#,##0;("$"#,##0);"-"`
- **Middle rows (New, Upsell, Downsell, Churn, helper customer cells):** `#,##0;(#,##0);"-"`
- **Percentages:** `0.0%;(0.0%);"-"` — one decimal, parens for negative, dash for zero
- **Customer counts:** `#,##0;(#,##0);"-"`
- **Dates:** `mmm-yy` (matches "Jan-25" style)

### Headers

- Title row: navy fill (`#1F4E79`), white bold, `centerContinuous` alignment
- Date row: same navy fill, white bold, center-aligned
- Sub-header rows (e.g., "vs prior year"): light-blue fill (`#D9E1F2`), no bold

### Unit labels (always present, never inferred)

The ARR factor cell label reads `"ARR Factor (MRR × N):"` not just `"ARR Factor"`. The title says `"YoY ARR Corkscrew"` or `"Monthly ARR Corkscrew"` — the comparison period is in the title. Currency unit on the title or in a unit cell: `"All figures in $USD"`.

### Other

- **Never merge cells.** Use `Alignment(horizontal="centerContinuous", vertical="center")` applied to every cell in the span, with text written only to the leftmost. Merged cells break selection, sorting, filtering, copy/paste.
- **Borders:** thick (1.5pt) around the rollforward block, retention metrics block, reconciliation block; thin (0.5pt) on data tables.
- **Column widths:** label column ~38, data columns ~13.
- **Freeze panes** on the Corkscrew at the first data column / first data row (typically `C7`).

---

## Customer Spot-Check Trace

A spot-check is the analyst's "show your work" — pick 2-3 real customers from THIS dataset (one stable, one churned, one expanded), trace each end-to-end, let the user verify against the deliverable's cells.

**Three rules:**

1. **Read values from the deliverable, not Python.** After recalc, re-open with `openpyxl.load_workbook(path, data_only=True)` and read the cached values from the helper sheet. The user's question is "does the deliverable say what you think it says" — so the trace must read the deliverable.

2. **Cite specific cells in both Raw Data and the helper.** Illustrative example: `"This customer lives in Raw Data with Analysis row 55. N55 (Jan-22 in-scope MRR) = 2,669; Z55 (Jan-23) = 10,659. The Recurring line driving the expansion is Raw Data row 244."` The user must be able to click directly to the cell.

3. **MRR and ARR are different units — show the multiplication step.** Write: `MRR change: +$7,990. ARR impact: +$7,990 × 12 = +$95,880.` Never label an MRR delta as "ARR" or skip the `× ARR_factor` step.

**Trace format** (per customer):

| Row in Raw Data | Product | Type | MRR period A | MRR period B |
|---|---|---|---:|---:|
| 241 | MergeLogic | Non-recurring → **excluded** | ~~$0~~ | ~~$0~~ |
| 243 | FlowBuilder | Re-occurring | $764 | $764 |
| 244 | InsightDash | Recurring | $1,905 | $9,895 |
| | **In-scope total** | | **$2,669** | **$10,659** |
| | **Helper cell** | | **N55 = 2,669** ✓ | **Z55 = 10,659** ✓ |

Then narrate the classification: `ΔMRR = +$7,990; ΔARR = +$7,990 × 12 = +$95,880 → Upsell.`

If the trace and the helper disagree, do not ship. The mismatch is the bug — find it (column-letter off-by-one, type filter mismatch, customer ID typo, row range cutoff) and fix the formula.

---

## Common Mistakes

**Hardcoded numbers in the corkscrew.** Writing `ws['G8'] = 1127918.40` makes the model a screenshot. Use a formula referencing the helper.

**Treating "Total" rows as customers.** A row with `Customer ID = "Total MRR"` will roughly double the corkscrew. Filter out summary labels (`Total`, `Grand Total`, `Sum`, `Subtotal`, `ACV`, etc.) case-insensitively. Report dropped rows.

**Silently coercing negatives.** A negative MRR (refund or sign-flip) silently inflates the next period's churn. Flag every negative to the user before computing.

**Assuming MRR vs ARR from the column header alone.** "Revenue" could be either. Use the numerical scale as the stronger signal ($2,500/month → MRR; $30,000/month → ARR).

**Tautological self-referencing checks.** A check that subtracts `(Beginning + New + Upsell + Downsell + Churn)` from `Ending` is zero by construction. Compare against an independent path — Raw Data or a pre-validated helper sheet total.

**First-period retention shown as 0% or "N/A".** The first N periods have no prior — there is no retention rate. Leave those cells truly empty (no value, no text).

**Color sprawl.** Blues + grey + white only. Reserve red for failed checks (don't ship them) and green for passing checks and cross-sheet references.

**Building end-to-end before confirming scope.** A wrong customer-column guess at the start wastes the user's review at the end. The single upfront confirmation question is the cheapest catch.

---

## Bundled Scripts

Three Python scripts in `scripts/`. Run with `python3 scripts/<name>.py <args>`.

### `scripts/survey.py` — Phase 1 helper

Opens an `.xlsx` or `.csv`. Two-pass design:

1. **Cheap structural scan** of every sheet by name (sheet-name hints like "Raw" / "MRR" lift the score; "Dashboard" / "Notes" / "Cover" lower it) plus row counts and date-header detection.
2. **Deep inspect** only the highest-scoring sheet (plus any runner-up scoring within 80% of the top).

Outputs hypotheses for the model to confirm with the user.

CLI: `python3 scripts/survey.py <path> [--json]`

### `scripts/compute.py` — Phase 2 worker

Input: long-format CSV `(customer_id, period, mrr)` plus config (ARR factor, exclusions, period definitions). Output: JSON with per-customer-period classifications, period-level aggregates, retention metrics, LTM metrics (if ≥13 periods), verification status.

CLI: `python3 scripts/compute.py <long-format-csv> [--arr-factor 12] [--output result.json]`

### `scripts/deliver.py` — Phase 3 worker

Input: `compute.py` output JSON + raw long-format data + output path. Writes the three-sheet workbook per the layouts above. Includes the helper self-validation block (when applicable), the decomposed reconciliation block (when multiple in-scope revenue types), and cell comments on every hardcoded input.

CLI:
```
python3 scripts/deliver.py <compute-output.json> <long-format-csv> <output.xlsx> \
    --source path/to/source.xlsx \
    --source-sheet "Raw Data" \
    --source-customer-col B \
    --source-first-data-row 8 \
    --source-first-date-col C
```

---

## Final Checklist

Before claiming done:

- [ ] Scope confirmed via one upfront consolidated question (Phase 1)
- [ ] Sheets in the deliverable: Corkscrew + Raw Data, plus the Raw Data with Analysis helper *only if* the raw data needed transformation
- [ ] **Raw Data sheet is a verbatim copy — no edits, no reformatting, no color changes**
- [ ] All Corkscrew formulas live (no hardcoded sums or rates)
- [ ] Corkscrew external check = 0 in every period (vs. Raw Data or pre-validated helper)
- [ ] Helper self-validation rows = 0 (if helper requires transformation)
- [ ] Decomposed reconciliation variance = 0 (if multiple in-scope revenue types)
- [ ] Cell comments on every hardcoded input
- [ ] 2-3 real customers spot-checked end-to-end against the deliverable's cells
- [ ] Negatives in source: reported and resolved
- [ ] First-period retention cells: empty (not 0, not "N/A")
- [ ] Color palette: blues + grey + white only; blue text = hardcode, green text = cross-sheet, black = formula
- [ ] Number formats: `$` on top and bottom of a numeric block, `#,##0` in the middle
- [ ] Markdown summary written for the user with the reconciliation results and any caveats

**File naming:** `<Company>_Retention_<YYYY-MM>_to_<YYYY-MM>.xlsx`
