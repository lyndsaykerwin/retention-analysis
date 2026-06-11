# retention-analysis

An agent skill that turns customer-level revenue into an **investor-grade SaaS retention analysis**. Works in any agent harness that can run a skill — it isn't tied to a specific tool.

Give it customer-level recurring revenue in whatever shape you have it:

- **Wide** — one row per customer, one column per month (the typical exported spreadsheet), or
- **Long / "tidy"** — one row per customer per month (columns like `customer_id, month, mrr`),

as an **Excel file or a CSV**, holding either MRR (monthly recurring revenue — what each customer pays per month) or ARR (annual recurring revenue — the same thing annualized). The skill returns an Excel deliverable: a retention rollforward (a period-by-period walk from starting revenue to ending revenue, showing what came in, what churned, and what expanded), plus Gross, Net, and Logo retention rates, with LTM (last-twelve-month) views where the data supports them. Every number is a live formula traceable back to the source data — no hardcoded values, full audit trail.

## What it's for / what it's NOT for

**For:**
- SaaS subscription retention analysis from customer-level revenue data

**Not for:**
- Forecasting future revenue
- LTV/CAC (lifetime value vs. customer acquisition cost) analysis
- Cohort-by-acquisition-month retention curves (the "triangle" view)
- Consumer or transactional churn (e-commerce, one-time purchases)

## Install

**As an agent skill** — clone it into your agent's skills folder so the agent can discover it. For example:

```
git clone https://github.com/lyndsaykerwin/retention-analysis.git <your-agent-skills-folder>/retention-analysis
```

**As a standalone Python tool** — clone anywhere and run the scripts directly:

```
git clone https://github.com/lyndsaykerwin/retention-analysis.git
cd retention-analysis
```

## Dependencies

**Required:**
- Python 3.9 or newer
- `openpyxl` (the only non-stdlib Python package) — install with `pip install openpyxl`

**Required for the final recalc step:**
- LibreOffice — provides the `soffice` command used in headless mode (no GUI) to recompute every formula.
  - macOS: `brew install --cask libreoffice`
  - Ubuntu/Debian: `sudo apt install libreoffice`
  - Windows: download installer from libreoffice.org

Why LibreOffice: after the skill writes all the Excel formulas, it runs `soffice --headless --calc --convert-to xlsx` to force every formula to compute and cache its result. Without this, you'd open the file and see formulas but empty result cells until you manually press F9 to recalc.

## Quickstart — run the self-tests

```
cd retention-analysis   # or wherever you cloned it
python3 scripts/survey.py --self-test
python3 scripts/compute.py --self-test
```

Both should print `PASS` and exit 0. They run against the bundled fixture at `scripts/fixtures/sample_retention_data.xlsx` — 10 synthetic SaaS customers across 18 months of revenue, covering every retention bucket (new, churned, expanded, contracted, resurrected).

## Using it inside an agent

Once the skill is installed in your agent's skills folder, point the agent at a folder containing a real retention workbook and ask something like:

> "Analyze the customer retention in this workbook"

The agent triggers the skill, runs the survey step, asks you ONE upfront question to confirm schema and methodology, then builds the deliverable end-to-end. The skill explicitly tells the agent not to drip-feed checkpoints — one confirmation up front, then it runs.

## What's in the repo

```
SKILL.md                              # the spec the agent follows
scripts/
  survey.py                           # phase 1: inspect the workbook, hypothesize structure
  compute.py                          # phase 2: compute metrics from long-format CSV
                                      #          (one row per customer-month, the tidy shape)
  deliver.py                          # phase 3: write the Excel deliverable
  fixtures/
    sample_retention_data.xlsx        # synthetic test data
    EXPECTED_VALUES.md                # hand-computed values the self-tests check against
README.md
.gitignore
```

## Methodology (brief)

**Comparison period.** By default the analysis compares **equivalent periods one year apart** (year-over-year) — this period versus the same period twelve months earlier. That's the right lens for businesses on **annual contracts**, where the renewal decision happens once a year. For businesses on **monthly contracts**, ask for **month-over-month** comparisons (this period versus the immediately prior one) instead.

Retention metrics follow the SaaS Metrics Standards Board definitions for Gross retention (revenue kept, capped at 100%), Net retention (revenue kept plus expansion from existing customers, can exceed 100%), and Logo retention (count of customers kept, regardless of dollars). LTM retention uses the direct-cohort point-in-time method — take the set of customers active 12 months ago, sum their revenue today, divide by their revenue then.

Full methodology, formula patterns, edge-case handling, and visual conventions are in `SKILL.md`.
