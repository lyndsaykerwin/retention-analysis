# retention-analysis

A Claude Code skill that turns a customer-by-month revenue workbook into a finished SaaS retention analysis. Drop in an Excel file where rows are customers and columns are months of MRR (monthly recurring revenue — what each customer pays per month) or ARR (annual recurring revenue — the same thing annualized), and the skill returns a three-tab Excel deliverable: a retention rollforward (a month-by-month walk from starting revenue to ending revenue, showing what came in, what churned, and what expanded), plus monthly and LTM (last-twelve-month) Gross, Net, and Logo retention rates. Every number is a live formula traceable back to the source data — no hardcoded values, full audit trail.

## What it's for / what it's NOT for

**For:**
- SaaS subscription retention analysis from customer-level revenue data

**Not for:**
- Forecasting future revenue
- LTV/CAC (lifetime value vs. customer acquisition cost) analysis
- Cohort-by-acquisition-month retention curves (the "triangle" view)
- Consumer or transactional churn (e-commerce, one-time purchases)

## Install

**As a Claude Code skill** — clone directly into the skills folder so Claude can find it:

```
git clone https://github.com/lyndsaykerwin/retention-analysis-skill.git ~/.claude/skills/retention-analysis
```

**As a standalone Python tool** — clone anywhere and run the scripts directly:

```
git clone https://github.com/lyndsaykerwin/retention-analysis-skill.git
cd retention-analysis-skill
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
cd ~/.claude/skills/retention-analysis   # or wherever you cloned it
python3 scripts/survey.py --self-test
python3 scripts/compute.py --self-test
```

Both should print `PASS` and exit 0. They run against the bundled fixture at `scripts/fixtures/sample_retention_data.xlsx` — 10 synthetic SaaS customers across 18 months of revenue, covering every retention bucket (new, churned, expanded, contracted, resurrected).

## Using it inside Claude Code

Once installed in `~/.claude/skills/`, open Claude Code in a folder containing a real retention workbook and ask something like:

> "Analyze the customer retention in this workbook"

Claude will trigger the skill, run the survey step, ask you ONE upfront question to confirm schema and methodology, then build the deliverable end-to-end. The skill explicitly tells Claude not to drip-feed checkpoints — one confirmation up front, then it runs.

## What's in the repo

```
SKILL.md                              # the spec Claude follows
scripts/
  survey.py                           # phase 1: inspect the workbook, hypothesize structure
  compute.py                          # phase 2: compute metrics from long-format CSV
                                      #          (one row per customer-month, the tidy shape)
  deliver.py                          # phase 3: write the three-tab Excel deliverable
  fixtures/
    sample_retention_data.xlsx        # synthetic test data
    EXPECTED_VALUES.md                # hand-computed values the self-tests check against
README.md
.gitignore
```

## Methodology (brief)

Retention metrics follow the SaaS Metrics Standards Board definitions for Gross retention (revenue kept, capped at 100%), Net retention (revenue kept plus expansion from existing customers, can exceed 100%), and Logo retention (count of customers kept, regardless of dollars). LTM retention uses the direct-cohort point-in-time method — take the set of customers active 12 months ago, sum their revenue today, divide by their revenue then.

Full methodology, formula patterns, edge-case handling, and visual conventions are in `SKILL.md`.
