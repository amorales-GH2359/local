# resident-aged-receivables (skill)

Analyzes **Resident Aged Receivables** reports (Entrata-style `.xlsx` exports)
and produces a per-property summary of observations: aging spread, largest
balances, credit balances, collections-note narrative, and auto-flagged
data-quality issues, plus a cross-property comparison.

## Files
- `SKILL.md` — the skill definition and workflow Claude follows.
- `scripts/analyze.py` — parses the workbooks, re-sums totals (warns on
  mismatch with the sheet total), and prints auto-flagged anomalies.

## Quick use
```bash
pip install openpyxl          # once, if not present
python3 scripts/analyze.py FILE1.xlsx [FILE2.xlsx ...]
```

## What gets auto-flagged
- Credit (negative) balances — prepayments, not past-due.
- 61-90 and 90+ day aged debt.
- Lease statuses that shouldn't carry a balance (Cancelled / Applicant / Future).
- Duplicate `Bldg-Unit` numbers across rows.
- Large balances with no delinquency note.
- Identical note *bodies* copied across residents (timestamp-insensitive).
- Entity/corporate accounts (`LLC, ...`, `Workforce Solutions ...`).
