---
name: yardi-ar-aging
description: Analyze a Yardi "Aging Detail" (AR aging) Excel export for a property and produce a full summary — net balance, aging buckets, charge composition, who owes vs. who's in credit, and flags (aged debt, lease-break/move-out concentration, large prepayments to verify, reversal mismatches). Use when the user shares a Yardi AR aging .xlsx/.xls export (often named ARAnalytics, or after a property name like Legends/Legacy/Gateway) and asks to analyze entries, summarize observations, or review receivables. Handles one property or several at once (portfolio rollup).
---

# Yardi AR Aging analysis

Turn a Yardi **Aging Detail** export into a clear receivables summary. A helper
script does the deterministic parsing and math; you layer on the judgment.

## When this applies

The user uploads a Yardi AR aging Excel file (sheet `Report1`, cell A1 =
`Aging Detail`, A2 = `Property = … Status: … Month From: …`) and asks for an
analysis, summary, or review. Files often arrive named `ARAnalytics_*.xlsx` or
renamed after a property (e.g. `Legends.xlsx`, `Gateway.xlsx`).

> Note: you cannot reach the user's local Chrome/Yardi session from this
> environment. You can only analyze files they upload. If they ask you to "run
> the report in Yardi," explain that and ask them to export and upload it.

## Steps

1. **Locate the file(s).** Uploads usually land under
   `/root/.claude/uploads/<session>/`. Confirm the path(s) with the user if
   unclear.
2. **Ensure openpyxl is available:** `pip install openpyxl -q` (only if the
   import fails).
3. **Run the analyzer** (pass every file in one call to also get a portfolio
   rollup):
   ```bash
   python3 .claude/skills/yardi-ar-aging/analyze_ar_aging.py FILE1.xlsx [FILE2.xlsx ...]
   ```
   Add `--json` if you want the raw structured numbers to reason over instead of
   the markdown report.
4. **Present the report and add observations.** The script output is the
   factual backbone (totals, buckets, composition, who owes, flags). Rewrite it
   into a tight narrative and add the interpretation the script can't:
   - **Lead with the story**, not the table. Is this property net-owed or
     net-credit? What drives it?
   - **Concentration & collectibility:** call out when one or two ledgers are
     most of the receivable, especially `Past`/`Notice` tenants or move-out /
     `termfee` situations — that AR is at risk; recommend deposit application
     and a collections/write-off path.
   - **Commercial vs. residential:** if the property has both books (e.g.
     Gateway), analyze them separately — commercial aging (rent + CAM/TAX/INS
     recoveries) is a different risk than residential.
   - **Verify the big credits:** large or heavily fragmented prepayments (one
     receipt split into many lines) are normal for furnished/STR units but
     worth confirming against the deposit/bank record and lease term.
   - **Clean-up items:** reversal mismatches, offsetting wash entries, tiny
     residual credits — note them as quick fixes, not alarms.
   - **Aging:** emphasize whether anything is past 30/60/90 days. These exports
     often show all-current AR, which is the healthy case worth stating plainly.
5. **Confirm the math ties out** (the script verifies internally; gross +
   prepayments = net) and offer next steps: a written/Excel memo, or a
   cross-check against QuickBooks AR (`qbo_accounting_get_ar_aging_*` tools) if
   that MCP is connected.

## Report layout notes

- Columns: `Property | Tenant | Unit | Status | Tran# | Charge | Date | Month |
  Current | 0-30 | 31-60 | 61-90 | Over 90 | Pre-payments | Total`.
- Column **Current** ("Current Owed") equals the **sum of the aging buckets**
  (gross owed), and **Total** = Current + Pre-payments. Pre-payments are
  negative (credits).
- Per-tenant `Total` rows and section/`Grand Total` rows are subtotals — the
  script derives its own totals from line items and ignores these, so don't
  double-count them when reading the raw sheet.
- Charge codes seen so far: `rent-res`, `rent-com`, `latefee`, `termfee`,
  `MoveOut`, `CAM-REC/EST`, `TAX-EST`, `INS-EST`, `Water`, `TrashFee`,
  `Pest-Cnt`, `ElecFee`, `STFee`, `W&D`, `Prepay`, `RentCons`, `LLLins`,
  `misc`. New codes fall into "Other" — extend `CHARGE_CATEGORY` in the script
  if a new one becomes common.
