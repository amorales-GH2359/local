---
name: security-deposit-review
description: Review a Yardi "Security Deposit Activity" report (Excel/.xlsx export) — verify the math reconciles, then summarize takeaways per property. Use when the user shares a Yardi security deposit return/report and asks to check it, audit it, reconcile it, or summarize deposits by property.
---

# Security Deposit Review (Yardi)

Audit a Yardi **Security Deposit Activity** export and produce a per-property summary with flagged action items. The report is one sheet with property sections, a `Total` row per property, and a final `Grand Total`.

## Report structure

Header rows (top of sheet): report title, `Property = ...` (list of properties), `Period = MM/YYYY-MM/YYYY`.

Column layout (data starts ~row 6). The header spans two rows, so the columns are:

| Col | Meaning |
|-----|---------|
| Property | Property name (repeats per section) |
| Unit | Unit number |
| Tenant | `Name (Status)` — status is `Current`, `Notice`, or `Past`. May include tags like `FURNISHED`, `MTM`, `KASAI`, or notes |
| Prior Deposit Billed | Beginning-of-period billed |
| Prior Receipts | Beginning-of-period received |
| Current Dep.Billed | Deposits billed this period (negative = move-out reversal) |
| Current Receipts | Deposits received this period (negative = refund/release) |
| Deposits On Hand | Ending balance held |
| (Prpd)/Delnq Deposits | Prepaid or delinquent deposit (billed vs received gap) |
| Deposits Forfeited | Amount forfeited this period |

`Total` rows have the literal string `Total` in the Property column; the final summary row is `Grand Total`.

## How to review

1. **Load the workbook** with `openpyxl` (`data_only=True` to get computed values, not formulas). Install if missing: `pip install openpyxl -q`. Iterate rows from the first data row; treat rows where Property is `Total`/`Grand Total` as section subtotals.

2. **Verify row-level math** for every tenant line:
   - `Deposits On Hand == Prior Receipts + Current Receipts` (ending = prior + period movement).
   - Billed should equal Received in both prior and current columns unless `(Prpd)/Delnq` is non-zero. Any row where `Billed != Receipts` should show a matching delinquent/prepaid amount.
   - Report the count of mismatches. Zero mismatches = report ties out internally.

3. **Reconcile section and grand totals** by summing each column per property and comparing to the printed `Total` rows, then summing those to the `Grand Total`. Confirm On Hand, Prior, Net activity, and Forfeited all reconcile.

4. **Surface exceptions** — these are the takeaways:
   - **Forfeitures**: list every line with `Deposits Forfeited > 0` (tenant, unit, amount).
   - **Delinquent/prepaid deposits**: any non-zero `(Prpd)/Delnq` — deposits billed but not collected.
   - **Past tenants still holding a deposit**: rows where status is `Past` **and** `Deposits On Hand > 0`. A moved-out tenant's deposit should be refunded, transferred, or forfeited — not still on hand. These are the primary action items. Distinguish genuinely stale balances (no current-period activity) from label-lag (deposit billed/collected within the period for a furnished/corporate stay that just ended).
   - **Turnover**: units with two tenant lines (a move-out with negative current + a new move-in) indicate a flip during the period.

5. **Summarize per property**: unit count, Current/Notice/Past breakdown, ending Deposits On Hand, net period activity (inflow/outflow), forfeitures, and any flags. Note net-outflow properties (churn) and furnished/corporate concentration. Lead with what's stable, end with what needs action.

6. **Close with a consolidated action list**: a single table of all Past-tenant-held deposits (and any delinquencies) that need disposition, with a total. Offer to draft a follow-up worksheet or a note to deposit accounting.

## Verification script

Adapt this to the uploaded file path. It checks row math, reconciles per-property stats, and flags Past tenants holding deposits.

```python
import openpyxl
from collections import defaultdict

PATH = "<uploaded .xlsx path>"
wb = openpyxl.load_workbook(PATH, data_only=True)
ws = wb.worksheets[0]

def num(x): return x if isinstance(x, (int, float)) else 0

stats = defaultdict(lambda: dict(units=set(), lines=0, current=0, notice=0, past=0,
                                 oh=0, forf=0, delq=0, prior=0, net=0))
bad_math, forfeits, delinquent, past_holding = [], [], [], []

for r in ws.iter_rows(min_row=6, values_only=True):
    prop, unit, ten = r[0], r[1], r[2]
    if not prop or prop in ("Total", "Grand Total"):
        continue
    pb, pr, cb, cr, oh, delq, forf = (num(x) for x in r[3:10])
    s = stats[prop]
    s['units'].add(str(unit).strip()); s['lines'] += 1
    s['oh'] += oh; s['forf'] += forf; s['delq'] += delq; s['prior'] += pr; s['net'] += cr
    status = ten.split("(")[-1].rstrip(") ").strip() if "(" in (ten or "") else "?"
    s[status.lower()] = s.get(status.lower(), 0) + 1
    if abs((pr + cr) - oh) > 0.01:
        bad_math.append((prop, unit, ten, pr, cr, oh))
    if forf:  forfeits.append((prop, unit, ten, forf))
    if delq:  delinquent.append((prop, unit, ten, delq))
    if status == "Past" and oh > 0:
        past_holding.append((prop, unit, ten, oh))

print("Row-math mismatches:", len(bad_math), bad_math)
for p, s in stats.items():
    print(f"\n{p}: {len(s['units'])} units | Cur {s['current']} / Notice {s['notice']} / Past {s['past']}"
          f" | OnHand ${s['oh']:,.2f} | Net ${s['net']:,.2f} | Forfeited ${s['forf']:,.2f} | Delinq ${s['delq']:,.2f}")
print("\nForfeitures:", forfeits)
print("Delinquent:", delinquent)
print("Past tenants still holding a deposit:", past_holding)
```

## Output style

Keep it business-focused and scannable for a property manager:
- Open with a one-line verdict on whether the return ties out (row math + totals).
- A reconciliation table (per property: Prior, Net, On Hand, Forfeited).
- Per-property takeaways, leading with stability and ending with flags.
- A final consolidated action list of deposits needing disposition.
Don't dump every row — surface the exceptions and the numbers that matter.
