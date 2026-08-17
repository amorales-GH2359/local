---
name: collection-rate-analysis
description: Calculate monthly collection rates from CS3 property "Rent Roll" AR Excel exports (Gateway, Legacy, Legends, etc.). Use when the user attaches one or more Rent Roll / AR .xlsx reports and asks for the collection rate, optionally split by Residential vs Commercial. Counts only positive AR balances against actual rent + misc, excluding future leases.
---

# Collection Rate Analysis (CS3 Rent Roll AR reports)

Compute the monthly collection rate for a property from its "Rent Roll" Excel export.

**Definition used:**
```
Billed          = Actual Rent total + Misc total   (current tenants only)
Uncollected AR  = sum of POSITIVE balances only     (current tenants only)
Collected       = Billed - Uncollected AR
Collection Rate = Collected / Billed * 100
```

Key rules (confirmed with the user):
- **Only positive balances** count as uncollected AR. Negative balances are credits/prepayments and are ignored (they must NOT offset/inflate the rate).
- Compare against **Actual Rent + Misc** for the month.
- **Exclude future leases** — the "Future Tenants/Applicants" section is skipped for both the billed totals and the AR.
- Report **Residential and Commercial separately** when a property has both sections. Some properties are residential-only (see below).

## Report layout

These are openpyxl-readable `.xlsx` files, sheet name `Report1`. Read with `data_only=True`.

Header rows 1-4: `Rent Roll`, `Property = ...`, `As Of = MM/DD/YYYY`, `Month = MM/YYYY`.
Always read the `As Of` / `Month` header and state it in the answer — the same month is
sent as a series of snapshots (e.g. as-of the 6th, 13th, 20th, 27th), and an early-month
snapshot naturally shows a low rate because rent is charged on the 1st.

Column map (1-indexed):
| Col | Letter | Meaning |
|----|----|----|
| 1 | A | Unit (also holds section headers / "Total" label) |
| 3 | C | Tenant Name (also the property/section name on Total rows) |
| 4 | D | **Actual Rent** |
| 8 | H | **Misc** |
| 13 | M | **Balance** (positive = owed/uncollected AR; negative = credit) |

Section structure within a property:
- `Current/Notice/Vacant Tenants` header → tenant rows → optional `Future Tenants/Applicants` header → future rows → `Total` row for that section.
- A property with both Residential and Commercial has TWO such blocks (two `Total` rows), then a `Total = All Properties` row.

### Picking the billed rent/misc totals
The script sums Actual Rent + Misc **directly from the current-tenant rows** (not the section
`Total` row). This is deliberate: when the Future section carries rent (e.g. a future applicant
with actual rent), the section `Total` row includes it, which would overstate billed and distort
the rate. Summing current rows excludes future tenants automatically and works per-section.

### Known property types
- **Gateway Village** — has BOTH `Gateway Village Residential` and `Gateway Village Commercial`. Report separately.
- **CS3 Legacy Townhomes** — residential only.
- **Legends at Armour Ave** — residential only.

### Not this skill
A different export, the **"Receivable Summary"** (`_AR` filename, columns Owner/Property/Unit/
Charge To/Opening Balance/Charges/Receipts/Closing Balance), is a *different* report and this
skill does NOT apply to it. If A1 says `Receivable Summary` rather than `Rent Roll`, stop and
confirm with the user which report they want analyzed.

## Procedure

1. `pip install openpyxl` if not present (the sandbox may reset between sessions and drop it).
2. Run `scripts/collection_rate.py FILE1.xlsx [FILE2.xlsx ...]`. It handles section detection,
   current-vs-future split, positive-AR summing, and the rate math.
3. Present a markdown table per property/section: Property | Actual Rent | Misc | Total Billed |
   Positive AR (uncollected) | Collected | Collection Rate.
4. List the tenants behind the positive AR (unit, name, balance) so delinquents are visible, and
   note any property that is residential-only, plus the as-of date.
5. When multiple snapshots of the same month exist, show the trend and flag balances that persist
   or grow across snapshots (genuine aging receivables) vs. those that are early-month timing.

## Verified reference numbers
- 06/2026 (as of 06/29): Gateway Res 98.55%, Gateway Com 88.39%, Legacy 95.39%, Legends 96.14%.
- Script reproduces these exactly; use as a regression check after editing.
