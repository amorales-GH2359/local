---
name: collection-rate
description: Calculate monthly collection rate for properties from Yardi Rent Roll AR exports (.xlsx). Use when asked to check the collection rate, collections, or AR performance for one or more properties given Rent Roll / AR report spreadsheets.
---

# Collection Rate (from Yardi Rent Roll AR exports)

Compute the monthly **collection rate** for one or more properties from their
Yardi "Rent Roll" Excel exports.

## Definition

```
collection rate = (Total Actual Charges − Sum of POSITIVE AR balances) / Total Actual Charges
```

- **Total Actual Charges** — the "Actual Charges" column total for the month (the
  amount actually billed to current residents).
- **Positive AR balances** — money residents still owe (uncollected). Only
  balances **> 0** count.

The numerator `(Charges − Positive balances)` is the amount collected.

## Rules (must follow)

1. **Only positive balances** in the AR count as outstanding. Negative balances
   are credits / prepayments — **ignore them, do not net them** against the
   positive balances. (Net balance totals are often large negatives because of
   prepaid accounts; that does not improve the collection rate.)
2. **Exclude future leases.** Yardi Rent Roll reports have a separate
   **"Future Resident Details"** section near the bottom, plus a "Status Summary"
   and "Charge Code Summary". Only the **current rent roll** (the unit rows above
   the property `…Total:` line) is in scope.
3. **Verify the tie-out.** The summed "Actual Charges" should match the report's
   property `…Total:` row. If it doesn't, something was mis-bounded — investigate
   before reporting.

## Report layout (for reference)

First worksheet = the rent roll. Header row (~row 7) columns:
`Bldg-Unit | Unit Type | SQFT | Unit Status | Resident | Move-In | Lease Start |
Lease End | Expected Move-Out | Market Rent | Actual Charges | Scheduled Charges |
Balance | Deposit Held`. Unit rows run from just below the header down to the
first `…Total:` row in column B. Everything after that row is out of scope.

> Note: in the "Future Resident Details" section the columns are shifted by one
> (no "Expected Move-Out"), so never read those rows with the main-section column
> positions. The script avoids this by stopping at the `…Total:` row.

## How to run

```bash
pip install openpyxl   # if not already installed
python3 scripts/collection_rate.py path/to/Property_AR.xlsx [more.xlsx ...]
python3 scripts/collection_rate.py *.xlsx --detail   # also list who owes
```

The script auto-detects the header and the current-rent-roll boundary, sums
charges and positive balances, confirms the charges tie out to the report total,
and prints the rate per file. Use `--detail` to list the residents/units carrying
positive balances.

## Reporting

Present a short table per property: Actual Charges, Positive Balances (and # of
residents), Collected, and the **Collection Rate %**. Note that the current rent
roll only was used and future residents were excluded. Offer to break out the
specific units carrying positive balances.
