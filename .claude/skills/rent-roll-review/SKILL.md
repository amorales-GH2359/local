---
name: rent-roll-review
description: Review a Yardi rent roll export (.xlsx) and surface what needs attention — delinquencies, large credit balances, expiring/expired leases, move-outs, vacancies, occupancy, and per-property summaries. Use when the user uploads or points to a rent roll spreadsheet and asks to check it, review it, summarize it, or find issues.
---

# Rent Roll Review

Analyzes a **Yardi Rent Roll** export and produces a prioritized review: a portfolio
snapshot, a per-property breakdown, and exception flags the property manager should
act on.

## When to use

Trigger this skill when the user provides a rent roll spreadsheet (typically named
like `Yardi_Rent_Rolls.xlsx`) and asks anything like "check this rent roll",
"anything I need to look at?", "summarize by property", or "find issues".

## How to run

A helper script does the parsing and flagging. It works on any standard Yardi rent
roll — it keys off the report's section headers and per-property `Total` rows rather
than hardcoded property names, so it handles any number of properties.

```bash
pip install openpyxl -q   # if not already installed
python3 scripts/analyze_rent_roll.py <path-to-rent-roll.xlsx>
# optional: override the report date with --as-of MM/DD/YYYY
```

The script prints:

- **Portfolio snapshot** — units, occupied, vacant, occupancy %, total monthly rent, total owed.
- **Per-property summary** — for each property: unit counts, occupancy %, monthly rent, amount owed, credit balances, move-outs on notice, future applicants, and vacant unit list.
- **Flags needing attention:**
  1. **Delinquencies** — positive balances, sorted high to low; tenants who have moved/are moving out are marked (collect before keys turn over).
  2. **Large credit balances** — unusually large negative balances (default ≤ −$2,000) that should be verified with accounting; they distort the net balance.
  3. **Expired / expiring leases** — within the next 60 days (and any already expired), with renewal/notice status.
  4. **Move-outs on notice** — with dates and any outstanding balance.
  5. **Occupied units with $0 security deposit.**
  6. **Occupied units with $0 rent.**
  7. **Negative misc charges** — concessions/credits worth confirming.

## How to present results

Run the script, then write up the findings for the user in prose — **lead with what
matters most** (delinquencies that are about to walk, unverified large credits, and
renewal clusters), not just a data dump. Key judgment points:

- The report's net **Balance** total is misleading: large credit balances offset real
  delinquencies. Always state the **gross amount actually owed** separately.
- Call out the **weakest property** (lowest occupancy / highest delinquency) and the
  **strongest** so the user knows where to focus.
- A lease long past expiration with no notice (e.g. holdover/MTM) is worth flagging
  explicitly — it usually means no current signed lease is on file.
- Flag deposit/rent/charge anomalies as housekeeping, not emergencies.

## Report format reference

Standard Yardi Rent Roll column layout (first worksheet):

| Col | Field | Col | Field |
|----|-------|----|-------|
| A | Unit | H | Misc charge |
| B | Unit SqFt | I | Misc per SqFt |
| C | Tenant Name | J | Move In |
| D | Actual Rent | K | Lease Expiration |
| E | Actual Rent per SqFt | L | Move Out |
| F | Tenant (Security) Deposit | M | Balance |
| G | Other Deposit | | |

Rows are grouped under section headers `Current/Notice/Vacant Tenants` and
`Future Tenants/Applicants`. Each property block ends with a `Total` row whose
property name sits in the Tenant-Name column; the report ends with an
`All Properties` total and a `Summary Groups` footer.

## Follow-ups to offer

After the review, offer to turn findings into action — a collections call list for the
delinquencies, a renewal tracker for the expiring leases, an accountant flag list for
the large credits, or ClickUp tasks / calendar reminders if those integrations are
connected.
