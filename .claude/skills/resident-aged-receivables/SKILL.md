---
name: resident-aged-receivables
description: Analyze Resident Aged Receivables reports (Entrata-style .xlsx exports) and produce a per-property summary of observations. Use when the user shares one or more "Resident Aged Receivables" workbooks and asks to analyze entries, summarize delinquencies, flag aging/credit balances, or review collections notes per property.
---

# Resident Aged Receivables Analyzer

Analyze one or more Resident Aged Receivables workbooks and produce a clear,
per-property summary of every entry plus cross-property takeaways.

## When to use

- The user shares `*_Resident_Aged_Receivables.xlsx` files (Entrata report v3.x).
- They ask to "analyze every entry," "summarize per property," review
  delinquencies/aging, or flag anomalies in these reports.

## Report structure (what to expect)

Each workbook has two sheets:

1. **`<Property Name>`** — the data sheet. Layout:
   - Row 2: `Resident Aged Receivables` (title)
   - Row 3: property name
   - Row 4: post month (e.g. `Jun 2026`)
   - Row 6: column headers:
     `Bldg-Unit | Resident | Lease Status | Unallocated Charges / Credits |
      0-30 Days | 31-60 Days | 61-90 Days | 90+ Days | Pre-Payments | Balance |
      Last Delinquency Note`
   - Rows 7..N: one resident per row.
   - Final row: `<Property> Total:` in column C with column totals.
2. **`Report Parameters`** — metadata: report version, property group, post
   month, min unpaid balance, generation timestamp (`data as of`).

Column meaning:
- **Balance (J)** = net amount owed. Negative = credit balance (prepayment
  exceeds charges), NOT a delinquency.
- **Pre-Payments (I)** = credits on the ledger (negative numbers).
- **0-30 / 31-60 / 61-90 / 90+** = aging buckets. Older buckets (61-90, 90+)
  signal serious delinquency and likely eviction/legal action.
- **Lease Status (C)** — watch for statuses that shouldn't carry a balance:
  `Cancelled`, `Applicant`, `Future`, `Past`, `Past - Eviction`, `Notice`.

## Workflow

1. **Load the data.** Ensure `openpyxl` is available
   (`pip install openpyxl -q` if not). Use `data_only=True` so formula cells
   return computed values. Run `scripts/analyze.py` with the file paths — it
   prints, per property: every row, the totals, and an auto-flagged list of
   anomalies. Read its output, then write the narrative summary yourself.
2. **Summarize per property.** For each property report:
   - State resident count, total balance, and the aging spread
     (0-30 / 31-60 / 61-90 / 90+ / pre-payments).
   - Call out the **largest balances** and their lease status + note.
   - List **credit (negative) balances** separately — these are prepayments,
     not past-due; exclude them from "true past-due" framing.
   - Highlight **oldest debt** (61-90 and 90+) and any eviction/move-out/notice
     activity from the Last Delinquency Note.
   - Summarize the collections narrative from notes (promises to pay, bounced
     payments, disputes, deposit applications).
3. **Flag data-quality issues** (these recur and matter to the user):
   - Lease status that shouldn't carry aged debt (e.g. `Future` lease with a
     61-90 balance; `Cancelled`/`Applicant` with a balance).
   - **Duplicate unit numbers** (same `Bldg-Unit` on two rows — roommate vs.
     assignment error).
   - **Duplicated/stale note text** copied across residents (e.g. a note that
     names a different resident, or a "deposit applied" note where the balance
     is unchanged).
   - Large balances with **no delinquency note** (no documented collection
     attempt).
   - Entity/corporate accounts (name entered as `LLC, <name>` or
     "Workforce Solutions ...").
4. **Cross-property summary.** A comparison table (residents / balance / worst
   aging / credits / biggest account), which property to prioritize, a
   consolidated list of items to reconcile with accounting, and follow-up gaps.

## Output format

Markdown. One `##` section per property, then a `## Cross-property summary`
with a comparison table and prioritized takeaways. Lead each property with the
headline numbers, then observations as bullets. Reference residents by
`Bldg-Unit` + name so they're traceable to the sheet. End by offering to
produce a formatted deliverable (Excel/PDF summary, consolidated workbook, or
prioritized collections action list).

## Notes

- Report is a point-in-time snapshot (`data as of` timestamp in Report
  Parameters) — cite it.
- Negative balances are credits; never describe them as amounts owed.
- Don't invent totals — the sheet's total row is authoritative; the script
  re-sums and warns if a computed total disagrees with the sheet.
