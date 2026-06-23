# Lease Audit Tab — QA Review (data sanity check)

Source: `Ascend_Lease_Audit_as_of_06_01_2026.xlsx` → tab **Lease Audit** (289 unit rows).
Scope: internal consistency of the Lease Audit tab only (not compared to the rentable-items report).

---

## Tier 1 — Structural errors that undermine the audit

### 1. Building-10 column shift (rows 10-105 → 10-208, ~17 units)
For this block, the charge columns are each displaced **one column to the right**:

| Column (should hold) | Actually contains |
|---|---|
| P = Base Rent flag | Parking/garage note |
| Q = Garage/Parking | Storage note |
| R = Storage | Washer & Dryer note |
| S = Washer & Dryer | Utility/RUBS note |

Example (10-107): Base Rent = "Parking Addendum in packet", Garage = "Enclosed Garage/Storage Addendum (storage unit…)", Storage = "Washer and Dryer Addendum in packet", W&D = "RUBS…".
Compare clean rows 10-101…10-104 and 10-209, which line up correctly. **The Base Rent verification is effectively missing for the whole block**, and every charge note is in the wrong column.

### 2. Columns J (Unit Type) and K (SQFT) are FALSE for all 289 rows
Everywhere else TRUE = "lease matches rent roll." These two columns are uniformly FALSE, which would mean every lease mismatches its floorplan and square footage — implausible. They were almost certainly never completed and should be treated as **unverified**, not as findings.

### 3. Column P (Base Rent) contaminated with garage/parking text
Beyond the Building-10 shift, several Base Rent cells hold unrelated notes — e.g. 03-101 "N/A - no garage charges", 05-203 "No garage addendum". Base Rent isn't actually verified for those rows.

---

## Tier 2 — Consistency / convention problems

### 4. Storage units recorded in the Garage/Parking column
Storage units (S#) are repeatedly labeled "Garage space" in column Q:
04-102 (S4-1), 04-106 (S4-3), 04-203 (S4-5), 05-102 (S5-1), 05-308 (S5-4), 05-306 ("unit 5-C"), 02-204 (S2-4), plus the Building-10 block. Garage vs storage are not cleanly separated, so neither column can be totaled reliably.

### 5. Mixed TRUE/FALSE convention
Flag columns N/O/P mix bare booleans (`True`/`False`) with text (`"TRUE - …"`, `"FALSE - …"`). Notably **06-307** has a bare `False` in Base Rent with no explanation, unlike every other flagged discrepancy. This makes the columns impossible to filter/sort cleanly.

### 6. Column M (Market Rent) = "Not on the lease" for all 289 rows
Identical text in every row — provides no audit value.

---

## Tier 3 — Genuine discrepancies the audit flagged (worth acting on)

### Lease date mismatches (rent roll vs lease document)
- **01-104** — rent roll 05/15/2026–08/13/2027 vs lease 04/15/2025–05/14/2026 (**off by ~1 year**; largest)
- **10-209** — end 09/30/2026 (lease) vs 11/30/2026 (rent roll) — 2 months
- **10-309** — end 01/20/2027 vs 11/20/2026 — 2 months
- **10-304** — end 04/30/2027 vs 05/31/2027 — 1 month
- **10-307** — end 12/31/2026 vs 01/09/2027
- **10-308** — start 01/01/2026 vs 12/01/2025
- **10-306** — end 09/01/2026 vs 08/31/2026 — 1 day
- **09-303** — end 11/30/2026 vs 12/02/2026 — 2 days
- **02-309** — start & end each off by 1 day

### Base rent mismatches
- **04-109** — lease $1,835 vs rent roll $1,989 (**$154**; lease expired, renewal not in Drive)
- **06-102** — lease $2,030 vs rent roll $2,089 ($59)
- **04-204** — lease $1,419 vs rent roll $1,409 ($10)
- **09-303** — lease $1,419 vs rent roll $1,429 ($10)
- **03-108** — $1,610 vs $1,545; **03-201** — $1,825 vs $1,820 (both: renewal not in Drive)

### Audit coverage gaps (lease packet not found)
- **01-304** — entirely unverifiable
- **08-309** — renewal packet missing
- **10-204** — lease not found
- Renewals missing in Drive: 03-108, 03-201, 04-109

---

## Tier 4 — Rent-roll observations surfaced in the tab

### 7. $0 Market Rent on 6 occupied units
08-301, 08-306, 08-308, 08-309, 08-310, 09-305 show Market Rent = $0 — implausible for occupied units; likely a rent-roll extract issue.

### 8. Inconsistent storage pricing
Same storage item type is priced at **$17 / $26 / $40 / $50** across units (e.g., 02-303 $17, 01-305 $26, 10-207 $50, most others $40). Worth a pricing review.
