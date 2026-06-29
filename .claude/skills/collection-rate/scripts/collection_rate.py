#!/usr/bin/env python3
"""Compute monthly collection rate from Yardi Rent Roll AR exports (.xlsx).

Collection rate = (Total Actual Charges - Sum of POSITIVE AR balances) / Total Actual Charges

Rules baked in (see SKILL.md):
  - Only POSITIVE balances count as outstanding/uncollected. Negative balances
    (credits / prepayments) are ignored, not netted against the positive ones.
  - Only the CURRENT rent roll is used. The separate "Future Resident Details"
    section at the bottom of each report is excluded automatically (the script
    stops at the property "...Total:" row that closes the current rent roll).

Usage:
    python3 collection_rate.py FILE1.xlsx [FILE2.xlsx ...] [--label "Name=FILE"]
    python3 collection_rate.py *.xlsx

Requires: openpyxl  (pip install openpyxl)
"""
import sys
import argparse
import openpyxl


def compute(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]  # rent roll is always the first sheet

    # Locate the header row (the one carrying both "Actual Charges" and "Balance").
    hdr = None
    cols = {}
    for r in range(1, 20):
        vals = {ws.cell(row=r, column=c).value: c for c in range(1, 20)}
        if "Actual Charges" in vals and "Balance" in vals:
            hdr, cols = r, vals
            break
    if hdr is None:
        raise ValueError(f"{path}: could not find a header row with 'Actual Charges' and 'Balance'")

    kcol = cols["Actual Charges"]
    mcol = cols["Balance"]

    # The current rent roll ends at the first "...Total:" row in column B.
    # Everything after it (Status Summary, Charge Code Summary, Future Resident
    # Details) is excluded.
    total_row = None
    for r in range(hdr + 1, ws.max_row + 1):
        b = ws.cell(row=r, column=2).value
        if isinstance(b, str) and b.strip().endswith("Total:"):
            total_row = r
            break
    if total_row is None:
        total_row = ws.max_row + 1

    total_charges = 0.0
    positive_balances = 0.0
    negative_balances = 0.0
    units = 0
    owing = []  # (unit, resident, balance)

    for r in range(hdr + 1, total_row):
        unit = ws.cell(row=r, column=1).value
        if not isinstance(unit, str) or "-" not in unit:
            continue  # skip blanks / non-unit rows; real unit ids look like "102-1A"
        units += 1
        k = ws.cell(row=r, column=kcol).value
        m = ws.cell(row=r, column=mcol).value
        if isinstance(k, (int, float)):
            total_charges += k
        if isinstance(m, (int, float)):
            if m > 0:
                positive_balances += m
                resident = ws.cell(row=r, column=cols.get("Resident", 5)).value
                owing.append((unit, resident, m))
            elif m < 0:
                negative_balances += m

    rate = (total_charges - positive_balances) / total_charges * 100 if total_charges else 0.0
    return {
        "path": path,
        "units": units,
        "total_charges": total_charges,
        "positive_balances": positive_balances,
        "negative_balances": negative_balances,
        "collected": total_charges - positive_balances,
        "rate": rate,
        "owing": owing,
        "report_total_charges": ws.cell(row=total_row, column=kcol).value if total_row <= ws.max_row else None,
    }


def main():
    ap = argparse.ArgumentParser(description="Collection rate from Yardi Rent Roll AR exports")
    ap.add_argument("files", nargs="+", help="One or more Rent Roll .xlsx files")
    ap.add_argument("--detail", action="store_true", help="List residents with positive balances")
    args = ap.parse_args()

    print(f"{'Property / File':<45} {'Charges':>14} {'Owed(+bal)':>12} {'Rate':>8}")
    print("-" * 81)
    for f in args.files:
        try:
            res = compute(f)
        except Exception as e:  # noqa: BLE001
            print(f"{f:<45} ERROR: {e}")
            continue
        rpt = res["report_total_charges"]
        tied = rpt is None or abs(float(rpt) - res["total_charges"]) < 0.01
        tie = "" if tied else "  [!charges differ from report total]"
        name = f.split("/")[-1]
        print(f"{name:<45} {res['total_charges']:>14,.2f} {res['positive_balances']:>12,.2f} {res['rate']:>7.2f}%{tie}")
        if args.detail and res["owing"]:
            for unit, resident, bal in sorted(res["owing"], key=lambda x: -x[2]):
                print(f"    {unit:<12} {str(resident):<35} {bal:>12,.2f}")
    print("-" * 81)
    print("Rate = (Charges - positive balances) / Charges. Future-resident section excluded.")


if __name__ == "__main__":
    main()
