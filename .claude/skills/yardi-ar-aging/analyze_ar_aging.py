#!/usr/bin/env python3
"""
Parse a Yardi "Aging Detail" (AR Aging) Excel export and emit structured
numbers + a markdown summary. Deterministic math only — the narrative
observations are layered on by the SKILL.md instructions.

Usage:
    python3 analyze_ar_aging.py <file.xlsx> [more.xlsx ...] [--json] [--md]

Defaults to --md (markdown report to stdout). --json emits a JSON blob
instead. Pass both to get JSON then the markdown report.

Tested against Yardi "Aging Detail" exports with this column layout:
    A Property | B Tenant | C Unit | D Status | E Tran# | F Charge code
    G Date | H Month | I Current(=gross owed) | J 0-30 | K 31-60
    L 61-90 | M Over 90 | N Pre-payments(neg) | O Total(=I+N)

Note: column I ("Current Owed") in these exports equals the SUM of the
aging buckets (gross owed), and O (Total) = I + N (prepayments). The
script treats I as gross owed and verifies the totals tie out.
"""
import sys
import json
from collections import defaultdict

try:
    import openpyxl
except ImportError:
    sys.stderr.write(
        "openpyxl is required. Install with: pip install openpyxl\n"
    )
    sys.exit(2)

# Column indexes (1-based) for the Yardi Aging Detail layout.
COL = dict(prop=1, tenant=2, unit=3, status=4, tran=5, charge=6,
           date=7, month=8, current=9, b0_30=10, b31_60=11,
           b61_90=12, over90=13, prepay=14, total=15)

# Charge-code -> category for composition analysis. Unknown codes fall
# back to "other". Extend as new codes appear in exports.
CHARGE_CATEGORY = {
    "rent-res": "Rent", "rent-com": "Rent",
    "latefee": "Late fees",
    "termfee": "Termination fees",
    "MoveOut": "Move-out charges",
    "CAM-REC": "CAM/Tax/Ins recoveries", "CAM-EST": "CAM/Tax/Ins recoveries",
    "TAX-EST": "CAM/Tax/Ins recoveries", "INS-EST": "CAM/Tax/Ins recoveries",
    "Water": "Utilities & fees", "TrashFee": "Utilities & fees",
    "Pest-Cnt": "Utilities & fees", "ElecFee": "Utilities & fees",
    "STFee": "Utilities & fees", "W&D": "Utilities & fees",
    "Prepay": "Prepayment", "RentCons": "Concession/credit",
    "LLLins": "Concession/credit", "misc": "Misc",
}


def _num(v):
    return float(v) if isinstance(v, (int, float)) else 0.0


def _is_total(v):
    return isinstance(v, str) and v.strip().lower() == "total"


def parse(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    meta = {"property": None, "month_from": None, "statuses": None,
            "run": None}
    # Row 2 carries property/status/month; last row carries the run stamp.
    for r in range(1, min(4, ws.max_row) + 1):
        a = ws.cell(row=r, column=1).value
        if isinstance(a, str) and a.startswith("Property ="):
            meta["property"] = a.split("Property =")[1].split("Status:")[0].strip()
            if "Status:" in a:
                meta["statuses"] = a.split("Status:")[1].split("Month From:")[0].strip()
            if "Month From:" in a:
                meta["month_from"] = a.split("Month From:")[1].strip()
    last = ws.cell(row=ws.max_row, column=1).value
    if isinstance(last, str) and "UserId" in last:
        meta["run"] = last.strip()

    items = []   # individual charge/receipt lines
    for r in range(5, ws.max_row + 1):
        tenant = ws.cell(row=r, column=COL["tenant"]).value
        unit = ws.cell(row=r, column=COL["unit"]).value
        charge = ws.cell(row=r, column=COL["charge"]).value
        if _is_total(tenant):
            continue
        if not tenant or not charge:
            continue
        items.append({
            "tenant": str(tenant).strip(),
            "unit": str(unit).strip() if unit is not None else "",
            "status": (ws.cell(row=r, column=COL["status"]).value or "").strip(),
            "tran": (ws.cell(row=r, column=COL["tran"]).value or "").strip(),
            "charge": str(charge).strip(),
            "current": _num(ws.cell(row=r, column=COL["current"]).value),
            "b0_30": _num(ws.cell(row=r, column=COL["b0_30"]).value),
            "b31_60": _num(ws.cell(row=r, column=COL["b31_60"]).value),
            "b61_90": _num(ws.cell(row=r, column=COL["b61_90"]).value),
            "over90": _num(ws.cell(row=r, column=COL["over90"]).value),
            "prepay": _num(ws.cell(row=r, column=COL["prepay"]).value),
            "total": _num(ws.cell(row=r, column=COL["total"]).value),
        })
    return meta, items


def summarize(meta, items):
    # Roll up per ledger (tenant + unit).
    ledgers = defaultdict(lambda: {
        "tenant": "", "unit": "", "status": "", "lines": 0,
        "gross": 0.0, "prepay": 0.0, "net": 0.0,
        "b0_30": 0.0, "b31_60": 0.0, "b61_90": 0.0, "over90": 0.0,
        "trans": defaultdict(float),
    })
    composition = defaultdict(float)   # positive charges only
    buckets = dict(b0_30=0.0, b31_60=0.0, b61_90=0.0, over90=0.0)
    gross = prepay = 0.0

    for it in items:
        key = (it["tenant"], it["unit"])
        L = ledgers[key]
        L["tenant"], L["unit"] = it["tenant"], it["unit"]
        if it["status"]:
            L["status"] = it["status"]
        L["lines"] += 1
        L["gross"] += it["current"]
        L["prepay"] += it["prepay"]
        L["net"] += it["total"]
        for b in ("b0_30", "b31_60", "b61_90", "over90"):
            L[b] += it[b]
            buckets[b] += it[b]
        L["trans"][it["tran"]] += it["total"]
        gross += it["current"]
        prepay += it["prepay"]

        cat = CHARGE_CATEGORY.get(it["charge"], "Other")
        if it["charge"] != "Prepay" and it["current"] > 0:
            composition[cat] += it["current"]

    net = gross + prepay

    owed, credits, zero = [], [], []
    for L in ledgers.values():
        rec = dict(tenant=L["tenant"], unit=L["unit"], status=L["status"],
                   gross=round(L["gross"], 2), prepay=round(L["prepay"], 2),
                   net=round(L["net"], 2),
                   b31_60=round(L["b31_60"], 2), b61_90=round(L["b61_90"], 2),
                   over90=round(L["over90"], 2), lines=L["lines"])
        if round(L["net"], 2) > 0:
            owed.append(rec)
        elif round(L["net"], 2) < 0:
            credits.append(rec)
        else:
            zero.append(rec)
    owed.sort(key=lambda x: -x["net"])
    credits.sort(key=lambda x: x["net"])

    # ---- Anomaly / observation flags ----
    flags = []

    aged = round(buckets["b31_60"] + buckets["b61_90"] + buckets["over90"], 2)
    if aged > 0:
        aged_who = [r for r in owed
                    if (r["b31_60"] + r["b61_90"] + r["over90"]) > 0]
        flags.append({"type": "aged_debt", "amount": aged,
                      "ledgers": [(r["tenant"], r["unit"],
                                   round(r["b31_60"] + r["b61_90"] + r["over90"], 2))
                                  for r in aged_who]})

    # Same tenant name across multiple units.
    by_name = defaultdict(set)
    for L in ledgers.values():
        by_name[L["tenant"]].add(L["unit"])
    multi = {n: sorted(u) for n, u in by_name.items() if len(u) > 1}
    if multi:
        flags.append({"type": "multi_unit_tenant", "tenants": multi})

    # Status of Notice/Past while still owing.
    notice_past = [(r["tenant"], r["unit"], r["status"], r["net"])
                   for r in owed if r["status"] in ("Notice", "Past")]
    if notice_past:
        flags.append({"type": "notice_past_with_balance", "ledgers": notice_past})

    # Concentration: top owed ledger vs total gross.
    if owed and gross > 0:
        top = owed[0]
        flags.append({"type": "concentration",
                      "top": (top["tenant"], top["unit"], top["net"]),
                      "pct_of_gross": round(100 * top["net"] / gross, 1)})

    # Large prepayments fragmented across many lines of one receipt.
    receipt_lines = defaultdict(lambda: {"n": 0, "amt": 0.0, "who": None})
    for it in items:
        if it["charge"] == "Prepay" and it["tran"]:
            k = (it["tenant"], it["unit"], it["tran"])
            receipt_lines[k]["n"] += 1
            receipt_lines[k]["amt"] += it["total"]
            receipt_lines[k]["who"] = (it["tenant"], it["unit"])
    big_prepays = [(k[2], v["who"][0], v["who"][1], v["n"], round(v["amt"], 2))
                   for k, v in receipt_lines.items()
                   if abs(v["amt"]) >= 5000 or v["n"] >= 10]
    big_prepays.sort(key=lambda x: x[4])
    if big_prepays:
        flags.append({"type": "large_fragmented_prepay", "receipts": big_prepays})

    # Charge + near-equal reversal within a ledger (possible over/under correction).
    for L in ledgers.values():
        pos = [it for it in items
               if it["tenant"] == L["tenant"] and it["unit"] == L["unit"]
               and it["total"] > 0]
        neg = [it for it in items
               if it["tenant"] == L["tenant"] and it["unit"] == L["unit"]
               and it["total"] < 0 and it["charge"] != "Prepay"]
        for p in pos:
            for n in neg:
                if p["charge"] == n["charge"] and abs(p["total"] + n["total"]) < max(1.0, 0.1 * p["total"]) and abs(p["total"] + n["total"]) != 0:
                    flags.append({"type": "charge_reversal_mismatch",
                                  "tenant": L["tenant"], "unit": L["unit"],
                                  "charge": p["charge"],
                                  "charged": round(p["total"], 2),
                                  "reversed": round(n["total"], 2),
                                  "residual": round(p["total"] + n["total"], 2)})

    return {
        "meta": meta,
        "totals": {"gross": round(gross, 2), "prepay": round(prepay, 2),
                   "net": round(net, 2),
                   "buckets": {k: round(v, 2) for k, v in buckets.items()},
                   "aged_over_30": aged},
        "composition": {k: round(v, 2) for k, v in
                        sorted(composition.items(), key=lambda x: -x[1])},
        "owed": owed, "credits": credits, "zero": zero,
        "ledger_count": len(ledgers),
        "flags": flags,
    }


def _money(x):
    return f"${x:,.2f}" if x >= 0 else f"-${abs(x):,.2f}"


def to_markdown(s):
    m, t = s["meta"], s["totals"]
    out = []
    out.append(f"# AR Aging — {m.get('property') or 'Unknown property'}")
    sub = []
    if m.get("month_from"):
        sub.append(f"Month from {m['month_from']}")
    if m.get("statuses"):
        sub.append(f"Statuses: {m['statuses']}")
    if sub:
        out.append("_" + " · ".join(sub) + "_")
    out.append("")
    out.append("## Bottom line")
    out.append(f"- **Net balance: {_money(t['net'])}**"
               + ("  _(credit)_" if t["net"] < 0 else ""))
    out.append(f"- Gross charges owed: {_money(t['gross'])}")
    out.append(f"- Prepayments / credits: {_money(t['prepay'])}")
    out.append(f"- Ledgers: {s['ledger_count']} "
               f"({len(s['owed'])} owe, {len(s['credits'])} credit, "
               f"{len(s['zero'])} zero)")
    out.append("")
    out.append("## Aging buckets")
    b = t["buckets"]
    out.append("| Bucket | Amount |")
    out.append("|---|---|")
    out.append(f"| 0–30 days | {_money(b['b0_30'])} |")
    out.append(f"| 31–60 days | {_money(b['b31_60'])} |")
    out.append(f"| 61–90 days | {_money(b['b61_90'])} |")
    out.append(f"| Over 90 days | {_money(b['over90'])} |")
    out.append(f"| Prepayments | {_money(t['prepay'])} |")
    out.append(f"| **Aged > 30 days** | **{_money(t['aged_over_30'])}** |")
    out.append("")
    if s["composition"]:
        out.append("## Charge composition (gross owed)")
        out.append("| Type | Amount | Share |")
        out.append("|---|---|---|")
        for k, v in s["composition"].items():
            pct = (100 * v / t["gross"]) if t["gross"] else 0
            out.append(f"| {k} | {_money(v)} | {pct:.1f}% |")
        out.append("")
    if s["owed"]:
        out.append("## Who owes")
        out.append("| Tenant | Unit | Status | Net owed | Aged>30 |")
        out.append("|---|---|---|---|---|")
        for r in s["owed"]:
            aged = r["b31_60"] + r["b61_90"] + r["over90"]
            out.append(f"| {r['tenant']} | {r['unit']} | {r['status']} | "
                       f"{_money(r['net'])} | {_money(aged)} |")
        out.append("")
    if s["credits"]:
        out.append("## Credit / prepayment balances")
        out.append("| Tenant | Unit | Net |")
        out.append("|---|---|---|")
        for r in s["credits"]:
            out.append(f"| {r['tenant']} | {r['unit']} | {_money(r['net'])} |")
        out.append("")
    if s["flags"]:
        out.append("## Flags to investigate")
        for f in s["flags"]:
            if f["type"] == "aged_debt":
                who = ", ".join(f"{t_} {u} ({_money(a)})"
                                for t_, u, a in f["ledgers"])
                out.append(f"- **Aged debt > 30 days: {_money(f['amount'])}** — {who}")
            elif f["type"] == "multi_unit_tenant":
                for n, u in f["tenants"].items():
                    out.append(f"- **Same tenant on multiple units:** {n} → units {', '.join(u)}")
            elif f["type"] == "notice_past_with_balance":
                for t_, u, st, nt in f["ledgers"]:
                    out.append(f"- **{st} status with balance:** {t_} {u} owes {_money(nt)}")
            elif f["type"] == "concentration":
                t_, u, nt = f["top"]
                out.append(f"- **Concentration:** {t_} {u} = {_money(nt)} "
                           f"({f['pct_of_gross']}% of gross owed)")
            elif f["type"] == "large_fragmented_prepay":
                for tran, t_, u, n, amt in f["receipts"]:
                    span = f", split across {n} lines" if n > 1 else ""
                    out.append(f"- **Large prepayment to verify:** {t_} {u} — "
                               f"{_money(amt)} on receipt {tran}{span}")
            elif f["type"] == "charge_reversal_mismatch":
                out.append(f"- **Reversal mismatch:** {f['tenant']} {f['unit']} "
                           f"{f['charge']} charged {_money(f['charged'])}, "
                           f"reversed {_money(f['reversed'])} "
                           f"(residual {_money(f['residual'])})")
        out.append("")
    return "\n".join(out)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = [a for a in sys.argv[1:] if a.startswith("--")]
    if not args:
        sys.stderr.write(__doc__)
        sys.exit(2)
    want_json = "--json" in flags
    want_md = "--md" in flags or not want_json

    results = []
    for path in args:
        meta, items = parse(path)
        results.append(summarize(meta, items))

    if want_json:
        print(json.dumps(results if len(results) > 1 else results[0], indent=2))
    if want_md:
        for i, s in enumerate(results):
            if i:
                print("\n---\n")
            print(to_markdown(s))
        if len(results) > 1:
            g = dict(gross=0.0, prepay=0.0, net=0.0, aged=0.0)
            print("\n---\n## Portfolio rollup")
            print("| Property | Gross owed | Prepayments | Net | Aged>30 |")
            print("|---|---|---|---|---|")
            for s in results:
                t = s["totals"]
                g["gross"] += t["gross"]; g["prepay"] += t["prepay"]
                g["net"] += t["net"]; g["aged"] += t["aged_over_30"]
                print(f"| {s['meta'].get('property')} | {_money(t['gross'])} | "
                      f"{_money(t['prepay'])} | {_money(t['net'])} | "
                      f"{_money(t['aged_over_30'])} |")
            print(f"| **Total** | **{_money(g['gross'])}** | "
                  f"**{_money(g['prepay'])}** | **{_money(g['net'])}** | "
                  f"**{_money(g['aged'])}** |")


if __name__ == "__main__":
    main()
