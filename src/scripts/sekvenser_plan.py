#!/usr/bin/env python3
"""Sekvenser produksjon per arbeidssenter x uke (minimerer changeover, FTI-hierarki)."""
import argparse, json, sys
from pathlib import Path
_src = Path(__file__).resolve().parent.parent
if str(_src) not in sys.path: sys.path.insert(0, str(_src))
from collections import defaultdict
from kostberegning import SqliteData
from optimization_engine import _fti_family, _changeover_minutes


def dims(fam):
    d = "".join(c for c in fam if c.isdigit())
    return (int(d[:2]), int(d[-3:])) if len(d) >= 4 else (None, None)


def tsp(nodes, cost):
    n = len(nodes)
    if n <= 1: return list(nodes), 0.0
    best_o, best_c = None, None
    for s in range(n):
        unv, o, cur = set(range(n)), [s], s
        unv.remove(s)
        while unv:
            nxt = min(unv, key=lambda j: cost[o[-1]][j])
            o.append(nxt); unv.remove(nxt)
        imp = True
        while imp:
            imp = False
            for i in range(1, n-1):
                for k in range(i+2, n):
                    d = cost[o[i-1]][o[k]] + cost[o[k-1]][o[i]] - cost[o[i-1]][o[i]] - cost[o[k-1]][o[k]]
                    if d < -1e-9:
                        o[i:k+1] = reversed(o[i:k+1]); imp = True
        c = sum(cost[o[i]][o[(i+1) % n]] for i in range(n))
        if best_c is None or c < best_c: best_c, best_o = c, list(o)
    return [nodes[i] for i in (best_o or [])], (best_c or 0.0)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", default="resultat.json")
    ap.add_argument("--db", default=None)
    ap.add_argument("--output", default=None)
    a = ap.parse_args()
    res = json.load(open(a.input, encoding="utf-8"))
    data = SqliteData(a.db or str(_src / "produksjonskalkyle.db"))
    iw = {}
    for rl in data.routing_lines:
        wc = data.work_center(rl.work_center_code)
        if wc: iw.setdefault((rl.item_no, wc.location_code), wc.code)
    fd = {}
    for rl in data.routing_lines:
        f = _fti_family(rl.item_no)
        if f not in fd and dims(f)[0] is not None: fd[f] = dims(f)
    g = defaultdict(list)
    for r in res.get("production_plan", []):
        wc = iw.get((r["product"], r["location"]), "KJØP")
        if wc == "KJØP": continue
        g[(wc, r["period"])].append({"p": r["product"], "f": _fti_family(r["product"]),
                                     "q": r["quantity"], "l": r["location"],
                                     "d": r.get("product_desc", "")})
    print("=" * 90)
    print(" SEKVENS per arbeidssenter x uke (FTI-hierarki, changeover minimert)")
    print("=" * 90)
    tot_min = tot_kr = 0.0
    for (wc, t) in sorted(g):
        rows = g[(wc, t)]
        byf = defaultdict(list)
        for r in rows: byf[r["f"]].append(r)
        fams = sorted(byf)
        if len(fams) <= 1:
            order, pen = fams, 0.0
        else:
            cm = [[0.0] * len(fams) for _ in range(len(fams))]
            for i, x in enumerate(fams):
                for j, y in enumerate(fams):
                    if i != j: cm[i][j] = _changeover_minutes(x, y, fd)
            order, pen = tsp(fams, cm)
        seq = []
        for f in order:
            seq.extend(sorted(byf[f], key=lambda r: r["p"]))
        print(f"\n[{wc}] UKE {t} - {len(fams)} familier, {len(rows)} produkter, omstilling {pen:.1f} min")
        prev = None
        for i, r in enumerate(seq, 1):
            if prev and r["f"] != prev:
                print(f"      -- changeover {prev} -> {r['f']}: {_changeover_minutes(prev, r['f'], fd):.1f} min")
            print(f"  {i:<3} {r['p']:<28} {r['f']:<12} {r['q']:>12,.1f} LM")
            prev = r["f"]
        tot_min += pen
        wo = data.work_center(wc)
        if wo and wo.total_cost_hour: tot_kr += (pen / 60.0) * wo.total_cost_hour
    print(f"\n{'=' * 90}")
    print(f" TOTALT omstilling: {tot_min:,.1f} min ({tot_min/60:.1f} t) ~ {tot_kr:,.0f} kr")
    print("=" * 90)
    if a.output:
        out = {"kilde": a.input, "total_min": round(tot_min, 2), "total_kr": round(tot_kr, 2),
               "sekvens": [{"wc": wc, "uke": t, "rekkefolge": [
                   {"pos": i+1, "produkt": r["p"], "familie": r["f"], "lokasjon": r["l"],
                    "kvantum": round(r["q"], 2), "beskrivelse": r["d"]}
                   for i, r in enumerate(rows)]} for (wc, t), rows in sorted(g.items())]}
        json.dump(out, open(a.output, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
        print(f"  [OK] -> {a.output}")


if __name__ == "__main__":
    main()