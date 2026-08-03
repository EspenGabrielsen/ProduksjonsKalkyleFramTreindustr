#!/usr/bin/env python3
"""Undersøk hvorfor JJ16123EH gjør modellen infeasible."""

import sys
from pathlib import Path

_src = Path(__file__).resolve().parent.parent
if str(_src) not in sys.path:
    sys.path.insert(0, str(_src))

from optimization_engine import OptimizationEngine, _finn_lm_per_m3_robust


def main():
    engine = OptimizationEngine()
    engine.load_all()

    # HYPOTESE: demand på JJ16123EH tvinger produksjon av JJ16123GH og JJ16123.
    # Sjekk om de 79 første demand-linjene har JJ16123*-produkter på andre lokasjoner.
    print("=== De første 80 validerte demand-linjene med JJ16123* ===")
    for i, d in enumerate(engine.demand[:80], 1):
        if d.product_id.startswith("JJ16123"):
            print(f"  #{i}: {d.product_id}, lokasjon={d.location_code}, uke={d.period}, qty={d.quantity}")

    print()
    print("=== Alle JJ16123* demand-linjer i hele demand (med lokasjon) ===")
    by_item_loc: dict[tuple[str, str], float] = {}
    for d in engine.demand:
        if d.product_id.startswith("JJ16123"):
            by_item_loc[(d.product_id, d.location_code)] = (
                by_item_loc.get((d.product_id, d.location_code), 0) + d.quantity
            )
    for (pid, loc), qty in sorted(by_item_loc.items()):
        print(f"  {pid:<15} lok={loc:<5} total_qty={qty}")

    print()
    print("=== Materialbalansen for JJ16123GH (krevd av JJ16123EH) ===")
    # Finn BOM-avhengigheter
    parents_of = {}
    for bs in engine.bom_structure:
        parents_of.setdefault(bs.child_product_id, []).append((bs.parent_product_id, bs.yield_factor))
    for child in ("JJ16123", "JJ16123GH", "JJ16123EH"):
        parents = parents_of.get(child, [])
        print(f"  {child}: brukes av {parents}")

    prod_locs = {}
    for rl in engine.data.routing_lines:
        wc = engine.data.work_center(rl.work_center_code)
        if wc and wc.location_code:
            prod_locs.setdefault(rl.item_no, set()).add(wc.location_code)

    print("=== BOM for JJ16123EH ===")
    for bs in engine.bom_structure:
        if bs.parent_product_id == "JJ16123EH":
            print(f"  JJ16123EH -> {bs.child_product_id}  (yield={bs.yield_factor})")

    # Finn alle BOM-linjer som involverer JJ16123
    print("\n=== Alle BOM-linjer som involverer JJ16123* ===")
    for bs in engine.bom_structure:
        if bs.parent_product_id.startswith("JJ16123") or bs.child_product_id.startswith("JJ16123"):
            print(f"  {bs.parent_product_id} -> {bs.child_product_id}  (yield={bs.yield_factor})")

    print("\n=== Routing (produksjonslokasjoner) ===")
    for item in sorted(prod_locs):
        if item.startswith("JJ16123"):
            print(f"  {item}: produserbar på {sorted(prod_locs[item])}")

    print("\n=== Item costs for JJ16123-kjeden ===")
    import sqlite3
    conn = sqlite3.connect("src/produksjonskalkyle.db")
    rows = conn.execute(
        "SELECT item_no, unit_cost FROM item_costs WHERE item_no LIKE '%JJ16123%'"
    ).fetchall()
    for r in rows:
        print(f"  {r[0]}: {r[1]}")
    conn.close()

    print("\n=== LM/M3 for JJ16123EH ===")
    print(f"  {_finn_lm_per_m3_robust(engine.data, 'JJ16123EH')}")

    engine.db.close()


if __name__ == "__main__":
    main()