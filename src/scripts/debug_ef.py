#!/usr/bin/env python3
"""Undersøk hvorfor JJ19148EF gjør modellen infeasible (ny arkitektur)."""

import sys
from pathlib import Path

_src = Path(__file__).resolve().parent.parent
if str(_src) not in sys.path:
    sys.path.insert(0, str(_src))

from optimization_engine import OptimizationEngine


def main():
    engine = OptimizationEngine()
    engine.load_all()

    print("=== JJ19148EF BOM-kjede ===")
    # Traverser BOM rekursivt og vis alle nivåer med qty_per/scrap
    def show(item, indent, visited):
        if indent > 5 or item in visited:
            return
        visited.add(item)
        for bl in engine.data.bom_for(item):
            print("  " * indent + f"{item} -> {bl.component_item_no} "
                  f"(qty_per={bl.quantity_per}, scrap={bl.scrap_pct}%, "
                  f"co_pct={bl.co_product_pct}%)")
            show(bl.component_item_no, indent + 1, visited.copy())

    show("JJ19148EF", 1, set())

    print("\n=== Kan hver komponent dekkes på KOD? ===")
    routing_items = {rl.item_no for rl in engine.data.routing_lines}
    purchasable = {ic.item_no for ic in engine.data.item_costs}
    prod_locs = {}
    for rl in engine.data.routing_lines:
        wc = engine.data.work_center(rl.work_center_code)
        if wc and wc.location_code:
            prod_locs.setdefault(rl.item_no, set()).add(wc.location_code)

    def check(item, indent, visited):
        if indent > 5 or item in visited:
            return
        visited.add(item)
        # Lager ingen ny oppføring her — kjører i hovedløkka nedenfor
        for bl in engine.data.bom_for(item):
            check(bl.component_item_no, indent + 1, visited.copy())

    alle = set()
    def collect(item, visited):
        if item in visited:
            return
        visited.add(item)
        for bl in engine.data.bom_for(item):
            alle.add(bl.component_item_no)
            collect(bl.component_item_no, visited.copy())
    collect("JJ19148EF", set())
    alle.add("JJ19148EF")

    import sqlite3
    conn = sqlite3.connect("src/produksjonskalkyle.db")
    for item in sorted(alle):
        has_routing = item in routing_items
        has_price = item in purchasable
        locs = prod_locs.get(item, set())
        prod = conn.execute("SELECT item_type, base_uom FROM products WHERE item_no=?", (item,)).fetchone()
        lok = "KOD i prod_locs" if "KOD" in locs else ("ANDRE: " + str(sorted(locs)) if locs else "IKKE produserbar")
        status = "✅" if (has_routing and "KOD" in locs) or has_price else "❌"
        print(f"  {status} {item:<30} routing={has_routing} pris={has_price} "
              f"{'type=' + str(prod[0]) if prod else 'MANGLER PRODUKT'}")
    conn.close()

    # VIS ekspandert demand for JJ19148EF-kjeden
    print("\n=== Ekspandert demand (per produkt/period/lok) ===")
    expanded = engine._expand_demand()
    for key, qty in sorted(expanded.items()):
        if key[0] in alle:
            print(f"  {key[0]:<25} lok={key[1]:<5} uke={key[2]:<4} behov={qty:.4f}")

    # VIS kapasitet per arbeidssenter på KOD
    print("\n=== Kapasitet per arbeidssenter på KOD ===")
    for wc in engine.data.work_centers:
        if wc.location_code == "KOD":
            eff = (wc.effective_capacity_pct or 100) / 100
            cap = wc.capacity_hours_day * eff * 5
            print(f"  {wc.code:<20} kapasitet/uke={cap:.2f} timer")

    engine.db.close()


if __name__ == "__main__":
    main()