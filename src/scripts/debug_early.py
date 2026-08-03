#!/usr/bin/env python3
"""Finn hvilken tidlig demand-linje som kolliderer med #80 (JJ16123EH)."""

import sys
from pathlib import Path

import pulp

_src = Path(__file__).resolve().parent.parent
if str(_src) not in sys.path:
    sys.path.insert(0, str(_src))

from optimization_engine import OptimizationEngine


def test(engine, demand_list, label: str) -> bool:
    engine.demand = list(demand_list)
    model = engine.build_model()
    model.solve(pulp.PULP_CBC_CMD(msg=0, timeLimit=15))
    status = pulp.LpStatus[model.status]
    ok = status == "Optimal"
    print(f"  {label}: {status}")
    return ok


def main():
    engine = OptimizationEngine()
    engine.load_all()
    full = list(engine.demand)

    # Unik: fjern dupliserte (produkt, lok, periode) — demand_dict summerer uansett
    unique = []
    seen = set()
    for d in full:
        key = (d.product_id, d.location_code, d.period)
        if key not in seen:
            seen.add(key)
            unique.append(d)

    print(f"Totalt {len(full)} linjer, {len(unique)} unike")

    # EH-kandidater (alle med JJ16123EH)
    eh_indices = [i for i, d in enumerate(full) if d.product_id == "JJ16123EH"]
    print(f"JJ16123EH-linjer: {eh_indices}")
    if not eh_indices:
        print("Fant ikke JJ16123EH!")
        return
    eh_line = full[eh_indices[0]]
    print(f"  {eh_line.product_id}, lok={eh_line.location_code}, uke={eh_line.period}, qty={eh_line.quantity}")

    # Test: de første N (uten EH) + EH-linjen
    for N in [10, 20, 30, 40, 50, 60, 70, 79]:
        baseline = [d for d in full[:N] if d.product_id != "JJ16123EH"]
        test(engine, baseline + [eh_line], f"de {N:>2} første + EH")

    engine.db.close()


if __name__ == "__main__":
    main()