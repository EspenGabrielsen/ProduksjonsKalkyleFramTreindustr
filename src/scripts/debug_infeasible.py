#!/usr/bin/env python3
"""Finn årsaken til infeasible modell (dekomponering av demand)."""

import sys
from pathlib import Path

import pulp

_src = Path(__file__).resolve().parent.parent
if str(_src) not in sys.path:
    sys.path.insert(0, str(_src))

from optimization_engine import OptimizationEngine


def _is_feasible(engine, N: int) -> bool:
    """Test om de første N demand-linjene er løsbare."""
    engine.demand = engine.demand[:N]
    model = engine.build_model()
    model.solve(pulp.PULP_CBC_CMD(msg=0, timeLimit=20))
    return pulp.LpStatus[model.status] == "Optimal"


def main():
    engine = OptimizationEngine()
    engine.load_all()
    full_demand = list(engine.demand)  # lagre original — IKKE muter
    total = len(full_demand)
    print(f"Totalt {total} validerte demand-linjer")

    def is_feasible(N: int) -> bool:
        """Test om de første N demand-linjene er løsbare (uten å mutere)."""
        engine.demand = full_demand[:N]
        model = engine.build_model()
        model.solve(pulp.PULP_CBC_CMD(msg=0, timeLimit=20))
        return pulp.LpStatus[model.status] == "Optimal"

    # Test kjente deler først
    for N in [50, 100, 150, 200, 300, 500, total]:
        ok = is_feasible(N)
        print(f"  N={N:>4}: {'Optimal' if ok else 'Infeasible'}")
        if not ok:
            break

    # Binærsøk etter brytningspunktet (fra 1 til total)
    low, high = 1, total
    if is_feasible(high):
        print(f"\n[OK] Alle {total} linjer er løsbare!")
        return
    while low < high:
        mid = (low + high) // 2
        if is_feasible(mid):
            low = mid + 1
        else:
            high = mid
    print(f"\n[FUNNET] Første infeasible demand-linje er #{low}")
    d = full_demand[low - 1] if low - 1 < len(full_demand) else None
    if d:
        print(f"  Produkt: {d.product_id}, uke: {d.period}, lokasjon: {d.location_code}, qty: {d.quantity}")
    # Fullt demand etter at vi fant punktet
    engine.demand = full_demand


if __name__ == "__main__":
    main()