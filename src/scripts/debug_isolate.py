#!/usr/bin/env python3
"""Isoler om spesifikke demand-linjer alene er infeasible."""

import sys
from pathlib import Path

import pulp

_src = Path(__file__).resolve().parent.parent
if str(_src) not in sys.path:
    sys.path.insert(0, str(_src))

from optimization_engine import OptimizationEngine


def test(engine, demand_slice, label: str):
    engine.demand = list(demand_slice)
    model = engine.build_model()
    model.solve(pulp.PULP_CBC_CMD(msg=0, timeLimit=20))
    status = pulp.LpStatus[model.status]
    print(f"  {label}: {status}")
    return status == "Optimal"


def main():
    engine = OptimizationEngine()
    engine.load_all()
    full_demand = list(engine.demand)  # lagre original — ikke muter
    total = len(full_demand)
    print(f"Totalt {total} validerte demand-linjer")

    print("=== Isolasjonstester ===")
    # Bare #79+#80 (JJ16123EF + JJ16123EH)
    test(engine, full_demand[78:80], "bare #79+#80 (JJ16123EF+EH)")
    # Bare #78+#79 (uten JJ16123EH)
    test(engine, full_demand[77:79], "bare #78+#79 (JJ16123+JJ16123EF)")
    # Bare #78+#79+#80
    test(engine, full_demand[77:80], "#78+#79+#80 (JJ16123+EF+EH)")
    # Bare #77+#78+#79+#80
    test(engine, full_demand[76:80], "#77-#80")
    # Bare #71-#80
    test(engine, full_demand[70:80], "#71-#80")
    # De 79 første (uten #80)
    test(engine, full_demand[:79], "de 79 første (uten EH)")
    # De 80 første (med EH)
    test(engine, full_demand[:80], "de 80 første (med EH)")

    # Hva er linje #77?
    print("\n=== Demand-linjene 76-80 ===")
    for i, d in enumerate(engine.demand[76:80], 77):
        print(f"  #{i}: {d.product_id}, lok={d.location_code}, uke={d.period}, qty={d.quantity}")

    engine.db.close()


if __name__ == "__main__":
    main()