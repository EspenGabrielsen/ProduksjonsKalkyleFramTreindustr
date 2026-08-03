#!/usr/bin/env python3
"""Finn hvilken tidlig demand-linje som kolliderer med JJ16123-kjeden.

Kjeden: JJ16123EH (#80) + JJ16123EF (#79) + JJ16123 (#78) er OPTIMALE alene.
"de 79 første + EH" er Infeasible. Vi finner den nøyaktige linjen som bryter.
"""

import sys
from pathlib import Path

import pulp

_src = Path(__file__).resolve().parent.parent
if str(_src) not in sys.path:
    sys.path.insert(0, str(_src))

from optimization_engine import OptimizationEngine


def test(engine, demand_list, label: str) -> str:
    engine.demand = list(demand_list)
    model = engine.build_model()
    model.solve(pulp.PULP_CBC_CMD(msg=0, timeLimit=15))
    return pulp.LpStatus[model.status]


def main():
    engine = OptimizationEngine()
    engine.load_all()
    full = list(engine.demand)

    # Kjernen: de tre JJ16123-linjene (#77-#80, 0-indeksert 77,78,79)
    kjede = full[76:80]  # JG21145, JJ16123, JJ16123EF, JJ16123EH
    print("=== Kjerne (4 linjer) ===")
    for d in kjede:
        print(f"  {d.product_id}, lok={d.location_code}, uke={d.period}, qty={d.quantity}")

    print(f"\nKjerne alene: {test(engine, kjede, 'kjerne')}")

    # Legg til de første N linjene én etter én til det brekker
    for i, d in enumerate(full):
        if i >= 76:
            break
        test_set = full[:i+1] + kjede
        status = test(engine, test_set, f"de {i+1} første + kjerne")
        if status != "Optimal":
            print(f"\n[FUNNET] Linje #{i+1} bryter med kjeden:")
            print(f"  {d.product_id}, lok={d.location_code}, uke={d.period}, qty={d.quantity}")
            # Vis de 10 før brytningspunktet
            print(f"\n=== De siste 10 før brytningspunktet ===")
            for j in range(max(0, i-9), i+1):
                print(f"  #{j+1}: {full[j].product_id}, lok={full[j].location_code}, uke={full[j].period}, qty={full[j].quantity}")
            break

    engine.db.close()


if __name__ == "__main__":
    main()