#!/usr/bin/env python3
"""Sammenligner batch-storrelse med og uten prognose for JD19123-familien."""

import sys
from pathlib import Path

_src = Path(__file__).resolve().parent.parent
if str(_src) not in sys.path:
    sys.path.insert(0, str(_src))

from optimization_engine import OptimizationEngine

DB = str(_src / "produksjonskalkyle.db")
FOKUS = "JD19123"


def kjor(forecast_weeks, holding_cost_pct):
    engine = OptimizationEngine(
        db_path=DB,
        days_per_period=500.0,
        holding_cost_pct=holding_cost_pct,
        forecast_weeks=forecast_weeks,
    )
    engine.load_all()
    result = engine.solve(time_limit_seconds=120)

    prod = [r for r in result.production_plan if r["product"].startswith(FOKUS)]
    batches = [r for r in result.batch_decisions if r["product"].startswith(FOKUS)]
    lager = [r for r in result.inventory_levels if r["product"].startswith(FOKUS)]

    perioder = sorted({r["period"] for r in prod})
    return {
        "status": result.status,
        "total_kost": result.total_cost,
        "prod_linjer": len(prod),
        "batch_linjer": len(batches),
        "lager_linjer": len(lager),
        "total_prod": sum(r["quantity"] for r in prod),
        "total_batch_kost": sum(r["setup_cost"] for r in batches),
        "total_lager": sum(r["quantity"] for r in lager),
        "perioder": perioder,
        "storste_batch": {
            t: max((r["quantity"] for r in prod if r["period"] == t), default=0)
            for t in perioder
        },
    }


def main():
    print("=" * 74)
    print(f" BATCH-STØRRELSESAVVEINING - fokusfamilie {FOKUS}")
    print("=" * 74)

    cases = [
        ("UTEN prognose, 0% lagerkost", 0, 0.0),
        ("UTEN prognose, 5% lagerkost", 0, 5.0),
        ("MED prognose,   0% lagerkost", 12, 0.0),
        ("MED prognose,   5% lagerkost", 12, 5.0),
    ]

    for navn, fw, hcp in cases:
        print(f"\n{'-' * 74}")
        print(f" {navn}  (forecast_weeks={fw}, holding_cost_pct={hcp})")
        print(f"{'-' * 74}")
        r = kjor(fw, hcp)
        print(f"  Status:            {r['status']}")
        print(f"  Total kost:        {r['total_kost']:>12,.2f} kr")
        print(f"  Produksjonslinjer: {r['prod_linjer']}")
        print(f"  Batch-starter:     {r['batch_linjer']}  (setup {r['total_batch_kost']:,.0f} kr)")
        print(f"  Lagerlinjer:       {r['lager_linjer']}  (tot. lager {r['total_lager']:,.0f} LM)")
        print(f"  Perioder:          {r['perioder']}")
        print(f"  Største batch/uke: {r['storste_batch']}")
        print(f"  Sum produsert:     {r['total_prod']:>12,.1f} LM")


if __name__ == "__main__":
    main()