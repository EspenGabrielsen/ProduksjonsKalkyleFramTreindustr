"""
lag_baseline.py - Generer baseline-referanse for verifikasjon av optimaliseringer.

Bruk:
  python lag_baseline.py                  # Generer baseline (før optimalisering)
  python lag_baseline.py --sammenlign     # Sammenlign med baseline (etter optimalisering)

Output:
  baseline_kalkyle.json       - ProductCostResult for alle produkter
  baseline_simulering.json    - SimulationComparison for alle produkter med overstyringer
"""

import json
import sys
import os
import io

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
elif hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from datetime import date, datetime
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from kostberegning import (
    ExcelData, CostCalculator, SimulationEngine, SimulationOverride,
    ProductCostResult, SimulationComparison, MaterialCostDetail,
    OperationCostDetail, ByProductDetail,
)

# Testdata-Excel ligger i src/ (ved siden av kostberegning.py)
_SRC_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXCEL_PATH = os.path.join(_SRC_DIR, "Produksjonsmodell_Testdata_v3.xlsx")

# Baseline-output legges i output/
_OUTPUT_DIR = os.path.join(os.path.dirname(_SRC_DIR), "output")
BASELINE_KALKYLE = os.path.join(_OUTPUT_DIR, "baseline_kalkyle.json")
BASELINE_SIMULERING = os.path.join(_OUTPUT_DIR, "baseline_simulering.json")


def _serialize(obj):
    """Serialiser dataclass-objekter til dict (samme som i kostberegning.py)."""
    if hasattr(obj, '__dataclass_fields__'):
        return {k: _serialize(v) for k, v in obj.__dict__.items()}
    if isinstance(obj, (date, datetime)):
        return obj.isoformat()
    if isinstance(obj, dict):
        return {k: _serialize(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_serialize(v) for v in obj]
    # Håndter tupler som dict-nøkler (f.eks. (parent, component))
    if isinstance(obj, tuple):
        return list(obj)
    return obj


def serialiser_resultater(results):
    """Serialiser liste med resultatobjekter til JSON-klar dict."""
    return _serialize(results)


def lag_baseline():
    """Kjør beregning og simulering, lagre til JSON-filer."""
    print(f"  >>> Laster data fra {EXCEL_PATH}")
    data = ExcelData(EXCEL_PATH)
    
    print(f"     {len(data.products)} produkter, {len(data.work_centers)} arbeidssentre, "
          f"{len(data.bom_lines)} BOM-linjer, {len(data.routing_lines)} routing-linjer")

    # === 1. CostCalculator.calculate_all() ===
    print(f"\n  >>> Kjører CostCalculator.calculate_all()...")
    calculator = CostCalculator(data)
    kalkyle_results = calculator.calculate_all()
    print(f"     {len(kalkyle_results)} resultater")
    
    with open(BASELINE_KALKYLE, "w", encoding="utf-8") as f:
        json.dump(serialiser_resultater(kalkyle_results), f, indent=2, ensure_ascii=False)
    print(f"     -> Lagret {BASELINE_KALKYLE}")

    # === 2. SimulationEngine.compare_all() med overstyringer ===
    print(f"\n  >>> Kjører SimulationEngine.compare_all() med overstyringer...")
    
    overrides = SimulationOverride(
        planned_quantity=50000,
    )
    
    # Overstyr én råvarepris
    rm_items = [p for p in data.products if p.item_type == 'Raw Material']
    if rm_items:
        overrides.item_costs[rm_items[0].item_no] = 3500.0
    
    # Overstyr svinn for første BOM-linje
    if data.bom_lines:
        bl = data.bom_lines[0]
        overrides.bom_scrap[(bl.parent_item_no, bl.component_item_no)] = 8.0
    
    # Overstyr work center lønn for første aktive WC
    active_wcs = [wc for wc in data.work_centers if wc.active]
    if active_wcs:
        overrides.work_centers[active_wcs[0].code] = {"labor_cost_hour": 600.0}
    
    # Overstyr routing for første routing-linje
    if data.routing_lines:
        rl = data.routing_lines[0]
        overrides.routing[(rl.item_no, rl.operation_no, rl.work_center_code)] = {
            "run_time_minutes": 0.20
        }
    
    engine = SimulationEngine(data)
    sim_results = engine.compare_all(overrides)
    print(f"     {len(sim_results)} sammenligningsresultater")
    
    with open(BASELINE_SIMULERING, "w", encoding="utf-8") as f:
        json.dump(serialiser_resultater(sim_results), f, indent=2, ensure_ascii=False)
    print(f"     -> Lagret {BASELINE_SIMULERING}")
    
    print(f"\n  >>> Baseline generert! Referansefilene er klare.")
    return True


def sammenlign_med_baseline():
    """Sammenlign nåværende output med baseline."""
    # Sjekk at baseline-filer finnes
    if not Path(BASELINE_KALKYLE).exists() or not Path(BASELINE_SIMULERING).exists():
        print(f"  >>> FEIL: Baseline-filer finnes ikke. Kjør uten --sammenlign først.")
        return False
    
    # Last baseline
    with open(BASELINE_KALKYLE, "r", encoding="utf-8") as f:
        baseline_kalkyle = json.load(f)
    with open(BASELINE_SIMULERING, "r", encoding="utf-8") as f:
        baseline_simulering = json.load(f)
    
    # Kjør på nytt med dagens kode
    data = ExcelData(EXCEL_PATH)
    
    # Kalkyle
    calculator = CostCalculator(data)
    kalkyle_results = calculator.calculate_all()
    nye_kalkyle = serialiser_resultater(kalkyle_results)
    
    # Simulering
    overrides = SimulationOverride(planned_quantity=50000)
    rm_items = [p for p in data.products if p.item_type == 'Raw Material']
    if rm_items:
        overrides.item_costs[rm_items[0].item_no] = 3500.0
    if data.bom_lines:
        bl = data.bom_lines[0]
        overrides.bom_scrap[(bl.parent_item_no, bl.component_item_no)] = 8.0
    active_wcs = [wc for wc in data.work_centers if wc.active]
    if active_wcs:
        overrides.work_centers[active_wcs[0].code] = {"labor_cost_hour": 600.0}
    if data.routing_lines:
        rl = data.routing_lines[0]
        overrides.routing[(rl.item_no, rl.operation_no, rl.work_center_code)] = {
            "run_time_minutes": 0.20
        }
    
    engine = SimulationEngine(data)
    sim_results = engine.compare_all(overrides)
    nye_simulering = serialiser_resultater(sim_results)
    
    # Sammenlign
    feil = 0
    
    if baseline_kalkyle != nye_kalkyle:
        print(f"  >>> FEIL: kalkyle-resultater er forskjellige!")
        print(f"      Baseline: {len(baseline_kalkyle)} resultater")
        print(f"      Nåværende: {len(nye_kalkyle)} resultater")
        feil += 1
    else:
        print(f"  ✔ Kalkyle-resultater: IDENTISKE ({len(nye_kalkyle)} resultater)")
    
    if baseline_simulering != nye_simulering:
        print(f"  >>> FEIL: simuleringsresultater er forskjellige!")
        print(f"      Baseline: {len(baseline_simulering)} resultater")
        print(f"      Nåværende: {len(nye_simulering)} resultater")
        feil += 1
    else:
        print(f"  ✔ Simuleringsresultater: IDENTISKE ({len(nye_simulering)} resultater)")
    
    if feil == 0:
        print(f"\n  ✅ ALT OK! Output er identisk med baseline.")
        return True
    else:
        print(f"\n  ❌ {feil} feil funnet.")
        return False


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Lag og verifiser baseline-referanse")
    parser.add_argument("--sammenlign", action="store_true", 
                        help="Sammenlign nåværende output med baseline")
    args = parser.parse_args()
    
    if args.sammenlign:
        sammenlign_med_baseline()
    else:
        lag_baseline()