#!/usr/bin/env python3
"""
testdata_for_app.py - Seed testdata for Marimo-appen i test-modus.

Brukes når PRODUKSJONSKALKYLE_TEST=true er satt.
Fyller in-memory SQLite-databasen med utvalgte testprodukter
som representerer de vanligste kostnadsscenariene.

Relaterte filer:
  - test_transport.py  (samtlige 5 automatiske tester)
  - TRANSPORTVARE_PLAN.md (full implementeringsplan)
"""

from data_repo import DataRepo

# ── Testprodukter ────────────────────────────────────────────────
TEST_VARER = {
    "BL98520": {
        "description": "FURU 45X045 GLATTKANT UBEHANDLET",
        "item_type": "Finished Good",
        "product_group": "Byggelist",
        "base_uom": "LM",
    },
    "JD16073": {
        "description": "16x73 Just. Kledn. ca. 1250 m/pk",
        "item_type": "Finished Good",
        "product_group": "Panel",
        "base_uom": "LM",
    },
    "JD16098TF": {
        "description": "16X98*GR.VISIR+TREBITT JUST.KLEDN NR.100",
        "item_type": "Finished Good",
        "product_group": "Justert kledning",
        "base_uom": "LM",
    },
}

# ── Råvarer / komponenter ────────────────────────────────────────
KOMPONENTER = {
    "RM_50x75_US_V_Gran": {"description": "50x75 US/V Gran", "item_type": "Raw Material", "product_group": "Skrulast", "base_uom": "M3"},
    "KD_Trebitt": {"description": "Maling - Trebitt", "item_type": "Raw Material", "product_group": "Maling", "base_uom": "LTR"},
    "KD_Visir": {"description": "Maling - Visir", "item_type": "Raw Material", "product_group": "Maling", "base_uom": "LTR"},
    "JD16098": {"description": "16x98 Just. u/mal", "item_type": "Semi Finished", "product_group": "Justert kledning", "base_uom": "LM"},
}

# ── Arbeidssentre ────────────────────────────────────────────────
WORK_CENTERS = [
    {"code": "HOVEDHOVEL", "description": "Hovedhovel (HH)", "location_code": "KOD",
     "labor_cost_hour": 1800.0, "machine_cost_hour": 2000.0, "overhead_cost_hour": 803.0,
     "capacity_hours_day": 16.0, "effective_capacity_pct": 92.0},
    {"code": "SPESIALHOVEL", "description": "Spesialhovel (SH)", "location_code": "KOD",
     "labor_cost_hour": 600.0, "machine_cost_hour": 700.0, "overhead_cost_hour": 305.0,
     "capacity_hours_day": 16.0, "effective_capacity_pct": 94.0},
    {"code": "KVHOVEL", "description": "Kvaas hovel", "location_code": "KV",
     "labor_cost_hour": 500.0, "machine_cost_hour": 200.0, "overhead_cost_hour": 100.0,
     "capacity_hours_day": 16.0, "effective_capacity_pct": 92.0},
]

# ── Lokasjoner ───────────────────────────────────────────────────
LOCATIONS = [
    {"code": "KOD", "name": "Kodal Fabrikk", "location_type": "Factory"},
    {"code": "KV", "name": "Kvaas", "location_type": "Factory"},
    {"code": "EIK", "name": "Eikaas", "location_type": "Factory"},
]

# ── BOM-linjer ───────────────────────────────────────────────────
BOM_LINES = [
    {"parent": "BL98520", "comp": "RM_50x75_US_V_Gran", "qty": 266.67, "uom": "LM", "scrap": 0.0, "co_pct": 0.0, "co_item": ""},
    {"parent": "JD16073", "comp": "RM_50x75_US_V_Gran", "qty": 533.05, "uom": "LM", "scrap": 0.0, "co_pct": 0.5, "co_item": "JD16073B"},
    {"parent": "JD16098TF", "comp": "JD16098", "qty": 1.0, "uom": "LM", "scrap": 0.0, "co_pct": 0.0, "co_item": ""},
    {"parent": "JD16098TF", "comp": "KD_Trebitt", "qty": 87.42, "uom": "LTR", "scrap": 15.0, "co_pct": 0.0, "co_item": ""},
    {"parent": "JD16098TF", "comp": "KD_Visir", "qty": 77.53, "uom": "LTR", "scrap": 15.0, "co_pct": 0.0, "co_item": ""},
]

# ── Routing-linjer ───────────────────────────────────────────────
ROUTING_LINES = [
    {"item": "BL98520", "op_no": 40, "op_code": "HOVLING", "wc": "HOVEDHOVEL", "setup": 30.0, "run": 0.040, "batch": 3000.0},
    {"item": "JD16073", "op_no": 40, "op_code": "HOVLING", "wc": "SPESIALHOVEL", "setup": 30.0, "run": 0.029, "batch": 3200.0},
    {"item": "JD16098", "op_no": 40, "op_code": "HOVLING", "wc": "HOVEDHOVEL", "setup": 60.0, "run": 0.011, "batch": 25000.0},
]


def er_test_modus() -> bool:
    """Sjekk om appen kjører i test-modus."""
    import os
    return os.environ.get("PRODUKSJONSKALKYLE_TEST", "").lower() in ("true", "1", "yes")


def reset_test_db() -> DataRepo:
    """Tøm eksisterende test-database og opprett en ny, tom.

    Kalles kun én gang ved app-start i test-modus. Sikrer at testdata
    ikke akkumuleres mellom kjøringer.

    Bruker clear_all_data() i stedet for fil-sletting fordi temp-filen kan
    være låst av WAL-filer fra tidligere prosesser.

    Returns:
        Ny DataRepo med tom test-database
    """
    db = DataRepo()
    db.initialize()
    db.clear_all_data()
    db.clear_change_log()
    return db


def seed_test_db(db: DataRepo) -> bool:
    """Fyll in-memory SQLite-databasen med testdata.

    Args:
        db: DataRepo-instans (forventes å være :memory:)

    Returns:
        True hvis seeding lyktes
    """
    # Locations
    db.upsert_locations(LOCATIONS, source="seed")
    # Work centers
    db.upsert_work_centers(WORK_CENTERS, source="seed")
    # Produkter
    products = []
    for k, v in TEST_VARER.items():
        products.append({"item_no": k, **v})
    for k, v in KOMPONENTER.items():
        products.append({"item_no": k, **v})
    db.upsert_products(products, source="seed")
    # BOM
    for bl in BOM_LINES:
        db.upsert_bom_lines([{
            "parent_item_no": bl["parent"], "component_item_no": bl["comp"],
            "quantity_per": bl["qty"], "uom": bl["uom"],
            "scrap_pct": bl["scrap"], "co_product_pct": bl["co_pct"],
            "co_product_item_no": bl["co_item"],
        }], source="seed")
    # Routing
    for rl in ROUTING_LINES:
        db.upsert_routing_lines([{
            "item_no": rl["item"], "operation_no": rl["op_no"],
            "operation_code": rl["op_code"], "work_center_code": rl["wc"],
            "setup_time_minutes": rl["setup"], "run_time_minutes": rl["run"],
            "batch_size": rl["batch"],
        }], source="seed")

    return True