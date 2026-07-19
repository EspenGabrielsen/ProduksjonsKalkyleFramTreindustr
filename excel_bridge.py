#!/usr/bin/env python3
"""
excel_bridge.py - Import/eksport mellom Excel-filer og SQLite-databasen.

Brukes til:
  1. Importere data fra Excel til SQLite (med endringslogg)
  2. Eksportere data fra SQLite til Excel (for nedlasting / SharePoint)

Alle importfunksjonene sammenligner felt-for-felt med eksisterende data i SQLite
og loggfører kun faktiske endringer.
"""

import os
import sys
import io
from datetime import date, datetime
from pathlib import Path
from typing import Optional

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment as XLComment

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
elif hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from data_repo import DataRepo


# ──────────────────────────────────────────────────────────────────────
#  Kolonnebeskrivelser (fra oppdater_mal.py)
# ──────────────────────────────────────────────────────────────────────

KOLONNER = {
    "Product Master": {
        "Item No": "Unik identifikator for varen. Eksempel: RM001, FG001, BP001",
        "Description": "Beskrivende navn pa varen. Eksempel: Skrulast 48x198",
        "Item Type": "Type vare: Raw Material, Semi Finished, Finished Good, By Product, Trading Item",
        "Product Group": "Gruppering av varer. Eksempel: Skrulast, Panel, Kledning, Spon",
        "Base Unit of Measure": "Standard maleenhet. Eksempel: LM, M3, KG, PCS",
        "Active": "Er varen aktiv? Ja / Nei",
    },
    "Locations": {
        "Location Code": "Unik kode for lokasjonen. Eksempel: KOD",
        "Location Name": "Navn pa lokasjonen. Eksempel: Kodal Fabrikk",
        "Location Type": "Type lokasjon: Factory, Warehouse, Distribution Center, Sales Office",
        "Active": "Er lokasjonen aktiv? Ja / Nei",
    },
    "Work Centers": {
        "Work Center Code": "Unik identifikator for arbeidssenteret. Eksempel: HOVEDHOVEL",
        "Description": "Beskrivende navn pa arbeidssenteret",
        "Location Code": "Fabrikken arbeidssenteret tilhorer",
        "Labor Cost per Hour": "Arbeidskostnad per time (lonn, avgift, pensjon, ferie). Eksempel: 550",
        "Machine Cost per Hour": "Maskinkostnad per time (avskrivn., service, leasing, vedlikehold, energi). Eksempel: 900",
        "Overhead Cost per Hour": "Indirekte produksjonskostnader. Eksempel: 150",
        "Capacity Hours per Day": "Tilgjengelige timer per dag. Eksempel: 16",
        "Effective Capacity %": "Hvor stor del av tiden som kan brukes til produksjon. Eksempel: 85",
        "Active": "Er arbeidssenteret aktivt? Ja / Nei",
    },
    "Operation Master": {
        "Operation Code": "Unik kode for operasjonen. Eksempel: HOVLING, MALING, PACKING",
        "Description": "Beskrivelse av operasjonen",
        "Default Work Center": "Anbefalt arbeidssenter for operasjonen",
        "Standard Unit": "Maleenhet for produksjonstid. Eksempel: Minutes, Hours",
        "Active": "Er operasjonen aktiv? Ja / Nei",
    },
    "Item Costs": {
        "Item No": "Referanse til varen (Item No fra Product Master)",
        "Cost Type": "Type kostpris: Standard Cost, Last Direct Cost, Forecast Cost, Budget Cost",
        "Unit Cost": "Kostpris per enhet. Eksempel: 3000.00",
        "Currency": "Valuta. Eksempel: NOK, EUR",
        "Effective Date": "Dato kostprisen gjelder fra",
    },
    "BOM": {
        "Parent Item No": "Produktet som produseres (Item No)",
        "Component Item No": "Komponenten som forbrukes (Item No)",
        "Quantity Per": "Antall output-enheter per input-enhet. Eksempel: 400",
        "Unit of Measure": "Maleenhet. Eksempel: LM",
        "Scrap %": "Forventet materialsvinn i prosent. Eksempel: 5",
        "Co-Prod %": "Andel samprodukt (co-product). Eksempel: 6",
        "Co-Prod Item No": "Varenummer for samproduktet. Eksempel: JD16073-B",
        "Valid From": "Gyldig fra dato",
        "Valid To": "Gyldig til dato (tom = alltid)",
    },
    "Routing": {
        "Item No": "Produkt som produseres (Item No)",
        "Operation No": "Sekvensnummer. Eksempel: 10, 20, 30",
        "Operation Code": "Operasjon (ref. Operation Master)",
        "Work Center Code": "Arbeidssenter (ref. Work Centers)",
        "Setup Time Minutes": "Klargjoringstid i minutter. Eksempel: 15",
        "Run Time Minutes": "Produksjonstid per enhet i minutter. Eksempel: 0.15",
        "Batch Size": "Normal ordrestorrelse. Eksempel: 500",
        "Valid From": "Gyldig fra dato",
        "Valid To": "Gyldig til dato (tom = alltid)",
    },
    "By Product Rules": {
        "Parent Item No": "Produktet (ferdigvaren) som skaper biproduktet",
        "By Product Item No": "Biproduktet (Item No). Eksempel: BP001",
        "Expected Quantity": "Forventet mengde biprodukt per enhet ferdigvare",
        "Unit of Measure": "Maleenhet. Eksempel: KG",
        "Market Value": "Forventet markedspris per enhet. Eksempel: 1.50",
        "Allocation Method": "Reduce Main Product Cost, Separate Profit Center, Informational Only",
    },
    "Capacity Calendar": {
        "Work Center": "Arbeidssenter (ref. Work Centers)",
        "Date": "Dato",
        "Available Hours": "Tilgjengelige timer for dagen",
        "Planned Downtime": "Planlagte stopp i timer (vedlikehold, ferie, ombygging)",
    },
    "Production Scenario": {
        "Scenario Name": "Navn pa scenario. Eksempel: Normal Produksjon",
        "Product": "Produktet som simuleres (Item No)",
        "Planned Quantity": "Planlagt produksjonsmengde. Eksempel: 100000",
        "Start Date": "Startdato for scenario",
        "End Date": "Sluttdato for scenario",
    },
}

STATISKE_DROPDOWNS = [
    ("Product Master", "C", '"Raw Material,Semi Finished,Finished Good,By Product,Trading Item"'),
    ("Product Master", "F", '"Ja,Nei"'),
    ("Locations", "A", ""),
    ("Locations", "C", '"Factory,Warehouse,Distribution Center,Sales Office"'),
    ("Locations", "D", '"Ja,Nei"'),
    ("Work Centers", "A", ""),
    ("Work Centers", "I", '"Ja,Nei"'),
    ("Operation Master", "D", '"Minutes,Hours"'),
    ("Operation Master", "E", '"Ja,Nei"'),
    ("Item Costs", "B", '"Standard Cost,Last Direct Cost,Forecast Cost,Budget Cost"'),
    ("Item Costs", "D", '"NOK,EUR,USD,SEK,DKK"'),
    ("By Product Rules", "F", '"Reduce Main Product Cost,Separate Profit Center,Informational Only"'),
]

# ──────────────────────────────────────────────────────────────────────
#  Hjelpefunksjoner for datakonvertering
# ──────────────────────────────────────────────────────────────────────

def _s(value) -> str:
    """Trygg streng-konvertering (samme som i ExcelData)."""
    if pd.isna(value):
        return ""
    return str(value).strip()


def _f(value) -> float:
    """Trygg float-konvertering (samme som i ExcelData)."""
    if pd.isna(value):
        return 0.0
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0.0


def _d(value) -> Optional[str]:
    """Trygg dato-konvertering til ISO-streng."""
    if pd.isna(value):
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, str) and value.strip():
        return value.strip()
    return None


def _b(value) -> bool:
    """Trygg bool-konvertering (Ja/Nei → True/False)."""
    s = _s(value).lower()
    return s in ("ja", "yes", "true", "1", "y")


# ──────────────────────────────────────────────────────────────────────
#  Validering (uten å skrive til DB)
# ──────────────────────────────────────────────────────────────────────

def validate_excel(excel_path: str, db: Optional[DataRepo] = None) -> dict:
    """Valider Excel-fil uten å skrive til SQLite.
    
    Sjekker:
      - Alle påkrevde ark finnes
      - Alle påkrevde kolonner i hvert ark
      - Kryss-referanser (Item No, Work Center, etc.)
      - Numeriske felt har gyldige verdier
    
    Args:
        excel_path: Sti til Excel-filen (.xlsx)
        db: DataRepo-instans (for å sammenligne med eksisterende data)
    
    Returns:
        dict: {
            "valid": bool,
            "errors": list[str],
            "warnings": list[str],
            "stats": dict[str, int],
            "diff": {"nye_rader": int, "endrede_felt": int, "nye_tabeller": list[str]}
        }
    """
    result = {
        "valid": True,
        "errors": [],
        "warnings": [],
        "stats": {},
        "diff": {"nye_rader": 0, "endrede_felt": 0, "slettede_rader": 0},
    }

    # Les Excel
    try:
        xls = pd.ExcelFile(excel_path)
        sheet_names = xls.sheet_names
    except Exception as e:
        result["valid"] = False
        result["errors"].append(f"Kunne ikke lese Excel-fil: {e}")
        return result

    # Forventede ark
    expected_sheets = {
        "Product Master": ["Item No"],
        "Locations": ["Location Code"],
        "Work Centers": ["Work Center Code"],
        "Operation Master": ["Operation Code"],
        "Item Costs": ["Item No"],
        "BOM": ["Parent Item No", "Component Item No"],
        "Routing": ["Item No", "Operation No", "Work Center Code"],
        "By Product Rules": ["Parent Item No", "By Product Item No"],
        "Capacity Calendar": ["Work Center", "Date"],
        "Production Scenario": ["Scenario Name", "Product"],
    }

    # Samle opp items for kryss-referanser
    all_item_nos: set[str] = set()
    all_wc_codes: set[str] = set()
    all_op_codes: set[str] = set()
    all_loc_codes: set[str] = set()
    parent_bom_set: set[str] = set()
    component_bom_set: set[str] = set()

    for sheet_name, required_cols in expected_sheets.items():
        if sheet_name not in sheet_names:
            result["warnings"].append(f"Ark '{sheet_name}' mangler — vil bli hoppet over")
            continue

        df = xls.parse(sheet_name)
        df = df.dropna(how="all").reset_index(drop=True)

        if df.empty:
            result["warnings"].append(f"Ark '{sheet_name}' er tomt")
            continue

        # Sjekk påkrevde kolonner
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            result["valid"] = False
            result["errors"].append(f"Ark '{sheet_name}' mangler kolonner: {missing}")
            continue

        result["stats"][sheet_name] = len(df)

        # Samle opp referanser basert på arktype
        if sheet_name == "Product Master":
            for _, row in df.iterrows():
                item_no = _s(row.get("Item No", ""))
                if item_no:
                    all_item_nos.add(item_no)

        elif sheet_name == "Work Centers":
            for _, row in df.iterrows():
                code = _s(row.get("Work Center Code", ""))
                loc = _s(row.get("Location Code", ""))
                if code:
                    all_wc_codes.add(code)
                if loc:
                    all_loc_codes.add(loc)

        elif sheet_name == "Operation Master":
            for _, row in df.iterrows():
                code = _s(row.get("Operation Code", ""))
                wc = _s(row.get("Default Work Center", ""))
                if code:
                    all_op_codes.add(code)
                if wc:
                    all_wc_codes.add(wc)

        elif sheet_name == "Locations":
            for _, row in df.iterrows():
                code = _s(row.get("Location Code", ""))
                if code:
                    all_loc_codes.add(code)

        elif sheet_name == "BOM":
            for _, row in df.iterrows():
                parent = _s(row.get("Parent Item No", ""))
                comp = _s(row.get("Component Item No", ""))
                if parent:
                    parent_bom_set.add(parent)
                if comp:
                    component_bom_set.add(comp)

                # Valider numeriske verdier
                qty = _f(row.get("Quantity Per", 1))
                if qty <= 0:
                    result["warnings"].append(f"BOM: {parent}/{comp} har Quantity Per <= 0 ({qty})")

        elif sheet_name == "Routing":
            for _, row in df.iterrows():
                item = _s(row.get("Item No", ""))
                wc = _s(row.get("Work Center Code", ""))
                op = _s(row.get("Operation Code", ""))
                if item:
                    pass  # sjekkes mot all_item_nos senere
                if wc:
                    all_wc_codes.add(wc)
                if op:
                    all_op_codes.add(op)

                # Valider numeriske verdier
                run_time = _f(row.get("Run Time Minutes", 0))
                if run_time < 0:
                    result["warnings"].append(f"Routing {item}: Run Time Minutes er negativ ({run_time})")

        elif sheet_name == "By Product Rules":
            for _, row in df.iterrows():
                parent = _s(row.get("Parent Item No", ""))
                bp = _s(row.get("By Product Item No", ""))
                if parent:
                    parent_bom_set.add(parent)
                if bp:
                    component_bom_set.add(bp)

    # ── Kryss-referanser ───────────────────────────────────────
    # Disse sjekkene krever at Product Master er lest først
    if all_item_nos:
        # BOM: parent må finnes i Product Master
        for p in parent_bom_set:
            if p not in all_item_nos:
                result["errors"].append(f"BOM: Parent Item No '{p}' finnes ikke i Product Master")

        # BOM: component må finnes i Product Master
        for c in component_bom_set:
            if c not in all_item_nos:
                result["errors"].append(f"BOM: Component Item No '{c}' finnes ikke i Product Master")

        # Item Costs: item_no må finnes i Product Master
        if "Item Costs" in result["stats"]:
            df_ic = xls.parse("Item Costs")
            for _, row in df_ic.iterrows():
                item = _s(row.get("Item No", ""))
                if item and item not in all_item_nos:
                    result["errors"].append(f"Item Costs: '{item}' finnes ikke i Product Master")

    # Operation codes
    if all_op_codes and "Routing" in result["stats"]:
        df_rt = xls.parse("Routing")
        for _, row in df_rt.iterrows():
            op = _s(row.get("Operation Code", ""))
            if op and op not in all_op_codes:
                result["errors"].append(f"Routing: Operation Code '{op}' finnes ikke i Operation Master")

    # Work centers
    if all_wc_codes and "Routing" in result["stats"]:
        df_rt = xls.parse("Routing")
        for _, row in df_rt.iterrows():
            wc = _s(row.get("Work Center Code", ""))
            if wc and wc not in all_wc_codes:
                result["errors"].append(f"Routing: Work Center Code '{wc}' finnes ikke i Work Centers")

    # Location codes
    if all_loc_codes and "Work Centers" in result["stats"]:
        df_wc = xls.parse("Work Centers")
        for _, row in df_wc.iterrows():
            loc = _s(row.get("Location Code", ""))
            if loc and loc not in all_loc_codes:
                result["errors"].append(f"Work Centers: Location Code '{loc}' finnes ikke i Locations")

    # ── Sammenlign med SQLite (diff) ──────────────────────────
    if db is None:
        db = DataRepo()
        db.initialize()
        egen_db = True
    else:
        egen_db = False

    if not db.is_empty():
        # Sammenlign produkter
        existing_products = {r["item_no"] for r in db.conn.execute("SELECT item_no FROM products").fetchall()}
        excel_products = all_item_nos
        nye = excel_products - existing_products
        slettet = existing_products - excel_products
        result["diff"]["nye_rader"] = len(nye)
        result["diff"]["slettede_rader"] = len(slettet)
        result["diff"]["nye_tabeller"] = list(nye)[:10]  # vis max 10

    if egen_db:
        db.close()

    # ── Oppsummer ─────────────────────────────────────────────
    if result["errors"]:
        result["valid"] = False
    if result["warnings"]:
        result["valid"] = True  # warnings gjør ikke import ugyldig

    return result


# ──────────────────────────────────────────────────────────────────────
#  Hoved-importfunksjon
# ──────────────────────────────────────────────────────────────────────

def import_excel_to_sqlite(excel_path: str, db: Optional[DataRepo] = None,
                           excel_blob: Optional[bytes] = None,
                           comment: str = "") -> dict:
    """Importer alle ark fra Excel-fil til SQLite.

    Args:
        excel_path: Sti til Excel-filen (.xlsx)
        db: DataRepo-instans (opprettes automatisk hvis None)
        excel_blob: Innholdet av Excel-filen som bytes (lagres i uploaded_files)
        comment: Brukerens kommentar for denne importen

    Returns:
        dict med statistikk: {
            "tables_updated": {"products": 5, "item_costs": 8, ...},
            "total_changes": 12,
            "errors": []
        }
    """
    if db is None:
        db = DataRepo()
        db.initialize()
        egen_db = True
    else:
        egen_db = False

    stats = {"tables_updated": {}, "total_changes": 0, "errors": []}

    # Blob: les fra fil hvis ikke oppgitt
    if excel_blob is None:
        with open(excel_path, "rb") as _f:
            excel_blob = _f.read()

    try:
        xls = pd.ExcelFile(excel_path)
        sheet_names = xls.sheet_names
    except Exception as e:
        stats["errors"].append(f"Kunne ikke lese Excel-fil: {e}")
        return stats

    # Mapping: Excel-arknavn → (funksjon, required_columns)
    sheet_map = {
        "Product Master": (_import_products, ["Item No"]),
        "Locations": (_import_locations, ["Location Code"]),
        "Work Centers": (_import_work_centers, ["Work Center Code"]),
        "Operation Master": (_import_operations, ["Operation Code"]),
        "Item Costs": (_import_item_costs, ["Item No"]),
        "BOM": (_import_bom, ["Parent Item No", "Component Item No"]),
        "Routing": (_import_routing, ["Item No", "Operation No", "Work Center Code"]),
        "By Product Rules": (_import_byproduct_rules, ["Parent Item No", "By Product Item No"]),
        "Capacity Calendar": (_import_capacity, ["Work Center", "Date"]),
        "Production Scenario": (_import_scenarios, ["Scenario Name", "Product"]),
    }

    for sheet_name, (import_func, required_cols) in sheet_map.items():
        if sheet_name not in sheet_names:
            continue

        try:
            df = xls.parse(sheet_name)
            df = df.dropna(how="all").reset_index(drop=True)

            if df.empty:
                continue

            # Sjekk at nødvendige kolonner finnes
            missing = [c for c in required_cols if c not in df.columns]
            if missing:
                stats["errors"].append(
                    f"Ark '{sheet_name}': mangler kolonner: {missing}"
                )
                continue

            rows = df.to_dict("records")
            count = import_func(db, rows)
            stats["tables_updated"][sheet_name] = count
            stats["total_changes"] += count

        except Exception as e:
            stats["errors"].append(f"Feil ved import av ark '{sheet_name}': {e}")

    # Lagre blob og kommentar i uploaded_files
    try:
        _filename = os.path.basename(excel_path)
        _row_count = stats["total_changes"]
        db.save_upload(_filename, excel_blob, comment=comment, row_count=_row_count)
    except Exception:
        pass

    if egen_db:
        db.close()

    return stats


# ──────────────────────────────────────────────────────────────────────
#  Importere hvert ark
# ──────────────────────────────────────────────────────────────────────

def _import_products(db: DataRepo, rows: list[dict]) -> int:
    """Importer produkter. Returner antall endringer."""
    changes = 0
    product_list = []
    for row in rows:
        item_no = _s(row.get("Item No", ""))
        if not item_no:
            continue
        product_list.append({
            "item_no": item_no,
            "description": _s(row.get("Description", "")),
            "item_type": _s(row.get("Item Type", "")),
            "product_group": _s(row.get("Product Group", "")),
            "base_uom": _s(row.get("Base Unit of Measure", "")),
        })
    if product_list:
        db.upsert_products(product_list, source="import")
        # Telle endringer: vi kunne sjekket change_log, men returantall er et estimat
        changes = len(product_list)
    return changes


def _import_locations(db: DataRepo, rows: list[dict]) -> int:
    """Importer lokasjoner."""
    loc_list = []
    for row in rows:
        code = _s(row.get("Location Code", ""))
        if not code:
            continue
        loc_list.append({
            "code": code,
            "name": _s(row.get("Location Name", "")),
            "location_type": _s(row.get("Location Type", "")),
        })
    if loc_list:
        db.upsert_locations(loc_list, source="import")
    return len(loc_list)


def _import_work_centers(db: DataRepo, rows: list[dict]) -> int:
    """Importer arbeidssentre."""
    wc_list = []
    for row in rows:
        code = _s(row.get("Work Center Code", ""))
        if not code:
            continue
        wc_list.append({
            "code": code,
            "description": _s(row.get("Description", "")),
            "location_code": _s(row.get("Location Code", "")),
            "labor_cost_hour": _f(row.get("Labor Cost per Hour", 0)),
            "machine_cost_hour": _f(row.get("Machine Cost per Hour", 0)),
            "overhead_cost_hour": _f(row.get("Overhead Cost per Hour", 0)),
            "capacity_hours_day": _f(row.get("Capacity Hours per Day", 0)),
            "effective_capacity_pct": _f(row.get("Effective Capacity %", 100)),
        })
    if wc_list:
        db.upsert_work_centers(wc_list, source="import")
    return len(wc_list)


def _import_operations(db: DataRepo, rows: list[dict]) -> int:
    """Importer operasjoner."""
    op_list = []
    for row in rows:
        code = _s(row.get("Operation Code", ""))
        if not code:
            continue
        op_list.append({
            "code": code,
            "description": _s(row.get("Description", "")),
            "default_work_center": _s(row.get("Default Work Center", "")),
            "standard_unit": _s(row.get("Standard Unit", "Minutes")),
        })
    if op_list:
        db.upsert_operations(op_list, source="import")
    return len(op_list)


def _import_item_costs(db: DataRepo, rows: list[dict]) -> int:
    """Importer kostpriser."""
    cost_list = []
    for row in rows:
        item_no = _s(row.get("Item No", ""))
        if not item_no:
            continue
        cost_list.append({
            "item_no": item_no,
            "cost_type": _s(row.get("Cost Type", "Standard Cost")),
            "unit_cost": _f(row.get("Unit Cost", 0)),
            "currency": _s(row.get("Currency", "NOK")),
            "effective_date": _d(row.get("Effective Date")),
        })
    if cost_list:
        db.upsert_item_costs(cost_list, source="import")
    return len(cost_list)


def _import_bom(db: DataRepo, rows: list[dict]) -> int:
    """Importer BOM-linjer."""
    bom_list = []
    for row in rows:
        parent = _s(row.get("Parent Item No", ""))
        component = _s(row.get("Component Item No", ""))
        if not parent or not component:
            continue
        bom_list.append({
            "parent_item_no": parent,
            "component_item_no": component,
            "quantity_per": _f(row.get("Quantity Per", 1)),
            "uom": _s(row.get("Unit of Measure", "")),
            "scrap_pct": _f(row.get("Scrap %", 0)),
            "co_product_pct": _f(row.get("Co-Prod %", 0)),
            "co_product_item_no": _s(row.get("Co-Prod Item No", "")),
        })
    if bom_list:
        db.upsert_bom_lines(bom_list, source="import")
    return len(bom_list)


def _import_routing(db: DataRepo, rows: list[dict]) -> int:
    """Importer routing-linjer."""
    rt_list = []
    for row in rows:
        item_no = _s(row.get("Item No", ""))
        if not item_no:
            continue
        rt_list.append({
            "item_no": item_no,
            "operation_no": int(_f(row.get("Operation No", 0))),
            "operation_code": _s(row.get("Operation Code", "")),
            "work_center_code": _s(row.get("Work Center Code", "")),
            "setup_time_minutes": _f(row.get("Setup Time Minutes", 0)),
            "run_time_minutes": _f(row.get("Run Time Minutes", 0)),
            "batch_size": _f(row.get("Batch Size", 1)),
        })
    if rt_list:
        db.upsert_routing_lines(rt_list, source="import")
    return len(rt_list)


def _import_byproduct_rules(db: DataRepo, rows: list[dict]) -> int:
    """Importer biproduktregler."""
    bp_list = []
    for row in rows:
        parent = _s(row.get("Parent Item No", ""))
        byprod = _s(row.get("By Product Item No", ""))
        if not parent or not byprod:
            continue
        bp_list.append({
            "parent_item_no": parent,
            "by_product_item_no": byprod,
            "expected_quantity": _f(row.get("Expected Quantity", 0)),
            "uom": _s(row.get("Unit of Measure", "")),
            "market_value": _f(row.get("Market Value", 0)),
            "allocation_method": _s(row.get("Allocation Method", "Reduce Main Product Cost")),
        })
    if bp_list:
        db.upsert_byproduct_rules(bp_list, source="import")
    return len(bp_list)


def _import_capacity(db: DataRepo, rows: list[dict]) -> int:
    """Importer kapasitetskalender."""
    cap_list = []
    for row in rows:
        wc = _s(row.get("Work Center", ""))
        d = _d(row.get("Date"))
        if not wc or not d:
            continue
        cap_list.append({
            "work_center": wc,
            "date": d,
            "available_hours": _f(row.get("Available Hours", 0)),
            "planned_downtime": _f(row.get("Planned Downtime", 0)),
        })
    if cap_list:
        db.upsert_capacity_days(cap_list, source="import")
    return len(cap_list)


def _import_scenarios(db: DataRepo, rows: list[dict]) -> int:
    """Importer produksjonsscenarioer."""
    sc_list = []
    for row in rows:
        name = _s(row.get("Scenario Name", ""))
        product = _s(row.get("Product", ""))
        if not name or not product:
            continue
        sc_list.append({
            "scenario_name": name,
            "product": product,
            "planned_quantity": _f(row.get("Planned Quantity", 0)),
        })
    if sc_list:
        db.upsert_scenarios(sc_list, source="import")
    return len(sc_list)


# ──────────────────────────────────────────────────────────────────────
#  Eksport: SQLite → Excel
# ──────────────────────────────────────────────────────────────────────

def export_sqlite_to_excel(output_path: str, db: Optional[DataRepo] = None) -> None:
    """Eksporter all data fra SQLite til en Excel-fil.
    
    Genererer de samme 10 arkene som den originale Excel-malen.
    Inkluderer et 11. ark: 'Endringslogg'.

    Args:
        output_path: Sti til output Excel-fil (.xlsx)
        db: DataRepo-instans (opprettes automatisk hvis None)
    """

    if db is None:
        db = DataRepo()
        db.initialize()
        egen_db = True
    else:
        egen_db = False

    data = db.export_all_data()
    wb = openpyxl.Workbook()

    # Stiler
    header_font = Font(name='Calibri', bold=True, color='FFFFFF', size=11)
    header_fill = PatternFill(start_color='14532D', end_color='14532D', fill_type='solid')
    data_font = Font(name='Calibri', size=10)
    thin_border = Border(
        left=Side(style='thin', color='48BB78'),
        right=Side(style='thin', color='48BB78'),
        top=Side(style='thin', color='48BB78'),
        bottom=Side(style='thin', color='48BB78'),
    )

    def _write_sheet(ws, title: str, rows: list, field_names: list[str], db_keys: Optional[list[str]] = None):
        """Skriv data til et Excel-ark med header, kolonnebeskrivelser og dropdowns.

        Args:
            ws: openpyxl worksheet
            title: Ark-tittel (Excel-fane)
            rows: Liste med dicts/rad-objekter
            field_names: Kolonneoverskrifter (Excel-visning)
            db_keys: Database-nøkler for data-aksess. Hvis None, konverteres field_names
                     til snake_case (f.eks. "Item No" → "item_no").
        """
        ws.title = title
        if not rows:
            ws.cell(row=1, column=1, value="(Ingen data)").font = data_font
            return

        # Hvis db_keys ikke er oppgitt, konverter field_names til snake_case
        if db_keys is None:
            db_keys = [name.lower().replace(" ", "_") for name in field_names]

        # Headere med kolonnebeskrivelser
        kol_desc = KOLONNER.get(title, {})
        for col_idx, name in enumerate(field_names, 1):
            cell = ws.cell(row=1, column=col_idx, value=name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = thin_border
            if name in kol_desc:
                cell.comment = XLComment(kol_desc[name], "System", width=300, height=100)

        # Data — konverter sqlite3.Row til dict for sikker aksess
        for row_idx, row in enumerate(rows, 2):
            if hasattr(row, 'keys'):
                row = dict(row)
            for col_idx, key in enumerate(db_keys, 1):
                value = row.get(key, "")
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.font = data_font
                cell.border = thin_border

        # Auto-width
        for col_idx in range(1, len(field_names) + 1):
            max_len = len(str(field_names[col_idx - 1]))
            for row_idx in range(2, len(rows) + 2):
                val = ws.cell(row=row_idx, column=col_idx).value
                if val is not None:
                    max_len = max(max_len, min(len(str(val)), 40))
            ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 2

        # Lås header-raden (freeze panes)
        ws.freeze_panes = "A2"

        # Legg til statiske dropdowns
        _legg_til_dropdowns(ws, title)

    def _legg_til_dropdowns(ws, ark_navn: str):
        """Legg til statiske dropdowns for et ark (samme som oppdater_mal.py)."""
        for sheet, col, liste in STATISKE_DROPDOWNS:
            if sheet != ark_navn:
                continue
            if not liste:
                continue
            dv = DataValidation(
                type="list",
                formula1=liste,
                allow_blank=True,
                showErrorMessage=True,
                errorTitle="Ugyldig verdi",
                error=f"Verdien må være en av: {liste.replace(chr(34), '')}",
            )
            ws.add_data_validation(dv)
            dv.add(f"{col}2:{col}1048576")

    # ── Ark 1: Product Master ────────────────────────────────
    ws1 = wb.active
    _write_sheet(ws1, "Product Master", data.get("products", []),
                 ["Item No", "Description", "Item Type", "Product Group", "Base Unit of Measure"],
                 db_keys=["item_no", "description", "item_type", "product_group", "base_uom"])

    # ── Ark 2: Locations ──────────────────────────────────────
    ws2 = wb.create_sheet()
    _write_sheet(ws2, "Locations", data.get("locations", []),
                 ["Location Code", "Location Name", "Location Type"],
                 db_keys=["code", "name", "location_type"])

    # ── Ark 3: Work Centers ───────────────────────────────────
    ws3 = wb.create_sheet()
    _write_sheet(ws3, "Work Centers", data.get("work_centers", []),
                 ["Work Center Code", "Description", "Location Code",
                  "Labor Cost per Hour", "Machine Cost per Hour", "Overhead Cost per Hour",
                  "Capacity Hours per Day", "Effective Capacity %"],
                 db_keys=["code", "description", "location_code",
                          "labor_cost_hour", "machine_cost_hour", "overhead_cost_hour",
                          "capacity_hours_day", "effective_capacity_pct"])

    # ── Ark 4: Operation Master ──────────────────────────────
    ws4 = wb.create_sheet()
    _write_sheet(ws4, "Operation Master", data.get("operations", []),
                 ["Operation Code", "Description", "Default Work Center", "Standard Unit"],
                 db_keys=["code", "description", "default_work_center", "standard_unit"])

    # ── Ark 5: Item Costs ────────────────────────────────────
    ws5 = wb.create_sheet()
    _write_sheet(ws5, "Item Costs", data.get("item_costs", []),
                 ["Item No", "Cost Type", "Unit Cost", "Currency", "Effective Date"],
                 db_keys=["item_no", "cost_type", "unit_cost", "currency", "effective_date"])

    # ── Ark 6: BOM ───────────────────────────────────────────
    ws6 = wb.create_sheet()
    _write_sheet(ws6, "BOM", data.get("bom_lines", []),
                 ["Parent Item No", "Component Item No", "Quantity Per", "Unit of Measure",
                  "Scrap %", "Co-Prod %", "Co-Prod Item No"],
                 db_keys=["parent_item_no", "component_item_no", "quantity_per", "uom",
                          "scrap_pct", "co_product_pct", "co_product_item_no"])

    # ── Ark 7: Routing ─────────────────────────────────────────
    ws7 = wb.create_sheet()
    _write_sheet(ws7, "Routing", data.get("routing_lines", []),
                 ["Item No", "Operation No", "Operation Code", "Work Center Code",
                  "Setup Time Minutes", "Run Time Minutes", "Batch Size"],
                 db_keys=["item_no", "operation_no", "operation_code", "work_center_code",
                          "setup_time_minutes", "run_time_minutes", "batch_size"])

    # ── Ark 8: By Product Rules ──────────────────────────────
    ws8 = wb.create_sheet()
    _write_sheet(ws8, "By Product Rules", data.get("byproduct_rules", []),
                 ["Parent Item No", "By Product Item No", "Expected Quantity",
                  "Unit of Measure", "Market Value", "Allocation Method"],
                 db_keys=["parent_item_no", "by_product_item_no", "expected_quantity",
                          "uom", "market_value", "allocation_method"])

    # ── Ark 9: Capacity Calendar ─────────────────────────────
    ws9 = wb.create_sheet()
    _write_sheet(ws9, "Capacity Calendar", data.get("capacity_days", []),
                 ["Work Center", "Date", "Available Hours", "Planned Downtime"],
                 db_keys=["work_center", "date", "available_hours", "planned_downtime"])

    # ── Ark 10: Production Scenario ──────────────────────────
    ws10 = wb.create_sheet()
    _write_sheet(ws10, "Production Scenario", data.get("production_scenarios", []),
                 ["Scenario Name", "Product", "Planned Quantity"],
                 db_keys=["scenario_name", "product", "planned_quantity"])

    # ── Ark 11: Endringslogg ─────────────────────────────────
    ws11 = wb.create_sheet()
    change_log = db.get_changes(limit=1000)
    if change_log:
        _write_sheet(ws11, "Endringslogg", change_log,
                     ["ID", "Tidspunkt", "Bruker", "Kilde", "Tabell", "Nokkel",
                      "Felt", "Gammel verdi", "Ny verdi"],
                     db_keys=["id", "timestamp", "user", "source", "table_name",
                              "record_key", "field_name", "old_value", "new_value"])
    else:
        ws11.title = "Endringslogg"
        ws11.cell(row=1, column=1, value="(Ingen endringer logget)").font = data_font

    wb.save(output_path)

    if egen_db:
        db.close()


# ──────────────────────────────────────────────────────────────────────
#  Kommandolinje
# ──────────────────────────────────────────────────────────────────────

def main():
    """Kjøres frittstående: python excel_bridge.py --import <excel> [--export <output>]"""
    import argparse

    parser = argparse.ArgumentParser(description="Import/eksport mellom Excel og SQLite")
    parser.add_argument("--import", dest="import_path", type=str,
                        help="Importer Excel-fil til SQLite")
    parser.add_argument("--export", type=str,
                        help="Eksporter SQLite til Excel-fil")
    parser.add_argument("--stats", action="store_true",
                        help="Vis status for SQLite-databasen")

    args = parser.parse_args()

    db = DataRepo()
    db.initialize()

    if args.import_path:
        if not os.path.exists(args.import_path):
            print(f"[FEIL] Finner ikke filen: {args.import_path}")
            sys.exit(1)

        print(f"[*] Importerer: {args.import_path}")
        stats = import_excel_to_sqlite(args.import_path, db)

        print(f"[OK] Ferdig!")
        for sheet, count in stats["tables_updated"].items():
            print(f"   {sheet}: {count} rader")
        if stats["errors"]:
            print(f"[ADV] Feil ({len(stats['errors'])}):")
            for e in stats["errors"]:
                print(f"   - {e}")

    if args.export:
        print(f"[*] Eksporterer til: {args.export}")
        export_sqlite_to_excel(args.export, db)
        print(f"[OK] Ferdig!")

    if args.stats:
        print(f"[DB] Database: {db.db_path}")
        for table, count in db.stats.items():
            print(f"   {table}: {count} rader")


if __name__ == "__main__":
    main()