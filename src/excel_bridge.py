#!/usr/bin/env python3
"""
excel_bridge.py - Import/eksport mellom Excel-filer og SQLite-databasen.

Brukes til:
  1. Importere data fra Excel til SQLite (med endringslogg)
  2. Eksportere data fra SQLite til Excel (for nedlasting / SharePoint)

Alle importfunksjonene sammenligner felt-for-felt med eksisterende data i SQLite
og loggfører kun faktiske endringer.

Excel-eksport inkluderer ACTION (col A) og Rad ID (col B) for hvert ark.
ACTION-støttede verdier: CREATE, UPDATE, DELETE (eller tom for å inferere).
"""

import json
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
        "Is Transport": "Transportvare? 1=Ja (frakt mellom høvlerier), 0=Nei",
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
    "Transport Ruter": {
        "From Loc": "Fra-lokasjon. Eksempel: KOD",
        "To Loc": "Til-lokasjon. Eksempel: KV",
        "Cost Per M3": "Fraktpris per M3 på denne ruten (eneste beregningsfelt)",
        "Distance Km": "Distanse i kilometer (informasjon — påvirker ikke kost)",
        "Hours": "Kjøretid i timer (informasjon — påvirker ikke kost)",
    },
}

STATISKE_DROPDOWNS = [
    ("Product Master", "C", '"Raw Material,Semi Finished,Finished Good,By Product,Trading Item"'),
    # "F" peker på Active sin posisjon i ark-uten-action, men Active eksporteres ikke.
    # Eksport-flytter F→H, som er "Is Transport"-kolonnen i Product Master-eksporten.
    ("Product Master", "F", '"1,0"'),
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

# ACTION-kolonne: Kolonne A er alltid ACTION, kolonne B er alltid Rad ID
ACTION_VALUES = '"CREATE,UPDATE,DELETE"'

# Mapping: Excel-arknavn → navn på DB-tabell (for sletting)
SHEET_TO_TABLE = {
    "Product Master": "products",
    "Locations": "locations",
    "Work Centers": "work_centers",
    "Operation Master": "operations",
    "Item Costs": "item_costs",
    "BOM": "bom_lines",
    "Routing": "routing_lines",
    "By Product Rules": "byproduct_rules",
    "Capacity Calendar": "capacity_days",
    "Production Scenario": "production_scenarios",
}

# Mapping: Excel-arknavn → navn på delete-metode i DataRepo
SHEET_TO_DELETE_METHOD = {
    "Product Master": "delete_product",
    "Locations": "delete_location",
    "Work Centers": "delete_work_center",
    "Operation Master": "delete_operation",
    "Item Costs": "delete_item_cost",
    "BOM": "delete_bom_line",
    "Routing": "delete_routing_line",
    "By Product Rules": "delete_byproduct_rule",
    "Capacity Calendar": "delete_capacity_day",
    "Production Scenario": "delete_scenario",
}

# Gyldige ACTION-verdier
VALID_ACTIONS = {"CREATE", "UPDATE", "DELETE", ""}

# Mapping: Excel-arknavn → (tabellnavn, nøkkelkolonne, sql_where) for DELETE uten Rad ID
SHEET_DELETE_KEY = {
    "Product Master": ("products", "Item No", "item_no = ?"),
    "Locations": ("locations", "Location Code", "code = ?"),
    "Work Centers": ("work_centers", "Work Center Code", "code = ?"),
    "Operation Master": ("operations", "Operation Code", "code = ?"),
    "Item Costs": ("item_costs", "Item No + Cost Type", "item_no = ? AND cost_type = ?"),
    "BOM": ("bom_lines", "Parent Item No + Component Item No", "parent_item_no = ? AND component_item_no = ?"),
    "Routing": ("routing_lines", "Item No + Operation No + Work Center", "item_no = ? AND operation_no = ? AND work_center_code = ?"),
    "By Product Rules": ("byproduct_rules", "Parent + By Product", "parent_item_no = ? AND by_product_item_no = ?"),
    "Capacity Calendar": ("capacity_days", "Work Center + Date", "work_center = ? AND date = ?"),
    "Production Scenario": ("production_scenarios", "Scenario + Product", "scenario_name = ? AND product = ?"),
}


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


def _i(value) -> Optional[int]:
    """Trygg integer-konvertering for Rad ID."""
    if pd.isna(value):
        return None
    try:
        return int(float(value))
    except (ValueError, TypeError):
        return None


def _get_action(row: dict) -> str:
    """Hent ACTION-verdi fra en rad. Returner uppercase eller tom string."""
    action = row.get("ACTION", "")
    if pd.isna(action):
        return ""
    action = str(action).strip().upper()
    if action in VALID_ACTIONS:
        return action
    return ""


# ──────────────────────────────────────────────────────────────────────
#  Validering (uten å skrive til DB)
# ──────────────────────────────────────────────────────────────────────

def validate_excel(excel_path: str, db: Optional[DataRepo] = None) -> dict:
    """Valider Excel-fil uten å skrive til SQLite.
    
    Sjekker:
      - Alle påkrevde ark finnes
      - Alle påkrevde kolonner i hvert ark
      - ACTION-kolonnen har gyldige verdier
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

        # Sjekk at ACTION og Rad ID kolonner finnes (advarsel hvis ikke)
        # Merk: ACTION og ID er kolonne A og B, men vi sjekker på navn
        if "ACTION" not in df.columns and "Rad ID" not in df.columns:
            # Sjekk om første/tomme kolonner fungerer som ACTION/ID
            # Dette er en myk sjekk — vi gir bare en advarsel
            pass

        # Sjekk påkrevde kolonner (etter ACTION/ID)
        # Først må vi finne datakolonnene (de som ikke er ACTION eller Rad ID)
        data_cols = [c for c in df.columns if c not in ("ACTION", "Rad ID")]
        missing = [c for c in required_cols if c not in data_cols]
        if missing:
            result["valid"] = False
            result["errors"].append(f"Ark '{sheet_name}' mangler kolonner: {missing}")
            continue

        # Valider ACTION-verdier
        if "ACTION" in df.columns:
            action_col = df["ACTION"]
            for i, val in action_col.items():
                if pd.isna(val) or str(val).strip() == "":
                    continue
                action_str = str(val).strip().upper()
                if action_str not in VALID_ACTIONS:
                    result["errors"].append(
                        f"Ark '{sheet_name}', rad {i+2}: Ugyldig ACTION-verdi '{val}'. "
                        f"Tillatte verdier: CREATE, UPDATE, DELETE"
                    )

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

    Håndterer ACTION-kolonnen:
      - CREATE: Ny rad (INSERT)
      - UPDATE: Oppdater eksisterende rad (match på Rad ID)
      - DELETE: Slett rad (via Rad ID)
      - Tom: Inferer (Rad ID finnes → UPDATE, ellers CREATE)

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
    "Transport Ruter": (_import_transport_ruter, ["From Loc", "To Loc"]),
    }

    for sheet_name, (import_func, required_cols) in sheet_map.items():
        if sheet_name not in sheet_names:
            continue

        try:
            df = xls.parse(sheet_name)
            df = df.dropna(how="all").reset_index(drop=True)

            if df.empty:
                continue

            # Sjekk at nødvendige kolonner finnes (etter å ha fjernet ACTION/Rad ID)
            data_cols = [c for c in df.columns if c not in ("ACTION", "Rad ID")]
            missing = [c for c in required_cols if c not in data_cols]
            if missing:
                stats["errors"].append(
                    f"Ark '{sheet_name}': mangler kolonner: {missing}"
                )
                continue

            rows = df.to_dict("records")
            count = import_func(db, rows, sheet_name)
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

    # NB! sync_transport_varer() er LEGACY og er bevisst IKKE kalt her.
    # Transport vises nå KUN i simuleringen via expand_*_with_transport()
    # i kostberegning.py — datamodellen (semi-finished) muteres aldri.
    # Se sync_transport_varer() for legacy-guard.

    if egen_db:
        db.close()

    return stats


# ──────────────────────────────────────────────────────────────────────
#  Transportvare-synkronisering (fler-høvleri-produksjon)
# ──────────────────────────────────────────────────────────────────────

def _aktive_factory_locations(db: DataRepo) -> list[str]:
    """Hent alle aktive fabrikk-lokasjoner (location_type = 'Factory')."""
    rows = db.conn.execute(
        "SELECT code FROM locations WHERE location_type = 'Factory' ORDER BY code"
    ).fetchall()
    return [r["code"] for r in rows]


def _produksjons_locations(db: DataRepo, item_no: str) -> set[str]:
    """Finn hvilke lokasjoner som produserer en vare (via routing → work_centers)."""
    rows = db.conn.execute(
        """SELECT DISTINCT wc.location_code
           FROM routing_lines rl
           JOIN work_centers wc ON rl.work_center_code = wc.code
           WHERE rl.item_no = ?""",
        (item_no,),
    ).fetchall()
    return {r["location_code"] for r in rows}


def _ensure_workcenter(db: DataRepo, code: str, description: str, location: str,
                       labor: float, machine: float, overhead: float,
                       source: str = "transport_sync"):
    """Opprett et arbeidssenter hvis det ikke finnes."""
    existing = db.conn.execute(
        "SELECT id FROM work_centers WHERE code = ?", (code,)
    ).fetchone()
    if not existing:
        db.upsert_work_centers([{
            "code": code, "description": description, "location_code": location,
            "labor_cost_hour": labor, "machine_cost_hour": machine,
            "overhead_cost_hour": overhead, "capacity_hours_day": 24,
            "effective_capacity_pct": 100,
        }], source=source)


def _ensure_operation(db: DataRepo, code: str, description: str,
                      default_wc: str, source: str = "transport_sync"):
    """Opprett TRANSPORT-operasjonen hvis den ikke finnes."""
    existing = db.conn.execute(
        "SELECT id FROM operations WHERE code = ?", (code,)
    ).fetchone()
    if not existing:
        db.upsert_operations([{
            "code": code, "description": description,
            "default_work_center": default_wc, "standard_unit": "Minutes",
        }], source=source)


def _get_or_create_product(db: DataRepo, item_no: str, description: str,
                           item_type: str, product_group: str, base_uom: str,
                           source: str = "transport_sync") -> Optional[tuple[int, bool]]:
    """Hent eller opprett et produkt.

    Returns:
        (id, var_ny) — id til produktet, og True hvis det ble nyopprettet,
        ellers None hvis opprettelsen mislyktes.
    """
    existing = db.conn.execute(
        "SELECT id FROM products WHERE item_no = ?", (item_no,)
    ).fetchone()
    if existing:
        return (existing["id"], False)
    db.upsert_products([{
        "item_no": item_no, "description": description,
        "item_type": item_type, "product_group": product_group,
        "base_uom": base_uom,
    }], source=source)
    row = db.conn.execute("SELECT id FROM products WHERE item_no = ?", (item_no,)).fetchone()
    if row is None:
        return None
    return (row["id"], True)


def _delete_product_by_no(db: DataRepo, item_no: str, source: str = "transport_sync"):
    """Slett et produkt basert på varenummer (slett kun hvis det er semi-finished/transport-generert)."""
    row = db.conn.execute("SELECT id FROM products WHERE item_no = ?", (item_no,)).fetchone()
    if row:
        db.delete_product(row["id"], source=source)


def _delete_bom_by_parent(db: DataRepo, parent_item_no: str, component_item_no: str,
                          source: str = "transport_sync"):
    """Slett en BOM-linje basert på parent+component."""
    row = db.conn.execute(
        "SELECT id FROM bom_lines WHERE parent_item_no = ? AND component_item_no = ?",
        (parent_item_no, component_item_no),
    ).fetchone()
    if row:
        db.delete_bom_line(row["id"], source=source)


def _delete_routing_by_item(db: DataRepo, item_no: str, operation_no: int, wc: str,
                            source: str = "transport_sync"):
    """Slett en routing-linje basert på item+op+wc."""
    row = db.conn.execute(
        "SELECT id FROM routing_lines WHERE item_no = ? AND operation_no = ? AND work_center_code = ?",
        (item_no, operation_no, wc),
    ).fetchone()
    if row:
        db.delete_routing_line(row["id"], source=source)


def sync_transport_varer(db: DataRepo, source: str = "import",
                         raise_if_called: bool = True) -> dict:
    """Synkroniser transportvarer — generer/fjern semi-finished varianter og transport-routing.

    ⚠️ LEGACY — IKKE kalt fra import eller app lenger.
    Transport vises nå KUN i simuleringen via expand_*_with_transport()
    i kostberegning.py — datamodellen (semi-finished) muteres aldri.

    Denne funksjonen beholdes urørt for fremtidig Business Central-integrasjon.
    For å aktivere den på nytt: bruk raise_if_called=False og koble den inn
    der det er ønskelig (f.eks. i import_excel_to_sqlite).

    Args:
        db: DataRepo-instans
        source: Kilde for endringslogg
        raise_if_called: True (default) kaster RuntimeError hvis funksjonen
            kalles — beskytter mot uønsket gjenoppretting av semi-finished.

    Returns:
        dict med statistikk: {"opprettet": int, "slettet": int, "produkter": [...]}

    Raises:
        RuntimeError: hvis raise_if_called=True (default) — funksjonen er legacy.
    """
    if raise_if_called:
        raise RuntimeError(
            "sync_transport_varer() er LEGACY og inaktiv. "
            "Transport vises nå kun i simuleringen via expand_*_with_transport(). "
            "For å bruke legacy-logikken: kall med raise_if_called=False."
        )

    stats = {"opprettet": 0, "slettet": 0, "produkter": []}

    # Sørg for at tabellen finnes
    db.conn.execute(
        """CREATE TABLE IF NOT EXISTS transport_flagg (
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               item_no TEXT NOT NULL UNIQUE,
               is_transport INTEGER NOT NULL DEFAULT 0,
               updated_at TEXT NOT NULL DEFAULT (datetime('now'))
        )"""
    )
    db.conn.commit()

    # Les alle flagg-varer
    flagged = db.conn.execute(
        "SELECT item_no, is_transport FROM transport_flagg"
    ).fetchall()

    for row in flagged:
        item_no = row["item_no"]
        is_transport = bool(row["is_transport"])

        # Hent produktdata for hovedvaren
        prod = db.conn.execute(
            "SELECT * FROM products WHERE item_no = ?", (item_no,)
        ).fetchone()
        if not prod:
            continue

        if is_transport:
            # ── Generer semi-finished for fabrikker som IKKE produserer varen ──
            # NB! Hovedproduktet (item_no) forblir fullstendig URØRT.
            # Semi-finished refererer TIL hovedproduktet og legger kun på
            # TRANSPORT-routing — original BOM/routing på hovedproduktet endres aldri.
            #
            # Regler:
            #   - Fabrikklokasjoner som allerede har routing for varen produserer
            #     den selv → ingen semi-finished, ingen transport.
            #   - Øvrige fabrikklokasjoner får semi-finished med TRANSPORT fra
            #     FØRSTE produksjonslokasjon (sortert alfabetisk).
            created_here = 0
            factory_locs = _aktive_factory_locations(db)
            prod_locs = _produksjons_locations(db, item_no)
            _from_loc = sorted(prod_locs)[0] if prod_locs else (factory_locs[0] if factory_locs else "KOD")

            _ensure_workcenter(db, "TRANSPORT", "Frakt mellom høvlerier", _from_loc,
                               300.0, 300.0, 200.0, source=source)
            _ensure_operation(db, "TRANSPORT", "Frakt mellom høvlerier", "TRANSPORT", source=source)

            for loc in factory_locs:
                if loc in prod_locs:
                    continue  # Lokasjonen produserer allerede — trenger ikke transport

                semi_no = f"{item_no}-{loc}"
                semi_res = _get_or_create_product(
                    db, semi_no,
                    f"{prod['description']} ({loc})",
                    "Semi Finished", prod["product_group"], prod["base_uom"],
                    source=source,
                )
                if semi_res is None:
                    continue
                semi_id, var_ny = semi_res
                if var_ny:
                    created_here += 1

                # BOM: semi-finished → hovedprodukt (Qty Per = 1)
                # Materialkost rulles dynamisk opp fra hovedproduktets netto produksjonskost
                db.upsert_bom_lines([{
                    "parent_item_no": semi_no,
                    "component_item_no": item_no,
                    "quantity_per": 1.0,
                    "uom": prod["base_uom"],
                    "scrap_pct": 0.0,
                    "co_product_pct": 0.0,
                    "co_product_item_no": "",
                }], source=source)

                # TRANSPORT-routing på semi-finished (ikke hovedprodukt)
                # NB: transport_ruter har nå kun cost_per_m3/distance_km/hours.
                # Legacy-estimat: hours → run_time minutter, standard setup/batch.
                rute = db.conn.execute(
                    "SELECT * FROM transport_ruter WHERE from_loc = ? AND to_loc = ?",
                    (_from_loc, loc)
                ).fetchone()
                if rute:
                    _hours = float(rute["hours"] or 0) if "hours" in rute.keys() else 0
                    _run = _hours * 60.0 if _hours > 0 else 45.0
                    _ensure_routing_entry(db, semi_no, "TRANSPORT", "TRANSPORT",
                                          setup=30.0, run=_run, batch=prod_batch(db, item_no),
                                          source=source)
                else:
                    _ensure_routing_entry(db, semi_no, "TRANSPORT", "TRANSPORT",
                                          setup=30.0, run=45.0, batch=prod_batch(db, item_no),
                                          source=source)

            stats["opprettet"] += created_here
            if created_here > 0:
                stats["produkter"].append(f"{item_no}: transport flagg satt")
        else:
            # ── Fjern alle genererte semi-finished for denne varen ──
            # Hovedproduktet (item_no) er aldri blitt rørt — kun semi-finished
            # og deres BOM/routing slettes her.
            deleted_here = 0
            # Finn alle semi-finished for varen
            semi_items = db.conn.execute(
                "SELECT item_no FROM products WHERE item_no LIKE ?",
                (f"{item_no}-%",)
            ).fetchall()

            for semi in semi_items:
                semi_no = semi["item_no"]
                # Slett BOM-linjer der semi-finished er parent
                rows = db.conn.execute(
                    "SELECT id FROM bom_lines WHERE parent_item_no = ?", (semi_no,)
                ).fetchall()
                for r in rows:
                    db.delete_bom_line(r["id"], source=source)
                # Slett routing-linjer for semi-finished
                rt_rows = db.conn.execute(
                    "SELECT id FROM routing_lines WHERE item_no = ?", (semi_no,)
                ).fetchall()
                for r in rt_rows:
                    db.delete_routing_line(r["id"], source=source)
                # Slett selve produktet
                _delete_product_by_no(db, semi_no, source=source)
                deleted_here += 1

            stats["slettet"] += deleted_here
            if deleted_here > 0:
                stats["produkter"].append(f"{item_no}: transport flagg fjernet")

    return stats


def _ensure_routing_entry(db: DataRepo, item_no: str, op_code: str, wc: str,
                          setup: float, run: float, batch: float,
                          source: str = "transport_sync"):
    """Legg til routing-linje hvis den ikke finnes."""
    existing = db.conn.execute(
        "SELECT id FROM routing_lines WHERE item_no = ? AND operation_code = ? AND work_center_code = ?",
        (item_no, op_code, wc)
    ).fetchone()
    if not existing:
        # Finn neste ledige operation_no
        max_op = db.conn.execute(
            "SELECT MAX(operation_no) as mx FROM routing_lines WHERE item_no = ?",
            (item_no,)
        ).fetchone()
        next_op = int(max_op["mx"] or 0) + 10
        db.upsert_routing_lines([{
            "item_no": item_no,
            "operation_no": next_op,
            "operation_code": op_code,
            "work_center_code": wc,
            "setup_time_minutes": setup,
            "run_time_minutes": run,
            "batch_size": batch,
        }], source=source)


def prod_batch(db: DataRepo, item_no: str) -> float:
    """Hent typisk batch-størrelse for et produkt."""
    row = db.conn.execute(
        "SELECT batch_size FROM routing_lines WHERE item_no = ? ORDER BY operation_no LIMIT 1",
        (item_no,)
    ).fetchone()
    return row["batch_size"] if row else 500.0


# ──────────────────────────────────────────────────────────────────────
#  Hjelper: prosesser en rad med ACTION
# ──────────────────────────────────────────────────────────────────────

def _deletions_from_action(rows: list[dict], sheet_name: str, db: DataRepo) -> int:
    """Prosesser og utfør DELETE-ACTIONer. Returner antall slettinger."""
    table_name = SHEET_TO_TABLE.get(sheet_name)
    if not table_name:
        return 0

    delete_method_name = SHEET_TO_DELETE_METHOD.get(sheet_name)
    delete_key_info = SHEET_DELETE_KEY.get(sheet_name)
    delete_count = 0
    for row in rows:
        action = _get_action(row)
        if action != "DELETE":
            continue

        row_id = _i(row.get("Rad ID"))

        # Hvis Rad ID mangler, slå opp id via naturlig nøkkel
        if row_id is None and delete_key_info:
            _tbl, _, _where = delete_key_info
            # Bygg WHERE-parametere basert på sheet-type
            if sheet_name == "Product Master":
                params = (_s(row.get("Item No", "")),)
            elif sheet_name in ("Locations", "Work Centers", "Operation Master"):
                col = "Location Code" if sheet_name == "Locations" else ("Work Center Code" if sheet_name == "Work Centers" else "Operation Code")
                params = (_s(row.get(col, "")),)
            elif sheet_name == "Item Costs":
                params = (_s(row.get("Item No", "")), _s(row.get("Cost Type", "Standard Cost")))
            elif sheet_name == "BOM":
                params = (_s(row.get("Parent Item No", "")), _s(row.get("Component Item No", "")))
            elif sheet_name == "Routing":
                params = (_s(row.get("Item No", "")), int(_f(row.get("Operation No", 0))), _s(row.get("Work Center Code", "")))
            elif sheet_name == "By Product Rules":
                params = (_s(row.get("Parent Item No", "")), _s(row.get("By Product Item No", "")))
            elif sheet_name == "Capacity Calendar":
                params = (_s(row.get("Work Center", "")), _s(row.get("Date", "")))
            elif sheet_name == "Production Scenario":
                params = (_s(row.get("Scenario Name", "")), _s(row.get("Product", "")))
            else:
                params = ()

            if params and params[0]:
                existing = db.conn.execute(
                    f"SELECT id FROM {_tbl} WHERE {_where}",
                    params,
                ).fetchone()
                if existing:
                    row_id = existing["id"]

        if row_id is None:
            continue

        # Kall riktig delete-metode basert på tabell
        if delete_method_name:
            delete_method = getattr(db, delete_method_name, None)
            if delete_method:
                delete_method(row_id, source="import")
                delete_count += 1

    return delete_count


# ──────────────────────────────────────────────────────────────────────
#  Importere hvert ark
# ──────────────────────────────────────────────────────────────────────

def _import_products(db: DataRepo, rows: list[dict], sheet_name: str) -> int:
    """Importer produkter. Returner antall endringer."""
    # Først: håndter DELETE
    deletions = _deletions_from_action(rows, sheet_name, db)

    # Så: upsert CREATE/UPDATE/tom
    product_list = []
    for row in rows:
        action = _get_action(row)
        if action == "DELETE":
            continue

        item_no = _s(row.get("Item No", ""))
        if not item_no:
            continue

        # Slå opp id basert på naturlig nøkkel hvis Rad ID mangler
        row_id = _i(row.get("Rad ID"))
        if row_id is None:
            existing = db.conn.execute(
                "SELECT id FROM products WHERE item_no = ?",
                (item_no,),
            ).fetchone()
            if existing:
                row_id = existing["id"]

        entry = {
            "id": row_id,
            "item_no": item_no,
            "description": _s(row.get("Description", "")),
            "item_type": _s(row.get("Item Type", "")),
            "product_group": _s(row.get("Product Group", "")),
            "base_uom": _s(row.get("Base Unit of Measure", "")),
        }
        product_list.append(entry)

        # Is Transport: skriv til transport_flagg-tabellen.
        # Kun når flag=1, eller når raden allerede finnes (for å kunne nedgradere 1→0).
        # Nye varer med is_transport=0 trenger ingen rad — fravær av rad = ikke transportvare.
        is_transport = row.get("Is Transport")
        if is_transport is not None and not pd.isna(is_transport):
            flag = 1 if str(is_transport).strip().lower() in ("1", "ja", "true", "yes") else 0
            har_rad = db.conn.execute(
                "SELECT 1 FROM transport_flagg WHERE item_no = ?", (item_no,)
            ).fetchone()
            if flag == 1 or har_rad:
                db.conn.execute(
                    """INSERT INTO transport_flagg (item_no, is_transport)
                       VALUES (?, ?)
                       ON CONFLICT(item_no) DO UPDATE SET is_transport = excluded.is_transport, updated_at = datetime('now')""",
                    (item_no, flag),
                )

    if product_list:
        db.upsert_products(product_list, source="import")
    db.conn.commit()
    return len(product_list) + deletions


def _import_locations(db: DataRepo, rows: list[dict], sheet_name: str) -> int:
    """Importer lokasjoner."""
    deletions = _deletions_from_action(rows, sheet_name, db)

    loc_list = []
    for row in rows:
        action = _get_action(row)
        if action == "DELETE":
            continue

        code = _s(row.get("Location Code", ""))
        if not code:
            continue

        # Slå opp id basert på naturlig nøkkel hvis Rad ID mangler
        row_id = _i(row.get("Rad ID"))
        if row_id is None:
            existing = db.conn.execute(
                "SELECT id FROM locations WHERE code = ?",
                (code,),
            ).fetchone()
            if existing:
                row_id = existing["id"]

        loc_list.append({
            "id": row_id,
            "code": code,
            "name": _s(row.get("Location Name", "")),
            "location_type": _s(row.get("Location Type", "")),
        })
    if loc_list:
        db.upsert_locations(loc_list, source="import")
    return len(loc_list) + deletions


def _import_work_centers(db: DataRepo, rows: list[dict], sheet_name: str) -> int:
    """Importer arbeidssentre."""
    deletions = _deletions_from_action(rows, sheet_name, db)

    wc_list = []
    for row in rows:
        action = _get_action(row)
        if action == "DELETE":
            continue

        code = _s(row.get("Work Center Code", ""))
        if not code:
            continue

        # Slå opp id basert på naturlig nøkkel hvis Rad ID mangler
        row_id = _i(row.get("Rad ID"))
        if row_id is None:
            existing = db.conn.execute(
                "SELECT id FROM work_centers WHERE code = ?",
                (code,),
            ).fetchone()
            if existing:
                row_id = existing["id"]

        wc_list.append({
            "id": row_id,
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
    return len(wc_list) + deletions


def _import_operations(db: DataRepo, rows: list[dict], sheet_name: str) -> int:
    """Importer operasjoner."""
    deletions = _deletions_from_action(rows, sheet_name, db)

    op_list = []
    for row in rows:
        action = _get_action(row)
        if action == "DELETE":
            continue

        code = _s(row.get("Operation Code", ""))
        if not code:
            continue

        # Slå opp id basert på naturlig nøkkel hvis Rad ID mangler
        row_id = _i(row.get("Rad ID"))
        if row_id is None:
            existing = db.conn.execute(
                "SELECT id FROM operations WHERE code = ?",
                (code,),
            ).fetchone()
            if existing:
                row_id = existing["id"]

        op_list.append({
            "id": row_id,
            "code": code,
            "description": _s(row.get("Description", "")),
            "default_work_center": _s(row.get("Default Work Center", "")),
            "standard_unit": _s(row.get("Standard Unit", "Minutes")),
        })
    if op_list:
        db.upsert_operations(op_list, source="import")
    return len(op_list) + deletions


def _import_item_costs(db: DataRepo, rows: list[dict], sheet_name: str) -> int:
    """Importer kostpriser."""
    deletions = _deletions_from_action(rows, sheet_name, db)

    cost_list = []
    for row in rows:
        action = _get_action(row)
        if action == "DELETE":
            continue

        item_no = _s(row.get("Item No", ""))
        if not item_no:
            continue

        cost_type = _s(row.get("Cost Type", "Standard Cost"))

        entry = {
            "item_no": item_no,
            "cost_type": cost_type,
            "unit_cost": _f(row.get("Unit Cost", 0)),
            "currency": _s(row.get("Currency", "NOK")),
            "effective_date": _d(row.get("Effective Date")),
        }

        # Hvis Rad ID mangler, slå opp id basert på naturlig nøkkel
        row_id = _i(row.get("Rad ID"))
        if row_id is None:
            existing = db.conn.execute(
                "SELECT id FROM item_costs WHERE item_no = ? AND cost_type = ?",
                (item_no, cost_type),
            ).fetchone()
            if existing:
                row_id = existing["id"]
        entry["id"] = row_id
        cost_list.append(entry)

    if cost_list:
        db.upsert_item_costs(cost_list, source="import")
    return len(cost_list) + deletions


def _import_bom(db: DataRepo, rows: list[dict], sheet_name: str) -> int:
    """Importer BOM-linjer."""
    deletions = _deletions_from_action(rows, sheet_name, db)

    bom_list = []
    for row in rows:
        action = _get_action(row)
        if action == "DELETE":
            continue

        parent = _s(row.get("Parent Item No", ""))
        component = _s(row.get("Component Item No", ""))
        if not parent or not component:
            continue

        entry = {
            "parent_item_no": parent,
            "component_item_no": component,
            "quantity_per": _f(row.get("Quantity Per", 1)),
            "uom": _s(row.get("Unit of Measure", "")),
            "scrap_pct": _f(row.get("Scrap %", 0)),
            "co_product_pct": _f(row.get("Co-Prod %", 0)),
            "co_product_item_no": _s(row.get("Co-Prod Item No", "")),
        }

        # Hvis Rad ID mangler, slå opp id basert på naturlig nøkkel
        row_id = _i(row.get("Rad ID"))
        if row_id is None:
            existing = db.conn.execute(
                "SELECT id FROM bom_lines WHERE parent_item_no = ? AND component_item_no = ?",
                (parent, component),
            ).fetchone()
            if existing:
                row_id = existing["id"]
        entry["id"] = row_id
        bom_list.append(entry)

    if bom_list:
        db.upsert_bom_lines(bom_list, source="import")
    return len(bom_list) + deletions


def _import_routing(db: DataRepo, rows: list[dict], sheet_name: str) -> int:
    """Importer routing-linjer."""
    deletions = _deletions_from_action(rows, sheet_name, db)

    rt_list = []
    for row in rows:
        action = _get_action(row)
        if action == "DELETE":
            continue

        item_no = _s(row.get("Item No", ""))
        if not item_no:
            continue

        operation_no = int(_f(row.get("Operation No", 0)))
        wc_code = _s(row.get("Work Center Code", ""))

        entry = {
            "item_no": item_no,
            "operation_no": operation_no,
            "operation_code": _s(row.get("Operation Code", "")),
            "work_center_code": wc_code,
            "setup_time_minutes": _f(row.get("Setup Time Minutes", 0)),
            "run_time_minutes": _f(row.get("Run Time Minutes", 0)),
            "batch_size": _f(row.get("Batch Size", 1)),
        }

        # Hvis Rad ID mangler, slå opp id basert på naturlig nøkkel
        row_id = _i(row.get("Rad ID"))
        if row_id is None:
            existing = db.conn.execute(
                "SELECT id FROM routing_lines WHERE item_no = ? AND operation_no = ? AND work_center_code = ?",
                (item_no, operation_no, wc_code),
            ).fetchone()
            if existing:
                row_id = existing["id"]
        entry["id"] = row_id
        rt_list.append(entry)

    if rt_list:
        db.upsert_routing_lines(rt_list, source="import")
    return len(rt_list) + deletions


def _import_byproduct_rules(db: DataRepo, rows: list[dict], sheet_name: str) -> int:
    """Importer biproduktregler."""
    deletions = _deletions_from_action(rows, sheet_name, db)

    bp_list = []
    for row in rows:
        action = _get_action(row)
        if action == "DELETE":
            continue

        parent = _s(row.get("Parent Item No", ""))
        byprod = _s(row.get("By Product Item No", ""))
        if not parent or not byprod:
            continue

        entry = {
            "parent_item_no": parent,
            "by_product_item_no": byprod,
            "expected_quantity": _f(row.get("Expected Quantity", 0)),
            "uom": _s(row.get("Unit of Measure", "")),
            "market_value": _f(row.get("Market Value", 0)),
            "allocation_method": _s(row.get("Allocation Method", "Reduce Main Product Cost")),
        }

        # Hvis Rad ID mangler, slå opp id basert på naturlig nøkkel
        row_id = _i(row.get("Rad ID"))
        if row_id is None:
            existing = db.conn.execute(
                "SELECT id FROM byproduct_rules WHERE parent_item_no = ? AND by_product_item_no = ?",
                (parent, byprod),
            ).fetchone()
            if existing:
                row_id = existing["id"]
        entry["id"] = row_id
        bp_list.append(entry)

    if bp_list:
        db.upsert_byproduct_rules(bp_list, source="import")
    return len(bp_list) + deletions


def _import_capacity(db: DataRepo, rows: list[dict], sheet_name: str) -> int:
    """Importer kapasitetskalender."""
    deletions = _deletions_from_action(rows, sheet_name, db)

    cap_list = []
    for row in rows:
        action = _get_action(row)
        if action == "DELETE":
            continue

        wc = _s(row.get("Work Center", ""))
        d = _d(row.get("Date"))
        if not wc or not d:
            continue

        entry = {
            "work_center": wc,
            "date": d,
            "available_hours": _f(row.get("Available Hours", 0)),
            "planned_downtime": _f(row.get("Planned Downtime", 0)),
        }

        # Hvis Rad ID mangler, slå opp id basert på naturlig nøkkel
        row_id = _i(row.get("Rad ID"))
        if row_id is None:
            existing = db.conn.execute(
                "SELECT id FROM capacity_days WHERE work_center = ? AND date = ?",
                (wc, d),
            ).fetchone()
            if existing:
                row_id = existing["id"]
        entry["id"] = row_id
        cap_list.append(entry)

    if cap_list:
        db.upsert_capacity_days(cap_list, source="import")
    return len(cap_list) + deletions


def _import_transport_ruter(db: DataRepo, rows: list[dict], sheet_name: str) -> int:
    """Importer transportruter mellom høvlerier."""
    deletions = 0
    for row in rows:
        action = _get_action(row)
        if action == "DELETE":
            # Slett rute basert på from_loc + to_loc
            existing = db.conn.execute(
                "SELECT id FROM transport_ruter WHERE from_loc = ? AND to_loc = ?",
                (_s(row.get("From Loc", "")), _s(row.get("To Loc", ""))),
            ).fetchone()
            if existing:
                db.conn.execute("DELETE FROM transport_ruter WHERE id = ?", (existing["id"],))
                db.conn.commit()
                deletions += 1
            continue

        from_loc = _s(row.get("From Loc", ""))
        to_loc = _s(row.get("To Loc", ""))
        if not from_loc or not to_loc:
            continue

        existing = db.conn.execute(
            "SELECT id FROM transport_ruter WHERE from_loc = ? AND to_loc = ?",
            (from_loc, to_loc),
        ).fetchone()

        cost_per_m3 = _f(row.get("Cost Per M3", 0))
        distance = _f(row.get("Distance Km", 0))
        hours = _f(row.get("Hours", 0))

        if existing:
            # Sammenlign for endringslogg
            old = dict(db.conn.execute(
                "SELECT * FROM transport_ruter WHERE id = ?", (existing["id"],)
            ).fetchone())
            if abs(old.get("cost_per_m3", 0) - cost_per_m3) > 0.001:
                db.log_change("transport_ruter", str(existing["id"]), "cost_per_m3",
                              old.get("cost_per_m3"), cost_per_m3, source="import")
            if abs(old.get("distance_km", 0) - distance) > 0.001:
                db.log_change("transport_ruter", str(existing["id"]), "distance_km",
                              old.get("distance_km"), distance, source="import")
            if abs(old.get("hours", 0) - hours) > 0.001:
                db.log_change("transport_ruter", str(existing["id"]), "hours",
                              old.get("hours"), hours, source="import")
            db.conn.execute(
                """UPDATE transport_ruter SET cost_per_m3 = ?, distance_km = ?, hours = ? WHERE id = ?""",
                (cost_per_m3, distance, hours, existing["id"]),
            )
        else:
            db.conn.execute(
                """INSERT INTO transport_ruter (from_loc, to_loc, cost_per_m3, distance_km, hours)
                   VALUES (?, ?, ?, ?, ?)""",
                (from_loc, to_loc, cost_per_m3, distance, hours),
            )
            db.log_change("transport_ruter", f"{from_loc}:{to_loc}", "_created",
                          None, f"{from_loc}→{to_loc}", source="import")
        db.conn.commit()

    return len(rows) - deletions


def _import_scenarios(db: DataRepo, rows: list[dict], sheet_name: str) -> int:
    """Importer produksjonsscenarioer."""
    deletions = _deletions_from_action(rows, sheet_name, db)

    sc_list = []
    for row in rows:
        action = _get_action(row)
        if action == "DELETE":
            continue

        name = _s(row.get("Scenario Name", ""))
        product = _s(row.get("Product", ""))
        if not name or not product:
            continue

        entry = {
            "scenario_name": name,
            "product": product,
            "planned_quantity": _f(row.get("Planned Quantity", 0)),
        }

        # Hvis Rad ID mangler, slå opp id basert på naturlig nøkkel
        row_id = _i(row.get("Rad ID"))
        if row_id is None:
            existing = db.conn.execute(
                "SELECT id FROM production_scenarios WHERE scenario_name = ? AND product = ?",
                (name, product),
            ).fetchone()
            if existing:
                row_id = existing["id"]
        entry["id"] = row_id
        sc_list.append(entry)

    if sc_list:
        db.upsert_scenarios(sc_list, source="import")
    return len(sc_list) + deletions


# ──────────────────────────────────────────────────────────────────────
#  Eksport: SQLite → Excel
# ──────────────────────────────────────────────────────────────────────

def export_sqlite_to_excel(output_path: str, db: Optional[DataRepo] = None) -> None:
    """Eksporter all data fra SQLite til en Excel-fil.
    
    Genererer de samme 10 arkene som den originale Excel-malen.
    Hvert ark har ACTION (col A) og Rad ID (col B) før datakolonnene.
    Inkluderer et 11. ark: 'Endringslogg' (uten ACTION/ID).

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
    action_font = Font(name='Calibri', size=10, italic=True, color='666666')
    thin_border = Border(
        left=Side(style='thin', color='48BB78'),
        right=Side(style='thin', color='48BB78'),
        top=Side(style='thin', color='48BB78'),
        bottom=Side(style='thin', color='48BB78'),
    )

    def _write_sheet(ws, title: str, rows: list, field_names: list[str],
                     db_keys: Optional[list[str]] = None,
                     include_action: bool = True):
        """Skriv data til et Excel-ark med header, kolonnebeskrivelser og dropdowns.

        Args:
            ws: openpyxl worksheet
            title: Ark-tittel (Excel-fane)
            rows: Liste med dicts/rad-objekter
            field_names: Kolonneoverskrifter (Excel-visning)
            db_keys: Database-nøkler for data-aksess. Hvis None, konverteres field_names
                     til snake_case (f.eks. "Item No" → "item_no").
            include_action: Om ACTION/ID skal inkluderes (True for data-ark, False for Endringslogg)
        """
        ws.title = title

        # Hvis db_keys ikke er oppgitt, konverter field_names til snake_case
        if db_keys is None:
            db_keys = [name.lower().replace(" ", "_") for name in field_names]

        kol_desc = KOLONNER.get(title, {})

        # -- Header --
        col_offset = 1
        if include_action:
            # Kolonne A: ACTION
            cell = ws.cell(row=1, column=1, value="ACTION")
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = thin_border
            cell.comment = XLComment(
                "Handling for denne raden ved import. CREATE=ny, UPDATE=endre, DELETE=slett. "
                "Tom verdi = inferer (CREATE hvis ny, UPDATE hvis eksisterende).",
                "System", width=300, height=100,
            )

            # Kolonne B: Rad ID
            cell = ws.cell(row=1, column=2, value="Rad ID")
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = thin_border
            cell.comment = XLComment(
                "Unik database-ID for raden. Skrivebeskyttet - brukes kun for matching ved import.",
                "System", width=300, height=100,
            )
            col_offset = 3

        # Headere for datakolonner
        for col_idx, name in enumerate(field_names, col_offset):
            cell = ws.cell(row=1, column=col_idx, value=name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = thin_border
            if name in kol_desc:
                cell.comment = XLComment(kol_desc[name], "System", width=300, height=100)

        if not rows:
            if not include_action:
                ws.cell(row=2, column=1, value="(Ingen data)").font = data_font
            return

        # -- Data --
        for row_idx, row in enumerate(rows, 2):
            if hasattr(row, 'keys'):
                row = dict(row)

            if include_action:
                # ACTION: default tom
                cell = ws.cell(row=row_idx, column=1, value="")
                cell.font = action_font
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center')

                # Rad ID
                row_id = row.get("id", "")
                cell = ws.cell(row=row_idx, column=2, value=row_id)
                cell.font = data_font
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center')

            for col_idx, key in enumerate(db_keys, col_offset):
                value = row.get(key, "")
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.font = data_font
                cell.border = thin_border

        # -- Auto-width --
        total_cols = len(field_names) + (2 if include_action else 0)
        for col_idx in range(1, total_cols + 1):
            max_len = len(str(ws.cell(row=1, column=col_idx).value or ""))
            for row_idx in range(2, len(rows) + 2):
                val = ws.cell(row=row_idx, column=col_idx).value
                if val is not None:
                    max_len = max(max_len, min(len(str(val)), 40))
            ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 2

        # -- Freeze panes --
        ws.freeze_panes = "A2"

        # -- Dropdowns --
        _legg_til_dropdowns(ws, title, include_action)

        # -- Høyde på rad 1 --
        ws.row_dimensions[1].height = 40

    def _legg_til_dropdowns(ws, ark_navn: str, include_action: bool = True):
        """Legg til statiske dropdowns for et ark."""
        # ACTION-dropdown (kolonne A) - alltid hvis include_action
        if include_action:
            dv_action = DataValidation(
                type="list",
                formula1=ACTION_VALUES,
                allow_blank=True,
                showErrorMessage=True,
                errorTitle="Ugyldig ACTION",
                error="Verdien må være: CREATE, UPDATE, DELETE, eller tom",
            )
            ws.add_data_validation(dv_action)
            dv_action.add("A2:A1048576")

        # Eksisterende statiske dropdowns (justert for ACTION/ID offset)
        col_adjust = 2 if include_action else 0
        for sheet, col, liste in STATISKE_DROPDOWNS:
            if sheet != ark_navn:
                continue
            if not liste:
                continue
            # Konverter kolonnebokstav til index, juster, og tilbake til bokstav
            col_idx = ord(col.upper()) - 65  # A=0
            new_col = col_idx + col_adjust + 1  # +1 for 1-indexed
            new_col_letter = get_column_letter(new_col)
            dv = DataValidation(
                type="list",
                formula1=liste,
                allow_blank=True,
                showErrorMessage=True,
                errorTitle="Ugyldig verdi",
                error=f"Verdien må være en av: {liste.replace(chr(34), '')}",
            )
            ws.add_data_validation(dv)
            dv.add(f"{new_col_letter}2:{new_col_letter}1048576")

    # ── Ark 1: Product Master ────────────────────────────────
    ws1 = wb.active
    # Slå opp transport_flagg for å vise Is Transport-kolonnen
    _transport_flagg = {r["item_no"]: r["is_transport"] for r in data.get("transport_flagg", [])}
    _products_with_transport = []
    for p in data.get("products", []):
        _row = dict(p)
        _row["is_transport"] = _transport_flagg.get(_row.get("item_no", ""), 0)
        _products_with_transport.append(_row)
    _write_sheet(ws1, "Product Master", _products_with_transport,
                 ["Item No", "Description", "Item Type", "Product Group", "Base Unit of Measure", "Is Transport"],
                 db_keys=["item_no", "description", "item_type", "product_group", "base_uom", "is_transport"])

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

    # ── Ark 11: Transport Ruter ─────────────────────────────
    # (Før Endringslogg så antallet ruter er lett å finne)
    ws11 = wb.create_sheet()
    _write_sheet(ws11, "Transport Ruter", data.get("transport_ruter", []),
                 ["From Loc", "To Loc", "Cost Per M3", "Distance Km", "Hours"],
                 db_keys=["from_loc", "to_loc", "cost_per_m3", "distance_km", "hours"])

    # ── Ark 12: Endringslogg ────────────────────────────────
    # Endringsloggen har IKKE ACTION/ID kolonner
    ws12 = wb.create_sheet()
    change_log = db.get_changes(limit=1000)
    if change_log:
        _write_sheet(ws12, "Endringslogg", change_log,
                     ["ID", "Tidspunkt", "Bruker", "Kilde", "Tabell", "Nokkel",
                      "Felt", "Gammel verdi", "Ny verdi"],
                     db_keys=["id", "timestamp", "user", "source", "table_name",
                              "record_key", "field_name", "old_value", "new_value"],
                     include_action=False)
    else:
        ws12.title = "Endringslogg"
        ws12.cell(row=1, column=1, value="(Ingen endringer logget)").font = data_font

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