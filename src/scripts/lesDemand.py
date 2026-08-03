#!/usr/bin/env python3
"""
lesDemand.py — Les demand-data fra PowerBI-eksport og importer til SQLite.

Excel-filen (`docs/Demand_fromPowerBI.xlsx`) kobles mot PowerBI og inneholder
ordrelinjer med shipment date, varenr (FTI), ordrevolum i LM og lokasjon.

Skriptet:
  1. Leser Excel-filen
  2. Beregner kalederuke fra shipment date (ISO-uke-år + ukenummer)
  3. Identifiserer hvilke ordrelinjer som er ÅPNE ORDRE (de neste DEMAND_UKER ukene)
     og hvilke som er HISTORISK SALG (de siste 52 ukene)
  4. Grupperer per uke × vare × lokasjon
  5. Konverterer kalederuker til løpende periodenummer:
       - historical_sales: periode 1..52 (1 = eldste, 52 = nyeste fullførte uke)
       - demand (åpne ordre): periode 53.. (fortsetter etter historikken)
     Dette gir en sammenhengende tidslinje som optimeringsmotoren kan bruke.
  6. Filtrerer bort varer som ikke finnes i products-tabellen (Product Master)
  7. Importerer til demand- og historical_sales-tabellene via DataRepo

Bruk:
  python src/scripts/lesDemand.py
"""

import datetime
import sys
from pathlib import Path

import pandas as pd

# Sørg for at src/ er på sys.path
_src = Path(__file__).resolve().parent.parent
if str(_src) not in sys.path:
    sys.path.insert(0, str(_src))

from data_repo import DataRepo

# Antall uker historikk som beholdes (12 måneder ≈ 52 uker)
ANTALL_HISTORISKE_UKER = 52

# Antall uker fremover som regnes som "åpne ordre" (kunden har booket).
# Dette havner i demand-tabellen. Resten av de neste 52 ukene kunne i
# prinsippet også vært booket, men vi skiller på ordrehorisonten.
DEMAND_UKER = 2


def _iso_key(row_date: pd.Timestamp) -> tuple[int, int]:
    """Returner (ISO-år, ISO-uke) for en dato."""
    iso = row_date.isocalendar()
    return iso.year, iso.week


def main():
    today = datetime.datetime.today()
    this_year, this_week = today.isocalendar().year, today.isocalendar().week

    # Les PowerBI-eksport
    excel_path = Path(__file__).resolve().parent.parent.parent / "docs" / "Demand_fromPowerBI.xlsx"
    df = pd.read_excel(excel_path, skiprows=2)

    shipment = df["Ordrelinje[Shipment Date]"]
    df["Aar"] = shipment.dt.isocalendar().year
    df["Uke"] = shipment.dt.isocalendar().week

    # ── Finn historisk uke-referanse: "nøyaktig 52 uker siden" ──
    # Vi finner den kalenderuken som ligger 52 uker før nåværende uke.
    hist_start = today - datetime.timedelta(weeks=ANTALL_HISTORISKE_UKER)
    hist_start_iso = hist_start.isocalendar()
    hist_start_key = (hist_start_iso.year, hist_start_iso.week)

    # ── Del dataene i to: historisk salg og åpne ordre ──
    now_key = (this_year, this_week)

    # Åpne ordre: de neste DEMAND_UKER ukene (nåværende uke + D-1 neste)
    open_week_keys = set()
    for offset in range(DEMAND_UKER):
        d = today + datetime.timedelta(weeks=offset)
        iso = d.isocalendar()
        open_week_keys.add((iso.year, iso.week))

    # Historisk salg: fra hist_start_key (inkl.) til uken FØR nåværende uke
    hist_week_keys = set()
    d = today - datetime.timedelta(weeks=1)
    while True:
        iso = d.isocalendar()
        key = (iso.year, iso.week)
        if key < hist_start_key:
            break
        # Ikke inkluder nåværende uke (den er ikke fullført ennå)
        if key < now_key:
            hist_week_keys.add(key)
        d -= datetime.timedelta(weeks=1)

    # Merk hver linje med "type"
    df["_key"] = list(zip(df["Aar"], df["Uke"]))
    df["_type"] = "ignorert"
    df.loc[df["_key"].isin(open_week_keys), "_type"] = "demand"
    df.loc[df["_key"].isin(hist_week_keys), "_type"] = "historisk"

    # ── Grupper per uke × vare × lokasjon, separat per type ──
    def _group(sub: pd.DataFrame) -> pd.DataFrame:
        g = (
            sub.groupby(["Aar", "Uke", "Varer2[HA_HasasItemNo_PIL]", "Ordrelinje[Location Code]"])["[SumOrdrevolum_LM]"]
            .sum()
            .reset_index()
        )
        g.columns = ["Aar", "Uke", "item_no", "location_raw", "qty"]
        g = g[g["qty"] > 0]
        return g

    hist_df = _group(df.loc[df["_type"] == "historisk"])
    demand_df = _group(df.loc[df["_type"] == "demand"])

    # ── Konverter kalenderuker til løpende periodenummer ──
    # Sorter historiske uker kronologisk eldst → nyest og tilordne 1..52.
    # Dette gir sammenhengende perioder på tvers av årsskifte.
    hist_weeks_sorted = sorted(hist_week_keys)
    week_to_period = {key: i + 1 for i, key in enumerate(hist_weeks_sorted)}
    hist_start_period = len(hist_weeks_sorted)  # siste historiske periode

    if "Aar" in hist_df.columns and len(hist_df) > 0:
        hist_df["Periode"] = hist_df.apply(
            lambda r: week_to_period[(int(r["Aar"]), int(r["Uke"]))], axis=1
        )

    # Åpne ordre fortsetter der historikken slutter (periode 53, 54, ...).
    # Vi sorterer ordre-ukene kronologisk for å få konsistente periodenumre.
    open_weeks_sorted = sorted(open_week_keys)
    open_week_to_period = {
        key: hist_start_period + i + 1 for i, key in enumerate(open_weeks_sorted)
    }

    if "Aar" in demand_df.columns and len(demand_df) > 0:
        demand_df["Periode"] = demand_df.apply(
            lambda r: open_week_to_period[(int(r["Aar"]), int(r["Uke"]))], axis=1
        )

    print(f"[READ] Historiske uker: {len(hist_weeks_sorted)} ({hist_weeks_sorted[0] if hist_weeks_sorted else '-'} → {hist_weeks_sorted[-1] if hist_weeks_sorted else '-'})")
    print(f"[READ] Åpne ordre-uker: {sorted(open_week_keys)}")
    print(f"[READ] Historiske linjer: {len(hist_df)}, åpne ordre-linjer: {len(demand_df)}")
    print(f"[READ] Unike varer (hist): {hist_df['item_no'].nunique() if len(hist_df) else 0}, unike varer (demand): {demand_df['item_no'].nunique() if len(demand_df) else 0}")

    # Importer via DataRepo
    db = DataRepo()
    db.initialize()

    # Felles lokasjonsmapping
    loc_map = {
        "HOVEDLAGER": "KOD",
        "EIKÅS": "EIK",
        "EIKAAS": "EIK",
        "EIK": "EIK",
        "KVÅS": "KV",
        "KVAAS": "KV",
        "KV": "KV",
        "LH": "LH",
        "MOIRANA": "MOIRANA",
        "VRAK": "VRAK",
    }

    if len(hist_df) > 0:
        hist_inserted, hist_filtered = db.upsert_historical_sales_from_df(
            hist_df,
            item_col="item_no",
            period_col="Periode",
            qty_col="qty",
            location_col="location_raw",
            location_mapping=loc_map,
        )
        print(f"[OK] {hist_inserted} historiske salgslinjer importert (periode 1-{hist_start_period}).")
        print(f"[OK] {hist_filtered} historiske varer filtrert bort (finnes ikke i Product Master).")
    else:
        hist_inserted = 0
        print("[OK] Ingen historiske linjer funnet.")

    if len(demand_df) > 0:
        demand_inserted, demand_filtered = db.upsert_demand_from_df(
            demand_df,
            item_col="item_no",
            period_col="Periode",
            qty_col="qty",
            location_col="location_raw",
            location_mapping=loc_map,
        )
        print(f"[OK] {demand_inserted} åpne ordre-linjer importert til demand-tabellen (periode {hist_start_period + 1}+).")
        print(f"[OK] {demand_filtered} demand-varer filtrert bort (finnes ikke i Product Master).")
    else:
        demand_inserted = 0
        print("[OK] Ingen åpne ordre-linjer funnet.")

    # Vis statistikk
    stats = db.stats
    print(f"[DB ] demand-tabellen har nå {stats.get('demand', 0)} rader.")
    print(f"[DB ] historical_sales-tabellen har nå {stats.get('historical_sales', 0)} rader.")


if __name__ == "__main__":
    main()
