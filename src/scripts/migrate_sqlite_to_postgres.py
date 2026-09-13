#!/usr/bin/env python3
"""Migrer data fra SQLite til PostgreSQL på en kontrollert måte.

Bruk:
    $env:DATABASE_URL="postgresql://user:pass@host:5432/db?sslmode=require"
    python src/scripts/migrate_sqlite_to_postgres.py --sqlite-path src/produksjonskalkyle.db

Standardoppførsel:
- SQLite-kilden åpnes uten schema-migreringer, slik at kildefilen ikke endres.
- PostgreSQL-målet tømmes før migrering.
- Stamdata, demand, transportdata, endringslogg og opplastede filer migreres.
- Kilde- og måltellinger sammenlignes og migreringen feiler ved avvik.
- Ved feil tømmes målet igjen for å unngå å etterlate en delvis migrering.

--keep-target-data er ment for spesielle merge-scenarier. Da tømmes ikke målet,
og streng radtallsvalidering deaktiveres fordi eksisterende måldata kan være legitime.
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path


_src = Path(__file__).resolve().parents[1]
if str(_src) not in sys.path:
    sys.path.insert(0, str(_src))

from data_repo import DataRepo


MIGRATED_TABLES = [
    "products",
    "locations",
    "work_centers",
    "operations",
    "item_costs",
    "bom_lines",
    "routing_lines",
    "byproduct_rules",
    "capacity_days",
    "production_scenarios",
    "demand",
    "historical_sales",
    "changeover_matrix",
    "transport_flagg",
    "transport_ruter",
    "change_log",
    "uploaded_files",
]


def _rows_to_dicts(rows) -> list[dict]:
    return [dict(row) for row in rows]


def _read_source_rows(source: DataRepo) -> dict[str, list]:
    """Les alle data som skal flyttes, også historikktabellene.

    DataRepo.export_all_data() brukes av Excel-eksport og inneholder med vilje
    ikke blob-/audit-historikken. Migreringsscriptet må derfor hente disse to
    tabellene eksplisitt.
    """
    rows = source.export_all_data()
    rows["change_log"] = source.execute(
        'SELECT timestamp, "user", source, table_name, record_key, '
        'field_name, old_value, new_value FROM change_log ORDER BY id'
    ).fetchall()
    rows["uploaded_files"] = source.execute(
        "SELECT filename, uploaded_at, blob, comment, row_count "
        "FROM uploaded_files ORDER BY id"
    ).fetchall()
    return rows


def _clear_target(target: DataRepo) -> None:
    """Tøm alle data som migreringsscriptet eier."""
    target.clear_all_data()
    target.clear_change_log()
    target.execute("DELETE FROM uploaded_files")
    target.conn.commit()


def _migrate_main_tables(source_rows: dict[str, list], target: DataRepo) -> None:
    target.upsert_products(_rows_to_dicts(source_rows.get("products", [])), source="migration")
    target.upsert_locations(_rows_to_dicts(source_rows.get("locations", [])), source="migration")
    target.upsert_work_centers(_rows_to_dicts(source_rows.get("work_centers", [])), source="migration")
    target.upsert_operations(_rows_to_dicts(source_rows.get("operations", [])), source="migration")
    target.upsert_item_costs(_rows_to_dicts(source_rows.get("item_costs", [])), source="migration")
    target.upsert_bom_lines(_rows_to_dicts(source_rows.get("bom_lines", [])), source="migration")
    target.upsert_routing_lines(_rows_to_dicts(source_rows.get("routing_lines", [])), source="migration")
    target.upsert_byproduct_rules(_rows_to_dicts(source_rows.get("byproduct_rules", [])), source="migration")
    target.upsert_capacity_days(_rows_to_dicts(source_rows.get("capacity_days", [])), source="migration")
    target.upsert_scenarios(_rows_to_dicts(source_rows.get("production_scenarios", [])), source="migration")


def _migrate_transport_flagg(source_rows: dict[str, list], target: DataRepo) -> None:
    rows = _rows_to_dicts(source_rows.get("transport_flagg", []))
    if not rows:
        return
    params = [
        (r.get("item_no", ""), int(r.get("is_transport", 0) or 0))
        for r in rows
    ]
    target.executemany(
        """INSERT INTO transport_flagg (item_no, is_transport)
           VALUES (?, ?)
           ON CONFLICT(item_no) DO UPDATE SET
               is_transport = excluded.is_transport,
               updated_at = CURRENT_TIMESTAMP""",
        params,
    )
    target.conn.commit()


def _migrate_transport_ruter(source_rows: dict[str, list], target: DataRepo) -> None:
    rows = _rows_to_dicts(source_rows.get("transport_ruter", []))
    if not rows:
        return
    params = [
        (
            r.get("from_loc", ""),
            r.get("to_loc", ""),
            float(r.get("cost_per_m3", 0) or 0),
            float(r.get("distance_km", 0) or 0),
            float(r.get("hours", 0) or 0),
        )
        for r in rows
    ]
    target.executemany(
        """INSERT INTO transport_ruter (from_loc, to_loc, cost_per_m3, distance_km, hours)
           VALUES (?, ?, ?, ?, ?)
           ON CONFLICT(from_loc, to_loc) DO UPDATE SET
               cost_per_m3 = excluded.cost_per_m3,
               distance_km = excluded.distance_km,
               hours = excluded.hours""",
        params,
    )
    target.conn.commit()


def _migrate_demand(source_rows: dict[str, list], target: DataRepo) -> None:
    rows = _rows_to_dicts(source_rows.get("demand", []))
    if not rows:
        return
    params = [
        (
            r.get("product_id", ""),
            int(r.get("period", 0) or 0),
            float(r.get("quantity", 0) or 0),
            r.get("location_code", ""),
            r.get("customer_region", ""),
        )
        for r in rows
    ]
    target.executemany(
        """INSERT INTO demand (product_id, period, quantity, location_code, customer_region)
           VALUES (?, ?, ?, ?, ?)
           ON CONFLICT(product_id, period, location_code) DO UPDATE SET
               quantity = excluded.quantity,
               customer_region = excluded.customer_region""",
        params,
    )
    target.conn.commit()


def _migrate_historical_sales(source_rows: dict[str, list], target: DataRepo) -> None:
    rows = _rows_to_dicts(source_rows.get("historical_sales", []))
    if not rows:
        return
    params = [
        (
            r.get("product_id", ""),
            int(r.get("period", 0) or 0),
            float(r.get("quantity", 0) or 0),
            r.get("location_code", ""),
        )
        for r in rows
    ]
    target.executemany(
        """INSERT INTO historical_sales (product_id, period, quantity, location_code)
           VALUES (?, ?, ?, ?)
           ON CONFLICT(product_id, period, location_code) DO UPDATE SET
               quantity = excluded.quantity""",
        params,
    )
    target.conn.commit()


def _migrate_changeover_matrix(source_rows: dict[str, list], target: DataRepo) -> None:
    rows = _rows_to_dicts(source_rows.get("changeover_matrix", []))
    if not rows:
        return
    params = [
        (
            r.get("work_center_code", ""),
            r.get("from_family", ""),
            r.get("to_family", ""),
            float(r.get("changeover_minutes", 0) or 0),
        )
        for r in rows
    ]
    target.executemany(
        """INSERT INTO changeover_matrix (work_center_code, from_family, to_family, changeover_minutes)
           VALUES (?, ?, ?, ?)
           ON CONFLICT(work_center_code, from_family, to_family) DO UPDATE SET
               changeover_minutes = excluded.changeover_minutes""",
        params,
    )
    target.conn.commit()


def _migrate_change_log(source_rows: dict[str, list], target: DataRepo) -> None:
    rows = _rows_to_dicts(source_rows.get("change_log", []))
    if not rows:
        return
    params = [
        (
            r.get("timestamp"),
            r.get("user"),
            r.get("source", "web_form"),
            r.get("table_name", ""),
            r.get("record_key", ""),
            r.get("field_name", ""),
            r.get("old_value"),
            r.get("new_value"),
        )
        for r in rows
    ]
    target.executemany(
        'INSERT INTO change_log (timestamp, "user", source, table_name, record_key, '
        'field_name, old_value, new_value) VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
        params,
    )
    target.conn.commit()


def _migrate_uploaded_files(source_rows: dict[str, list], target: DataRepo) -> None:
    rows = _rows_to_dicts(source_rows.get("uploaded_files", []))
    if not rows:
        return
    params = [
        (
            r.get("filename", ""),
            r.get("uploaded_at"),
            r.get("blob"),
            r.get("comment", ""),
            int(r.get("row_count", 0) or 0),
        )
        for r in rows
    ]
    target.executemany(
        """INSERT INTO uploaded_files (filename, uploaded_at, blob, comment, row_count)
           VALUES (?, ?, ?, ?, ?)""",
        params,
    )
    target.conn.commit()


def _counts(repo: DataRepo) -> dict[str, int]:
    return repo.stats


def _validate_counts(source_stats: dict[str, int], target_stats: dict[str, int]) -> list[str]:
    """Returner menneskelesbare avvik mellom kilde og mål."""
    mismatches: list[str] = []
    for table in MIGRATED_TABLES:
        source_count = int(source_stats.get(table, 0))
        target_count = int(target_stats.get(table, 0))
        if source_count != target_count:
            mismatches.append(
                f"{table}: kilde={source_count}, mål={target_count}"
            )
    return mismatches


def main() -> int:
    parser = argparse.ArgumentParser(description="Migrer ProduksjonsKalkyle-data fra SQLite til PostgreSQL")
    parser.add_argument("--sqlite-path", required=True, help="Sti til SQLite-kildedatabase")
    parser.add_argument(
        "--keep-target-data",
        action="store_true",
        help="Ikke tøm måldatabasen før migrering (deaktiverer eksakt radtallsvalidering)",
    )
    args = parser.parse_args()

    database_url = os.environ.get("DATABASE_URL", "").strip()
    if not database_url:
        print("[ERROR] DATABASE_URL er ikke satt.")
        return 1

    sqlite_path = Path(args.sqlite_path)
    if not sqlite_path.exists():
        print(f"[ERROR] Fant ikke SQLite-fil: {sqlite_path}")
        return 2

    old_database_url = os.environ.get("DATABASE_URL")
    source: DataRepo | None = None
    target: DataRepo | None = None
    source_stats: dict[str, int] = {}
    target_stats: dict[str, int] = {}

    try:
        # Kilden skal aldri endres av migreringsscriptet. Derfor connect(), ikke initialize().
        os.environ.pop("DATABASE_URL", None)
        source = DataRepo(str(sqlite_path))
        source.connect()
        source_rows = _read_source_rows(source)
        source_stats = _counts(source)
        source.close()
        source = None

        os.environ["DATABASE_URL"] = database_url
        target = DataRepo()
        target.initialize()
        if not args.keep_target_data:
            _clear_target(target)

        _migrate_main_tables(source_rows, target)
        _migrate_transport_flagg(source_rows, target)
        _migrate_transport_ruter(source_rows, target)
        _migrate_demand(source_rows, target)
        _migrate_historical_sales(source_rows, target)
        _migrate_changeover_matrix(source_rows, target)

        # CRUD-metodene over logger selve migreringen. Ved en ren migrering vil vi
        # bevare den opprinnelige revisjonshistorikken i stedet for migreringsstøy.
        if not args.keep_target_data:
            target.clear_change_log()
        _migrate_change_log(source_rows, target)
        _migrate_uploaded_files(source_rows, target)

        target_stats = _counts(target)

        if args.keep_target_data:
            print("[WARN] --keep-target-data: eksakt radtallsvalidering er hoppet over.")
        else:
            mismatches = _validate_counts(source_stats, target_stats)
            if mismatches:
                print("[ERROR] Migreringen ga radtallsavvik:")
                for mismatch in mismatches:
                    print(f"  - {mismatch}")
                _clear_target(target)
                return 3

    except Exception as exc:
        print(f"[ERROR] Migrering feilet: {exc}")
        if target is not None and not args.keep_target_data:
            try:
                target.conn.rollback()
                _clear_target(target)
                print("[INFO] Måldatabasen er tømt etter feilen.")
            except Exception as cleanup_exc:
                print(f"[WARN] Klarte ikke å rydde måldatabasen: {cleanup_exc}")
        return 4
    finally:
        if source is not None:
            source.close()
        if target is not None:
            target.close()
        if old_database_url is None:
            os.environ.pop("DATABASE_URL", None)
        else:
            os.environ["DATABASE_URL"] = old_database_url

    print("[OK] Migrering fullført og validert")
    print("[INFO] Kilde-statistikk:")
    for key, value in source_stats.items():
        print(f"  {key}: {value}")
    print("[INFO] Mål-statistikk:")
    for key, value in target_stats.items():
        print(f"  {key}: {value}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
