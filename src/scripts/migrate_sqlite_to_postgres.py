#!/usr/bin/env python3
"""Migrer data fra SQLite til PostgreSQL.

Bruk:
    $env:DATABASE_URL="postgresql://user:pass@host:5432/db?sslmode=require"
    python src/scripts/migrate_sqlite_to_postgres.py --sqlite-path src/produksjonskalkyle.db

Forutsetter at:
- kilde er SQLite-fil
- mål er PostgreSQL via DATABASE_URL
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


def _rows_to_dicts(rows) -> list[dict]:
    result = []
    for row in rows:
        if isinstance(row, dict):
            result.append(dict(row))
        else:
            result.append(dict(row))
    return result


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
        (
            r.get("item_no", ""),
            int(r.get("is_transport", 0) or 0),
        )
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


def _migrate_uploaded_files(source_rows: dict[str, list], target: DataRepo) -> None:
    rows = _rows_to_dicts(source_rows.get("uploaded_files", []))
    if not rows:
        return
    params = [
        (
            r.get("filename", ""),
            r.get("blob"),
            r.get("comment", ""),
            int(r.get("row_count", 0) or 0),
        )
        for r in rows
    ]
    target.executemany(
        """INSERT INTO uploaded_files (filename, blob, comment, row_count)
           VALUES (?, ?, ?, ?)""",
        params,
    )
    target.conn.commit()


def _counts(repo: DataRepo) -> dict[str, int]:
    return repo.stats


def main() -> int:
    parser = argparse.ArgumentParser(description="Migrer ProduksjonsKalkyle-data fra SQLite til PostgreSQL")
    parser.add_argument("--sqlite-path", required=True, help="Sti til SQLite-kildedatabase")
    parser.add_argument("--keep-target-data", action="store_true", help="Ikke tøm måldatabasen før migrering")
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
    try:
        os.environ.pop("DATABASE_URL", None)
        source = DataRepo(str(sqlite_path))
        source.initialize()
        source_rows = source.export_all_data()
        source_stats = _counts(source)
        source.close()

        os.environ["DATABASE_URL"] = database_url
        target = DataRepo()
        target.initialize()
        if not args.keep_target_data:
            target.clear_all_data()
            target.clear_change_log()

        _migrate_main_tables(source_rows, target)
        _migrate_transport_flagg(source_rows, target)
        _migrate_transport_ruter(source_rows, target)
        _migrate_demand(source_rows, target)
        _migrate_historical_sales(source_rows, target)
        _migrate_changeover_matrix(source_rows, target)
        _migrate_uploaded_files(source_rows, target)

        target_stats = _counts(target)
        target.close()
    finally:
        if old_database_url is None:
            os.environ.pop("DATABASE_URL", None)
        else:
            os.environ["DATABASE_URL"] = old_database_url

    print("[OK] Migrering fullført")
    print("[INFO] Kilde-statistikk:")
    for key, value in source_stats.items():
        print(f"  {key}: {value}")
    print("[INFO] Mål-statistikk:")
    for key, value in target_stats.items():
        print(f"  {key}: {value}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())