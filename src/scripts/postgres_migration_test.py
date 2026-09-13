#!/usr/bin/env python3
"""End-to-end test av SQLite -> PostgreSQL-migreringen.

Lager en midlertidig SQLite-database med representative data og historikk,
kjører det ekte migreringsscriptet og verifiserer PostgreSQL-resultatet.
"""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path
from tempfile import TemporaryDirectory


_src = Path(__file__).resolve().parents[1]
if str(_src) not in sys.path:
    sys.path.insert(0, str(_src))

from data_repo import DataRepo


ITEM_NO = "MIGRATION_TEST_001"
UPLOAD_NAME = "migration_test.xlsx"


def _build_sqlite_fixture(sqlite_path: Path, database_url: str) -> dict[str, int]:
    """Bygg testkilde uten at DATABASE_URL tvinger DataRepo til PostgreSQL."""
    old_url = os.environ.pop("DATABASE_URL", None)
    try:
        source = DataRepo(str(sqlite_path))
        source.initialize()
        source.upsert_products(
            [
                {
                    "item_no": ITEM_NO,
                    "description": "Migration test product",
                    "item_type": "Finished Good",
                    "product_group": "TEST",
                    "base_uom": "LM",
                }
            ],
            source="migration_test_source",
        )
        source.execute(
            "INSERT INTO transport_flagg (item_no, is_transport) VALUES (?, ?)",
            (ITEM_NO, 1),
        )
        source.execute(
            """INSERT INTO demand
               (product_id, period, quantity, location_code, customer_region)
               VALUES (?, ?, ?, ?, ?)""",
            (ITEM_NO, 37, 1234.5, "KOD", "TEST"),
        )
        source.conn.commit()
        source.save_upload(
            UPLOAD_NAME,
            b"migration-test-bytes",
            comment="Migration history test",
            row_count=1,
        )
        stats = source.stats
        source.close()
        return stats
    finally:
        if old_url is not None:
            os.environ["DATABASE_URL"] = old_url
        else:
            os.environ["DATABASE_URL"] = database_url


def main() -> int:
    database_url = os.environ.get("DATABASE_URL", "").strip()
    if not database_url:
        print("[ERROR] DATABASE_URL er ikke satt.")
        return 1

    with TemporaryDirectory(prefix="prodcalc-migration-test-") as tmpdir:
        sqlite_path = Path(tmpdir) / "source.db"
        source_stats = _build_sqlite_fixture(sqlite_path, database_url)

        env = os.environ.copy()
        env["DATABASE_URL"] = database_url
        cmd = [
            sys.executable,
            str(_src / "scripts" / "migrate_sqlite_to_postgres.py"),
            "--sqlite-path",
            str(sqlite_path),
        ]
        completed = subprocess.run(cmd, env=env, text=True, capture_output=True)
        print(completed.stdout, end="")
        if completed.stderr:
            print(completed.stderr, file=sys.stderr, end="")
        if completed.returncode != 0:
            print(f"[ERROR] Migreringsscriptet returnerte {completed.returncode}")
            return 2

        target = DataRepo()
        try:
            target.initialize()
            target_stats = target.stats

            row = target.execute(
                "SELECT description FROM products WHERE item_no = ?",
                (ITEM_NO,),
            ).fetchone()
            if row is None or row["description"] != "Migration test product":
                print("[ERROR] Produktet ble ikke migrert korrekt.")
                return 3

            demand = target.execute(
                "SELECT quantity FROM demand WHERE product_id = ? AND period = ?",
                (ITEM_NO, 37),
            ).fetchone()
            if demand is None or abs(float(demand["quantity"]) - 1234.5) > 0.001:
                print("[ERROR] Demand ble ikke migrert korrekt.")
                return 4

            flag = target.execute(
                "SELECT is_transport FROM transport_flagg WHERE item_no = ?",
                (ITEM_NO,),
            ).fetchone()
            if flag is None or not bool(flag["is_transport"]):
                print("[ERROR] Transportflagget ble ikke migrert korrekt.")
                return 5

            upload = target.execute(
                "SELECT blob, comment FROM uploaded_files WHERE filename = ?",
                (UPLOAD_NAME,),
            ).fetchone()
            if upload is None or bytes(upload["blob"]) != b"migration-test-bytes":
                print("[ERROR] Opplastet fil/blob ble ikke migrert korrekt.")
                return 6
            if upload["comment"] != "Migration history test":
                print("[ERROR] Kommentar på opplastet fil ble ikke bevart.")
                return 7

            source_change_count = int(source_stats.get("change_log", 0))
            target_change_count = int(target_stats.get("change_log", 0))
            if source_change_count != target_change_count:
                print(
                    "[ERROR] Endringsloggen ble ikke bevart: "
                    f"kilde={source_change_count}, mål={target_change_count}"
                )
                return 8

            upload_count = int(target_stats.get("uploaded_files", 0))
            if upload_count != int(source_stats.get("uploaded_files", 0)):
                print("[ERROR] uploaded_files-radtall avviker etter migrering.")
                return 9

            print("[OK] SQLite -> PostgreSQL migrering er validert end-to-end")
            return 0
        finally:
            target.clear_all_data()
            target.clear_change_log()
            target.execute("DELETE FROM uploaded_files")
            target.conn.commit()
            target.close()


if __name__ == "__main__":
    raise SystemExit(main())
