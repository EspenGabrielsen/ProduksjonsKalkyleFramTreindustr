#!/usr/bin/env python3
"""Smoke test for PostgreSQL-backend.

Testen er laget for lokal bruk og CI. Den verifiserer at PostgreSQL-sporet kan:

1. initialisere skjemaet
2. skrive og lese via DataRepo
3. lese de samme dataene via DatabaseData (kalkylemotorens datakilde)
4. rydde opp testdata etter seg

Bruk (PowerShell):
    $env:DATABASE_URL="postgresql://user:pass@host:5432/db?sslmode=require"
    python src/scripts/postgres_smoke_test.py
"""

import os
import sys
from pathlib import Path

_src = Path(__file__).resolve().parents[1]
if str(_src) not in sys.path:
    sys.path.insert(0, str(_src))

from data_repo import DataRepo
from kostberegning import DatabaseData


SMOKE_ITEM = "SMOKE_FG_001"


def _cleanup(db: DataRepo, row_id: int | None = None) -> None:
    """Fjern testdata uten å bruke CRUD-logging."""
    db.execute("DELETE FROM products WHERE item_no = ?", (SMOKE_ITEM,))
    if row_id is not None:
        db.execute(
            "DELETE FROM change_log WHERE source = ? AND record_key = ?",
            ("smoke_test", str(row_id)),
        )
    db.conn.commit()


def main() -> int:
    database_url = os.environ.get("DATABASE_URL", "").strip()
    if not database_url:
        print("[ERROR] DATABASE_URL er ikke satt.")
        return 1

    db = DataRepo()
    smoke_id: int | None = None
    data: DatabaseData | None = None

    try:
        if db.backend != "postgresql":
            print(f"[ERROR] Forventet PostgreSQL-backend, fikk: {db.backend}")
            return 2

        print(f"[INFO] Backend: {db.backend}")
        db.initialize()
        print("[OK] initialize() fullført")

        # Gjør testen idempotent etter eventuelle avbrutte tidligere kjøringer.
        _cleanup(db)

        ids = db.upsert_products([
            {
                "item_no": SMOKE_ITEM,
                "description": "Smoke test product",
                "item_type": "FG",
                "product_group": "TEST",
                "base_uom": "LM",
            }
        ], source="smoke_test")
        smoke_id = ids[0] if ids else None
        print("[OK] upsert_products() fullført")

        smoke_row = db.execute(
            "SELECT id, item_no, description FROM products WHERE item_no = ?",
            (SMOKE_ITEM,),
        ).fetchone()
        if smoke_row is None:
            print("[ERROR] Fant ikke smoke test-produkt etter insert.")
            return 3

        smoke_id = int(smoke_row["id"])
        print(f"[OK] DataRepo roundtrip: {smoke_row['item_no']} / {smoke_row['description']}")

        # Verifiser den faktiske datakilden som CostCalculator bruker i appen.
        data = DatabaseData()
        product = data.product(SMOKE_ITEM)
        if product is None:
            print("[ERROR] DatabaseData fant ikke smoke test-produktet.")
            return 4
        if product.description != "Smoke test product":
            print(f"[ERROR] Uventet produktbeskrivelse: {product.description}")
            return 5
        print("[OK] DatabaseData leser fra PostgreSQL")

        print("[OK] PostgreSQL smoke test fullført")
        return 0
    finally:
        # Dersom testen feilet midt i en PostgreSQL-transaksjon må vi rulle
        # tilbake før cleanup, ellers maskerer InFailedSqlTransaction rotfeilen.
        try:
            if db._conn is not None:
                db.conn.rollback()
            _cleanup(db, smoke_id)
        finally:
            if data is not None:
                data.db.close()
            db.close()


if __name__ == "__main__":
    raise SystemExit(main())
