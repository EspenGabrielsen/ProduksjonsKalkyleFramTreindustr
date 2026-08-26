#!/usr/bin/env python3
"""Enkel smoke test for PostgreSQL-backend.

Bruk:
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


def main() -> int:
    database_url = os.environ.get("DATABASE_URL", "").strip()
    if not database_url:
        print("[ERROR] DATABASE_URL er ikke satt.")
        return 1

    db = DataRepo()
    print(f"[INFO] Backend: {db.backend}")
    db.initialize()
    print("[OK] initialize() fullført")

    db.upsert_products([
        {
            "item_no": "SMOKE_FG_001",
            "description": "Smoke test product",
            "item_type": "FG",
            "product_group": "TEST",
            "base_uom": "LM",
        }
    ], source="smoke_test")
    db.conn.commit()
    print("[OK] upsert_products() fullført")

    rows = db.execute("SELECT COUNT(*) AS count FROM products").fetchone()
    count = rows["count"] if isinstance(rows, dict) else rows[0]
    print(f"[OK] products count: {count}")

    smoke_row = db.execute(
        "SELECT item_no, description FROM products WHERE item_no = ?",
        ("SMOKE_FG_001",),
    ).fetchone()
    if smoke_row is None:
        print("[ERROR] Fant ikke smoke test product etter insert.")
        db.close()
        return 2
    print(f"[OK] roundtrip: {smoke_row['item_no']} / {smoke_row['description']}")

    db.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())