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

    rows = db.execute("SELECT COUNT(*) AS count FROM products").fetchone()
    count = rows["count"] if isinstance(rows, dict) else rows[0]
    print(f"[OK] products count: {count}")
    db.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())