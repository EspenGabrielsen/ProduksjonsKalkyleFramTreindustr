#!/usr/bin/env python3
"""Verifiser schema_version på aktiv DataRepo-backend."""

from __future__ import annotations

import sys
from pathlib import Path


_src = Path(__file__).resolve().parents[1]
if str(_src) not in sys.path:
    sys.path.insert(0, str(_src))

from data_repo import CURRENT_SCHEMA_VERSION, DataRepo


def main() -> int:
    db = DataRepo()
    try:
        db.initialize()
        version = db.get_schema_version()
        if version != CURRENT_SCHEMA_VERSION:
            print(
                "[ERROR] Feil schema_version på "
                f"{db.backend}: database={version}, kode={CURRENT_SCHEMA_VERSION}"
            )
            return 1
        print(f"[OK] {db.backend} schema_version={version}")
        return 0
    finally:
        db.close()


if __name__ == "__main__":
    raise SystemExit(main())
