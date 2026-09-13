#!/usr/bin/env python3
"""Verifiser at dataendring og change_log er én atomisk transaksjon.

Testen fremprovoserer en UNIQUE-feil under produktoppdatering. Etter feilen skal
hverken produktet eller audit-loggen vise den mislykkede endringen.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path


_src = Path(__file__).resolve().parents[1]
if str(_src) not in sys.path:
    sys.path.insert(0, str(_src))

from data_repo import DataRepo


ITEM_A = "TX_ATOMIC_A"
ITEM_B = "TX_ATOMIC_B"


def _cleanup(db: DataRepo) -> None:
    db.conn.rollback()
    rows = db.execute(
        "SELECT id FROM products WHERE item_no IN (?, ?)",
        (ITEM_A, ITEM_B),
    ).fetchall()
    ids = [str(row["id"]) for row in rows]
    db.execute("DELETE FROM products WHERE item_no IN (?, ?)", (ITEM_A, ITEM_B))
    for row_id in ids:
        db.execute("DELETE FROM change_log WHERE record_key = ?", (row_id,))
    db.conn.commit()


def main() -> int:
    if not os.environ.get("DATABASE_URL", "").strip():
        print("[ERROR] DATABASE_URL er ikke satt")
        return 1

    db = DataRepo()
    try:
        db.initialize()
        _cleanup(db)

        ids = db.upsert_products(
            [
                {"item_no": ITEM_A, "description": "Atomic A"},
                {"item_no": ITEM_B, "description": "Atomic B"},
            ],
            source="transaction_test",
        )
        if len(ids) != 2:
            print("[ERROR] Klarte ikke å opprette transaksjonstest-data")
            return 2

        b_id = int(ids[1])
        db.execute("DELETE FROM change_log WHERE record_key = ?", (str(b_id),))
        db.conn.commit()

        try:
            db.upsert_products(
                [
                    {
                        "id": b_id,
                        "item_no": ITEM_A,
                        "description": "Should never commit",
                    }
                ],
                source="transaction_test_failed_update",
            )
        except Exception:
            # Metoden selv skal rulle tilbake. Denne rollbacken gjør testen robust
            # også mot eldre kode slik at vi kan inspisere resultatet etterpå.
            db.conn.rollback()
        else:
            print("[ERROR] Forventet UNIQUE-feil, men oppdateringen lyktes")
            return 3

        product = db.execute(
            "SELECT item_no, description FROM products WHERE id = ?", (b_id,)
        ).fetchone()
        if product is None or product["item_no"] != ITEM_B:
            print("[ERROR] Mislykket oppdatering endret produktdata")
            return 4

        audit_count = db.execute(
            """SELECT COUNT(*) AS count FROM change_log
               WHERE record_key = ? AND source = ?""",
            (str(b_id), "transaction_test_failed_update"),
        ).fetchone()["count"]
        if int(audit_count) != 0:
            print(
                "[ERROR] Audit-loggen ble committed selv om dataoppdateringen feilet "
                f"({audit_count} rad(er))"
            )
            return 5

        print("[OK] Data og audit-logg rulles tilbake atomisk")
        return 0
    finally:
        try:
            _cleanup(db)
        finally:
            db.close()


if __name__ == "__main__":
    raise SystemExit(main())
