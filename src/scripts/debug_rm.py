#!/usr/bin/env python3
"""Sjekk om RM_38x125 og KD_Extreem kan kjøpes/produseres."""

import sqlite3
import sys
from pathlib import Path

_src = Path(__file__).resolve().parent.parent
if str(_src) not in sys.path:
    sys.path.insert(0, str(_src))


def main():
    conn = sqlite3.connect(str(_src / "produksjonskalkyle.db"))

    for item in ["RM_38x125__US_V_Gran", "KD_Extreem", "KD_Opaque", "KD_Visir", "KD_Trebitt"]:
        prod = conn.execute(
            "SELECT item_no, item_type, base_uom FROM products WHERE item_no=?", (item,)
        ).fetchone()
        cost = conn.execute(
            "SELECT unit_cost FROM item_costs WHERE item_no=?", (item,)
        ).fetchone()
        routing = conn.execute(
            "SELECT COUNT(*) FROM routing_lines WHERE item_no=?", (item,)
        ).fetchone()
        print(f"{item}:")
        print(f"  type={prod[1] if prod else 'MANGLER'}, uom={prod[2] if prod else '-'}, "
              f"unit_cost={cost[0] if cost else 'INGEN'}, routing={routing[0] if routing else 0}")

    conn.close()


if __name__ == "__main__":
    main()