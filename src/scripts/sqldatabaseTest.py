import os
import sqlite3
from pathlib import Path

# Databasen ligger nå i src/ (ved siden av data_repo.py)
_db_path = Path(__file__).resolve().parents[1] / "produksjonskalkyle.db"
conn = sqlite3.connect(str(_db_path))

#resp = conn.execute("select * from products where item_no = 'JD29198'")
#resp = conn.execute("select * from routing_lines where item_no = 'JD29198'")
resp = conn.execute("SELECT * from work_centers")

for line in resp:
    print(line)

conn.close()

