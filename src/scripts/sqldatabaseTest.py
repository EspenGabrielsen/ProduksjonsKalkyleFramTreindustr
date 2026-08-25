import os
import sqlite3
from pathlib import Path

from data_repo import _get_db_path

# Bruk samme path-resolusjon som applikasjonen
_db_path = _get_db_path()
conn = sqlite3.connect(str(_db_path))

#resp = conn.execute("select * from products where item_no = 'JD29198'")
#resp = conn.execute("select * from routing_lines where item_no = 'JD29198'")
resp = conn.execute("SELECT * from work_centers")

for line in resp:
    print(line)

conn.close()

