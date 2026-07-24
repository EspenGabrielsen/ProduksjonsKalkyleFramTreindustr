#!/usr/bin/env python3
"""
data_repo.py - SQLite-database for produksjonskalkylen.

Autoritativ datakilde for all stamdata. Erstatter Excel som master.
Excel brukes kun som import/eksport-format.

Tabeller:
  - Stamdata: products, locations, work_centers, operations, item_costs,
    bom_lines, routing_lines, byproduct_rules, capacity_days, production_scenarios
  - Infrastruktur: change_log, uploaded_files
"""

import os
import sqlite3
from contextlib import contextmanager
from datetime import date, datetime
from pathlib import Path
from typing import Optional


# ──────────────────────────────────────────────────────────────────────
#  SSO-brukerdeteksjon
# ──────────────────────────────────────────────────────────────────────

def get_current_user() -> Optional[str]:
    """Forsøk å identifisere bruker via SSO-proxy-headere/miljøvariabler.
    
    Sjekker følgende i prioritert rekkefølge:
      1. REMOTE_USER         (CGI-standard / oauth2-proxy)
      2. X_FORWARDED_USER    (Azure App Proxy, nginx)
      3. OIDC_CLAIM_preferred_username (Keycloak, Dex)
      4. HTTP_X_FORWARDED_USER (Marimo behind reverse proxy)
    """
    candidates = [
        os.environ.get("REMOTE_USER"),
        os.environ.get("X_FORWARDED_USER"),
        os.environ.get("OIDC_CLAIM_preferred_username"),
        os.environ.get("HTTP_X_FORWARDED_USER"),
    ]
    for c in candidates:
        if c:
            return c.strip()
    return None


# ──────────────────────────────────────────────────────────────────────
#  Databasehåndtering
# ──────────────────────────────────────────────────────────────────────

#DB_FILENAME = "endringslogg.db"
DB_FILENAME = "produksjonskalkyle_copy.db"


def _get_db_path(db_path: Optional[str] = None) -> str:
    """Finn stien til databasen. Default: ved siden av data_repo.py."""
    if db_path:
        return db_path
    return str(Path(__file__).parent / DB_FILENAME)


SCHEMA_SQL = """
-- Stamdata-tabeller (speiler dagens Excel-ark)

CREATE TABLE IF NOT EXISTS products (
    item_no TEXT PRIMARY KEY,
    description TEXT NOT NULL DEFAULT '',
    item_type TEXT NOT NULL DEFAULT '',
    product_group TEXT NOT NULL DEFAULT '',
    base_uom TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS locations (
    code TEXT PRIMARY KEY,
    name TEXT NOT NULL DEFAULT '',
    location_type TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS work_centers (
    code TEXT PRIMARY KEY,
    description TEXT NOT NULL DEFAULT '',
    location_code TEXT NOT NULL DEFAULT '',
    labor_cost_hour REAL NOT NULL DEFAULT 0,
    machine_cost_hour REAL NOT NULL DEFAULT 0,
    overhead_cost_hour REAL NOT NULL DEFAULT 0,
    capacity_hours_day REAL NOT NULL DEFAULT 0,
    effective_capacity_pct REAL NOT NULL DEFAULT 100
);

CREATE TABLE IF NOT EXISTS operations (
    code TEXT PRIMARY KEY,
    description TEXT NOT NULL DEFAULT '',
    default_work_center TEXT NOT NULL DEFAULT '',
    standard_unit TEXT NOT NULL DEFAULT 'Minutes'
);

CREATE TABLE IF NOT EXISTS item_costs (
    item_no TEXT NOT NULL,
    cost_type TEXT NOT NULL DEFAULT 'Standard Cost',
    unit_cost REAL NOT NULL DEFAULT 0,
    currency TEXT NOT NULL DEFAULT 'NOK',
    effective_date TEXT,
    PRIMARY KEY (item_no, cost_type)
);

CREATE TABLE IF NOT EXISTS bom_lines (
    parent_item_no TEXT NOT NULL,
    component_item_no TEXT NOT NULL,
    quantity_per REAL NOT NULL DEFAULT 1,
    uom TEXT NOT NULL DEFAULT '',
    scrap_pct REAL NOT NULL DEFAULT 0,
    co_product_pct REAL NOT NULL DEFAULT 0,
    co_product_item_no TEXT NOT NULL DEFAULT '',
    PRIMARY KEY (parent_item_no, component_item_no)
);

CREATE TABLE IF NOT EXISTS routing_lines (
    item_no TEXT NOT NULL,
    operation_no INTEGER NOT NULL,
    operation_code TEXT NOT NULL DEFAULT '',
    work_center_code TEXT NOT NULL DEFAULT '',
    setup_time_minutes REAL NOT NULL DEFAULT 0,
    run_time_minutes REAL NOT NULL DEFAULT 0,
    batch_size REAL NOT NULL DEFAULT 1,
    changeover_time_minutes REAL NOT NULL DEFAULT 0,
    PRIMARY KEY (item_no, operation_no, work_center_code)
);

CREATE TABLE IF NOT EXISTS byproduct_rules (
    parent_item_no TEXT NOT NULL,
    by_product_item_no TEXT NOT NULL,
    expected_quantity REAL NOT NULL DEFAULT 0,
    uom TEXT NOT NULL DEFAULT '',
    market_value REAL NOT NULL DEFAULT 0,
    allocation_method TEXT NOT NULL DEFAULT 'Reduce Main Product Cost',
    PRIMARY KEY (parent_item_no, by_product_item_no)
);

CREATE TABLE IF NOT EXISTS capacity_days (
    work_center TEXT NOT NULL,
    date TEXT NOT NULL,
    available_hours REAL NOT NULL DEFAULT 0,
    planned_downtime REAL NOT NULL DEFAULT 0,
    PRIMARY KEY (work_center, date)
);

CREATE TABLE IF NOT EXISTS production_scenarios (
    scenario_name TEXT NOT NULL,
    product TEXT NOT NULL,
    planned_quantity REAL NOT NULL DEFAULT 0,
    PRIMARY KEY (scenario_name, product)
);

-- Infrastruktur-tabeller

CREATE TABLE IF NOT EXISTS change_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp TEXT NOT NULL DEFAULT (datetime('now')),
    user TEXT,
    source TEXT NOT NULL DEFAULT 'web_form',
    table_name TEXT NOT NULL,
    record_key TEXT NOT NULL,
    field_name TEXT NOT NULL,
    old_value TEXT,
    new_value TEXT
);

CREATE TABLE IF NOT EXISTS uploaded_files (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    filename TEXT NOT NULL,
    uploaded_at TEXT NOT NULL DEFAULT (datetime('now')),
    blob BLOB,
    comment TEXT NOT NULL DEFAULT '',
    row_count INTEGER DEFAULT 0
);

-- Indekser for raskere oppslag i change_log
CREATE INDEX IF NOT EXISTS idx_change_log_table ON change_log(table_name);
CREATE INDEX IF NOT EXISTS idx_change_log_timestamp ON change_log(timestamp DESC);
CREATE INDEX IF NOT EXISTS idx_change_log_record ON change_log(table_name, record_key);

-- Trigger for automatisk timestamp ved INSERT på change_log (fallback)
CREATE TRIGGER IF NOT EXISTS trg_change_log_timestamp
    AFTER INSERT ON change_log
    WHEN new.timestamp IS NULL
BEGIN
    UPDATE change_log SET timestamp = datetime('now') WHERE id = new.id;
END;
"""


class DataRepo:
    """Hovedklasse for all databaseinteraksjon.
    
    Bruk som context manager:
        with DataRepo() as db:
            db.log_change(...)
    
    Eller opprett og lukk manuelt:
        db = DataRepo()
        db.initialize()
        ...
        db.close()
    """

    def __init__(self, db_path: Optional[str] = None):
        self.db_path = _get_db_path(db_path)
        self._conn: Optional[sqlite3.Connection] = None

    def __enter__(self):
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    # ── Tilkobling ───────────────────────────────────────────────

    def connect(self):
        """Åpne forbindelse til SQLite-databasen med WAL-mode."""
        if self._conn is not None:
            return
        self._conn = sqlite3.connect(self.db_path)
        self._conn.row_factory = sqlite3.Row
        self._conn.execute("PRAGMA journal_mode=WAL")
        self._conn.execute("PRAGMA foreign_keys=ON")

    def close(self):
        """Lukk forbindelsen."""
        if self._conn:
            self._conn.close()
            self._conn = None

    @property
    def conn(self) -> sqlite3.Connection:
        if self._conn is None:
            self.connect()
        assert self._conn is not None
        return self._conn

    # ── Migrering ────────────────────────────────────────────────

    def initialize(self):
        """Opprett alle tabeller hvis de ikke finnes."""
        self.conn.executescript(SCHEMA_SQL)
        self.conn.commit()
        # Migrer: legg til comment-kolonne hvis den ikke finnes
        try:
            self.conn.execute("ALTER TABLE uploaded_files ADD COLUMN comment TEXT NOT NULL DEFAULT ''")
        except sqlite3.OperationalError:
            pass  # kolonnen finnes allerede

    def is_empty(self) -> bool:
        """Sjekk om databasen har data (products-tabellen tom)."""
        cur = self.conn.execute("SELECT COUNT(*) FROM products")
        return cur.fetchone()[0] == 0

    def get_table_row_count(self, table_name: str) -> int:
        """Hent antall rader i en tabell."""
        cur = self.conn.execute(f"SELECT COUNT(*) FROM {table_name}")
        return cur.fetchone()[0]

    @property
    def stats(self) -> dict:
        """Returner en oversikt over antall rader i alle tabeller."""
        tables = [
            "products", "locations", "work_centers", "operations",
            "item_costs", "bom_lines", "routing_lines", "byproduct_rules",
            "capacity_days", "production_scenarios", "change_log", "uploaded_files",
        ]
        stats = {}
        for t in tables:
            try:
                stats[t] = self.get_table_row_count(t)
            except sqlite3.OperationalError:
                stats[t] = 0
        return stats

    # ── Endringslogg ─────────────────────────────────────────────

    def log_change(
        self,
        table_name: str,
        record_key: str,
        field_name: str,
        old_value: object,
        new_value: object,
        source: str = "web_form",
        user: Optional[str] = None,
    ):
        """Loggfør en endring i change_log.
        
        Args:
            table_name: Hvilken tabell ble endret (f.eks. "item_costs")
            record_key: Hvilken rad (f.eks. "RM001")
            field_name: Hvilket felt (f.eks. "unit_cost")
            old_value: Gammel verdi (konverteres til str)
            new_value: Ny verdi (konverteres til str)
            source: Hvor kom endringen fra ("web_form" eller "import")
            user: Bruker-id fra SSO (auto-detekteres hvis None)
        """
        if user is None:
            user = get_current_user()

        self.conn.execute(
            """INSERT INTO change_log (user, source, table_name, record_key, field_name, old_value, new_value)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (user, source, table_name, record_key, field_name,
             str(old_value) if old_value is not None else None,
             str(new_value) if new_value is not None else None),
        )
        self.conn.commit()

    def log_batch_changes(
        self,
        changes: list[dict],
        source: str = "import",
        user: Optional[str] = None,
    ):
        """Loggfør mange endringer på én gang (bulk-import).
        
        Args:
            changes: Liste med dicts som har nøkler:
                     table_name, record_key, field_name, old_value, new_value
            source: "import" eller "web_form"
            user: Bruker-id (auto hvis None)
        """
        if not changes:
            return
        if user is None:
            user = get_current_user()

        rows = []
        for c in changes:
            rows.append((
                user,
                source,
                c["table_name"],
                c["record_key"],
                c["field_name"],
                str(c["old_value"]) if c.get("old_value") is not None else None,
                str(c["new_value"]) if c.get("new_value") is not None else None,
            ))

        self.conn.executemany(
            """INSERT INTO change_log (user, source, table_name, record_key, field_name, old_value, new_value)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            rows,
        )
        self.conn.commit()

    def get_changes(
        self,
        table_name: Optional[str] = None,
        record_key: Optional[str] = None,
        limit: int = 50,
        offset: int = 0,
    ) -> list[sqlite3.Row]:
        """Hent endringer fra change_log.
        
        Args:
            table_name: Filtrer på tabell (optional)
            record_key: Filtrer på radnøkkel (optional)
            limit: Maks antall rader
            offset: Paginering
            
        Returns:
            Liste med sqlite3.Row-objekter
        """
        where_clauses = []
        params = []

        if table_name:
            where_clauses.append("table_name = ?")
            params.append(table_name)
        if record_key:
            where_clauses.append("record_key = ?")
            params.append(record_key)

        where = ""
        if where_clauses:
            where = "WHERE " + " AND ".join(where_clauses)

        query = f"SELECT * FROM change_log {where} ORDER BY id DESC LIMIT ? OFFSET ?"
        params.extend([limit, offset])

        return self.conn.execute(query, params).fetchall()

    def get_recent_changes(self, limit: int = 20) -> list[sqlite3.Row]:
        """Hent de siste endringene (for dashboard)."""
        return self.get_changes(limit=limit)

    # ── CRUD-hjelpere ────────────────────────────────────────────

    def upsert_products(self, products: list[dict], source: str = "web_form"):
        """Oppdater eller sett inn produkter. Logger endringer."""
        for p in products:
            key = p.get("item_no", "")
            existing = self.conn.execute(
                "SELECT * FROM products WHERE item_no = ?", (key,)
            ).fetchone()

            if existing:
                # Sammenlign felt-for-felt og loggfør endringer
                fields = {
                    "description": (str, ""),
                    "item_type": (str, ""),
                    "product_group": (str, ""),
                    "base_uom": (str, ""),
                }
                for field, (ftype, default) in fields.items():
                    old = existing[field]
                    new = ftype(p.get(field, default))
                    if old != new:
                        self.log_change(
                            "products", key, field, old, new,
                            source=source,
                        )

            # Upsert: INSERT OR REPLACE
            self.conn.execute(
                """INSERT OR REPLACE INTO products
                   (item_no, description, item_type, product_group, base_uom)
                   VALUES (?, ?, ?, ?, ?)""",
                (
                    p.get("item_no", ""),
                    p.get("description", ""),
                    p.get("item_type", ""),
                    p.get("product_group", ""),
                    p.get("base_uom", ""),
                ),
            )
        self.conn.commit()

    def upsert_work_centers(self, wcs: list[dict], source: str = "web_form"):
        """Oppdater eller sett inn arbeidssentre. Logger endringer."""
        for wc in wcs:
            key = wc.get("code", "")
            existing = self.conn.execute(
                "SELECT * FROM work_centers WHERE code = ?", (key,)
            ).fetchone()

            if existing:
                fields = {
                    "description": (str, ""),
                    "location_code": (str, ""),
                    "labor_cost_hour": (float, 0),
                    "machine_cost_hour": (float, 0),
                    "overhead_cost_hour": (float, 0),
                    "capacity_hours_day": (float, 0),
                    "effective_capacity_pct": (float, 100),
                }
                for field, (ftype, default) in fields.items():
                    old = existing[field]
                    new = ftype(wc.get(field, default))
                    if old != new:
                        self.log_change(
                            "work_centers", key, field, old, new,
                            source=source,
                        )

            self.conn.execute(
                """INSERT OR REPLACE INTO work_centers
                   (code, description, location_code, labor_cost_hour, machine_cost_hour,
                    overhead_cost_hour, capacity_hours_day, effective_capacity_pct)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                (
                    wc.get("code", ""),
                    wc.get("description", ""),
                    wc.get("location_code", ""),
                    float(wc.get("labor_cost_hour", 0)),
                    float(wc.get("machine_cost_hour", 0)),
                    float(wc.get("overhead_cost_hour", 0)),
                    float(wc.get("capacity_hours_day", 0)),
                    float(wc.get("effective_capacity_pct", 100)),
                ),
            )
        self.conn.commit()

    def upsert_item_costs(self, costs: list[dict], source: str = "web_form"):
        """Oppdater eller sett inn kostpriser. Logger endringer."""
        for c in costs:
            key = f"{c.get('item_no', '')}:{c.get('cost_type', 'Standard Cost')}"
            existing = self.conn.execute(
                """SELECT * FROM item_costs
                   WHERE item_no = ? AND cost_type = ?""",
                (c.get("item_no", ""), c.get("cost_type", "Standard Cost")),
            ).fetchone()

            if existing:
                old = existing["unit_cost"]
                new = float(c.get("unit_cost", 0))
                if abs(old - new) > 0.001:
                    self.log_change(
                        "item_costs", c.get("item_no", ""),
                        "unit_cost", old, new, source=source,
                    )

            self.conn.execute(
                """INSERT OR REPLACE INTO item_costs
                   (item_no, cost_type, unit_cost, currency, effective_date)
                   VALUES (?, ?, ?, ?, ?)""",
                (
                    c.get("item_no", ""),
                    c.get("cost_type", "Standard Cost"),
                    float(c.get("unit_cost", 0)),
                    c.get("currency", "NOK"),
                    c.get("effective_date"),
                ),
            )
        self.conn.commit()

    def upsert_bom_lines(self, lines: list[dict], source: str = "web_form"):
        """Oppdater eller sett inn BOM-linjer. Logger endringer."""
        for bl in lines:
            parent = bl.get("parent_item_no", "")
            component = bl.get("component_item_no", "")
            key = f"{parent}:{component}"
            existing = self.conn.execute(
                """SELECT * FROM bom_lines
                   WHERE parent_item_no = ? AND component_item_no = ?""",
                (parent, component),
            ).fetchone()

            if existing:
                for field in ("quantity_per", "scrap_pct", "co_product_pct"):
                    old = existing[field]
                    new = float(bl.get(field, 0))
                    if abs(old - new) > 0.001:
                        self.log_change(
                            "bom_lines", key, field, old, new, source=source,
                        )
            else:
                # Ny rad
                self.log_change(
                    "bom_lines", key, "_created", None, component, source=source,
                )

            self.conn.execute(
                """INSERT OR REPLACE INTO bom_lines
                   (parent_item_no, component_item_no, quantity_per, uom,
                    scrap_pct, co_product_pct, co_product_item_no)
                   VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (
                    parent, component,
                    float(bl.get("quantity_per", 1)),
                    bl.get("uom", ""),
                    float(bl.get("scrap_pct", 0)),
                    float(bl.get("co_product_pct", 0)),
                    bl.get("co_product_item_no", ""),
                ),
            )
        self.conn.commit()

    def upsert_routing_lines(self, lines: list[dict], source: str = "web_form"):
        """Oppdater eller sett inn routing-linjer. Logger endringer."""
        for rl in lines:
            item = rl.get("item_no", "")
            op_no = int(rl.get("operation_no", 0))
            wc = rl.get("work_center_code", "")
            key = f"{item}:{op_no}:{wc}"
            existing = self.conn.execute(
                """SELECT * FROM routing_lines
                   WHERE item_no = ? AND operation_no = ? AND work_center_code = ?""",
                (item, op_no, wc),
            ).fetchone()

            if existing:
                for field in ("run_time_minutes", "setup_time_minutes", "batch_size"):
                    old = existing[field]
                    new = float(rl.get(field, 0))
                    if abs(old - new) > 0.001:
                        self.log_change(
                            "routing_lines", key, field, old, new, source=source,
                        )
            else:
                self.log_change(
                    "routing_lines", key, "_created", None, item, source=source,
                )

            self.conn.execute(
                """INSERT OR REPLACE INTO routing_lines
                   (item_no, operation_no, operation_code, work_center_code,
                    setup_time_minutes, run_time_minutes, batch_size,
                    changeover_time_minutes)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                (
                    item, op_no,
                    rl.get("operation_code", ""),
                    wc,
                    float(rl.get("setup_time_minutes", 0)),
                    float(rl.get("run_time_minutes", 0)),
                    float(rl.get("batch_size", 1)),
                    float(rl.get("changeover_time_minutes", 0)),
                ),
            )
        self.conn.commit()

    # ── Bulk-oppdatering (for enkelthets skyld) ──────────────────

    def upsert_locations(self, locations: list[dict], source: str = "web_form"):
        """Erstatt alle locations (enkel import)."""
        for loc in locations:
            self.conn.execute(
                """INSERT OR REPLACE INTO locations
                   (code, name, location_type)
                   VALUES (?, ?, ?)""",
                (
                    loc.get("code", ""),
                    loc.get("name", ""),
                    loc.get("location_type", ""),
                ),
            )
        self.conn.commit()

    def upsert_operations(self, operations: list[dict], source: str = "web_form"):
        """Erstatt alle operations (enkel import)."""
        for op in operations:
            self.conn.execute(
                """INSERT OR REPLACE INTO operations
                   (code, description, default_work_center, standard_unit)
                   VALUES (?, ?, ?, ?)""",
                (
                    op.get("code", ""),
                    op.get("description", ""),
                    op.get("default_work_center", ""),
                    op.get("standard_unit", "Minutes"),
                ),
            )
        self.conn.commit()

    def upsert_byproduct_rules(self, rules: list[dict], source: str = "web_form"):
        """Erstatt alle byproduct rules (enkel import)."""
        for r in rules:
            self.conn.execute(
                """INSERT OR REPLACE INTO byproduct_rules
                   (parent_item_no, by_product_item_no, expected_quantity, uom,
                    market_value, allocation_method)
                   VALUES (?, ?, ?, ?, ?, ?)""",
                (
                    r.get("parent_item_no", ""),
                    r.get("by_product_item_no", ""),
                    float(r.get("expected_quantity", 0)),
                    r.get("uom", ""),
                    float(r.get("market_value", 0)),
                    r.get("allocation_method", "Reduce Main Product Cost"),
                ),
            )
        self.conn.commit()

    def upsert_capacity_days(self, days: list[dict], source: str = "web_form"):
        """Erstatt alle capacity days (enkel import)."""
        for d in days:
            self.conn.execute(
                """INSERT OR REPLACE INTO capacity_days
                   (work_center, date, available_hours, planned_downtime)
                   VALUES (?, ?, ?, ?)""",
                (
                    d.get("work_center", ""),
                    d.get("date", ""),
                    float(d.get("available_hours", 0)),
                    float(d.get("planned_downtime", 0)),
                ),
            )
        self.conn.commit()

    def upsert_scenarios(self, scenarios: list[dict], source: str = "web_form"):
        """Erstatt alle production scenarios (enkel import)."""
        for sc in scenarios:
            self.conn.execute(
                """INSERT OR REPLACE INTO production_scenarios
                   (scenario_name, product, planned_quantity)
                   VALUES (?, ?, ?)""",
                (
                    sc.get("scenario_name", ""),
                    sc.get("product", ""),
                    float(sc.get("planned_quantity", 0)),
                ),
            )
        self.conn.commit()

    # ── Tømming og tilbakestilling ──────────────────────────────

    def clear_all_data(self):
        """Slett all data fra alle stamdata-tabeller (bevar change_log)."""
        tables = [
            "products", "locations", "work_centers", "operations",
            "item_costs", "bom_lines", "routing_lines", "byproduct_rules",
            "capacity_days", "production_scenarios",
        ]
        for t in tables:
            self.conn.execute(f"DELETE FROM {t}")
        self.conn.commit()

    def clear_change_log(self):
        """Slett alle oppføringer i change_log."""
        self.conn.execute("DELETE FROM change_log")
        self.conn.commit()

    # ── Full eksport (for å bygge Excel) ─────────────────────────

    def get_last_uploads(self, limit: int = 10) -> list[dict]:
        """Hent de sist N opplastede filene (uten blob).

        Args:
            limit: Maks antall filer å returnere (default 10)

        Returns:
            Liste med dicts: id, filename, uploaded_at, comment
        """
        rows = self.conn.execute(
            """SELECT id, filename, uploaded_at, comment
               FROM uploaded_files
               ORDER BY id DESC
               LIMIT ?""",
            (limit,),
        ).fetchall()
        return [dict(r) for r in rows]

    def get_upload_blob(self, upload_id: int) -> Optional[bytes]:
        """Hent blob for en spesifikk opplastet fil.

        Args:
            upload_id: ID fra uploaded_files-tabellen

        Returns:
            Blob-data (bytes) eller None hvis ikke funnet
        """
        row = self.conn.execute(
            "SELECT blob FROM uploaded_files WHERE id = ?",
            (upload_id,),
        ).fetchone()
        if row and row["blob"]:
            return row["blob"]
        return None

    def save_upload(
        self,
        filename: str,
        blob: bytes,
        comment: str = "",
        row_count: int = 0,
    ) -> int:
        """Lagre en opplastet fil i uploaded_files-tabellen.

        Args:
            filename: Filnavn
            blob: Innholdet av filen (.xlsx)
            comment: Brukerens kommentar
            row_count: Antall rader/endringer

        Returns:
            ID-en til den nye raden
        """
        cur = self.conn.execute(
            """INSERT INTO uploaded_files (filename, blob, comment, row_count)
               VALUES (?, ?, ?, ?)""",
            (filename, blob, comment, row_count),
        )
        self.conn.commit()
        # cur.lastrowid kan være None på INSERT OR REPLACE, men vi har plain INSERT
        result = cur.lastrowid
        return result if result is not None else 0

    def export_all_data(self) -> dict[str, list[sqlite3.Row]]:
        """Eksporter all data fra alle tabeller.
        
        Returnerer dict med tabellnavn som nøkler.
        """
        tables = {
            "products": "SELECT * FROM products ORDER BY item_no",
            "locations": "SELECT * FROM locations ORDER BY code",
            "work_centers": "SELECT * FROM work_centers ORDER BY code",
            "operations": "SELECT * FROM operations ORDER BY code",
            "item_costs": "SELECT * FROM item_costs ORDER BY item_no, cost_type",
            "bom_lines": "SELECT * FROM bom_lines ORDER BY parent_item_no, component_item_no",
            "routing_lines": "SELECT * FROM routing_lines ORDER BY item_no, operation_no",
            "byproduct_rules": "SELECT * FROM byproduct_rules ORDER BY parent_item_no, by_product_item_no",
            "capacity_days": "SELECT * FROM capacity_days ORDER BY work_center, date",
            "production_scenarios": "SELECT * FROM production_scenarios ORDER BY scenario_name, product",
        }
        result = {}
        for name, query in tables.items():
            result[name] = self.conn.execute(query).fetchall()
        return result


# ── Kommandolinje-hjelper ─────────────────────────────────────────

def main():
    """Kjøres frittstående: python data_repo.py [--stats] [--clear] [--changes]"""
    import argparse

    parser = argparse.ArgumentParser(description="Administrasjon av endringslogg-databasen")
    parser.add_argument("--stats", action="store_true", help="Vis statistikk over alle tabeller")
    parser.add_argument("--clear", action="store_true", help="Tøm all data (bevar endringslogg)")
    parser.add_argument("--clear-log", action="store_true", help="Tøm endringsloggen")
    parser.add_argument("--changes", type=int, nargs="?", const=20, help="Vis siste N endringer")
    parser.add_argument("--init", action="store_true", help="Initialiser databasen (opprett tabeller)")

    args = parser.parse_args()

    db = DataRepo()
    db.initialize()

    if args.init:
        print(f"[OK] Database initialisert: {db.db_path}")
        for table, count in db.stats.items():
            print(f"   {table}: {count} rader")

    if args.stats:
        print(f"[DB] Database: {db.db_path}")
        for table, count in db.stats.items():
            print(f"   {table}: {count} rader")

    if args.clear:
        confirm = input("[ADV] T\u00f8mme all data? (ja/nei): ")
        if confirm.lower() in ("ja", "yes", "y"):
            db.clear_all_data()
            print("[OK] All data slettet")
        else:
            print("Avbrutt")

    if args.clear_log:
        confirm = input("[ADV] T\u00f8mme endringsloggen? (ja/nei): ")
        if confirm.lower() in ("ja", "yes", "y"):
            db.clear_change_log()
            print("[OK] Endringslogg slettet")
        else:
            print("Avbrutt")

    if args.changes is not None:
        changes = db.get_changes(limit=args.changes)
        print(f"[LOG] Siste {len(changes)} endringer:")
        print(f"{'ID':>4} {'Dato':<20} {'Bruker':<20} {'Tabell':<18} {'Nokkel':<25} {'Felt':<20} {'Gammel':<15} {'Ny':<15}")
        print("-" * 140)
        for c in changes:
            print(f"{c['id']:>4} {c['timestamp']:<20} {(c['user'] or '-'):<20} "
                  f"{c['table_name']:<18} {c['record_key']:<25} {c['field_name']:<20} "
                  f"{(c['old_value'] or '-'):<15} {(c['new_value'] or '-'):<15}")

    if not any([args.stats, args.clear, args.clear_log, args.changes is not None, args.init]):
        parser.print_help()


if __name__ == "__main__":
    main()