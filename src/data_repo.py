#!/usr/bin/env python3
"""
data_repo.py - SQLite-database for produksjonskalkylen.

Autoritativ datakilde for all stamdata. Erstatter Excel som master.
Excel brukes kun som import/eksport-format.

Tabeller:
  - Stamdata: products, locations, work_centers, operations, item_costs,
    bom_lines, routing_lines, byproduct_rules, capacity_days, production_scenarios
  - Optimering: demand (sluttetterspørsel), changeover_matrix (omstillingstid mellom produktfamilier)
  - Infrastruktur: change_log, uploaded_files

Alle tabeller har id INTEGER PRIMARY KEY AUTOINCREMENT, med unike constraints
for å bevare integriteten til de naturlige nøklene.
"""

import base64
import json
import os
import sqlite3
from contextlib import contextmanager
from datetime import date, datetime
from pathlib import Path
from typing import Optional


# ──────────────────────────────────────────────────────────────────────
#  SSO-brukerdeteksjon
# ──────────────────────────────────────────────────────────────────────

_CURRENT_USER: Optional[str] = None
_CURRENT_TENANT_ID: Optional[str] = None


def set_current_user(user: Optional[str]) -> None:
    """Lagre gjeldende bruker for aktiv Marimo-kernel/økt."""
    global _CURRENT_USER
    _CURRENT_USER = user.strip() if isinstance(user, str) and user.strip() else None


def set_current_tenant_id(tenant_id: Optional[str]) -> None:
    """Lagre tenant-id for aktiv Marimo-kernel/økt."""
    global _CURRENT_TENANT_ID
    _CURRENT_TENANT_ID = tenant_id.strip() if isinstance(tenant_id, str) and tenant_id.strip() else None


def _decode_easy_auth_claims(encoded_principal: str) -> dict:
    """Dekod Azure EasyAuth X-MS-CLIENT-PRINCIPAL til claims-dict."""
    if not encoded_principal:
        return {}
    try:
        padded = encoded_principal + "=" * (-len(encoded_principal) % 4)
        payload = base64.b64decode(padded).decode("utf-8")
        data = json.loads(payload)
    except Exception:
        return {}

    claims = data.get("claims", []) if isinstance(data, dict) else []
    result = {}
    for claim in claims:
        if not isinstance(claim, dict):
            continue
        typ = str(claim.get("typ", "")).strip()
        val = claim.get("val")
        if typ and val is not None and typ not in result:
            result[typ] = str(val).strip()
    return result


def user_from_headers(headers: Optional[dict]) -> Optional[str]:
    """Forsøk å finne innlogget bruker fra reverse-proxy/EasyAuth-headere."""
    if not headers:
        return None

    normalized = {str(k).lower(): str(v) for k, v in headers.items() if v is not None}
    direct_candidates = [
        normalized.get("x-ms-client-principal-name"),
        normalized.get("x-forwarded-user"),
        normalized.get("remote-user"),
        normalized.get("oidc-claim-preferred_username"),
    ]
    for candidate in direct_candidates:
        if candidate and candidate.strip():
            return candidate.strip()

    claims = _decode_easy_auth_claims(normalized.get("x-ms-client-principal", ""))
    claim_candidates = [
        claims.get("preferred_username"),
        claims.get("http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn"),
        claims.get("upn"),
        claims.get("email"),
        claims.get("name"),
    ]
    for candidate in claim_candidates:
        if candidate and candidate.strip():
            return candidate.strip()
    return None


def tenant_id_from_headers(headers: Optional[dict]) -> Optional[str]:
    """Forsøk å hente tenant-id fra Azure EasyAuth-headere."""
    if not headers:
        return None

    normalized = {str(k).lower(): str(v) for k, v in headers.items() if v is not None}
    claims = _decode_easy_auth_claims(normalized.get("x-ms-client-principal", ""))
    claim_candidates = [
        claims.get("tid"),
        claims.get("http://schemas.microsoft.com/identity/claims/tenantid"),
    ]
    for candidate in claim_candidates:
        if candidate and candidate.strip():
            return candidate.strip()
    return None


def is_allowed_tenant(tenant_id: Optional[str]) -> bool:
    """Valider tenant-id mot ALLOWED_TENANT_ID hvis satt.

    Hvis miljøvariabelen ikke er satt, tillates alle tenants som allerede er
    sluppet gjennom EasyAuth. Dette gjør funksjonen bakoverkompatibel lokalt.
    """
    allowed_tenant_id = os.environ.get("ALLOWED_TENANT_ID", "").strip()
    if not allowed_tenant_id:
        return True
    if not tenant_id:
        return False
    return tenant_id.strip().lower() == allowed_tenant_id.lower()


def get_current_tenant_id() -> Optional[str]:
    """Hent tenant-id for aktiv økt eller miljøvariabel."""
    if _CURRENT_TENANT_ID:
        return _CURRENT_TENANT_ID
    return os.environ.get("AZURE_TENANT_ID") or os.environ.get("ALLOWED_TENANT_ID")

def get_current_user() -> Optional[str]:
    """Forsøk å identifisere bruker via SSO-proxy-headere/miljøvariabler.
    
    Sjekker følgende i prioritert rekkefølge:
      1. set_current_user() / aktiv Marimo-økt
      2. X_MS_CLIENT_PRINCIPAL_NAME / HTTP_X_MS_CLIENT_PRINCIPAL_NAME (Azure EasyAuth)
      3. REMOTE_USER / X_FORWARDED_USER / HTTP_X_FORWARDED_USER
      4. OIDC_CLAIM_preferred_username (Keycloak, Dex)
    """
    if _CURRENT_USER:
        return _CURRENT_USER

    candidates = [
        os.environ.get("X_MS_CLIENT_PRINCIPAL_NAME"),
        os.environ.get("HTTP_X_MS_CLIENT_PRINCIPAL_NAME"),
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

DB_FILENAME = "produksjonskalkyle.db"


def _get_db_path(db_path: Optional[str] = None) -> str:
    """Finn stien til databasen.

    Spesiell logikk for test-modus:
      - PRODUKSJONSKALKYLE_TEST=true → midlertidig fil-database i temp-mappe.
        Bruker en EKTE fil (ikke :memory:) slik at DataRepo og SqliteData
        (som har to separate tilkoblinger) deler samme data.
      - Ellers: ved siden av data_repo.py

    Args:
        db_path: Eksplisitt sti (valgfri) — prioriteres før test-modus
    """
    if db_path:
        return db_path
    if os.environ.get("PRODUKSJONSKALKYLE_TEST", "").lower() in ("true", "1", "yes"):
        import tempfile
        return os.path.join(tempfile.gettempdir(), "produksjonskalkyle_test.db")
    return str(Path(__file__).parent / DB_FILENAME)


SCHEMA_SQL = """
-- Stamdata-tabeller (speiler dagens Excel-ark)
-- Alle tabeller har id INTEGER PRIMARY KEY AUTOINCREMENT + UNIQUE constraints

CREATE TABLE IF NOT EXISTS products (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    item_no TEXT NOT NULL UNIQUE,
    description TEXT NOT NULL DEFAULT '',
    item_type TEXT NOT NULL DEFAULT '',
    product_group TEXT NOT NULL DEFAULT '',
    base_uom TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS locations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT NOT NULL UNIQUE,
    name TEXT NOT NULL DEFAULT '',
    location_type TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS work_centers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT NOT NULL UNIQUE,
    description TEXT NOT NULL DEFAULT '',
    location_code TEXT NOT NULL DEFAULT '',
    labor_cost_hour REAL NOT NULL DEFAULT 0,
    machine_cost_hour REAL NOT NULL DEFAULT 0,
    overhead_cost_hour REAL NOT NULL DEFAULT 0,
    capacity_hours_day REAL NOT NULL DEFAULT 0,
    effective_capacity_pct REAL NOT NULL DEFAULT 100
);

CREATE TABLE IF NOT EXISTS operations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT NOT NULL UNIQUE,
    description TEXT NOT NULL DEFAULT '',
    default_work_center TEXT NOT NULL DEFAULT '',
    standard_unit TEXT NOT NULL DEFAULT 'Minutes'
);

CREATE TABLE IF NOT EXISTS item_costs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    item_no TEXT NOT NULL,
    cost_type TEXT NOT NULL DEFAULT 'Standard Cost',
    unit_cost REAL NOT NULL DEFAULT 0,
    currency TEXT NOT NULL DEFAULT 'NOK',
    effective_date TEXT,
    UNIQUE(item_no, cost_type)
);

CREATE TABLE IF NOT EXISTS bom_lines (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    parent_item_no TEXT NOT NULL,
    component_item_no TEXT NOT NULL,
    quantity_per REAL NOT NULL DEFAULT 1,
    uom TEXT NOT NULL DEFAULT '',
    scrap_pct REAL NOT NULL DEFAULT 0,
    co_product_pct REAL NOT NULL DEFAULT 0,
    co_product_item_no TEXT NOT NULL DEFAULT '',
    UNIQUE(parent_item_no, component_item_no)
);

CREATE TABLE IF NOT EXISTS routing_lines (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    item_no TEXT NOT NULL,
    operation_no INTEGER NOT NULL,
    operation_code TEXT NOT NULL DEFAULT '',
    work_center_code TEXT NOT NULL DEFAULT '',
    setup_time_minutes REAL NOT NULL DEFAULT 0,
    run_time_minutes REAL NOT NULL DEFAULT 0,
    batch_size REAL NOT NULL DEFAULT 1,
    changeover_time_minutes REAL NOT NULL DEFAULT 0,
    UNIQUE(item_no, operation_no, work_center_code)
);

CREATE TABLE IF NOT EXISTS byproduct_rules (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    parent_item_no TEXT NOT NULL,
    by_product_item_no TEXT NOT NULL,
    expected_quantity REAL NOT NULL DEFAULT 0,
    uom TEXT NOT NULL DEFAULT '',
    market_value REAL NOT NULL DEFAULT 0,
    allocation_method TEXT NOT NULL DEFAULT 'Reduce Main Product Cost',
    UNIQUE(parent_item_no, by_product_item_no)
);

CREATE TABLE IF NOT EXISTS capacity_days (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    work_center TEXT NOT NULL,
    date TEXT NOT NULL,
    available_hours REAL NOT NULL DEFAULT 0,
    planned_downtime REAL NOT NULL DEFAULT 0,
    UNIQUE(work_center, date)
);

CREATE TABLE IF NOT EXISTS production_scenarios (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    scenario_name TEXT NOT NULL,
    product TEXT NOT NULL,
    planned_quantity REAL NOT NULL DEFAULT 0,
    UNIQUE(scenario_name, product)
);

-- Transportflagg for fler-høvleri-produksjon (transportvare-modul)
CREATE TABLE IF NOT EXISTS transport_flagg (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    item_no TEXT NOT NULL UNIQUE,
    is_transport INTEGER NOT NULL DEFAULT 0,
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- Transportruter mellom høvlerier (from → to, med kost per M3 som eneste beregningsfelt)
CREATE TABLE IF NOT EXISTS transport_ruter (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    from_loc TEXT NOT NULL,
    to_loc TEXT NOT NULL,
    cost_per_m3 REAL NOT NULL DEFAULT 0,
    distance_km REAL NOT NULL DEFAULT 0,
    hours REAL NOT NULL DEFAULT 0,
    UNIQUE(from_loc, to_loc)
);

-- Optimerings-tabeller (brukes av optimization_engine.py — MILP-optimeringsmotor)
-- Eksisterende kode/app rører ALDRI disse tabellene.

-- Sluttetterspørsel: hva kunden/markedet har bestilt per periode
CREATE TABLE IF NOT EXISTS demand (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    product_id TEXT NOT NULL,
    period INTEGER NOT NULL,
    quantity REAL NOT NULL DEFAULT 0,
    location_code TEXT NOT NULL,
    customer_region TEXT NOT NULL DEFAULT '',
    UNIQUE(product_id, period, location_code)
);

-- Historisk salg: brukes som prognose for batch-størrelsesvalg.
-- Modellen slår sammen faktisk demand (åpne ordrer) med historisk
-- salgsmønster for å forutse fremtidig etterspørsel per produkt×uke.
CREATE TABLE IF NOT EXISTS historical_sales (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    product_id TEXT NOT NULL,
    period INTEGER NOT NULL,
    quantity REAL NOT NULL DEFAULT 0,
    location_code TEXT NOT NULL,
    UNIQUE(product_id, period, location_code)
);

-- Omstillingstid mellom produktfamilier per arbeidssenter.
-- Familie = FTI prefiks+siffer (f.eks. 'JD19073'), dvs. samme dimensjon/profil.
-- Kun TID lagres her — kostnad beregnes ALLTID i koden:
--   changeover_cost = (changeover_minutes / 60) × work_centers.total_cost_hour
-- Dette unngår dobbelt vedlikehold av kostnadsdata.
-- Fallback-generator i kode bruker FTI-nummerstrukturen (se CLINE.md) når
-- denne tabellen er tom: suffiks-bytte ~0min, bredde-bytte 15-30min,
-- tykkelse-bytte 45-60min, prefiks-bytte ~90min.
CREATE TABLE IF NOT EXISTS changeover_matrix (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    work_center_code TEXT NOT NULL,
    from_family TEXT NOT NULL,
    to_family TEXT NOT NULL,
    changeover_minutes REAL NOT NULL DEFAULT 0,
    UNIQUE(work_center_code, from_family, to_family)
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
        """Opprett alle tabeller hvis de ikke finnes.
        
        Hvis det finnes en database med gammelt skjema (uten id-kolonne),
        vil den gamle databasen slettes og en ny opprettes.
        """
        # Sjekk om gammelt skjema finnes (products uten id-kolonne)
        try:
            test = self.conn.execute("SELECT id FROM products LIMIT 1").fetchone()
            # id-kolonne finnes → nytt skjema
        except sqlite3.OperationalError:
            # Gammelt skjema → steng tilkobling, slett DB-fil, opprett på nytt
            try:
                if self._conn:
                    self._conn.close()
                    self._conn = None
                if os.path.exists(self.db_path):
                    os.remove(self.db_path)
                    print(f"[*] Gammel database slettet: {self.db_path}")
            except PermissionError:
                # Kan hende WAL-filer eller -shm/-wal eksisterer; prøv å fjerne dem også
                for ext in ('', '-wal', '-shm'):
                    f = self.db_path + ext
                    if os.path.exists(f):
                        try:
                            os.remove(f)
                        except PermissionError:
                            print(f"[!] Kunne ikke slette: {f}")
            self.connect()

        self.conn.executescript(SCHEMA_SQL)
        self.conn.commit()

        # Migrering: oppgrader transport_ruter til ny struktur (cost_per_m3).
        # Gammel struktur hadde run_time_minutes/setup_time_minutes/batch_size.
        # Ny struktur: cost_per_m3 (eneste beregningsfelt) + distance_km/hours (info).
        try:
            cols = [r["name"] for r in self.conn.execute(
                "PRAGMA table_info(transport_ruter)").fetchall()]
            if "run_time_minutes" in cols:
                # Drop og bygg på nytt - gamle ruter må tastes inn på nytt med cost_per_m3
                self.conn.execute("DROP TABLE transport_ruter")
                self.conn.execute("""
                    CREATE TABLE transport_ruter (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        from_loc TEXT NOT NULL,
                        to_loc TEXT NOT NULL,
                        cost_per_m3 REAL NOT NULL DEFAULT 0,
                        distance_km REAL NOT NULL DEFAULT 0,
                        hours REAL NOT NULL DEFAULT 0,
                        UNIQUE(from_loc, to_loc)
                    )
                """)
                self.conn.commit()
                print("[*] transport_ruter migrert til ny struktur (cost_per_m3)")
        except sqlite3.OperationalError:
            pass  # tabellen finnes ikke (forste kjoring) - SCHEMA_SQL har allerede opprettet den

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
            "capacity_days", "production_scenarios",
            "demand", "historical_sales", "changeover_matrix",
            "transport_flagg", "transport_ruter",
            "change_log", "uploaded_files",
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
            record_key: Radens ID (string av id integer) eller tekst-nøkkel
            field_name: Hvilket felt (f.eks. "unit_cost"), eller "_deleted" for sletting
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
            (user, source, table_name, str(record_key), field_name,
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
                str(c["record_key"]),
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

    def upsert_products(self, products: list[dict], source: str = "web_form") -> list[int]:
        """Oppdater eller sett inn produkter. Logger endringer.
        
        Args:
            products: Liste med dicts. Hver dict kan ha 'id' (for UPDATE) eller ikke (for CREATE).
            source: Kilden for endringen

        Returns:
            Liste med id-er for de innsatte/oppdaterte radene
        """
        result_ids = []
        for p in products:
            row_id = p.get("id")
            if row_id:
                # UPDATE: match på id
                existing = self.conn.execute(
                    "SELECT * FROM products WHERE id = ?", (row_id,)
                ).fetchone()
                if existing:
                    fields = {
                        "description": (str, ""),
                        "item_type": (str, ""),
                        "product_group": (str, ""),
                        "base_uom": (str, ""),
                    }
                    # Også sjekk item_no (kan endres, men UNIQUE constraint sørger for integritet)
                    item_no_change = False
                    new_item_no = str(p.get("item_no", ""))
                    if new_item_no and new_item_no != existing["item_no"]:
                        self.log_change(
                            "products", str(row_id), "item_no",
                            existing["item_no"], new_item_no, source=source,
                        )
                        item_no_change = True

                    for field, (ftype, default) in fields.items():
                        old = existing[field]
                        new = ftype(p.get(field, default))
                        if old != new:
                            self.log_change(
                                "products", str(row_id), field, old, new,
                                source=source,
                            )

                    self.conn.execute(
                        """UPDATE products SET
                           item_no = ?, description = ?, item_type = ?,
                           product_group = ?, base_uom = ?
                           WHERE id = ?""",
                        (
                            p.get("item_no", existing["item_no"]),
                            p.get("description", ""),
                            p.get("item_type", ""),
                            p.get("product_group", ""),
                            p.get("base_uom", ""),
                            row_id,
                        ),
                    )
                    result_ids.append(row_id)
                else:
                    # id oppgitt men finnes ikke → INSERT
                    result_ids.append(self._insert_product(p, source))
            else:
                # CREATE: ny rad
                result_ids.append(self._insert_product(p, source))
        self.conn.commit()
        return result_ids

    def _insert_product(self, p: dict, source: str) -> int:
        """INSERT en ny product-rad og returner id."""
        cur = self.conn.execute(
            """INSERT INTO products (item_no, description, item_type, product_group, base_uom)
               VALUES (?, ?, ?, ?, ?)
               ON CONFLICT(item_no) DO UPDATE SET
                   description = excluded.description,
                   item_type = excluded.item_type,
                   product_group = excluded.product_group,
                   base_uom = excluded.base_uom
               RETURNING id""",
            (
                p.get("item_no", ""),
                p.get("description", ""),
                p.get("item_type", ""),
                p.get("product_group", ""),
                p.get("base_uom", ""),
            ),
        )
        row = cur.fetchone()
        new_id = row["id"] if row else 0
        # Logg opprettelse
        self.log_change(
            "products", str(new_id), "_created", None,
            f"{p.get('item_no', '')}: {p.get('description', '')}",
            source=source,
        )
        return new_id

    def delete_product(self, product_id: int, source: str = "web_form"):
        """Slett et produkt. Logger full rad i change_log for reversering."""
        existing = self.conn.execute(
            "SELECT * FROM products WHERE id = ?", (product_id,)
        ).fetchone()
        if existing:
            self.log_change(
                "products", str(product_id), "_deleted",
                json.dumps(dict(existing), ensure_ascii=False), None,
                source=source,
            )
            self.conn.execute("DELETE FROM products WHERE id = ?", (product_id,))
            self.conn.commit()

    # ── Locations ─────────────────────────────────────────────────

    def upsert_locations(self, locations: list[dict], source: str = "web_form") -> list[int]:
        """Oppdater eller sett inn lokasjoner."""
        result_ids = []
        for loc in locations:
            row_id = loc.get("id")
            if row_id:
                existing = self.conn.execute(
                    "SELECT * FROM locations WHERE id = ?", (row_id,)
                ).fetchone()
                if existing:
                    self._log_location_changes(existing, loc, row_id, source)
                    self.conn.execute(
                        """UPDATE locations SET code = ?, name = ?, location_type = ?
                           WHERE id = ?""",
                        (loc.get("code", existing["code"]),
                         loc.get("name", ""),
                         loc.get("location_type", ""),
                         row_id),
                    )
                    result_ids.append(row_id)
                else:
                    result_ids.append(self._insert_location(loc, source))
            else:
                result_ids.append(self._insert_location(loc, source))
        self.conn.commit()
        return result_ids

    def _log_location_changes(self, existing, loc: dict, row_id: int, source: str):
        fields = {
            "code": (str, ""),
            "name": (str, ""),
            "location_type": (str, ""),
        }
        for field, (ftype, default) in fields.items():
            old = existing[field]
            new = ftype(loc.get(field, default))
            if old != new:
                self.log_change(
                    "locations", str(row_id), field, old, new, source=source,
                )

    def _insert_location(self, loc: dict, source: str) -> int:
        cur = self.conn.execute(
            """INSERT INTO locations (code, name, location_type)
               VALUES (?, ?, ?)
               ON CONFLICT(code) DO UPDATE SET
                   name = excluded.name,
                   location_type = excluded.location_type
               RETURNING id""",
            (loc.get("code", ""), loc.get("name", ""), loc.get("location_type", "")),
        )
        row = cur.fetchone()
        new_id = row["id"] if row else 0
        self.log_change("locations", str(new_id), "_created", None,
                        loc.get("code", ""), source=source)
        return new_id

    def delete_location(self, location_id: int, source: str = "web_form"):
        existing = self.conn.execute(
            "SELECT * FROM locations WHERE id = ?", (location_id,)
        ).fetchone()
        if existing:
            self.log_change(
                "locations", str(location_id), "_deleted",
                json.dumps(dict(existing), ensure_ascii=False), None,
                source=source,
            )
            self.conn.execute("DELETE FROM locations WHERE id = ?", (location_id,))
            self.conn.commit()

    # ── Work Centers ──────────────────────────────────────────────

    def upsert_work_centers(self, wcs: list[dict], source: str = "web_form") -> list[int]:
        """Oppdater eller sett inn arbeidssentre. Logger endringer."""
        result_ids = []
        for wc in wcs:
            row_id = wc.get("id")
            if row_id:
                existing = self.conn.execute(
                    "SELECT * FROM work_centers WHERE id = ?", (row_id,)
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
                                "work_centers", str(row_id), field, old, new,
                                source=source,
                            )
                    self.conn.execute(
                        """UPDATE work_centers SET
                           code = ?, description = ?, location_code = ?,
                           labor_cost_hour = ?, machine_cost_hour = ?,
                           overhead_cost_hour = ?, capacity_hours_day = ?,
                           effective_capacity_pct = ?
                           WHERE id = ?""",
                        (
                            wc.get("code", existing["code"]),
                            wc.get("description", ""),
                            wc.get("location_code", ""),
                            float(wc.get("labor_cost_hour", 0)),
                            float(wc.get("machine_cost_hour", 0)),
                            float(wc.get("overhead_cost_hour", 0)),
                            float(wc.get("capacity_hours_day", 0)),
                            float(wc.get("effective_capacity_pct", 100)),
                            row_id,
                        ),
                    )
                    result_ids.append(row_id)
                else:
                    result_ids.append(self._insert_work_center(wc, source))
            else:
                result_ids.append(self._insert_work_center(wc, source))
        self.conn.commit()
        return result_ids

    def _insert_work_center(self, wc: dict, source: str) -> int:
        cur = self.conn.execute(
            """INSERT INTO work_centers
               (code, description, location_code, labor_cost_hour, machine_cost_hour,
                overhead_cost_hour, capacity_hours_day, effective_capacity_pct)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(code) DO UPDATE SET
                   description = excluded.description, location_code = excluded.location_code,
                   labor_cost_hour = excluded.labor_cost_hour,
                   machine_cost_hour = excluded.machine_cost_hour,
                   overhead_cost_hour = excluded.overhead_cost_hour,
                   capacity_hours_day = excluded.capacity_hours_day,
                   effective_capacity_pct = excluded.effective_capacity_pct
               RETURNING id""",
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
        row = cur.fetchone()
        new_id = row["id"] if row else 0
        self.log_change("work_centers", str(new_id), "_created", None,
                        wc.get("code", ""), source=source)
        return new_id

    def delete_work_center(self, wc_id: int, source: str = "web_form"):
        existing = self.conn.execute(
            "SELECT * FROM work_centers WHERE id = ?", (wc_id,)
        ).fetchone()
        if existing:
            self.log_change(
                "work_centers", str(wc_id), "_deleted",
                json.dumps(dict(existing), ensure_ascii=False), None,
                source=source,
            )
            self.conn.execute("DELETE FROM work_centers WHERE id = ?", (wc_id,))
            self.conn.commit()

    # ── Operations ────────────────────────────────────────────────

    def upsert_operations(self, operations: list[dict], source: str = "web_form") -> list[int]:
        """Oppdater eller sett inn operasjoner."""
        result_ids = []
        for op in operations:
            row_id = op.get("id")
            if row_id:
                existing = self.conn.execute(
                    "SELECT * FROM operations WHERE id = ?", (row_id,)
                ).fetchone()
                if existing:
                    fields = {
                        "description": (str, ""),
                        "default_work_center": (str, ""),
                        "standard_unit": (str, "Minutes"),
                    }
                    for field, (ftype, default) in fields.items():
                        old = existing[field]
                        new = ftype(op.get(field, default))
                        if old != new:
                            self.log_change(
                                "operations", str(row_id), field, old, new,
                                source=source,
                            )
                    self.conn.execute(
                        """UPDATE operations SET
                           code = ?, description = ?, default_work_center = ?, standard_unit = ?
                           WHERE id = ?""",
                        (op.get("code", existing["code"]),
                         op.get("description", ""),
                         op.get("default_work_center", ""),
                         op.get("standard_unit", "Minutes"),
                         row_id),
                    )
                    result_ids.append(row_id)
                else:
                    result_ids.append(self._insert_operation(op, source))
            else:
                result_ids.append(self._insert_operation(op, source))
        self.conn.commit()
        return result_ids

    def _insert_operation(self, op: dict, source: str) -> int:
        cur = self.conn.execute(
            """INSERT INTO operations (code, description, default_work_center, standard_unit)
               VALUES (?, ?, ?, ?)
               ON CONFLICT(code) DO UPDATE SET
                   description = excluded.description,
                   default_work_center = excluded.default_work_center,
                   standard_unit = excluded.standard_unit
               RETURNING id""",
            (op.get("code", ""), op.get("description", ""),
             op.get("default_work_center", ""), op.get("standard_unit", "Minutes")),
        )
        row = cur.fetchone()
        new_id = row["id"] if row else 0
        self.log_change("operations", str(new_id), "_created", None,
                        op.get("code", ""), source=source)
        return new_id

    def delete_operation(self, op_id: int, source: str = "web_form"):
        existing = self.conn.execute(
            "SELECT * FROM operations WHERE id = ?", (op_id,)
        ).fetchone()
        if existing:
            self.log_change(
                "operations", str(op_id), "_deleted",
                json.dumps(dict(existing), ensure_ascii=False), None,
                source=source,
            )
            self.conn.execute("DELETE FROM operations WHERE id = ?", (op_id,))
            self.conn.commit()

    # ── Item Costs ────────────────────────────────────────────────

    def upsert_item_costs(self, costs: list[dict], source: str = "web_form") -> list[int]:
        """Oppdater eller sett inn kostpriser. Logger endringer."""
        result_ids = []
        for c in costs:
            row_id = c.get("id")
            if row_id:
                existing = self.conn.execute(
                    "SELECT * FROM item_costs WHERE id = ?", (row_id,)
                ).fetchone()
                if existing:
                    # Logg unit_cost-endringer
                    old_unit_cost = existing["unit_cost"]
                    new_unit_cost = float(c.get("unit_cost", 0))
                    if abs(old_unit_cost - new_unit_cost) > 0.001:
                        self.log_change(
                            "item_costs", str(row_id),
                            "unit_cost", old_unit_cost, new_unit_cost,
                            source=source,
                        )
                    # Logg currency-endringer
                    old_currency = existing["currency"]
                    new_currency = c.get("currency", "NOK")
                    if old_currency != new_currency:
                        self.log_change(
                            "item_costs", str(row_id),
                            "currency", old_currency, new_currency,
                            source=source,
                        )
                    # Logg effective_date-endringer
                    old_date = existing["effective_date"]
                    new_date = c.get("effective_date")
                    if old_date != new_date:
                        self.log_change(
                            "item_costs", str(row_id),
                            "effective_date", old_date, new_date,
                            source=source,
                        )
                    # Logg cost_type-endringer
                    old_cost_type = existing["cost_type"]
                    new_cost_type = c.get("cost_type", existing["cost_type"])
                    if old_cost_type != new_cost_type:
                        self.log_change(
                            "item_costs", str(row_id),
                            "cost_type", old_cost_type, new_cost_type,
                            source=source,
                        )
                    self.conn.execute(
                        """UPDATE item_costs SET
                           item_no = ?, cost_type = ?, unit_cost = ?,
                           currency = ?, effective_date = ?
                           WHERE id = ?""",
                        (
                            c.get("item_no", existing["item_no"]),
                            new_cost_type,
                            new_unit_cost,
                            new_currency,
                            new_date,
                            row_id,
                        ),
                    )
                    result_ids.append(row_id)
                else:
                    result_ids.append(self._insert_item_cost(c, source))
            else:
                result_ids.append(self._insert_item_cost(c, source))
        self.conn.commit()
        return result_ids

    def _insert_item_cost(self, c: dict, source: str) -> int:
        cur = self.conn.execute(
            """INSERT INTO item_costs (item_no, cost_type, unit_cost, currency, effective_date)
               VALUES (?, ?, ?, ?, ?)
               ON CONFLICT(item_no, cost_type) DO UPDATE SET
                   unit_cost = excluded.unit_cost,
                   currency = excluded.currency,
                   effective_date = excluded.effective_date
               RETURNING id""",
            (
                c.get("item_no", ""),
                c.get("cost_type", "Standard Cost"),
                float(c.get("unit_cost", 0)),
                c.get("currency", "NOK"),
                c.get("effective_date"),
            ),
        )
        row = cur.fetchone()
        new_id = row["id"] if row else 0
        self.log_change("item_costs", str(new_id), "_created", None,
                        f"{c.get('item_no', '')}: {c.get('cost_type', '')}",
                        source=source)
        return new_id

    def delete_item_cost(self, cost_id: int, source: str = "web_form"):
        existing = self.conn.execute(
            "SELECT * FROM item_costs WHERE id = ?", (cost_id,)
        ).fetchone()
        if existing:
            self.log_change(
                "item_costs", str(cost_id), "_deleted",
                json.dumps(dict(existing), ensure_ascii=False), None,
                source=source,
            )
            self.conn.execute("DELETE FROM item_costs WHERE id = ?", (cost_id,))
            self.conn.commit()

    # ── BOM Lines ─────────────────────────────────────────────────

    def upsert_bom_lines(self, lines: list[dict], source: str = "web_form") -> list[int]:
        """Oppdater eller sett inn BOM-linjer. Logger endringer."""
        result_ids = []
        for bl in lines:
            row_id = bl.get("id")
            if row_id:
                existing = self.conn.execute(
                    "SELECT * FROM bom_lines WHERE id = ?", (row_id,)
                ).fetchone()
                if existing:
                    for field in ("quantity_per", "scrap_pct", "co_product_pct"):
                        old = existing[field]
                        new = float(bl.get(field, 0))
                        if abs(old - new) > 0.001:
                            self.log_change(
                                "bom_lines", str(row_id), field, old, new,
                                source=source,
                            )
                    self.conn.execute(
                        """UPDATE bom_lines SET
                           parent_item_no = ?, component_item_no = ?,
                           quantity_per = ?, uom = ?, scrap_pct = ?,
                           co_product_pct = ?, co_product_item_no = ?
                           WHERE id = ?""",
                        (
                            bl.get("parent_item_no", existing["parent_item_no"]),
                            bl.get("component_item_no", existing["component_item_no"]),
                            float(bl.get("quantity_per", 1)),
                            bl.get("uom", ""),
                            float(bl.get("scrap_pct", 0)),
                            float(bl.get("co_product_pct", 0)),
                            bl.get("co_product_item_no", ""),
                            row_id,
                        ),
                    )
                    result_ids.append(row_id)
                else:
                    result_ids.append(self._insert_bom_line(bl, source))
            else:
                result_ids.append(self._insert_bom_line(bl, source))
        self.conn.commit()
        return result_ids

    def _insert_bom_line(self, bl: dict, source: str) -> int:
        cur = self.conn.execute(
            """INSERT INTO bom_lines
               (parent_item_no, component_item_no, quantity_per, uom,
                scrap_pct, co_product_pct, co_product_item_no)
               VALUES (?, ?, ?, ?, ?, ?, ?)
               ON CONFLICT(parent_item_no, component_item_no) DO UPDATE SET
                   quantity_per = excluded.quantity_per, uom = excluded.uom,
                   scrap_pct = excluded.scrap_pct, co_product_pct = excluded.co_product_pct,
                   co_product_item_no = excluded.co_product_item_no
               RETURNING id""",
            (
                bl.get("parent_item_no", ""),
                bl.get("component_item_no", ""),
                float(bl.get("quantity_per", 1)),
                bl.get("uom", ""),
                float(bl.get("scrap_pct", 0)),
                float(bl.get("co_product_pct", 0)),
                bl.get("co_product_item_no", ""),
            ),
        )
        row = cur.fetchone()
        new_id = row["id"] if row else 0
        self.log_change("bom_lines", str(new_id), "_created", None,
                        f"{bl.get('parent_item_no', '')}:{bl.get('component_item_no', '')}",
                        source=source)
        return new_id

    def delete_bom_line(self, bom_id: int, source: str = "web_form"):
        existing = self.conn.execute(
            "SELECT * FROM bom_lines WHERE id = ?", (bom_id,)
        ).fetchone()
        if existing:
            self.log_change(
                "bom_lines", str(bom_id), "_deleted",
                json.dumps(dict(existing), ensure_ascii=False), None,
                source=source,
            )
            self.conn.execute("DELETE FROM bom_lines WHERE id = ?", (bom_id,))
            self.conn.commit()

    # ── Routing Lines ─────────────────────────────────────────────

    def upsert_routing_lines(self, lines: list[dict], source: str = "web_form") -> list[int]:
        """Oppdater eller sett inn routing-linjer. Logger endringer."""
        result_ids = []
        for rl in lines:
            row_id = rl.get("id")
            if row_id:
                existing = self.conn.execute(
                    "SELECT * FROM routing_lines WHERE id = ?", (row_id,)
                ).fetchone()
                if existing:
                    for field in ("run_time_minutes", "setup_time_minutes", "batch_size", "changeover_time_minutes"):
                        old = existing[field]
                        new = float(rl.get(field, 0))
                        if abs(old - new) > 0.001:
                            self.log_change(
                                "routing_lines", str(row_id), field, old, new,
                                source=source,
                            )
                    self.conn.execute(
                        """UPDATE routing_lines SET
                           item_no = ?, operation_no = ?, operation_code = ?,
                           work_center_code = ?, setup_time_minutes = ?,
                           run_time_minutes = ?, batch_size = ?, changeover_time_minutes = ?
                           WHERE id = ?""",
                        (
                            rl.get("item_no", existing["item_no"]),
                            int(rl.get("operation_no", 0)),
                            rl.get("operation_code", ""),
                            rl.get("work_center_code", ""),
                            float(rl.get("setup_time_minutes", 0)),
                            float(rl.get("run_time_minutes", 0)),
                            float(rl.get("batch_size", 1)),
                            float(rl.get("changeover_time_minutes", 0)),
                            row_id,
                        ),
                    )
                    result_ids.append(row_id)
                else:
                    result_ids.append(self._insert_routing_line(rl, source))
            else:
                result_ids.append(self._insert_routing_line(rl, source))
        self.conn.commit()
        return result_ids

    def _insert_routing_line(self, rl: dict, source: str) -> int:
        cur = self.conn.execute(
            """INSERT INTO routing_lines
               (item_no, operation_no, operation_code, work_center_code,
                setup_time_minutes, run_time_minutes, batch_size,
                changeover_time_minutes)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)
               ON CONFLICT(item_no, operation_no, work_center_code) DO UPDATE SET
                   operation_code = excluded.operation_code,
                   setup_time_minutes = excluded.setup_time_minutes,
                   run_time_minutes = excluded.run_time_minutes,
                   batch_size = excluded.batch_size,
                   changeover_time_minutes = excluded.changeover_time_minutes
               RETURNING id""",
            (
                rl.get("item_no", ""),
                int(rl.get("operation_no", 0)),
                rl.get("operation_code", ""),
                rl.get("work_center_code", ""),
                float(rl.get("setup_time_minutes", 0)),
                float(rl.get("run_time_minutes", 0)),
                float(rl.get("batch_size", 1)),
                float(rl.get("changeover_time_minutes", 0)),
            ),
        )
        row = cur.fetchone()
        new_id = row["id"] if row else 0
        self.log_change("routing_lines", str(new_id), "_created", None,
                        f"{rl.get('item_no', '')}:{rl.get('operation_no', 0)}",
                        source=source)
        return new_id

    def delete_routing_line(self, routing_id: int, source: str = "web_form"):
        existing = self.conn.execute(
            "SELECT * FROM routing_lines WHERE id = ?", (routing_id,)
        ).fetchone()
        if existing:
            self.log_change(
                "routing_lines", str(routing_id), "_deleted",
                json.dumps(dict(existing), ensure_ascii=False), None,
                source=source,
            )
            self.conn.execute("DELETE FROM routing_lines WHERE id = ?", (routing_id,))
            self.conn.commit()

    # ── By Product Rules ──────────────────────────────────────────

    def upsert_byproduct_rules(self, rules: list[dict], source: str = "web_form") -> list[int]:
        """Oppdater eller sett inn byproduct rules."""
        result_ids = []
        for r in rules:
            row_id = r.get("id")
            if row_id:
                existing = self.conn.execute(
                    "SELECT * FROM byproduct_rules WHERE id = ?", (row_id,)
                ).fetchone()
                if existing:
                    fields = {
                        "expected_quantity": (float, 0),
                        "market_value": (float, 0),
                        "allocation_method": (str, "Reduce Main Product Cost"),
                        "uom": (str, ""),
                    }
                    for field, (ftype, default) in fields.items():
                        old = existing[field]
                        new = ftype(r.get(field, default))
                        if old != new:
                            self.log_change(
                                "byproduct_rules", str(row_id), field, old, new,
                                source=source,
                            )
                    self.conn.execute(
                        """UPDATE byproduct_rules SET
                           parent_item_no = ?, by_product_item_no = ?,
                           expected_quantity = ?, uom = ?,
                           market_value = ?, allocation_method = ?
                           WHERE id = ?""",
                        (
                            r.get("parent_item_no", existing["parent_item_no"]),
                            r.get("by_product_item_no", existing["by_product_item_no"]),
                            float(r.get("expected_quantity", 0)),
                            r.get("uom", ""),
                            float(r.get("market_value", 0)),
                            r.get("allocation_method", "Reduce Main Product Cost"),
                            row_id,
                        ),
                    )
                    result_ids.append(row_id)
                else:
                    result_ids.append(self._insert_byproduct_rule(r, source))
            else:
                result_ids.append(self._insert_byproduct_rule(r, source))
        self.conn.commit()
        return result_ids

    def _insert_byproduct_rule(self, r: dict, source: str) -> int:
        cur = self.conn.execute(
            """INSERT INTO byproduct_rules
               (parent_item_no, by_product_item_no, expected_quantity, uom,
                market_value, allocation_method)
               VALUES (?, ?, ?, ?, ?, ?)
               ON CONFLICT(parent_item_no, by_product_item_no) DO UPDATE SET
                   expected_quantity = excluded.expected_quantity,
                   uom = excluded.uom,
                   market_value = excluded.market_value,
                   allocation_method = excluded.allocation_method
               RETURNING id""",
            (
                r.get("parent_item_no", ""),
                r.get("by_product_item_no", ""),
                float(r.get("expected_quantity", 0)),
                r.get("uom", ""),
                float(r.get("market_value", 0)),
                r.get("allocation_method", "Reduce Main Product Cost"),
            ),
        )
        row = cur.fetchone()
        new_id = row["id"] if row else 0
        self.log_change("byproduct_rules", str(new_id), "_created", None,
                        f"{r.get('parent_item_no', '')}:{r.get('by_product_item_no', '')}",
                        source=source)
        return new_id

    def delete_byproduct_rule(self, rule_id: int, source: str = "web_form"):
        existing = self.conn.execute(
            "SELECT * FROM byproduct_rules WHERE id = ?", (rule_id,)
        ).fetchone()
        if existing:
            self.log_change(
                "byproduct_rules", str(rule_id), "_deleted",
                json.dumps(dict(existing), ensure_ascii=False), None,
                source=source,
            )
            self.conn.execute("DELETE FROM byproduct_rules WHERE id = ?", (rule_id,))
            self.conn.commit()

    # ── Capacity Days ─────────────────────────────────────────────

    def upsert_capacity_days(self, days: list[dict], source: str = "web_form") -> list[int]:
        """Oppdater eller sett inn capacity days."""
        result_ids = []
        for d in days:
            row_id = d.get("id")
            if row_id:
                existing = self.conn.execute(
                    "SELECT * FROM capacity_days WHERE id = ?", (row_id,)
                ).fetchone()
                if existing:
                    fields = {
                        "available_hours": (float, 0),
                        "planned_downtime": (float, 0),
                    }
                    for field, (ftype, default) in fields.items():
                        old = existing[field]
                        new = ftype(d.get(field, default))
                        if old != new:
                            self.log_change(
                                "capacity_days", str(row_id), field, old, new,
                                source=source,
                            )
                    self.conn.execute(
                        """UPDATE capacity_days SET
                           work_center = ?, date = ?,
                           available_hours = ?, planned_downtime = ?
                           WHERE id = ?""",
                        (
                            d.get("work_center", existing["work_center"]),
                            d.get("date", existing["date"]),
                            float(d.get("available_hours", 0)),
                            float(d.get("planned_downtime", 0)),
                            row_id,
                        ),
                    )
                    result_ids.append(row_id)
                else:
                    result_ids.append(self._insert_capacity_day(d, source))
            else:
                result_ids.append(self._insert_capacity_day(d, source))
        self.conn.commit()
        return result_ids

    def _insert_capacity_day(self, d: dict, source: str) -> int:
        cur = self.conn.execute(
            """INSERT INTO capacity_days (work_center, date, available_hours, planned_downtime)
               VALUES (?, ?, ?, ?)
               ON CONFLICT(work_center, date) DO UPDATE SET
                   available_hours = excluded.available_hours,
                   planned_downtime = excluded.planned_downtime
               RETURNING id""",
            (
                d.get("work_center", ""),
                d.get("date", ""),
                float(d.get("available_hours", 0)),
                float(d.get("planned_downtime", 0)),
            ),
        )
        row = cur.fetchone()
        new_id = row["id"] if row else 0
        self.log_change("capacity_days", str(new_id), "_created", None,
                        f"{d.get('work_center', '')}:{d.get('date', '')}",
                        source=source)
        return new_id

    def delete_capacity_day(self, capacity_id: int, source: str = "web_form"):
        existing = self.conn.execute(
            "SELECT * FROM capacity_days WHERE id = ?", (capacity_id,)
        ).fetchone()
        if existing:
            self.log_change(
                "capacity_days", str(capacity_id), "_deleted",
                json.dumps(dict(existing), ensure_ascii=False), None,
                source=source,
            )
            self.conn.execute("DELETE FROM capacity_days WHERE id = ?", (capacity_id,))
            self.conn.commit()

    # ── Production Scenarios ──────────────────────────────────────

    def upsert_scenarios(self, scenarios: list[dict], source: str = "web_form") -> list[int]:
        """Oppdater eller sett inn production scenarios."""
        result_ids = []
        for sc in scenarios:
            row_id = sc.get("id")
            if row_id:
                existing = self.conn.execute(
                    "SELECT * FROM production_scenarios WHERE id = ?", (row_id,)
                ).fetchone()
                if existing:
                    old_qty = existing["planned_quantity"]
                    new_qty = float(sc.get("planned_quantity", 0))
                    if abs(old_qty - new_qty) > 0.001:
                        self.log_change(
                            "production_scenarios", str(row_id),
                            "planned_quantity", old_qty, new_qty,
                            source=source,
                        )
                    self.conn.execute(
                        """UPDATE production_scenarios SET
                           scenario_name = ?, product = ?, planned_quantity = ?
                           WHERE id = ?""",
                        (
                            sc.get("scenario_name", existing["scenario_name"]),
                            sc.get("product", existing["product"]),
                            new_qty,
                            row_id,
                        ),
                    )
                    result_ids.append(row_id)
                else:
                    result_ids.append(self._insert_scenario(sc, source))
            else:
                result_ids.append(self._insert_scenario(sc, source))
        self.conn.commit()
        return result_ids

    def _insert_scenario(self, sc: dict, source: str) -> int:
        cur = self.conn.execute(
            """INSERT INTO production_scenarios (scenario_name, product, planned_quantity)
               VALUES (?, ?, ?)
               ON CONFLICT(scenario_name, product) DO UPDATE SET
                   planned_quantity = excluded.planned_quantity
               RETURNING id""",
            (
                sc.get("scenario_name", ""),
                sc.get("product", ""),
                float(sc.get("planned_quantity", 0)),
            ),
        )
        row = cur.fetchone()
        new_id = row["id"] if row else 0
        self.log_change("production_scenarios", str(new_id), "_created", None,
                        f"{sc.get('scenario_name', '')}:{sc.get('product', '')}",
                        source=source)
        return new_id

    def delete_scenario(self, scenario_id: int, source: str = "web_form"):
        existing = self.conn.execute(
            "SELECT * FROM production_scenarios WHERE id = ?", (scenario_id,)
        ).fetchone()
        if existing:
            self.log_change(
                "production_scenarios", str(scenario_id), "_deleted",
                json.dumps(dict(existing), ensure_ascii=False), None,
                source=source,
            )
            self.conn.execute("DELETE FROM production_scenarios WHERE id = ?", (scenario_id,))
            self.conn.commit()

    # ── Demand-import fra DataFrame ─────────────────────────────

    def upsert_demand_from_df(
        self,
        df,
        item_col: str = "item_no",
        period_col: str = "Uke",
        qty_col: str = "qty",
        location_col: str = "location_code",
        location_mapping: Optional[dict[str, str]] = None,
    ) -> tuple[int, int]:
        """Importer demand-data fra en pandas DataFrame.

        Filtrerer automatisk bort varer som ikke finnes i products-tabellen.
        Overskriver eksisterende demand-data (tømmer tabellen først).

        Args:
            df: DataFrame med kolonner for varenr, uke, kvantum og lokasjon
            item_col: Kolonnenavn for varenummer (må matches mot products.item_no)
            period_col: Kolonnenavn for ukenummer/periode
            qty_col: Kolonnenavn for kvantum (løpemeter)
            location_col: Kolonnenavn for lokasjonsnavn (f.eks. "HOVEDLAGER")
            location_mapping: Valgfri mapping fra rå navn til våre location_code
                              (f.eks. {"HOVEDLAGER": "KOD"})

        Returns:
            (antall_importerte_rader, antall_filtrert_bort)
        """
        import pandas as pd

        # Hent alle gyldige product_id fra products-tabellen
        valid_items = {r[0] for r in self.conn.execute(
            "SELECT item_no FROM products"
        ).fetchall()}

        # Filtrer DataFrame mot gyldige varer
        df_filtered = df[df[item_col].isin(valid_items)]
        filtered_out = len(df) - len(df_filtered)

        # Default mapping hvis ikke angitt
        if location_mapping is None:
            location_mapping = {"HOVEDLAGER": "KOD"}

        # Tøm eksisterende demand-data
        self.conn.execute("DELETE FROM demand")

        # Sett inn nye rader
        rows = []
        for _, row in df_filtered.iterrows():
            raw_loc = str(row[location_col]).strip() if location_col in df_filtered.columns else ""
            mapped_loc = location_mapping.get(raw_loc, raw_loc)
            rows.append((
                str(row[item_col]),
                int(row[period_col]),
                float(row[qty_col]),
                mapped_loc,
                "",
            ))

        if rows:
            self.conn.executemany(
                """INSERT OR REPLACE INTO demand 
                   (product_id, period, quantity, location_code, customer_region)
                   VALUES (?, ?, ?, ?, ?)""",
                rows,
            )
        self.conn.commit()
        return len(rows), filtered_out

    # ── Historisk salg-import fra DataFrame ─────────────────────

    def upsert_historical_sales_from_df(
        self,
        df,
        item_col: str = "item_no",
        period_col: str = "Uke",
        qty_col: str = "qty",
        location_col: str = "location_code",
        location_mapping: Optional[dict[str, str]] = None,
    ) -> tuple[int, int]:
        """Importer historisk salg fra en pandas DataFrame.

        Filtrerer automatisk bort varer som ikke finnes i products-tabellen.
        Overskriver eksisterende historical_sales-data (tømmer tabellen først).

        Args:
            df: DataFrame med kolonner for varenr, uke, kvantum og lokasjon
            item_col: Kolonnenavn for varenummer (må matches mot products.item_no)
            period_col: Kolonnenavn for ukenummer/periode
            qty_col: Kolonnenavn for kvantum (løpemeter)
            location_col: Kolonnenavn for lokasjonsnavn (f.eks. "HOVEDLAGER")
            location_mapping: Valgfri mapping fra rå navn til våre location_code
                              (f.eks. {"HOVEDLAGER": "KOD"})

        Returns:
            (antall_importerte_rader, antall_filtrert_bort)
        """
        # Hent alle gyldige product_id fra products-tabellen
        valid_items = {r[0] for r in self.conn.execute(
            "SELECT item_no FROM products"
        ).fetchall()}

        # Filtrer DataFrame mot gyldige varer
        df_filtered = df[df[item_col].isin(valid_items)]
        filtered_out = len(df) - len(df_filtered)

        # Default mapping hvis ikke angitt
        if location_mapping is None:
            location_mapping = {"HOVEDLAGER": "KOD"}

        # Tøm eksisterende historical_sales-data
        self.conn.execute("DELETE FROM historical_sales")

        # Sett inn nye rader
        rows = []
        for _, row in df_filtered.iterrows():
            raw_loc = str(row[location_col]).strip() if location_col in df_filtered.columns else ""
            mapped_loc = location_mapping.get(raw_loc, raw_loc)
            rows.append((
                str(row[item_col]),
                int(row[period_col]),
                float(row[qty_col]),
                mapped_loc,
            ))

        if rows:
            self.conn.executemany(
                """INSERT OR REPLACE INTO historical_sales 
                   (product_id, period, quantity, location_code)
                   VALUES (?, ?, ?, ?)""",
                rows,
            )
        self.conn.commit()
        return len(rows), filtered_out

    # ── Tømming og tilbakestilling ──────────────────────────────

    def clear_all_data(self):
        """Slett all data fra alle stamdata-tabeller (bevar change_log)."""
        tables = [
            "products", "locations", "work_centers", "operations",
            "item_costs", "bom_lines", "routing_lines", "byproduct_rules",
            "capacity_days", "production_scenarios",
            "demand", "historical_sales", "changeover_matrix",
            "transport_flagg", "transport_ruter",
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
        result = cur.lastrowid
        return result if result is not None else 0

    def export_all_data(self) -> dict[str, list[sqlite3.Row]]:
        """Eksporter all data fra alle tabeller (inkluderer id).
        
        Returnerer dict med tabellnavn som nøkler.
        """
        tables = {
            "products": "SELECT * FROM products ORDER BY id",
            "locations": "SELECT * FROM locations ORDER BY id",
            "work_centers": "SELECT * FROM work_centers ORDER BY id",
            "operations": "SELECT * FROM operations ORDER BY id",
            "item_costs": "SELECT * FROM item_costs ORDER BY id",
            "bom_lines": "SELECT * FROM bom_lines ORDER BY id",
            "routing_lines": "SELECT * FROM routing_lines ORDER BY id",
            "byproduct_rules": "SELECT * FROM byproduct_rules ORDER BY id",
            "capacity_days": "SELECT * FROM capacity_days ORDER BY id",
            "production_scenarios": "SELECT * FROM production_scenarios ORDER BY id",
            "transport_flagg": "SELECT * FROM transport_flagg ORDER BY item_no",
            "transport_ruter": "SELECT * FROM transport_ruter ORDER BY from_loc, to_loc",
            "demand": "SELECT * FROM demand ORDER BY period, product_id",
            "historical_sales": "SELECT * FROM historical_sales ORDER BY period, product_id",
            "changeover_matrix": "SELECT work_center_code, from_family, to_family, changeover_minutes FROM changeover_matrix ORDER BY work_center_code, from_family, to_family",
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
        confirm = input("[ADV] Tømme all data? (ja/nei): ")
        if confirm.lower() in ("ja", "yes", "y"):
            db.clear_all_data()
            print("[OK] All data slettet")
        else:
            print("Avbrutt")

    if args.clear_log:
        confirm = input("[ADV] Tømme endringsloggen? (ja/nei): ")
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