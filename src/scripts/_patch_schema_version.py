from pathlib import Path

path = Path("src/data_repo.py")
text = path.read_text(encoding="utf-8")

# Baseline for the schema that azure-readiness currently creates/supports.
needle = 'DB_FILENAME = "produksjonskalkyle.db"\n'
if "CURRENT_SCHEMA_VERSION = 1" not in text:
    text = text.replace(
        needle,
        needle + 'CURRENT_SCHEMA_VERSION = 1\n',
        1,
    )

# Register/validate schema version after all legacy SQLite compatibility work.
call_marker = "    def is_empty(self) -> bool:\n"
if "        self._ensure_schema_version()\n\n    def is_empty" not in text:
    text = text.replace(
        call_marker,
        "        self._ensure_schema_version()\n\n" + call_marker,
        1,
    )

# Add backend-neutral schema metadata helpers immediately before is_empty().
if "    def _ensure_schema_version(self):" not in text:
    methods = '''    def _ensure_schema_version(self):
        """Registrer og valider database-skjemaets versjon.

        Eksisterende databaser uten metadata regnes som baseline v1 etter at
        initialize() har kjørt dagens idempotente skjema og legacy-migreringer.
        Fremtidige skjemaendringer skal øke CURRENT_SCHEMA_VERSION og få en
        eksplisitt nummerert migrering før denne kontrollen oppdateres.
        """
        self.execute(
            """CREATE TABLE IF NOT EXISTS schema_version (
                   singleton INTEGER PRIMARY KEY,
                   version INTEGER NOT NULL,
                   updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                   CHECK (singleton = 1)
               )"""
        )
        row = self.execute(
            "SELECT version FROM schema_version WHERE singleton = 1"
        ).fetchone()

        if row is None:
            self.execute(
                "INSERT INTO schema_version (singleton, version) VALUES (?, ?)",
                (1, CURRENT_SCHEMA_VERSION),
            )
            self._commit()
            return

        version = int(row["version"] if isinstance(row, dict) else row[0])
        if version > CURRENT_SCHEMA_VERSION:
            self.conn.rollback()
            raise RuntimeError(
                "Databasen bruker skjema-versjon "
                f"{version}, men denne appversjonen støtter bare "
                f"{CURRENT_SCHEMA_VERSION}. Oppgrader applikasjonen før databasen brukes."
            )
        if version < CURRENT_SCHEMA_VERSION:
            self.conn.rollback()
            raise RuntimeError(
                "Databasen bruker skjema-versjon "
                f"{version}, mens appen forventer {CURRENT_SCHEMA_VERSION}. "
                "En eksplisitt databasemigrering må kjøres før oppstart."
            )
        self._commit()

    def get_schema_version(self) -> int:
        """Returner registrert skjema-versjon; initialize() må være kjørt først."""
        row = self.execute(
            "SELECT version FROM schema_version WHERE singleton = 1"
        ).fetchone()
        if row is None:
            raise RuntimeError("Databasen har ingen registrert schema_version")
        return int(row["version"] if isinstance(row, dict) else row[0])

'''
    text = text.replace(call_marker, methods + call_marker, 1)

path.write_text(text, encoding="utf-8")
