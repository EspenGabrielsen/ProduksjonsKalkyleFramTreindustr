from pathlib import Path

path = Path("src/data_repo.py")
text = path.read_text(encoding="utf-8")
start = text.index("    def executemany(self, query: str, params_seq):")
end = text.index("    def executescript(self, sql: str):", start)

replacement = '''    def executemany(self, query: str, params_seq):
        """Backend-bevisst executemany med enkel placeholder-konvertering."""
        sql = _convert_sql_placeholders(query, self.backend)
        if self.backend == "sqlite":
            return self.conn.executemany(sql, params_seq)
        with self.conn.cursor() as cur:
            cur.executemany(sql, params_seq)
            return cur.rowcount

'''

path.write_text(text[:start] + replacement + text[end:], encoding="utf-8")
