from pathlib import Path
import re

path = Path("src/data_repo.py")
text = path.read_text(encoding="utf-8")

# 1) imports
if "from functools import wraps" not in text:
    text = text.replace("from contextlib import contextmanager\n", "from contextlib import contextmanager\nfrom functools import wraps\n", 1)

# 2) Convert existing direct commits to a depth-aware helper. Do this before
# inserting the helper/decorator so their real commits stay real commits.
text = text.replace("self.conn.commit()", "self._commit()")

# 3) Insert the atomic decorator before DataRepo.
class_marker = "class DataRepo:"
if "def _atomic_mutation(method):" not in text:
    decorator = '''def _atomic_mutation(method):
    """Kjør en mutasjon og tilhørende audit-logg som én transaksjon."""
    @wraps(method)
    def wrapped(self, *args, **kwargs):
        outermost = self._transaction_depth == 0
        self._transaction_depth += 1
        try:
            result = method(self, *args, **kwargs)
        except Exception:
            self._transaction_depth -= 1
            if outermost and self._conn is not None:
                self.conn.rollback()
            raise
        self._transaction_depth -= 1
        if outermost:
            self.conn.commit()
        return result
    return wrapped


'''
    text = text.replace(class_marker, decorator + class_marker, 1)

# 4) Add transaction depth to constructor.
needle = "        self._conn = None\n"
if "self._transaction_depth" not in text[text.index(class_marker):text.index(class_marker)+1200]:
    text = text.replace(needle, needle + "        self._transaction_depth = 0\n", 1)

# 5) Add depth-aware commit helper before execute().
execute_marker = "    def execute(self, query: str, params: Optional[tuple | list] = None):"
if "    def _commit(self):" not in text:
    helper = '''    def _commit(self):
        """Commit umiddelbart utenfor mutasjon, ellers utsett til ytterste mutasjon."""
        if self._transaction_depth == 0:
            self.conn.commit()

'''
    text = text.replace(execute_marker, helper + execute_marker, 1)

# 6) Make all public CRUD upserts/deletes atomic. Avoid duplicate decorators.
pattern = re.compile(r'(?m)^(    def (?:upsert_[A-Za-z0-9_]+|delete_[A-Za-z0-9_]+)\()')
text = pattern.sub(lambda m: "    @_atomic_mutation\n" + m.group(1), text)
text = text.replace("    @_atomic_mutation\n    @_atomic_mutation\n", "    @_atomic_mutation\n")

path.write_text(text, encoding="utf-8")
