"""Build test workbooks from FIXTURE lines (VCL 0.1 rev 3 rules, deterministic).

Seeding: each column gets random.Random(seed) where seed = first 8 bytes of
sha256("<file>|<sheet>|<column letter>"). values() lists are shuffled with it, every random rule
draws from it, and +blank(n) picks its rows from it last. The same spec always builds the same
workbook. (Reconcile with harness/fixtures.py in the repo when integrating.)
"""
import datetime as dt
import hashlib
import random
import string

from lark import Token, Tree


class Sheet:
    def __init__(self, name):
        self.name = name
        self.cells = {}    # (row, col) -> value
        self.fills = {}    # (row, col) -> "#RRGGBB"
        self.formats = {}  # (row, col) -> number format

    def get(self, r, c):
        return self.cells.get((int(r), int(c)))

    def set(self, r, c, v):
        key = (int(r), int(c))
        if v is None or v == "":
            self.cells.pop(key, None)
        else:
            self.cells[key] = v

    def max_row(self):
        return max((r for r, _ in self.cells), default=1)

    def max_col(self):
        return max((c for _, c in self.cells), default=1)

    def snapshot(self):
        s = Sheet(self.name)
        s.cells, s.fills, s.formats = dict(self.cells), dict(self.fills), dict(self.formats)
        return s


def col_number(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _seed(*parts):
    return int.from_bytes(hashlib.sha256("|".join(parts).encode()).digest()[:8], "big")


def _num(tok):
    s = str(tok)
    return float(s) if "." in s else int(s)


def _unquote(s):
    return str(s)[1:-1].replace('""', '"')


def parse_fixture(tree):
    """fixture_value tree -> dict(file, sheet, cols=[(col, header)], rows, data={col: (rule, args, blanks)})"""
    kids = tree.children
    fx = {"file": str(kids[0]), "sheet": str(kids[1]), "cols": [], "rows": 0, "data": {}}
    rest = kids[2:]
    for item in rest:
        if isinstance(item, Tree) and item.data == "col_list":
            fx["cols"] = [(col_number(str(cd.children[0])), str(cd.children[1]).strip()) for cd in item.children]
        elif isinstance(item, Token) and item.type == "INT":
            fx["rows"] = int(item)
        elif isinstance(item, Tree) and item.data == "data_list":
            for dd in item.children:
                col = col_number(str(dd.children[0]))
                rule = dd.children[1]
                blanks = int(dd.children[2].children[0]) if len(dd.children) > 2 else 0
                fx["data"][col] = (rule.data, rule.children, blanks)
    return fx


def _column_values(fx, col, letter):
    rows = fx["rows"]
    rng = random.Random(_seed(fx["file"], fx["sheet"], letter))
    if col not in fx["data"]:
        return [None] * rows
    kind, args, blanks = fx["data"][col]
    if kind == "values_rule":
        vals = []
        for item in args:
            vals += [_unquote(item.children[0])] * int(item.children[1])
        if len(vals) != rows:
            raise ValueError(f"{fx['sheet']}!{letter}: values() add up to {len(vals)}, rows={rows}")
        rng.shuffle(vals)
        vals = [None if v == "" else v for v in vals]
    elif kind == "seq_rule":
        start, step = _num(args[0]), _num(args[1])
        vals = [start + i * step for i in range(rows)]
    elif kind == "int_rule":
        lo, hi = int(_num(args[0])), int(_num(args[1]))
        vals = [rng.randint(lo, hi) for _ in range(rows)]
    elif kind == "num_rule":
        lo, hi = float(_num(args[0])), float(_num(args[1]))
        vals = [round(rng.uniform(lo, hi), 2) for _ in range(rows)]
    elif kind == "dates_rule":
        d1 = dt.date.fromisoformat(str(args[0]))
        d2 = dt.date.fromisoformat(str(args[1]))
        span = (d2 - d1).days
        vals = [d1 + dt.timedelta(days=rng.randint(0, span)) for _ in range(rows)]
    elif kind == "text_rule":
        n = int(args[0])
        alphabet = string.ascii_uppercase + string.digits
        vals = ["".join(rng.choice(alphabet) for _ in range(n)) for _ in range(rows)]
    else:
        raise ValueError(f"unknown fixture rule {kind}")
    if blanks:
        for i in rng.sample(range(rows), min(blanks, rows)):
            vals[i] = None
    return vals


def build_workbooks(fixture_trees):
    """All FIXTURE lines of a spec -> {file: {sheet name: Sheet}}"""
    books = {}
    for tree in fixture_trees:
        fx = parse_fixture(tree)
        sheet = Sheet(fx["sheet"])
        for col, header in fx["cols"]:
            sheet.set(1, col, header)
            letter = ""
            n = col
            while n:
                n, rem = divmod(n - 1, 26)
                letter = chr(65 + rem) + letter
            for i, v in enumerate(_column_values(fx, col, letter), start=2):
                sheet.set(i, col, v)
        books.setdefault(fx["file"], {})[fx["sheet"]] = sheet
    return books
