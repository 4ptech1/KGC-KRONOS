"""Build FIXTURE test data: the one fixture generator of VBA Forge.

Everything that needs fixture data gets it from here:
  - the reference interpreter (vcl_interp), for TEST runs and mutation scores
  - the expected results (dist/expected/*.json), which come from those interpreter runs
  - the batch builder (`vclc build`), which turns the same values into the fixture-builder
    VBA in dist/forge_fixtures.bas, so Excel can rebuild exactly the same workbooks

The data is computed here, never with Rnd in VBA, so Python and Excel see identical values.

Seed rule (spec.md, FIXTURE SPEC): each column has its own fixed seed, made from file name +
sheet + column, used for everything random in that column. This algorithm must never change,
because expected results and the dataset depend on it. It is the algorithm of the pilot's
harness/fixtures.py (PR #4), so both build bit-for-bit the same values:

    seed      first 8 bytes (big-endian) of SHA-256("<file>|<sheet>|<column letter>")
    PRNG      SplitMix64 on that seed; below(n) = uniform 0..n-1 by rejection sampling
    values()  every value in order, then a Fisher-Yates shuffle (i from last to 1, j = below(i+1))
    seq()     start + i * step, no randomness
    int()     lo + below(hi - lo + 1), row by row
    num()     lo + (hi - lo) * next() / 2**64, rounded to 2 decimals, row by row
    dates()   first date + below(days + 1) days, row by row
    reldates(lo, hi)  TODAY + lo + below(hi - lo + 1) days, row by row
    text(n)   n letters A-Z, each "ABC...Z"[below(26)], row by row
Then the modifiers, always in this order:
    +boundary(v1, ..., vk)  rows 2 to k+1 hold exactly v1 to vk ("" is a blank cell). For
              values() the values are moved to the top, so the counts stay the same; for the
              other rules they replace the generated values. No PRNG draws.
    +blank(n) a partial Fisher-Yates over the rows below the boundary rows (all rows when
              there are none) picks n rows to blank, still drawing from the column's PRNG
    +dupes(n) with its own PRNG, seeded from "<file>|<sheet>|<column letter>|dupes": a partial
              Fisher-Yates over the non-blank rows below the boundary rows picks n targets;
              each target, in pick order, takes the value of source[below(len(source))], where
              source is the other non-blank rows below the boundary rows
A values() entry of "" is a blank cell, and so is a column without a data rule.
stale=<n> on a FIXTURE line marks n rows as leftovers of an earlier run, with a red (#FF0000)
fill across the header's columns. On a header-only sheet (rows=0) they are new rows below the
header, every cell holding the text "stale". On a sheet with data, a partial Fisher-Yates
with its own PRNG, seeded from "<file>|<sheet>|stale", picks the n data rows; their values stay.
"""
import datetime
import hashlib
from dataclasses import dataclass, field

from lark import Token, Tree

MASK = (1 << 64) - 1
LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
ROWS_PER_PROC = 250  # keeps each generated VBA procedure far below Excel's 64 KB limit
MAX_LINE = 1000      # VBA allows 1023 characters per line


# ---------------------------------------------------------------- the data model

class Sheet:
    """One worksheet as the interpreter sees it: values, fills and number formats by (row, col)."""

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


@dataclass(frozen=True)
class RelDate:
    """A date written relative to TODAY: TODAY + days. Resolved when a workbook is built."""
    days: int


STALE_FILL = "#FF0000"
STALE_TEXT = "stale"


@dataclass
class Rule:
    """One FIXTURE data rule with its modifiers, e.g. C:int(1,50)+boundary(0)+blank(3)."""
    col: str      # column letter
    kind: str     # values, seq, int, num, dates, reldates or text
    args: list    # values: [(text, count)]; dates: [date, date]; the others: numbers
    blank: int = 0
    boundary: list = field(default_factory=list)  # the first data rows: text, number, date or ""
    dupes: int = 0


@dataclass
class Fixture:
    """One FIXTURE line: one sheet of one fixture file."""
    file: str
    sheet: str
    headers: list                               # [(column letter, header)], in column order
    rows: int
    rules: dict = field(default_factory=dict)   # column letter -> Rule
    stale: int = 0


def col_number(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def col_letter(n):
    letters = ""
    while n:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


# ---------------------------------------------------------------- the generator

class ColumnRandom:
    """SplitMix64: small, fast and the same on every Python version."""

    def __init__(self, file, sheet, col):
        digest = hashlib.sha256(f"{file}|{sheet}|{col}".encode("utf-8")).digest()
        self.state = int.from_bytes(digest[:8], "big")

    def next64(self):
        self.state = (self.state + 0x9E3779B97F4A7C15) & MASK
        z = self.state
        z = ((z ^ (z >> 30)) * 0xBF58476D1CE4E5B9) & MASK
        z = ((z ^ (z >> 27)) * 0x94D049BB133111EB) & MASK
        return z ^ (z >> 31)

    def below(self, n):
        limit = (1 << 64) - ((1 << 64) % n)
        while True:
            x = self.next64()
            if x < limit:
                return x % n

    def shuffle(self, items):
        for i in range(len(items) - 1, 0, -1):
            j = self.below(i + 1)
            items[i], items[j] = items[j], items[i]


def _boundary_value(kind, value):
    """One +boundary value as the cell value it becomes."""
    if value == "":
        return None
    if kind == "reldates":
        if not isinstance(value, int):
            raise ValueError(f"reldates() boundary values are whole numbers of days, got {value!r}")
        return RelDate(value)
    if kind == "dates" and not isinstance(value, datetime.date):
        raise ValueError(f"dates() boundary values are yyyy-mm-dd dates, got {value!r}")
    if kind in ("seq", "int", "num") and not isinstance(value, (int, float)):
        raise ValueError(f"{kind}() boundary values are numbers, got {value!r}")
    if kind in ("values", "text") and not isinstance(value, str):
        raise ValueError(f'{kind}() boundary values are "text", got {value!r}')
    return value


def column_values(file, sheet, rule, rows):
    """The values of one column (rows 2 onwards) for one data rule and its modifiers."""
    rng = ColumnRandom(file, sheet, rule.col)
    kind = rule.kind
    if kind == "values":
        values = [text if text != "" else None for text, count in rule.args for _ in range(count)]
        if len(values) != rows:
            raise ValueError(f"{sheet}!{rule.col}: values() counts add up to {len(values)}, but rows={rows}")
        rng.shuffle(values)
    elif kind == "seq":
        start, step = rule.args
        values = [start + i * step for i in range(rows)]
    elif kind == "int":
        lo, hi = (int(a) for a in rule.args)
        values = [lo + rng.below(hi - lo + 1) for _ in range(rows)]
    elif kind == "num":
        lo, hi = (float(a) for a in rule.args)
        values = [round(lo + (hi - lo) * rng.next64() / 2 ** 64, 2) for _ in range(rows)]
    elif kind == "dates":
        first, last = rule.args
        days = (last - first).days
        values = [first + datetime.timedelta(days=rng.below(days + 1)) for _ in range(rows)]
    elif kind == "reldates":
        lo, hi = (int(a) for a in rule.args)
        values = [RelDate(lo + rng.below(hi - lo + 1)) for _ in range(rows)]
    elif kind == "text":
        length = int(rule.args[0])
        values = ["".join(LETTERS[rng.below(26)] for _ in range(length)) for _ in range(rows)]
    else:
        raise ValueError(f"unknown data rule {kind}()")

    top = len(rule.boundary)   # +boundary: the first data rows
    if top > rows:
        raise ValueError(f"{sheet}!{rule.col}: +boundary has {top} values, but rows={rows}")
    for i, want in enumerate(_boundary_value(kind, v) for v in rule.boundary):
        if kind == "values":   # move the value up, so the counts stay the same
            j = next((j for j in range(i, rows) if values[j] == want), None)
            if j is None:
                raise ValueError(f"{sheet}!{rule.col}: +boundary needs one more {want or ''!r} than values() has")
            values[i], values[j] = values[j], values[i]
        else:
            values[i] = want

    if rule.blank > rows - top:  # +blank: a partial Fisher-Yates over the rows below the boundary
        raise ValueError(f"{sheet}!{rule.col}: +blank({rule.blank}) is more than the {rows - top} rows it can use")
    order = list(range(top, rows))
    for i in range(rule.blank):
        j = i + rng.below(len(order) - i)
        order[i], order[j] = order[j], order[i]
        values[order[i]] = None

    if rule.dupes:             # +dupes: its own PRNG; targets and sources are non-blank rows
        drng = ColumnRandom(file, sheet, rule.col + "|dupes")
        pool = [i for i in range(top, rows) if values[i] is not None]
        if rule.dupes >= len(pool):
            raise ValueError(f"{sheet}!{rule.col}: +dupes({rule.dupes}) needs more than {rule.dupes} "
                             f"non-blank rows below the boundary rows, and there are {len(pool)}")
        for i in range(rule.dupes):
            j = i + drng.below(len(pool) - i)
            pool[i], pool[j] = pool[j], pool[i]
        targets, sources = pool[:rule.dupes], pool[rule.dupes:]
        for t in targets:
            values[t] = values[sources[drng.below(len(sources))]]
    return values


def sheet_rows(fx):
    """The data rows (row 2 onwards) of one FIXTURE line, as lists in column order. Relative
    dates stay RelDate here; build_workbooks resolves them."""
    columns = [column_values(fx.file, fx.sheet, fx.rules[letter], fx.rows) if letter in fx.rules
               else [None] * fx.rows for letter, _ in fx.headers]
    rows = [list(row) for row in zip(*columns)] if columns else []
    if fx.rows == 0:
        rows += [[STALE_TEXT] * len(fx.headers) for _ in range(fx.stale)]
    return rows


def stale_rows(fx):
    """The sheet row numbers (2 = first data row) that stale=<n> marks with a red fill."""
    if not fx.stale:
        return []
    if fx.rows == 0:
        return list(range(2, 2 + fx.stale))
    if fx.stale > fx.rows:
        raise ValueError(f"{fx.sheet}: stale={fx.stale} is more than rows={fx.rows}")
    rng = ColumnRandom(fx.file, fx.sheet, "stale")
    order = list(range(fx.rows))
    for i in range(fx.stale):
        j = i + rng.below(fx.rows - i)
        order[i], order[j] = order[j], order[i]
    return sorted(r + 2 for r in order[:fx.stale])


def resolve(value, today):
    return today + datetime.timedelta(days=value.days) if isinstance(value, RelDate) else value


# ---------------------------------------------------------------- FIXTURE lines

def _number(tok):
    s = str(tok)
    return float(s) if "." in s else int(s)


def _unquote(tok):
    return str(tok)[1:-1].replace('""', '"')


def parse_fixture(tree):
    """A fixture_value tree from vcl/grammar.lark -> Fixture."""
    kids = tree.children
    fx = Fixture(file=str(kids[0]), sheet=str(kids[1]), headers=[], rows=0)
    for item in kids[2:]:
        if isinstance(item, Tree) and item.data == "col_list":
            fx.headers = [(str(cd.children[0]), str(cd.children[1]).strip()) for cd in item.children]
        elif isinstance(item, Token) and item.type == "INT":
            fx.rows = int(item)
        elif isinstance(item, Tree) and item.data == "stale_opt":
            fx.stale = int(item.children[0])
        elif isinstance(item, Tree) and item.data == "data_list":
            for dd in item.children:
                letter = str(dd.children[0])
                rule = dd.children[1]
                kind = rule.data[:-len("_rule")]
                if kind == "values":
                    args = [(_unquote(v.children[0]), int(v.children[1])) for v in rule.children]
                elif kind == "dates":
                    args = [datetime.date.fromisoformat(str(a)) for a in rule.children]
                else:
                    args = [_number(a) for a in rule.children]
                fx.rules[letter] = r = Rule(letter, kind, args)
                for mod in dd.children[2:]:
                    if mod.data == "blank_mod":
                        r.blank = int(mod.children[0])
                    elif mod.data == "dupes_mod":
                        r.dupes = int(mod.children[0])
                    elif mod.data == "boundary_mod":
                        r.boundary = [_boundary_token(t) for t in mod.children]
    return fx


def _boundary_token(tok):
    if tok.type == "STRING":
        return _unquote(tok)
    if tok.type == "ISO_DATE":
        return datetime.date.fromisoformat(str(tok))
    return _number(tok)


def build_workbooks(fixture_trees, today=None):
    """All FIXTURE lines of a spec -> {file: {sheet name: Sheet}}, in FIXTURE line order.
    today resolves reldates() (default: the real date)."""
    today = today or datetime.date.today()
    books = {}
    for tree in fixture_trees:
        fx = parse_fixture(tree)
        sheet = Sheet(fx.sheet)
        for c, (_, header) in enumerate(fx.headers, start=1):
            sheet.set(1, c, header)
        for r, row in enumerate(sheet_rows(fx), start=2):
            for c, v in enumerate(row, start=1):
                sheet.set(r, c, resolve(v, today))
        for r in stale_rows(fx):
            for c in range(1, len(fx.headers) + 1):
                sheet.fills[(r, c)] = STALE_FILL
        books.setdefault(fx.file, {})[fx.sheet] = sheet
    return books


# ---------------------------------------------------------------- the fixture-builder VBA

def needs_text_guard(text):
    """True if Excel would turn this text into something else when it is written to a cell:
    a number, date, time, TRUE/FALSE or formula. Those get a ' prefix so they stay text.
    Being careful: any text with a digit counts ("W-01" stays safe either way)."""
    return (any(ch.isdigit() for ch in text) or text[:1] in "=+-@'"
            or text.strip().upper() in ("TRUE", "FALSE"))


def vba_literal(value):
    if value is None:
        return "Empty"
    if isinstance(value, bool):
        raise ValueError("fixtures don't produce TRUE/FALSE values")
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        text = repr(value)
        if "e" in text or "E" in text or "inf" in text or "nan" in text:
            raise ValueError(f"number {value!r} can't be written as a plain VBA literal")
        return text
    if isinstance(value, RelDate):   # relative to the day Excel builds the workbook
        return "Date" if value.days == 0 else f"Date {'+' if value.days > 0 else '-'} {abs(value.days)}"
    if isinstance(value, datetime.date):
        return f"DateSerial({value.year}, {value.month}, {value.day})"
    if isinstance(value, str):
        if not value.isascii():
            raise ValueError(f"fixture text must be ASCII for the VBA editor: {value!r}")
        if needs_text_guard(value):
            value = "'" + value
        return '"' + value.replace('"', '""') + '"'
    raise ValueError(f"unexpected fixture value {value!r}")


def proc_name(file, sheet=None, part=None):
    name = "ForgeFixture_" + file.rsplit(".", 1)[0]
    if sheet:
        name += "_" + sheet
    if part:
        name += f"_{part}"
    return name


def _write_line(target, row_no, values):
    line = f"    FixtureWriteRow {target}, {row_no}, Array({', '.join(vba_literal(v) for v in values)})"
    if len(line) > MAX_LINE:
        raise ValueError(f"fixture row {row_no} makes a VBA line of {len(line)} characters (limit {MAX_LINE})")
    return line


def vba_for_file(file, fixtures):
    """VBA procedures that build one fixture file's sheets in a workbook.

    The entry point is ForgeFixture_<file>(wb). Text that Excel would convert gets a ' prefix
    (see needs_text_guard) and blank cells are left empty, so Excel holds exactly the values
    computed here."""
    fixtures = [fx for fx in fixtures if fx.file == file]
    out = [f"' FIXTURE {file}: generated by compiler/vcl_fixtures.py from the spec's FIXTURE lines.",
           f"Public Sub {proc_name(file)}(ByVal wb As Workbook)"]
    out += [f"    {proc_name(file, fx.sheet)} wb" for fx in fixtures]
    out += ["End Sub", ""]
    for fx in fixtures:
        rows = [(n, row) for n, row in enumerate(sheet_rows(fx), start=2) if any(v is not None for v in row)]
        chunks = [rows[i:i + ROWS_PER_PROC] for i in range(0, len(rows), ROWS_PER_PROC)] or [[]]
        out += [f"Private Sub {proc_name(file, fx.sheet)}(ByVal wb As Workbook)",
                "    Dim ws As Worksheet",
                f"    Set ws = FixtureSheet(wb, {vba_literal(fx.sheet)})",
                _write_line("ws", 1, [h for _, h in fx.headers])]
        out += [_write_line("ws", n, row) for n, row in chunks[0]]
        for part in range(2, len(chunks) + 1):
            out.append(f"    {proc_name(file, fx.sheet, part)} ws")
        out += [f"    FixtureStaleFill ws, {r}, {len(fx.headers)}" for r in stale_rows(fx)]
        out += ["End Sub", ""]
        for part, chunk in enumerate(chunks[1:], start=2):
            out += [f"Private Sub {proc_name(file, fx.sheet, part)}(ByVal ws As Worksheet)"]
            out += [_write_line("ws", n, row) for n, row in chunk]
            out += ["End Sub", ""]
    return "\n".join(out)


FIXTURE_HELPERS = '''Private Function FixtureSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    ' The sheet called sheetName in wb, added at the end if it is missing, and emptied.
    Dim ws As Worksheet
    Dim sh As Worksheet
    For Each sh In wb.Worksheets
        If StrComp(sh.Name, sheetName, vbTextCompare) = 0 Then Set ws = sh
    Next sh
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set FixtureSheet = ws
End Function

Private Sub FixtureWriteRow(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal rowValues As Variant)
    Dim i As Long
    Dim n As Long
    Dim rowData() As Variant
    n = UBound(rowValues) - LBound(rowValues) + 1
    ReDim rowData(1 To 1, 1 To n)
    For i = 1 To n
        rowData(1, i) = rowValues(LBound(rowValues) + i - 1)
    Next i
    ws.Range(ws.Cells(rowNo, 1), ws.Cells(rowNo, n)).Value = rowData
End Sub

Private Sub FixtureStaleFill(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal nCols As Long)
    ' stale=<n>: a leftover red fill from an earlier run, across the header's columns.
    ws.Range(ws.Cells(rowNo, 1), ws.Cells(rowNo, nCols)).Interior.Color = RGB(255, 0, 0)
End Sub
'''


def fixtures_module(fixture_trees_by_spec):
    """One paste-ready module with a ForgeFixture_<file>(wb) builder for every fixture file.

    fixture_trees_by_spec: [(spec name, [fixture_value trees])]. A fixture file name used by two
    specs must build the same sheets, or the build fails."""
    seen = {}
    parts = []
    for spec_name, trees in fixture_trees_by_spec:
        fixtures = [parse_fixture(t) for t in trees]
        for file in dict.fromkeys(fx.file for fx in fixtures):
            mine = [fx for fx in fixtures if fx.file == file]
            if file in seen:
                if seen[file][1] != mine:
                    raise ValueError(f"fixture {file} is defined differently in {seen[file][0]} and {spec_name}")
                continue
            seen[file] = (spec_name, mine)
            parts.append(f"' ---- {spec_name}")
            parts.append(vba_for_file(file, mine))
    head = ["' VBA Forge test fixtures: one builder per fixture file, ForgeFixture_<file>(wb).",
            "' Generated by vclc from the specs' FIXTURE lines. Do not edit by hand.",
            "' Each builder writes the same synthetic data the reference interpreter tests with.",
            "Option Explicit", "", ""]
    return "\n".join(head) + FIXTURE_HELPERS + "\n" + "\n".join(parts) + "\n"
