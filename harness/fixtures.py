"""Build FIXTURE data for VBA Forge tests (VCL 0.1 revision 3), and the VBA that writes it.

    from fixtures import build_file, vba_for_file
    sheets = build_file(spec.fixtures, "bom_grn_01.xlsx")   # {sheet name: SheetData}
    vba = vba_for_file("bom_grn_01.xlsx", sheets)            # paste-ready VBA procedures

The data is computed here, never with Rnd in VBA, so Python and Excel see identical values.
The VBA writes them as literals into a temporary workbook.

Seed rule (spec.md, FIXTURE SPEC): each column has its own fixed seed, made from file name +
sheet + column, used for everything random in that column. This algorithm must never change,
because expected results and the dataset depend on it:

    seed      first 8 bytes (big-endian) of SHA-256("<file>|<sheet>|<column letter>")
    PRNG      SplitMix64 on that seed; below(n) = uniform 0..n-1 by rejection sampling
    values()  every value in order, then a Fisher-Yates shuffle (i from last to 1, j = below(i+1))
    seq()     start + i * step, no randomness
    int()     lo + below(hi - lo + 1), row by row
    num()     lo + (hi - lo) * next() / 2**64, rounded to 2 decimals, row by row
    dates()   first date + below(days + 1) days, row by row
    text(n)   n letters A-Z, each "ABC...Z"[below(26)], row by row
    +blank(n) after the values, a partial Fisher-Yates over the row numbers picks n rows to blank
A values() entry of "" is a blank cell, and so is a column without a data rule.
"""
import datetime
import hashlib
import re
from dataclasses import dataclass

MASK = (1 << 64) - 1
LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
ROWS_PER_PROC = 250  # keeps each generated VBA procedure far below Excel's 64 KB limit
MAX_LINE = 1000      # VBA allows 1023 characters per line


@dataclass
class SheetData:
    file: str
    sheet: str
    headers: list  # header texts, in column order A, B, C, ...
    rows: list     # data rows (row 2 onwards); each value is int, float, str, datetime.date or None


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


def _number(text):
    return int(text) if text.lstrip("-").isdigit() else float(text)


def column_values(file, sheet, col, rule, rows):
    """The values of one column (rows 2 onwards) for one data rule."""
    rng = ColumnRandom(file, sheet, col)
    kind = rule.kind
    if kind == "values":
        values = [text if text != "" else None for text, count in rule.args for _ in range(count)]
        if len(values) != rows:
            raise ValueError(f"{sheet}!{col}: values() counts add up to {len(values)}, but rows={rows}")
        rng.shuffle(values)
    elif kind == "seq":
        start, step = (_number(a) for a in rule.args)
        values = [start + i * step for i in range(rows)]
    elif kind == "int":
        lo, hi = (int(a) for a in rule.args)
        values = [lo + rng.below(hi - lo + 1) for _ in range(rows)]
    elif kind == "num":
        lo, hi = (float(a) for a in rule.args)
        values = [round(lo + (hi - lo) * rng.next64() / 2 ** 64, 2) for _ in range(rows)]
    elif kind == "dates":
        first, last = (datetime.date.fromisoformat(a) for a in rule.args)
        days = (last - first).days
        values = [first + datetime.timedelta(days=rng.below(days + 1)) for _ in range(rows)]
    elif kind == "text":
        length = int(rule.args[0])
        values = ["".join(LETTERS[rng.below(26)] for _ in range(length)) for _ in range(rows)]
    else:
        raise ValueError(f"unknown data rule {kind}()")
    if rule.blank > rows:
        raise ValueError(f"{sheet}!{col}: +blank({rule.blank}) is more than rows={rows}")
    order = list(range(rows))
    for i in range(rule.blank):  # partial Fisher-Yates: the first n picks are the blank rows
        j = i + rng.below(rows - i)
        order[i], order[j] = order[j], order[i]
        values[order[i]] = None
    return values


def build_sheet(fixture):
    """Generate one FIXTURE line's sheet."""
    columns = []
    rules = {rule.col: rule for rule in fixture.rules}
    for col, _ in fixture.headers:
        rule = rules.get(col)
        columns.append(column_values(fixture.file, fixture.sheet, col, rule, fixture.rows) if rule
                       else [None] * fixture.rows)
    rows = [list(row) for row in zip(*columns)] if columns else []
    return SheetData(fixture.file, fixture.sheet, [h for _, h in fixture.headers], rows[:fixture.rows])


def build_file(fixtures, file):
    """All sheets of one fixture file, in FIXTURE line order: {sheet name: SheetData}."""
    sheets = {fx.sheet: build_sheet(fx) for fx in fixtures if fx.file == file}
    if not sheets:
        raise ValueError(f"no FIXTURE lines build {file}")
    return sheets


# ---------------------------------------------------------------- VBA

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
    line = f"    ForgeWriteRow {target}, {row_no}, Array({', '.join(vba_literal(v) for v in values)})"
    if len(line) > MAX_LINE:
        raise ValueError(f"fixture row {row_no} makes a VBA line of {len(line)} characters (limit {MAX_LINE})")
    return line


def vba_for_file(file, sheets):
    """VBA procedures that build one fixture file in a temporary workbook.

    The entry point is ForgeFixture_<file>(wb). It relies on ForgeFixtureSheet and ForgeWriteRow from
    harness/ForgeTest.bas. Text that Excel would convert gets a ' prefix (see needs_text_guard) and
    blank cells are left empty, so Excel holds exactly the values computed here.
    """
    out = [f"' FIXTURE {file}: generated by harness/fixtures.py from the spec's FIXTURE lines. Do not edit.",
           f"Private Sub {proc_name(file)}(ByVal wb As Workbook)"]
    for sheet in sheets:
        out.append(f"    {proc_name(file, sheet)} wb")
    out += ["End Sub", ""]
    for sheet, data in sheets.items():
        rows = [(n, row) for n, row in enumerate(data.rows, start=2) if any(v is not None for v in row)]
        chunks = [rows[i:i + ROWS_PER_PROC] for i in range(0, len(rows), ROWS_PER_PROC)] or [[]]
        out += [f"Private Sub {proc_name(file, sheet)}(ByVal wb As Workbook)",
                "    Dim ws As Worksheet",
                f"    Set ws = ForgeFixtureSheet(wb, {vba_literal(sheet)})",
                _write_line("ws", 1, data.headers)]
        out += [_write_line("ws", n, row) for n, row in chunks[0]]
        for part in range(2, len(chunks) + 1):
            out.append(f"    {proc_name(file, sheet, part)} ws")
        out += ["End Sub", ""]
        for part, chunk in enumerate(chunks[1:], start=2):
            out += [f"Private Sub {proc_name(file, sheet, part)}(ByVal ws As Worksheet)"]
            out += [_write_line("ws", n, row) for n, row in chunk]
            out += ["End Sub", ""]
    return "\n".join(out)
