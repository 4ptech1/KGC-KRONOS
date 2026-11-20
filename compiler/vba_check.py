"""The compiler's VBA, run in the VBA simulator, against the reference interpreter (D90 Phase 1, D99, D100).

For every TEST of a spec: build the TEST's practice workbook from its FIXTURE lines (the one fixture
generator, `vcl_fixtures.py`), run the compiled procedure on it in `vba_sim.py` as Excel would (the same
arguments, runs=n times), and compare what VBA leaves with what the interpreter's run of the same TEST
leaves (`vcl_interp.run_tests`): RESULT, ERROR_COUNT, the messages shown, and every sheet's cells (as
canonical text), fills, number formats and hidden rows, and the mock file system's folders, files and
texts. A BUILD runs through its modMain.RunHeadless, as the interpreter runs it.

    python compiler/vclc.py sim FILES... [--today YYYY-MM-DD]     (or python compiler/vba_check.py ...)

One line per difference, exit 1 when there is one. A spec that runs no TESTs (one that doesn't parse, or a
BUILD before VCL 0.5) is listed as skipped, never counted as proven.
"""
import datetime as dt
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "vcl"))

from lark import Token, Tree  # noqa: E402

import semantics as S  # noqa: E402
import vba_sim as V  # noqa: E402
from errcodes import coded  # noqa: E402
from vcl_codegen import assemble, compile_build, compile_text  # noqa: E402
from vcl_fixtures import SANDBOX, Formula, build_folders, build_workbooks  # noqa: E402
from vcl_interp import entry_of, files_of, run_tests, test_number, texts_of  # noqa: E402
from vcl_parse import parse  # noqa: E402

__all__ = ["check_text", "Outcome"]


class Outcome:
    """What one spec gave: proven, a list of differences, or a reason it was skipped."""

    def __init__(self, name):
        self.name, self.tests, self.differences, self.skipped = name, 0, [], None

    @property
    def proven(self):
        return self.skipped is None and not self.differences


# ------------------------------------------------------------------ values both ways

def to_sim(v):
    """A fixture value as Excel holds it."""
    if v is None:
        return V.EMPTY
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, (dt.date, dt.datetime)):
        return V.VDate.of(v)
    if isinstance(v, S.XlError):
        return V.VError({t: c for t, c in V._EXCEL_ERRORS.items()}[v.text])
    if isinstance(v, str):
        return v
    raise ValueError(f"fixture value {v!r}")


def from_sim(v):
    """An Excel value as the interpreter would hold it, for the canonical text."""
    if isinstance(v, V.VDate):
        d = v.to_datetime()
        return d.date() if v.serial == int(v.serial) else d
    if isinstance(v, V.VError):
        return S.XlError({c: t for t, c in V._EXCEL_ERRORS.items()}[v.code])
    if isinstance(v, float) and v.is_integer() and abs(v) < 1e15:
        return int(v)
    if v is V.EMPTY:
        return None
    return v


def colour_to_sim(hexcode):
    h = hexcode.lstrip("#")
    return int(h[0:2], 16) + int(h[2:4], 16) * 256 + int(h[4:6], 16) * 65536


def colour_from_sim(n):
    n = int(n)
    return f"#{n & 255:02X}{(n >> 8) & 255:02X}{(n >> 16) & 255:02X}"


def sim_book(fixture, sheets):
    book = V.Workbook(fixture)
    for name, sh in sheets.items():
        ws = book.add_sheet(name)
        for rc, v in sh.cells.items():
            if isinstance(v, Formula):   # a template's formula (D110): its text, and the value Excel saved
                ws.formulas[rc] = v.text
                if v.saved is not None:
                    ws.cells[rc] = to_sim(v.saved)
            else:
                ws.cells[rc] = to_sim(v)
        for rc, f in sh.fills.items():
            ws.fills[rc] = colour_to_sim(f)
        ws.formats.update(sh.formats)
        ws.hidden = set(sh.hidden)
    return book


def state_of(book):
    out = {}
    for ws in book.sheets:
        cells = {rc: S.text_of(from_sim(v)) for rc, v in ws.cells.items()}
        cells.update(ws.formulas)
        out[ws.name] = (sorted(cells.items()), sorted((rc, colour_from_sim(f)) for rc, f in ws.fills.items()),
                        sorted(ws.formats.items()), sorted(ws.hidden))
    return out


# ------------------------------------------------------------------ one spec

def _test_parts(t):
    kids = t.children
    fixture, rest = str(kids[0]), kids[1:]
    i, call_vals = 1, []
    while i < len(rest) and not (isinstance(rest[i], Token) and rest[i].type == "INT") and not (
            isinstance(rest[i], Tree) and rest[i].data == "assertion"):
        call_vals.append(rest[i])
        i += 1
    runs = int(rest[i]) if i < len(rest) and isinstance(rest[i], Token) and rest[i].type == "INT" else 1
    return fixture, call_vals, runs


def _same(a, b):
    if a is None and b in (None, V.EMPTY):
        return True
    try:
        return S.text_of(a) == S.text_of(from_sim(b))
    except Exception:
        return False


def check_text(text, name="spec", today=None, module=None):
    """module: the VBA to run instead of the compiler's (a test plants a bug in it; later, our model's VBA)."""
    today = today or dt.date.today()
    out = Outcome(name)
    spec, problems = parse(text)
    if problems or not spec:
        out.skipped = "it doesn't parse"
        return out
    if spec["kind"] == "BUILD" and spec.get("version") in ("VCL 0.1", "VCL 0.2", "VCL 0.3", "VCL 0.4"):
        out.skipped = "a BUILD before VCL 0.5 runs no TESTs (R044)"
        return out
    trees = [k["tree"] for k in spec["keys"] if k["key"] == "FIXTURE"]
    task, _, keys = entry_of(spec)
    proc = next(k["tree"] for k in task["keys"] if k["key"] == "PROC").children[0]
    if module is None:
        module = compile_build(text, name)[1] if spec["kind"] == "BUILD" else assemble([compile_text(text, name)])
    try:
        parsed = V.Module(module)
    except V.VbaUnsupported as e:
        out.differences.append(coded(f"{name}: the simulator can't read the module: {e}"))
        return out
    params = [str(p.children[0]) for k in task["keys"] if k["key"] == "ARGS" for p in k["tree"].children]
    expected = run_tests(text, today)
    tests = [k for k in keys if k["key"] == "TEST"]
    for k, want in zip(tests, expected):
        out.tests += 1
        fixture, call_vals, runs = _test_parts(k["tree"])
        books = build_workbooks(trees, today)
        book = sim_book(fixture, {n: sh.snapshot() for n, sh in books.get(fixture, {}).items()})
        fs = build_folders(trees)[fixture].snapshot()
        sim = V.Simulator([parsed], today=today, book=book, fs=fs)
        args = []
        for pname, v in zip(params, call_vals):
            if isinstance(v, Token) and v.type == "FILENAME":
                args.append(book)
            elif isinstance(v, Token) and v.type == "NAME":
                if any(ws.name == str(v) for ws in book.sheets):
                    args.append(book.sheet(str(v)))
                else:   # a FIXTURE folder (F6)
                    args.append(SANDBOX + "\\" + str(v))
            elif isinstance(v, Token) and v.type == "STRING":
                args.append(str(v)[1:-1].replace('""', '"'))
            else:
                n = test_number(v)
                args.append(float(n) if isinstance(n, float) else n)
        where = f"{name}: TEST on line {k['line']}"
        result = V.EMPTY
        try:
            for _ in range(runs):
                sim.call("ForgeClearLog")
                sim.messages.clear()
                sim.debug.clear()
                result = sim.call(str(proc), *args)
            errors = sim.call("ForgeErrorCount")
        except V.VbaUnsupported as e:
            out.differences.append(coded(f"{where}: the simulator can't run it: {e}"))
            continue
        except V.VbaError as e:
            out.differences.append(coded(f"{where}: the VBA stops with error {e.number} ({e.description})"))
            continue
        if not _same(want.result, result):
            out.differences.append(coded(f"{where}: RESULT is {S.text_of(from_sim(result))!r} in VBA, "
                                         f"{S.text_of(want.result)!r} in the interpreter"))
        if want.errors != errors:
            out.differences.append(coded(f"{where}: ERROR_COUNT is {errors} in VBA, {want.errors} in the interpreter"))
        want_msgs = [m for kind, m in want.messages if kind == "MSG"]
        want_logs = [m for kind, m in want.messages if kind == "LOG"]
        got_logs = [line.split(" | ", 1)[1] for line in sim.debug if line.startswith("ForgeInfo: ")]
        if want_msgs != sim.messages:
            out.differences.append(coded(f"{where}: MSG shows {sim.messages!r} in VBA, {want_msgs!r} in the interpreter"))
        if want_logs != got_logs:
            out.differences.append(coded(f"{where}: LOG writes {got_logs!r} in VBA, {want_logs!r} in the interpreter"))
        got_files, got_texts = files_of(fs), texts_of(fs)
        if [tuple(f) for f in want.files] != got_files:
            out.differences.append(coded(f"{where}: the files differ: {_first_difference_list(want.files, got_files)}"))
        if [tuple(f) for f in want.texts] != got_texts:
            out.differences.append(coded(f"{where}: the file texts differ: VBA has {got_texts!r}, "
                                         f"the interpreter {[tuple(f) for f in want.texts]!r}"))
        got = state_of(book)
        for sheet in sorted(set(want.state) | set(got)):
            if sheet not in got or sheet not in want.state:
                out.differences.append(coded(f"{where}: sheet {sheet} is only in "
                                             f"{'the interpreter' if sheet in want.state else 'VBA'}"))
                continue
            w_cells, w_fills, w_formats = want.state[sheet]
            g_cells, g_fills, g_formats, g_hidden = got[sheet]
            w_hidden = want.hidden.get(sheet, [])
            for what, w, g in (("cells", w_cells, g_cells), ("fills", [(rc, f.upper()) for rc, f in w_fills], g_fills),
                               ("number formats", w_formats, g_formats), ("hidden rows", w_hidden, g_hidden)):
                if list(w) != list(g):
                    out.differences.append(coded(f"{where}: {sheet} {what} differ: {_first_difference(w, g)}"))
    return out


def _first_difference(want, got):
    w, g = dict(want) if want and isinstance(want[0], tuple) else None, None
    if w is not None:
        g = dict(got)
        for key in sorted(set(w) | set(g)):
            if w.get(key) != g.get(key):
                return f"at {key}: VBA has {g.get(key)!r}, the interpreter {w.get(key)!r}"
    return f"VBA has {list(got)[:6]!r}, the interpreter {list(want)[:6]!r}"


def _first_difference_list(want, got):
    want = [tuple(f) for f in want]
    for w in sorted(set(want) ^ set(got)):
        return f"{w!r} is only in {'the interpreter' if w in want else 'VBA'}"
    return "the order differs"


def main(argv):
    today = None
    if "--today" in argv:
        i = argv.index("--today")
        today = dt.date.fromisoformat(argv[i + 1])
        argv = argv[:i] + argv[i + 2:]
    proven = skipped = 0
    bad = False
    for p in argv:
        name = Path(p).parent.name if Path(p).name == "spec.vcl" else Path(p).name
        o = check_text(Path(p).read_text(encoding="utf-8"), name, today)
        if o.skipped:
            skipped += 1
            print(f"SKIP {p}: {o.skipped}")
        elif o.differences:
            bad = True
            for d in o.differences:
                print(f"FAIL {d}")
        else:
            proven += 1
            print(f"OK   {p}: {o.tests} TESTs, the VBA ends where the interpreter does")
    print(f"{proven} proven, {skipped} skipped")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
