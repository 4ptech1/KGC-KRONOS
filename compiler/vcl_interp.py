"""Reference interpreter for VCL 0.1 and the VCL 0.2 value functions (milestone M3).

run_tests(text) -> list of TestResult. Executes the spec's OP tree on fixture workbooks built by
vcl_fixtures, with the meanings defined in vcl/semantics.py, and evaluates every TEST.
Error model: a failed PRE logs once, changes nothing and returns -1; a runtime error logs once
and returns -2 (Functions).
"""
import datetime as dt
import fnmatch
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

from lark import Token, Tree

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "vcl"))
import semantics as S  # noqa: E402
from vcl_fixtures import (EXCEL_SANDBOX_MAX, MAX_FILE_PATH, MAX_FOLDER_PATH, SANDBOX, MockFS,  # noqa: E402
                          Sheet, build_folders, build_workbooks, col_number)
from vcl_parse import parse  # noqa: E402


class VclError(Exception):
    pass


class _Return(Exception):
    def __init__(self, value):
        self.value = value


class _Exit(Exception):
    pass


class _Skip(Exception):
    """SKIP_IF: go on with the next pass of the innermost loop."""


class _Break(Exception):
    """EXIT_LOOP_IF: leave the innermost loop."""


@dataclass
class Rng:
    sheet: Sheet
    r1: int
    c1: int
    r2: int
    c2: int

    def cells(self):
        for r in range(self.r1, self.r2 + 1):
            for c in range(self.c1, self.c2 + 1):
                yield r, c


class FDict:
    """Dictionary with S10 keys: canonical text, case ignored; keeps the first spelling."""
    def __init__(self):
        self.items = {}

    def key(self, k):
        return S.dict_key(k)

    def exists(self, k):
        return self.key(k) in self.items

    def get(self, k):
        return self.items.get(self.key(k), (None, None))[1]

    def put(self, k, v):
        kk = self.key(k)
        orig = self.items[kk][0] if kk in self.items else S.text_of(k)
        self.items[kk] = (orig, v)

    def keys(self):
        return [orig for orig, _ in self.items.values()]


class Arr:
    """A 2-D array from READ_ARRAY: 1-based rows and columns, as Range.Value gives them (D36)."""

    def __init__(self, rows, cols, data):
        self.rows, self.cols, self.data = rows, cols, data   # data: {(row, col): value}

    def check(self, r, c):
        if not (1 <= r <= self.rows and 1 <= c <= self.cols):
            raise VclError(f"subscript out of range (VBA error 9): ARR({r}, {c}) of a {self.rows} x {self.cols} array")

    def get(self, r, c):
        self.check(r, c)
        return self.data.get((r, c))

    def set(self, r, c, v):
        self.check(r, c)
        self.data[(r, c)] = v


class Book:
    """THIS_BOOK and ACTIVE_BOOK in a TEST: the fixture workbook, its sheets in FIXTURE order."""

    def __init__(self, sheets):
        self.sheets = sheets   # name -> Sheet

    def sheet(self, name):
        for key, sh in self.sheets.items():
            if key.casefold() == S.text_of(name).casefold():   # Worksheets("name") ignores case
                return sh
        raise VclError(f"subscript out of range (VBA error 9): no sheet {S.text_of(name)!r} in the workbook")

    def exists(self, name):
        """SHEET_EXISTS: a sheet of that name, case ignored as Excel does."""
        return any(key.casefold() == S.text_of(name).casefold() for key in self.sheets)

    def add(self, name, after=None):
        """ADD_SHEET (D40): a new empty sheet after `after` (else after the last one). A bad name or one
        already used is an error before anything is added. The dict is changed in place, so the TEST's
        state and assertions see the new sheet."""
        text = S.text_of(name)
        if not S.sheet_name_ok(text):
            raise VclError(f"Excel error 1004: ADD_SHEET: {text!r} is not a valid sheet name (S23)")
        if self.exists(text):
            raise VclError(f"Excel error 1004: ADD_SHEET: a sheet named {text!r} already exists")
        items = list(self.sheets.items())
        at = len(items) if after is None else next((i + 1 for i, (_, sh) in enumerate(items) if sh is after), None)
        if at is None:
            raise VclError("Excel error 1004: ADD_SHEET: after= is not a sheet of this workbook")
        new = Sheet(text)
        items.insert(at, (text, new))
        self.sheets.clear()
        self.sheets.update(items)
        return new

    def copy(self, src, name, after=None):
        """COPY_SHEET (D41): a copy of src (values, fills and number formats) with this name, placed as
        ADD_SHEET places a sheet, and checked the same way before anything is copied."""
        if not any(sh is src for sh in self.sheets.values()):
            raise VclError("COPY_SHEET: the sheet to copy is not in this TEST's workbook")
        new = self.add(name, after)
        new.cells, new.fills, new.formats = dict(src.cells), dict(src.fills), dict(src.formats)
        return new


EXCEL_ROWS, EXCEL_COLS = 1048576, 16384   # the size of an Excel 2016 sheet (XFD1048576)
A1_REF = re.compile(r"\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?")


def a1_range(sheet, ref):
    """RANGE(ws, "B2") or RANGE(ws, "B2:D9") (D36). Other refs stop the TEST with a clear message."""
    m = A1_REF.fullmatch(S.text_of(ref).strip())
    if not m:
        raise VclError(f"RANGE {S.text_of(ref)!r}: the interpreter reads a cell (\"B2\") or a block (\"B2:D9\") only")
    c1, r1 = col_number(m.group(1).upper()), int(m.group(2))
    c2, r2 = (col_number(m.group(3).upper()), int(m.group(4))) if m.group(3) else (c1, r1)
    if min(r1, r2, c1, c2) < 1 or max(r1, r2) > EXCEL_ROWS or max(c1, c2) > EXCEL_COLS:
        raise VclError(f"RANGE {S.text_of(ref)!r}: outside an Excel sheet (rows 1 to {EXCEL_ROWS}, columns A to XFD); "
                       "Excel error 1004")
    return Rng(sheet, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))


def opts_of(args):
    return {str(a.children[0]): a.children[1] for a in args if isinstance(a, Tree) and a.data == "option"}


def pos_of(args):
    return [a for a in args if not (isinstance(a, Tree) and a.data == "option")]


def call_args(call):
    return call.children[1].children if len(call.children) > 1 else []


# ---------- VCL 0.3 file ops on the mock file system (E11, router Phase 2) ----------
WILDCARDS = "*?"        # FSO reads them as patterns, so a path with one is never used


def make_folder(fs, root, path):
    """MAKE_FOLDER's status (E11). It makes the folder and every missing folder between it and its
    root, never the root itself, and checks the whole way down before making anything."""
    if not path or not root or any(ch in path for ch in WILDCARDS):
        return "Bad name"
    if len(path) > MAX_FOLDER_PATH:
        return "Too long"
    entry = fs.get(path)
    if entry is not None:
        return "Exists" if entry["kind"] == "folder" else "Blocked by file"
    if not path.casefold().startswith(root.casefold() + "\\"):
        return "Bad name"
    if not fs.is_folder(root):
        return "No folder"
    levels, cur = [], root
    for part in path[len(root) + 1:].split("\\"):
        cur += "\\" + part
        levels.append(cur)
        if fs.is_file(cur):
            return "Blocked by file"
    for level in levels:
        fs.add_folder(level)
    return "Created"


def copy_file(fs, src, dst):
    """COPY_FILE's status (E11). One attempt; it never overwrites, deletes or makes folders."""
    if not src or not dst or any(ch in src + dst for ch in WILDCARDS):
        return "Bad name"
    if len(src) > MAX_FILE_PATH or len(dst) > MAX_FILE_PATH:
        return "Too long"
    source = fs.get(src)
    if source is None or source["kind"] != "file":
        return "Missing"
    if "\\" not in dst or not fs.is_folder(dst.rsplit("\\", 1)[0]):
        return "No folder"
    target = fs.get(dst)
    if target is not None:
        return "Already there" if target["kind"] == "file" and target["size"] == source["size"] else "Conflict"
    if source["locked"]:
        return "Locked"
    fs.add_file(dst, source["size"])
    return "Copied"


def delete_row(sh, row):
    """Delete a whole sheet row: the rows below move up, with their fills and formats."""
    sh.cells = {(r - (r > row), c): v for (r, c), v in sh.cells.items() if r != row}
    sh.fills = {(r - (r > row), c): v for (r, c), v in sh.fills.items() if r != row}
    sh.formats = {(r - (r > row), c): v for (r, c), v in sh.formats.items() if r != row}


class Machine:
    def __init__(self, spec, bound_args, log, today=None, fs=None, book=None):
        self.spec = spec
        self.log = log
        self.book = book if isinstance(book, Book) else Book(book if book is not None else {})
        self.messages = []   # LOG and MSG texts: never errors, never part of ERROR_COUNT (D36)
        self.fs = fs if fs is not None else MockFS()
        self.args = bound_args
        self.today = today or dt.date.today()
        self.vars = {}
        self.types = {}
        self.proc = ""
        self.pre_failed = False
        self.file_paths = []   # (path, its length limit) for every file op (D30)
        self.kind = "Sub"
        self.ret = None
        for k in spec["keys"]:
            t = k["tree"]
            if t is None:
                continue
            if k["key"] == "PROC":
                self.proc = str(t.children[0])
            elif k["key"] == "KIND":
                self.kind = str(t.children[0])
            elif k["key"] == "RET":
                self.ret = str(t.children[0])
            elif k["key"] == "VAR":
                name, typ = str(t.children[0]), str(t.children[1])
                self.types[name] = typ
                self.vars[name] = {"Long": 0, "Double": 0.0, "String": "", "Boolean": False,
                                   "Dictionary": FDict(), "Collection": []}.get(typ)

    # ---------- expressions ----------
    def ev(self, t):
        if isinstance(t, Token):
            k = t.type
            if k == "NUMBER":
                return S.to_number(str(t))
            if k == "STRING":
                return str(t)[1:-1].replace('""', '"')
            if k == "COLOR":
                return str(t).upper()
            if k == "ARG":
                return self.args[str(t)[1:]]
            name = str(t)
            if name == "TODAY":
                return self.today
            if name == "NOW":
                return dt.datetime.combine(self.today, dt.datetime.now().time())
            if name in ("TRUE", "FALSE"):
                return name == "TRUE"
            if name in ("THIS_BOOK", "ACTIVE_BOOK"):
                return self.book
            if name == "NOTHING":
                return None
            if name == "ACTIVE_SHEET":
                raise VclError("ACTIVE_SHEET depends on what was active in Excel: pass the sheet as an argument")
            if name in self.vars:
                return self.vars[name]
            return name  # an option word such as up, sum, desc
        d = t.data
        if d == "call":
            return self.ev_call(t)
        if d == "not_op":
            return not self.ev(t.children[0])
        if d == "and_expr":
            vals = [self.ev(c) for c in t.children]   # VBA evaluates every side
            return all(vals)
        if d == "or_expr":
            vals = [self.ev(c) for c in t.children]
            return any(vals)
        if d == "neg":
            return -self.num(self.ev(t.children[0]))
        if d == "concat":
            return "".join(S.text_of(self.ev(c)) for c in t.children)
        if d in ("sum", "product"):
            cur = self.ev(t.children[0])
            for i in range(1, len(t.children), 2):
                cur = self.arith(cur, str(t.children[i]), self.ev(t.children[i + 1]))
            return cur
        if d == "comparison":
            a, op, b = self.ev(t.children[0]), str(t.children[1]), self.ev(t.children[2])
            return self.compare(a, op, b)
        raise VclError(f"can't evaluate {d}")

    def index(self, v):
        """A row or column number. Excel raises error 1004 for anything below 1."""
        n = int(self.num(v))
        if n < 1:
            raise VclError("Excel error 1004: row or column number below 1")
        return n

    def num(self, v):
        if isinstance(v, (dt.date, dt.datetime)):
            return v
        if S.is_number(v):
            return S.to_number(v)
        raise VclError(f"type mismatch: {S.text_of(v)!r} is not a number")

    def arith(self, a, op, b):
        a, b = self.num(a), self.num(b)
        if isinstance(a, dt.date) and not isinstance(b, dt.date):
            return a + dt.timedelta(days=b) if op == "+" else a - dt.timedelta(days=b)
        if isinstance(a, dt.date) and isinstance(b, dt.date) and op == "-":
            return (a - b).days
        if op == "+":
            return a + b
        if op == "-":
            return a - b
        if op == "*":
            return a * b
        if b == 0:
            raise VclError("division by zero")
        return a / b

    def compare(self, a, op, b):
        if op in ("=", "<>"):
            eq = S.equals(a, b) if not (isinstance(a, str) or isinstance(b, str)) or True else None
            return eq if op == "=" else not eq
        if isinstance(a, str) and isinstance(b, str):
            x, y = a, b
        else:
            x, y = self.num(a), self.num(b)
        return {"<": x < y, ">": x > y, "<=": x <= y, ">=": x >= y}[op]

    def target(self, t):
        """A range used as the target of SET_VALUE, CLEAR, FORMAT, SORT."""
        v = self.ev_call(t, as_target=True) if isinstance(t, Tree) and t.data == "call" else self.ev(t)
        if v is None:   # NOTHING, as FIND gives when no cell matches
            raise VclError("object variable not set (VBA error 91): the range is NOTHING; test it with IS_NOTHING first")
        if not isinstance(v, Rng):
            raise VclError("expected a range")
        return v

    def ev_call(self, t, as_target=False):
        name = str(t.children[0])
        args = call_args(t)
        pos = pos_of(args)
        e = self.ev
        if name == "CELL":
            sh, r, c = e(pos[0]), self.index(e(pos[1])), self.index(e(pos[2]))
            return Rng(sh, r, c, r, c) if as_target else sh.get(r, c)
        if name == "ROW_RANGE":
            sh = e(pos[0])
            r, c1, c2 = self.index(e(pos[1])), self.index(e(pos[2])), self.index(e(pos[3]))
            return Rng(sh, r, min(c1, c2), r, max(c1, c2))
        if name == "USED_RANGE":
            sh = e(pos[0])
            return Rng(sh, 1, 1, sh.max_row(), sh.max_col())
        if name == "COL":
            sh = e(pos[0])
            return S.col([sh.get(1, c) for c in range(1, sh.max_col() + 1)], e(pos[1]))
        if name == "LASTROW":
            sh, c = e(pos[0]), self.index(e(pos[1]))
            return S.last_row([sh.get(r, c) for r in range(1, sh.max_row() + 1)])
        if name == "LASTCOL":
            sh, r = e(pos[0]), self.index(e(pos[1]))
            return S.last_row([sh.get(r, c) for c in range(1, sh.max_col() + 1)])
        if name == "TRIM":
            return S.trim(e(pos[0]))
        if name == "UPPER":
            return S.upper(e(pos[0]))
        if name == "LOWER":
            return S.lower(e(pos[0]))
        if name == "LEN":
            return S.length(e(pos[0]))
        if name == "LEFT":
            return S.left(e(pos[0]), self.num(e(pos[1])))
        if name == "RIGHT":
            return S.right(e(pos[0]), self.num(e(pos[1])))
        if name == "IS_EMPTY":
            return S.is_empty(e(pos[0]))
        if name == "IS_NUMBER":
            return S.is_number(e(pos[0]))
        if name == "IS_DATE":
            return S.is_date(e(pos[0]))
        if name == "SAME":
            return S.same(e(pos[0]), e(pos[1]))
        if name == "CLEAN_TEXT":
            return S.clean_text(e(pos[0]))
        if name == "IS_ERROR":
            return S.is_error(e(pos[0]))
        if name == "PARSE_DATE":
            return S.parse_date(e(pos[0]), e(pos[1]))
        if name == "IIF":
            cond, a, b = e(pos[0]), e(pos[1]), e(pos[2])   # VBA's IIf evaluates both values
            return S.iif(cond, a, b)
        if name == "MOD":
            a, b = self.num(e(pos[0])), self.num(e(pos[1]))
            if b == 0:
                raise VclError("division by zero")
            return S.mod(a, b)
        if name == "SAFE_NAME":
            return S.safe_name(e(pos[0]))
        if name == "PATH_JOIN":
            return S.path_join(*[e(a) for a in pos])
        if name == "FOLDER_EXISTS":
            return self.fs.is_folder(S.text_of(e(pos[0])))
        if name == "EXISTS":
            return e(pos[0]).exists(e(pos[1]))
        if name == "COUNT":
            d = e(pos[0])
            return len(d.items) if isinstance(d, FDict) else len(d)
        if name == "KEYS":
            return e(pos[0]).keys()
        # rev 3 functions the compiler already compiled (VEP 0.4 Phase 1, D36)
        if name == "RANGE":
            return a1_range(e(pos[0]), e(pos[1]))
        if name == "SHEET":
            return e(pos[0]).sheet(e(pos[1]))
        if name == "SHEETS":
            return list(e(pos[0]).sheets.values())
        if name == "CELLS":
            rng = self.target(pos[0])
            return [rng.sheet.get(r, c) for r, c in rng.cells()]   # row by row, as For Each
        if name == "ARR":
            return e(pos[0]).get(self.index(e(pos[1])), self.index(e(pos[2])))
        if name == "ROWS":
            return e(pos[0]).rows
        if name == "COLS":
            return e(pos[0]).cols
        if name == "FILE_NAME":
            return S.file_name(e(pos[0]))
        if name == "FIND":   # S22: the first matching cell, row by row, or NOTHING (D39)
            rng = self.target(pos[0])
            cells = list(rng.cells())
            i = S.find([rng.sheet.get(r, c) for r, c in cells], e(pos[1]), str(opts_of(args).get("match", "whole")))
            return Rng(rng.sheet, *cells[i - 1], *cells[i - 1]) if i else None
        if name == "IS_NOTHING":
            return e(pos[0]) is None
        if name == "SHEET_EXISTS":   # VCL 0.4 (D40)
            return e(pos[0]).exists(e(pos[1]))
        raise VclError(f"{name} is not supported by the MVP interpreter")

    # ---------- statements ----------
    def assign(self, name, value):
        typ = self.types.get(name)
        if typ == "Long" and not isinstance(value, bool):
            value = round(self.num(value)) if value is not None else 0   # VBA rounds half to even
        elif typ == "Double":
            value = float(self.num(value)) if value is not None else 0.0
        elif typ == "String":
            value = S.text_of(value)
        self.vars[name] = value

    def action(self, call):
        name = str(call.children[0])
        args = call_args(call)
        pos, opts = pos_of(args), opts_of(args)
        e = self.ev
        if name == "LET":
            self.assign(str(pos[0]), e(pos[1]))
        elif name == "SET_VALUE" and isinstance(pos[0], Tree) and str(pos[0].children[0]) == "ARR":
            ap = pos_of(call_args(pos[0]))
            e(ap[0]).set(self.index(e(ap[1])), self.index(e(ap[2])), e(pos[1]))
        elif name == "SET_VALUE":
            rng = self.target(pos[0])
            v = e(pos[1])
            for r, c in rng.cells():
                rng.sheet.set(r, c, v)
        elif name == "CLEAR":
            rng = self.target(pos[0])
            what = str(opts["what"]) if "what" in opts else "all"
            for r, c in rng.cells():
                if what in ("all", "contents"):
                    rng.sheet.set(r, c, None)
                if what in ("all", "formats"):
                    rng.sheet.fills.pop((r, c), None)
                    rng.sheet.formats.pop((r, c), None)
        elif name == "FORMAT":
            rng = self.target(pos[0])
            for r, c in rng.cells():
                if "fill" in opts:
                    f = e(opts["fill"])
                    if str(f).lower() == "none":
                        rng.sheet.fills.pop((r, c), None)
                    elif str(f).startswith("#"):
                        rng.sheet.fills[(r, c)] = str(f).upper()
                    else:
                        rng.sheet.fills[(r, c)] = S.FILL_COLORS[str(f).lower()]
                if "numfmt" in opts:
                    rng.sheet.formats[(r, c)] = e(opts["numfmt"])
        elif name == "WRITE_ROW":
            sh, r = e(pos[0]), self.index(e(pos[1]))
            c = self.index(e(opts["col"])) if "col" in opts else 1
            for i, v in enumerate([e(a) for a in pos[2:]]):
                sh.set(r, c + i, v)
        elif name == "COPY_ROW":
            src, r, dst, r2 = e(pos[0]), self.index(e(pos[1])), e(pos[2]), self.index(e(pos[3]))
            n = self.index(e(opts["cols"]))
            for c in range(1, n + 1):
                dst.set(r2, c, src.get(r, c))
        elif name == "INSERT_ROW":   # a new empty row, no fill, no number format (D36)
            sh, row = e(pos[0]), self.index(e(pos[1]))
            sh.cells = {(r + (r >= row), c): v for (r, c), v in sh.cells.items()}
            sh.fills = {(r + (r >= row), c): v for (r, c), v in sh.fills.items()}
            sh.formats = {(r + (r >= row), c): v for (r, c), v in sh.formats.items()}
        elif name == "READ_ARRAY":   # always 2-D and 1-based, even for one cell (ForgeReadArray)
            rng = self.target(pos[0])
            data = {(r - rng.r1 + 1, c - rng.c1 + 1): rng.sheet.get(r, c) for r, c in rng.cells()}
            self.assign(str(opts["into"]), Arr(rng.r2 - rng.r1 + 1, rng.c2 - rng.c1 + 1, data))
        elif name == "WRITE_ARRAY":   # from the top-left cell of to=; text stays text (S17)
            arr, top = e(pos[0]), self.target(opts["to"])
            for r in range(1, arr.rows + 1):
                for c in range(1, arr.cols + 1):
                    top.sheet.set(top.r1 + r - 1, top.c1 + c - 1, arr.data.get((r, c)))
        elif name in ("LOG", "MSG"):
            self.messages.append((name, S.text_of(e(pos[0]))))
        elif name == "DELETE_ROW":
            delete_row(e(pos[0]), int(e(pos[1])))
        elif name == "DEDUPE":   # S21: keep the first row of each key, delete the later whole rows (D38)
            rng, cols = self.target(pos[0]), opts["cols"]
            items = cols.children if isinstance(cols, Tree) and cols.data == "list_lit" else [cols]
            keys = [int(self.num(e(c))) for c in items]
            first = rng.r1 + (1 if str(opts.get("header", "yes")) == "yes" else 0)
            bad = [c for c in keys if not 1 <= c <= EXCEL_COLS]
            if bad and first <= rng.r2:   # Excel reads a key cell only when there is a row to check
                raise VclError(f"Excel error 1004: DEDUPE cols= column {bad[0]} is outside the sheet")
            seen, doomed = set(), []
            for r in range(first, rng.r2 + 1):   # marked top down ...
                key = S.row_key([rng.sheet.get(r, c) for c in keys])
                if key is not None and key in seen:
                    doomed.append(r)
                elif key is not None:
                    seen.add(key)
            for r in reversed(doomed):   # ... deleted bottom up, as ForgeDedupe does
                delete_row(rng.sheet, r)
        elif name in ("ADD_SHEET", "COPY_SHEET"):   # VCL 0.4 (D40, D41): never a replacement
            after = e(opts["after"]) if "after" in opts else None
            if name == "ADD_SHEET":
                new = e(pos[0]).add(e(opts["name"]), after)
            else:
                new = self.book.copy(e(pos[0]), e(opts["name"]), after)
            if "into" in opts:
                self.assign(str(opts["into"]), new)
        elif name == "DICT_ADD":
            d, k, v = e(pos[0]), e(pos[1]), e(pos[2])
            mode = str(opts["mode"])
            if mode == "overwrite":
                d.put(k, v)
            elif mode == "keep":
                if not d.exists(k):
                    d.put(k, v)
            elif mode == "sum":
                d.put(k, (d.get(k) or 0) + self.num(v))
            elif mode == "count":
                d.put(k, (d.get(k) or 0) + 1)
        elif name == "DICT_GET":
            self.assign(str(opts["into"]), e(pos[0]).get(e(pos[1])))
        elif name == "SORT":
            self.sort(self.target(pos[0]), int(self.num(e(opts["key"]))),
                      str(opts.get("order", "asc")) == "asc", str(opts.get("header", "yes")) == "yes")
        elif name == "MAKE_FOLDER":
            target = pos[0]
            is_join = isinstance(target, Tree) and target.data == "call" and str(target.children[0]) == "PATH_JOIN"
            root = S.text_of(e(pos_of(call_args(target))[0] if is_join else target)).rstrip("\\")
            path = S.text_of(e(target))
            self.file_paths.append((path, MAX_FOLDER_PATH))
            self.assign(str(opts["into"]), make_folder(self.fs, root, path))
        elif name == "COPY_FILE":
            src, dst = S.text_of(e(pos[0])), S.text_of(e(pos[1]))
            self.file_paths += [(src, MAX_FILE_PATH), (dst, MAX_FILE_PATH)]
            self.assign(str(opts["into"]), copy_file(self.fs, src, dst))
        elif name == "RETURN":
            raise _Return(e(pos[0]) if pos else None)
        else:
            raise VclError(f"{name} is not supported by the MVP interpreter")

    def sort(self, rng, key, ascending, header):
        sh = rng.sheet
        first = rng.r1 + (1 if header else 0)
        rows = [{c: sh.get(r, c) for c in range(rng.c1, rng.c2 + 1)} for r in range(first, rng.r2 + 1)]
        fills = [{c: sh.fills.get((r, c)) for c in range(rng.c1, rng.c2 + 1)} for r in range(first, rng.r2 + 1)]
        formats = [{c: sh.formats.get((r, c)) for c in range(rng.c1, rng.c2 + 1)} for r in range(first, rng.r2 + 1)]

        def rank(v):  # Excel: numbers and dates < text < booleans; blanks always last
            if isinstance(v, bool):
                return (2, v)
            if isinstance(v, (dt.date, dt.datetime)):
                return (0, v.toordinal())
            if isinstance(v, (int, float)):
                return (0, v)
            return (1, str(v).casefold())

        order = list(range(len(rows)))
        filled = [i for i in order if rows[i][key] is not None]
        empty = [i for i in order if rows[i][key] is None]
        filled.sort(key=lambda i: rank(rows[i][key]), reverse=not ascending)
        for offset, i in enumerate(filled + empty):
            r = first + offset
            for c in range(rng.c1, rng.c2 + 1):   # Range.Sort moves values and cell formats
                sh.set(r, c, rows[i][c])
                for store, saved in ((sh.fills, fills), (sh.formats, formats)):
                    if saved[i][c]:
                        store[(r, c)] = saved[i][c]
                    else:
                        store.pop((r, c), None)

    def block(self, nodes):
        for n in nodes:
            s, kind = n["stmt"], n["kind"]
            if kind == "call":
                self.action(s)
            elif kind == "if_inline":
                if s.children[1].data == "note":
                    raise VclError("NOTE steps need the local model; the interpreter can't run them")
                if self.ev(s.children[0]):
                    self.action(s.children[1])
            elif kind == "if_block":
                self.block(n["children"] if self.ev(s.children[0]) else n["else"])
            elif kind == "exit_if":
                if self.ev(s.children[0]):
                    raise _Exit()
            elif kind == "skip_if":
                if self.ev(s.children[0]):
                    raise _Skip()
            elif kind == "exit_loop_if":
                if self.ev(s.children[0]):
                    raise _Break()
            elif kind in ("loop_rows", "loop_cols"):
                o = {str(x.children[0]): x.children[1] for x in s.children}
                a, b = int(self.num(self.ev(o["from"]))), int(self.num(self.ev(o["to"])))
                step = -1 if str(o.get("dir", "down")) in ("up", "left") else 1
                var = str(o["var"])
                for i in range(a, b + step, step) if (b - a) * step >= 0 else []:
                    self.vars[var] = i
                    if not self.loop_pass(n["children"]):
                        break
            elif kind == "select_case":   # the first CASE with a value = the selector (S7); else CASE_ELSE
                sel = self.ev(s.children[0])
                for case in n["children"]:
                    if case["kind"] == "case_else" or any(self.compare(sel, "=", self.ev(v))
                                                          for v in case["stmt"].children):
                        self.block(case["children"])
                        break
            elif kind == "for_each":
                o = {str(x.children[0]): x.children[1] for x in s.children}
                var = str(o["var"])
                for item in list(self.ev(o["in"])):
                    self.vars[var] = item
                    if not self.loop_pass(n["children"]):
                        break
            else:
                raise VclError(f"{kind} is not supported by the MVP interpreter")

    def loop_pass(self, body):
        """One pass of a loop body. False when EXIT_LOOP_IF leaves the loop."""
        try:
            self.block(body)
        except _Skip:
            pass
        except _Break:
            return False
        return True

    def run(self):
        for k in self.spec["keys"]:
            if k["key"] == "PRE" and k["tree"] is not None:
                for i, cond in enumerate(k["tree"].children):
                    try:
                        ok = self.ev(cond)
                    except (VclError, AttributeError, TypeError, KeyError) as ex:
                        ok = False
                    if not ok:
                        self.pre_failed = True
                        self.log.append((self.proc, 0, f"PRE failed: condition {i + 1}"))
                        return -1 if self.kind == "Function" else None
        value, given = None, False
        try:
            self.block(self.spec["ops"])
        except _Return as r:
            value, given = r.value, True
        except _Exit:
            pass
        except (VclError, AttributeError, TypeError, KeyError, ValueError, ZeroDivisionError) as ex:
            self.log.append((self.proc, 13, f"runtime error: {ex}"))
            return -2 if self.kind == "Function" else None
        try:
            return self.result(value, given)
        except (VclError, TypeError, ValueError) as ex:
            self.log.append((self.proc, 13, f"runtime error: {ex}"))
            return -2 if self.kind == "Function" else None

    def result(self, value, given):
        """A Function's return value in its RET type, as VBA converts it. Without a RETURN
        (the end of the procedure, or EXIT_IF) VBA gives the type's default: 0, "" or FALSE."""
        if self.kind != "Function":
            return None
        if not given:
            return {"Long": 0, "Double": 0.0, "String": "", "Boolean": False}.get(self.ret)
        if self.ret == "Long":
            return round(self.num(value))   # VBA rounds half to even, like Python's round
        if self.ret == "Double":
            return float(self.num(value))
        if self.ret == "String":
            return S.text_of(value)
        if self.ret == "Boolean" and isinstance(value, bool):
            return value
        if self.ret == "Boolean":
            raise VclError(f"type mismatch: {S.text_of(value)!r} is not TRUE or FALSE")
        return value


# ---------- TEST assertions ----------
def _crit_match(v, crit):
    """Excel COUNTIF criteria: optional operator, then a number or text (* and ? wildcards)."""
    crit = S.text_of(crit)
    op = "="
    for o in ("<=", ">=", "<>", "<", ">", "="):
        if crit.startswith(o):
            op, crit = o, crit[len(o):]
            break
    if crit == "":
        blank = v is None or v == ""
        return blank if op == "=" else (not blank if op == "<>" else False)
    if S.is_number(crit) and op != "=" or (S.is_number(crit) and isinstance(v, (int, float)) and not isinstance(v, bool)):
        if not (isinstance(v, (int, float)) and not isinstance(v, bool)):
            return op == "<>"
        x, y = v, S.to_number(crit)
        return {"=": x == y, "<>": x != y, "<": x < y, ">": x > y, "<=": x <= y, ">=": x >= y}[op]
    text = S.text_of(v).casefold()
    hit = fnmatch.fnmatchcase(text, crit.casefold())
    return hit if op == "=" else (not hit if op == "<>" else False)


class Asserter:
    def __init__(self, books_after, books_before, result, errors, fs_after=None, fs_before=None):
        self.after, self.before, self.result, self.errors = books_after, books_before, result, errors
        self.fs_after = fs_after if fs_after is not None else MockFS()
        self.fs_before = fs_before if fs_before is not None else MockFS()

    def cells(self, ref, pre):
        book = self.before if pre else self.after
        sheet_name = str(ref.children[0])
        sh = book[sheet_name]
        r = str(ref.children[1])
        if r.startswith("["):
            header, _, row = r[1:].partition("]")
            c = S.col([sh.get(1, x) for x in range(1, sh.max_col() + 1)], header)
            if c == 0:
                raise VclError(f"no header [{header}] on {sheet_name}")
            if row:
                return sh, [(int(row), c)]
            return sh, [(rr, c) for rr in range(1, sh.max_row() + 1)]
        from vcl_fixtures import col_number
        import re
        parts = r.replace("$", "").split(":")
        def split(p):
            m = re.match(r"([A-Z]+)(\d*)", p)
            return col_number(m.group(1)), (int(m.group(2)) if m.group(2) else None)
        c1, r1 = split(parts[0])
        c2, r2 = split(parts[-1])
        r1 = r1 or 1
        r2 = r2 or sh.max_row()
        return sh, [(rr, cc) for rr in range(r1, r2 + 1) for cc in range(c1, c2 + 1)]

    def values(self, ref, pre):
        sh, pts = self.cells(ref, pre)
        return [sh.get(r, c) for r, c in pts]

    def ev(self, t):
        if isinstance(t, Token):
            if t.type == "NUMBER":
                return S.to_number(str(t))
            if t.type == "STRING":
                return str(t)[1:-1].replace('""', '"')
            if str(t) == "RESULT":
                return self.result
            if str(t) == "ERROR_COUNT":
                return self.errors
            raise VclError(f"unknown name {t}")
        d = t.data
        if d == "a_neg":
            return -self.ev(t.children[0])
        if d == "a_sum":
            cur = self.ev(t.children[0])
            for i in range(1, len(t.children), 2):
                rhs = self.ev(t.children[i + 1])
                cur = cur + rhs if str(t.children[i]) == "+" else cur - rhs
            return cur
        if d == "a_func":
            name = str(t.children[0])
            pre = name.startswith("PRE_")
            fn = name[4:] if pre else name
            args = t.children[1:]
            if fn == "COUNTA":
                return sum(1 for v in self.values(args[0], pre) if v is not None)
            if fn == "COUNTIF":
                return sum(1 for v in self.values(args[0], pre) if _crit_match(v, self.ev(args[1])))
            if fn == "SUM":
                return sum(v for v in self.values(args[0], pre) if isinstance(v, (int, float)) and not isinstance(v, bool))
            if fn in ("COUNTIFS", "SUMIFS"):
                if fn == "SUMIFS":
                    target, pairs = self.values(args[0], pre), args[1:]
                else:
                    target, pairs = None, args
                cols = [(self.values(pairs[i], pre), self.ev(pairs[i + 1])) for i in range(0, len(pairs), 2)]
                n = len(cols[0][0])
                hits = [i for i in range(n) if all(_crit_match(vals[i], crit) for vals, crit in cols)]
                if fn == "COUNTIFS":
                    return len(hits)
                return sum(target[i] for i in hits if isinstance(target[i], (int, float)) and not isinstance(target[i], bool))
            if fn == "VALUE":
                v = self.values(args[0], pre)[0]
                return v
            if fn == "FILL_COUNT":
                sh, pts = self.cells(args[0], pre)
                color = str(args[1]).upper()
                return sum(1 for p in pts if sh.fills.get(p, "").upper() == color)
            if fn == "ROWS_USED":
                return (self.before if pre else self.after)[str(args[0])].max_row()
            if fn == "SHEET_EXISTS":   # after the run, case ignored as Excel does (D40)
                return Book(self.before if pre else self.after).exists(self.ev(args[0]))
            if fn in ("FILE_EXISTS", "FILE_COUNT"):   # a path below the sandbox, starting with a FIXTURE folder
                fs = self.fs_before if pre else self.fs_after
                path = SANDBOX + "\\" + self.ev(args[0])
                return fs.is_file(path) if fn == "FILE_EXISTS" else fs.files_below(path)
            raise VclError(f"assertion function {name} is not supported by the MVP interpreter")
        if d == "a_cond":
            if len(t.children) == 1:
                return bool(self.ev(t.children[0]))
            a, op, b = self.ev(t.children[0]), str(t.children[1]), self.ev(t.children[2])
            if op in ("=", "<>"):
                eq = (a == b) if not (isinstance(a, str) or isinstance(b, str)) else S.text_of(a) == S.text_of(b)
                return eq if op == "=" else not eq
            if a is None or b is None:
                return False
            return {"<": a < b, ">": a > b, "<=": a <= b, ">=": a >= b}[op]
        raise VclError(f"can't evaluate assertion part {d}")


# ---------- MODIFIES: the automatic UNCHANGED check (VCL 0.2) ----------
def modifies_of(spec):
    """{arg: {"whole": True} or {"cols": [header, ...], "fills": True}} from the MODIFIES lines."""
    out = {}
    for k in spec["keys"]:
        if k["key"] != "MODIFIES" or k["tree"] is None:
            continue
        kids = k["tree"].children
        entry = out.setdefault(str(kids[0])[1:], {})
        if len(kids) == 1:
            entry["whole"] = True
        elif isinstance(kids[1], Tree):
            entry.setdefault("cols", []).extend(str(h)[1:-1].replace('""', '"') for h in kids[1].children)
        elif kids[1].type == "FILES_WORD":
            entry["files"] = True   # VCL 0.3: a root folder, checked on the mock file system (router Phase 2)
        elif kids[1].type == "SHEETS_WORD":
            entry["sheets"] = True   # VCL 0.4: new sheets may be added to a Workbook argument (D40)
        else:
            entry["fills"] = True
    return out


WRITE_OPS = {"SET_VALUE", "WRITE_ROW", "FORMAT", "CLEAR", "COPY_ROW", "INSERT_ROW", "DELETE_ROW", "SORT",
             "WRITE_ARRAY", "SET_FORMULA", "DEDUPE", "FILTER", "ADD_SHEET", "COPY_SHEET"}


def writes_anything(spec):
    """True when any op of the spec writes to a sheet. A VCL 0.2 spec that writes nothing gets
    the UNCHANGED check without MODIFIES lines: it may change nothing (D20)."""
    def calls(nodes):
        for n in nodes:
            stmt = n["stmt"]
            if n["kind"] == "call":
                yield str(stmt.children[0])
            elif n["kind"] == "if_inline" and isinstance(stmt.children[1], Tree) and stmt.children[1].data == "call":
                yield str(stmt.children[1].children[0])
            yield from calls(n["children"])
            yield from calls(n.get("else", []))
    return any(name in WRITE_OPS for name in calls(spec["ops"]))


def needs_unchanged_check(spec, modifies):
    """MODIFIES lines, or (VCL 0.2 and later, D20) a spec that writes nothing: it may change nothing."""
    return bool(modifies) or (spec.get("version", "VCL 0.1") != "VCL 0.1" and not writes_anything(spec))


def _letters(c):
    out = ""
    while c:
        c, rest = divmod(c - 1, 26)
        out = chr(65 + rest) + out
    return out


def _kind(v):
    if v is None:
        return "empty"
    if isinstance(v, bool):
        return "bool"
    if isinstance(v, (int, float)):
        return "num"
    if isinstance(v, (dt.date, dt.datetime)):
        return "date"
    if isinstance(v, S.XlError):
        return "error"
    return "text"


def _show(v):
    if v is None:
        return "(empty)"
    return '"' + v + '"' if isinstance(v, str) else S.text_of(v)


def unchanged_problems(modifies, sheet_args, before, after, stopped_at_pre, limit=5, book_args=()):
    """Cells MODIFIES doesn't allow to change that did change, as "UNCHANGED: ..." lines.
    sheet_args maps each sheet name to the arguments the TEST passed it as; book_args are the
    Workbook arguments (VCL 0.4), whose MODIFIES $book sheets lets new sheets appear (D40). When
    every run stopped at PRE, nothing may change at all."""
    found = []
    may_add = not stopped_at_pre and any(modifies.get(arg, {}).get("sheets") for arg in book_args)
    for name in sorted(set(before) | set(after)):
        if name not in before and may_add:
            continue   # a new sheet: everything on it is new
        if name not in before or name not in after:
            found.append(f"UNCHANGED: sheet {name} was {'added' if name in after else 'removed'}")
            continue
        allowed = {}
        if not stopped_at_pre:
            for arg in sheet_args.get(name, ()):
                for key, value in modifies.get(arg, {}).items():
                    allowed[key] = value if key != "cols" else allowed.get("cols", []) + value
        if allowed.get("whole"):
            continue
        b, a = before[name], after[name]
        wanted = {S.header_key(h) for h in allowed.get("cols", [])}
        cols = {c for c in range(1, b.max_col() + 1) if S.header_key(b.get(1, c)) in wanted and b.get(1, c) is not None}
        for (r, c) in sorted(set(b.cells) | set(a.cells)):
            bv, av = b.cells.get((r, c)), a.cells.get((r, c))
            if (r >= 2 and c in cols) or (_kind(bv), S.text_of(bv)) == (_kind(av), S.text_of(av)):
                continue
            found.append(f"UNCHANGED: {name}!{_letters(c)}{r} was {_show(bv)}, now {_show(av)}")
        stores = [("number format", b.formats, a.formats)]
        if not allowed.get("fills"):
            stores.insert(0, ("fill", b.fills, a.fills))
        for what, bs, as_ in stores:
            for (r, c) in sorted(set(bs) | set(as_)):
                if bs.get((r, c)) != as_.get((r, c)):
                    found.append(f"UNCHANGED: {name}!{_letters(c)}{r} {what} was {bs.get((r, c)) or 'none'}, "
                                 f"now {as_.get((r, c)) or 'none'}")
    if len(found) > limit:
        found = found[:limit] + [f"UNCHANGED: ... and {len(found) - limit} more changes"]
    return found


@dataclass
class TestResult:
    line: int
    passed: bool
    result: object
    errors: int
    failed: list = field(default_factory=list)
    log: list = field(default_factory=list)
    state: dict = field(default_factory=dict)
    files: list = field(default_factory=list)   # VCL 0.3: the mock file system after the last run
    messages: list = field(default_factory=list)   # LOG and MSG texts of the last run (D36)


def files_of(fs):
    """Every folder and file below the sandbox root, as (path below it, kind, size, locked), sorted."""
    prefix = SANDBOX.casefold() + "\\"
    return sorted((e["path"][len(prefix):], e["kind"], e["size"], e["locked"])
                  for k, e in fs.entries.items() if k.startswith(prefix))


def file_unchanged_problems(roots, before, after, stopped_at_pre, limit=5):
    """E10: nothing on disk is removed or changed, and new folders and files appear only below a
    root the TEST passed for a MODIFIES $root files argument (nowhere when every run stopped at PRE)."""
    allowed = [] if stopped_at_pre else [r.rstrip("\\").casefold() + "\\" for r in roots]
    prefix = len(SANDBOX) + 1
    found = []
    for key in sorted(set(before.entries) | set(after.entries)):
        b, a = before.entries.get(key), after.entries.get(key)
        if b == a:
            continue
        what = f"{(a or b)['kind']} {(a or b)['path'][prefix:]}"
        if b is None:
            if not any(key.startswith(root) for root in allowed):
                found.append(f"UNCHANGED: {what} was added")
        elif a is None:
            found.append(f"UNCHANGED: {what} was removed")
        else:
            def show(e):
                return f"{e['kind']}, {e['size']} bytes{', locked' if e['locked'] else ''}"
            found.append(f"UNCHANGED: {what} was {show(b)}, now {show(a)}")
    if len(found) > limit:
        found = found[:limit] + [f"UNCHANGED: ... and {len(found) - limit} more file changes"]
    return found


def sandbox_length_problems(paths):
    """D30: in Excel the sandbox root is longer than C:\\KRONOS_SANDBOX (up to EXCEL_SANDBOX_MAX
    characters), so a path below it is longer there too. A path whose Too long status could be
    different in Excel fails the TEST: it must be far over its limit or well under it."""
    found = []
    for path, limit in dict.fromkeys(paths):
        if not path.casefold().startswith(SANDBOX.casefold() + "\\"):
            continue
        most = len(path) - len(SANDBOX) + EXCEL_SANDBOX_MAX
        if len(path) <= limit < most:
            found.append(f"SANDBOX: {path[len(SANDBOX) + 1:]} has {len(path)} characters here but up to {most} in "
                         f"Excel's sandbox, where it would be Too long (over {limit}): make it far longer or shorter")
    return found


def run_tests(text, today=None):
    """Run every TEST of a TASK spec. today pins TODAY (default: the real date)."""
    spec, problems = parse(text)
    if problems or not spec or spec["kind"] != "TASK":
        raise VclError(f"spec doesn't parse: {problems[:3]}")
    today = today or dt.date.today()
    trees = [k["tree"] for k in spec["keys"] if k["key"] == "FIXTURE"]
    books = build_workbooks(trees, today)
    systems = build_folders(trees)
    results = []
    for k in spec["keys"]:
        if k["key"] != "TEST":
            continue
        t = k["tree"]
        kids = t.children
        fixture = str(kids[0])
        rest = kids[1:]
        target, i = rest[0], 1
        call_vals = []
        while i < len(rest) and not (isinstance(rest[i], Token) and rest[i].type == "INT") and not (isinstance(rest[i], Tree) and rest[i].data == "assertion"):
            call_vals.append(rest[i])
            i += 1
        runs = 1
        if i < len(rest) and isinstance(rest[i], Token) and rest[i].type == "INT":
            runs = int(rest[i])
            i += 1
        assertion = rest[i]
        if fixture not in systems:   # every FIXTURE line, sheet or folder, gives its file a mock file system
            raise VclError(f"TEST uses fixture {fixture}, but no FIXTURE line builds it")
        book = {name: sh.snapshot() for name, sh in books.get(fixture, {}).items()}
        before = {name: sh.snapshot() for name, sh in book.items()}
        fs = systems[fixture].snapshot()
        fs_before = fs.snapshot()
        params = [str(p.children[0]) for kk in spec["keys"] if kk["key"] == "ARGS" for p in kk["tree"].children]
        wb = Book(book)   # THIS_BOOK, ACTIVE_BOOK and every Workbook argument: this TEST's fixture file
        bound, book_args = {}, []
        for pname, v in zip(params, call_vals):
            if isinstance(v, Token) and v.type == "FILENAME":   # VCL 0.4 (D40)
                if str(v) != fixture:
                    raise VclError(f"TEST passes {v} for ${pname}, but its fixture is {fixture}")
                bound[pname] = wb
                book_args.append(pname)
            elif isinstance(v, Token) and v.type == "NAME" and str(v) in book:
                bound[pname] = book[str(v)]
            elif isinstance(v, Token) and v.type == "NAME":   # a FIXTURE folder (F6)
                if not fs.is_folder(SANDBOX + "\\" + str(v)):
                    raise VclError(f"TEST passes {v}, which is neither a sheet nor a folder of {fixture}")
                bound[pname] = SANDBOX + "\\" + str(v)
            elif isinstance(v, Token) and v.type == "STRING":
                bound[pname] = str(v)[1:-1]
            else:
                bound[pname] = S.to_number(str(v))
        log = []
        result = None
        stopped_at_pre = True
        file_paths = []
        for _ in range(runs):
            log.clear()
            machine = Machine(spec, bound, log, today, fs, wb)
            result = machine.run()
            stopped_at_pre = stopped_at_pre and machine.pre_failed
            file_paths += machine.file_paths
        asserter = Asserter(book, before, result, len(log), fs, fs_before)
        failed = []
        for cond in assertion.children:
            try:
                ok = asserter.ev(cond)
            except (VclError, KeyError, TypeError) as ex:
                ok = False
            if not ok:
                failed.append(cond)
        modifies = modifies_of(spec)
        if needs_unchanged_check(spec, modifies):
            sheet_args = {}
            for pname, v in zip(params, call_vals):
                if isinstance(v, Token) and v.type == "NAME":
                    sheet_args.setdefault(str(v), []).append(pname)
            failed += unchanged_problems(modifies, sheet_args, before, book, stopped_at_pre, book_args=book_args)
            roots = [bound[a] for a, kinds in modifies.items() if kinds.get("files") and isinstance(bound.get(a), str)]
            failed += file_unchanged_problems(roots, fs_before, fs, stopped_at_pre)
        failed += sandbox_length_problems(file_paths)
        state = {name: (sorted((rc, S.text_of(v)) for rc, v in sh.cells.items()), sorted(sh.fills.items()),
                        sorted(sh.formats.items()))
                 for name, sh in book.items()}
        results.append(TestResult(k["line"], not failed, result, len(log), failed, list(log), state, files_of(fs),
                                   list(machine.messages)))
    return results


if __name__ == "__main__":
    total_fail = 0
    for p in sys.argv[1:]:
        try:
            res = run_tests(Path(p).read_text(encoding="utf-8"))
        except VclError as ex:
            print(f"ERROR {p}: {ex}")
            total_fail += 1
            continue
        bad = [r for r in res if not r.passed]
        total_fail += len(bad)
        print(f"{'PASS' if not bad else 'FAIL'} {p}  ({len(res) - len(bad)}/{len(res)} tests; RESULTs {[r.result for r in res]})")
        for r in bad:
            print(f"     line {r.line}: RESULT={r.result} ERROR_COUNT={r.errors} log={r.log[:2]}")
    sys.exit(1 if total_fail else 0)
