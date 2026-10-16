"""Reference interpreter for VCL 0.1 rev 3 (milestone M3).

run_tests(text) -> list of TestResult. Executes the spec's OP tree on fixture workbooks built by
vcl_fixtures, with the meanings defined in vcl/semantics.py, and evaluates every TEST.
Error model: a failed PRE logs once, changes nothing and returns -1; a runtime error logs once
and returns -2 (Functions).
"""
import datetime as dt
import fnmatch
import sys
from dataclasses import dataclass, field
from pathlib import Path

from lark import Token, Tree

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "vcl"))
import semantics as S  # noqa: E402
from vcl_fixtures import Sheet, build_workbooks  # noqa: E402
from vcl_parse import parse  # noqa: E402


class VclError(Exception):
    pass


class _Return(Exception):
    def __init__(self, value):
        self.value = value


class _Exit(Exception):
    pass


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


def opts_of(args):
    return {str(a.children[0]): a.children[1] for a in args if isinstance(a, Tree) and a.data == "option"}


def pos_of(args):
    return [a for a in args if not (isinstance(a, Tree) and a.data == "option")]


def call_args(call):
    return call.children[1].children if len(call.children) > 1 else []


class Machine:
    def __init__(self, spec, bound_args, log, today=None):
        self.spec = spec
        self.log = log
        self.args = bound_args
        self.today = today or dt.date.today()
        self.vars = {}
        self.types = {}
        self.proc = ""
        self.kind = "Sub"
        for k in spec["keys"]:
            t = k["tree"]
            if t is None:
                continue
            if k["key"] == "PROC":
                self.proc = str(t.children[0])
            elif k["key"] == "KIND":
                self.kind = str(t.children[0])
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
        if name == "EXISTS":
            return e(pos[0]).exists(e(pos[1]))
        if name == "COUNT":
            d = e(pos[0])
            return len(d.items) if isinstance(d, FDict) else len(d)
        if name == "KEYS":
            return e(pos[0]).keys()
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
                    else:
                        rng.sheet.fills[(r, c)] = str(f).upper()
                if "numfmt" in opts:
                    rng.sheet.formats[(r, c)] = e(opts["numfmt"])
        elif name == "COPY_ROW":
            src, r, dst, r2 = e(pos[0]), self.index(e(pos[1])), e(pos[2]), self.index(e(pos[3]))
            n = self.index(e(opts["cols"]))
            for c in range(1, n + 1):
                dst.set(r2, c, src.get(r, c))
        elif name == "DELETE_ROW":
            sh, row = e(pos[0]), int(e(pos[1]))
            sh.cells = {(r - (r > row), c): v for (r, c), v in sh.cells.items() if r != row}
            sh.fills = {(r - (r > row), c): v for (r, c), v in sh.fills.items() if r != row}
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
        elif name == "RETURN":
            raise _Return(e(pos[0]) if pos else None)
        else:
            raise VclError(f"{name} is not supported by the MVP interpreter")

    def sort(self, rng, key, ascending, header):
        sh = rng.sheet
        first = rng.r1 + (1 if header else 0)
        rows = [{c: sh.get(r, c) for c in range(rng.c1, rng.c2 + 1)} for r in range(first, rng.r2 + 1)]
        fills = [{c: sh.fills.get((r, c)) for c in range(rng.c1, rng.c2 + 1)} for r in range(first, rng.r2 + 1)]

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
            for c in range(rng.c1, rng.c2 + 1):
                sh.set(r, c, rows[i][c])
                if fills[i][c]:
                    sh.fills[(r, c)] = fills[i][c]
                else:
                    sh.fills.pop((r, c), None)

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
            elif kind in ("loop_rows", "loop_cols"):
                o = {str(x.children[0]): x.children[1] for x in s.children}
                a, b = int(self.num(self.ev(o["from"]))), int(self.num(self.ev(o["to"])))
                step = -1 if str(o.get("dir", "down")) in ("up", "left") else 1
                var = str(o["var"])
                for i in range(a, b + step, step) if (b - a) * step >= 0 else []:
                    self.vars[var] = i
                    self.block(n["children"])
            elif kind == "for_each":
                o = {str(x.children[0]): x.children[1] for x in s.children}
                var = str(o["var"])
                for item in list(self.ev(o["in"])):
                    self.vars[var] = item
                    self.block(n["children"])
            else:
                raise VclError(f"{kind} is not supported by the MVP interpreter")

    def run(self):
        for k in self.spec["keys"]:
            if k["key"] == "PRE" and k["tree"] is not None:
                for i, cond in enumerate(k["tree"].children):
                    try:
                        ok = self.ev(cond)
                    except (VclError, AttributeError, TypeError, KeyError) as ex:
                        ok = False
                    if not ok:
                        self.log.append((self.proc, 0, f"PRE failed: condition {i + 1}"))
                        return -1 if self.kind == "Function" else None
        try:
            self.block(self.spec["ops"])
        except _Return as r:
            return r.value
        except _Exit:
            pass
        except (VclError, AttributeError, TypeError, KeyError, ValueError, ZeroDivisionError) as ex:
            self.log.append((self.proc, 13, f"runtime error: {ex}"))
            return -2 if self.kind == "Function" else None
        return 0 if self.kind == "Function" else None


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
    def __init__(self, books_after, books_before, result, errors):
        self.after, self.before, self.result, self.errors = books_after, books_before, result, errors

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


@dataclass
class TestResult:
    line: int
    passed: bool
    result: object
    errors: int
    failed: list = field(default_factory=list)
    log: list = field(default_factory=list)
    state: dict = field(default_factory=dict)


def run_tests(text, today=None):
    """Run every TEST of a TASK spec. today pins TODAY (default: the real date)."""
    spec, problems = parse(text)
    if problems or not spec or spec["kind"] != "TASK":
        raise VclError(f"spec doesn't parse: {problems[:3]}")
    books = build_workbooks([k["tree"] for k in spec["keys"] if k["key"] == "FIXTURE"])
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
        book = {name: sh.snapshot() for name, sh in books[fixture].items()}
        before = {name: sh.snapshot() for name, sh in book.items()}
        params = [str(p.children[0]) for kk in spec["keys"] if kk["key"] == "ARGS" for p in kk["tree"].children]
        bound = {}
        for pname, v in zip(params, call_vals):
            if isinstance(v, Token) and v.type == "NAME":
                bound[pname] = book[str(v)]
            elif isinstance(v, Token) and v.type == "STRING":
                bound[pname] = str(v)[1:-1]
            else:
                bound[pname] = S.to_number(str(v))
        log = []
        result = None
        for _ in range(runs):
            log.clear()
            result = Machine(spec, bound, log, today).run()
        asserter = Asserter(book, before, result, len(log))
        failed = []
        for cond in assertion.children:
            try:
                ok = asserter.ev(cond)
            except (VclError, KeyError, TypeError) as ex:
                ok = False
            if not ok:
                failed.append(cond)
        state = {name: (sorted((rc, S.text_of(v)) for rc, v in sh.cells.items()), sorted(sh.fills.items()))
                 for name, sh in book.items()}
        results.append(TestResult(k["line"], not failed, result, len(log), failed, list(log), state))
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
