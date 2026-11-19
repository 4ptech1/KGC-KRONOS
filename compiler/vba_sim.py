"""A simulator for the VBA that VBA Forge writes (D90 Phase 1, D98).

It runs, without Excel, the subset of VBA 7 that the compiler (`vcl_codegen.py`), ForgeLib
(`vba/ForgeLib.bas`) and the self-test (`vcl_selftest.py`) emit, against a small model of an Excel
2016 workbook. Its job is proof: the compiler's own VBA must end exactly where the reference
interpreter says, and from March 2027 our model's VBA must too (`docs/ai-vba-plan.md`). So it follows
VBA's own rules, not Python's: Variant subtypes, banker's rounding in CLng, Excel storing every number
as a Double, On Error GoTo and Resume, ByRef arguments, Erl and line numbers.

    import vba_sim
    sim = vba_sim.Simulator([module_text, ...], today=date(2026, 11, 19))
    sim.call("ForgeSelfTest")

What it does not run is refused with VbaUnsupported (the statement or name), never guessed. Excel
features outside the model raise VBA error 438, so a macro that needs one fails loudly.
"""
import datetime as dt
import math
import re

from vba_lint import VBA_RESERVED

__all__ = ["Simulator", "VbaError", "VbaUnsupported", "EMPTY", "NULL", "NOTHING", "VDate", "VError", "VArray",
           "Workbook", "Worksheet", "Range"]

ROWS, COLS = 1048576, 16384
EPOCH = dt.datetime(1899, 12, 30)


# ======================================================================= values

class _Special:
    def __init__(self, name):
        self.name = name

    def __repr__(self):
        return self.name


EMPTY, NULL, NOTHING, MISSING = _Special("Empty"), _Special("Null"), _Special("Nothing"), _Special("Missing")


class VDate:
    """A Date: VBA's serial number, days since 30 Dec 1899, the time as the fraction."""
    __slots__ = ("serial",)

    def __init__(self, serial):
        self.serial = float(serial)

    @classmethod
    def of(cls, d):
        if isinstance(d, dt.datetime):
            return cls((d - EPOCH).total_seconds() / 86400.0)
        return cls((dt.datetime(d.year, d.month, d.day) - EPOCH).days)

    def to_datetime(self):
        return EPOCH + dt.timedelta(days=self.serial)

    def __eq__(self, other):
        return isinstance(other, VDate) and other.serial == self.serial

    def __hash__(self):
        return hash(("date", self.serial))

    def __repr__(self):
        return f"VDate({self.to_datetime():%Y-%m-%d %H:%M:%S})"


class VError:
    """An Excel error value (CVErr): #N/A is 2042."""
    __slots__ = ("code",)

    def __init__(self, code):
        self.code = int(code)

    def __eq__(self, other):
        return isinstance(other, VError) and other.code == self.code

    def __hash__(self):
        return hash(("err", self.code))

    def __repr__(self):
        return f"VError({self.code})"


class VArray:
    """A VBA array: any number of dimensions, each with its own lower bound."""

    def __init__(self, bounds, fill=EMPTY):
        self.bounds = [(int(lo), int(hi)) for lo, hi in bounds]
        size = 1
        for lo, hi in self.bounds:
            size *= max(0, hi - lo + 1)
        self.data = [fill() if callable(fill) else fill for _ in range(size)]

    @classmethod
    def from_list(cls, items, lower=0):
        a = cls([(lower, lower + len(items) - 1)])
        a.data = list(items)
        return a

    def _offset(self, idx):
        if len(idx) != len(self.bounds):
            raise VbaError(9, "Subscript out of range")
        off = 0
        for i, (lo, hi) in zip(idx, self.bounds):
            i = to_long(i)
            if i < lo or i > hi:
                raise VbaError(9, "Subscript out of range")
            off = off * (hi - lo + 1) + (i - lo)
        return off

    def get(self, idx):
        return self.data[self._offset(idx)]

    def set(self, idx, value):
        self.data[self._offset(idx)] = value

    def copy(self):
        a = VArray(self.bounds)
        a.data = [v.copy() if isinstance(v, VArray) else v for v in self.data]
        return a

    def items(self):
        return list(self.data)


class VbaError(Exception):
    """A run-time error, as Err.Raise or VBA itself raises it."""

    def __init__(self, number, description="", source=""):
        super().__init__(f"error {number}: {description}")
        self.number, self.description, self.source = int(number), description, source


class VbaUnsupported(Exception):
    """VBA the simulator does not run. Never guessed: the run stops."""


class _Exit(Exception):
    pass


def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def to_double(v):
    if isinstance(v, bool):
        return -1.0 if v else 0.0
    if is_num(v):
        return float(v)
    if v is EMPTY:
        return 0.0
    if isinstance(v, VDate):
        return v.serial
    if isinstance(v, str):
        n = _parse_number_text(v)
        if n is None:
            raise VbaError(13, "Type mismatch")
        return float(n)
    if v is NULL:
        raise VbaError(94, "Invalid use of Null")
    raise VbaError(13, "Type mismatch")


def _round_even(x):
    return int(round(x))   # Python rounds half to even, as VBA's CLng and CInt do


def to_long(v):
    if isinstance(v, bool):
        return -1 if v else 0
    if isinstance(v, int):
        return v
    d = to_double(v)
    if math.isnan(d) or abs(d) > 2147483647.5:
        raise VbaError(6, "Overflow")
    return _round_even(d)


def to_bool(v):
    if isinstance(v, bool):
        return v
    if v is EMPTY:
        return False
    if isinstance(v, str):
        if v.strip().lower() == "true":
            return True
        if v.strip().lower() == "false":
            return False
    return to_double(v) != 0


_NUM_RE = re.compile(r"^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$")


def _parse_number_text(s):
    """CDbl and IsNumeric on text, read as on an English PC: spaces around, one sign, thousands commas."""
    t = s.strip(" \t")
    if t == "":
        return None
    neg = False
    if t.startswith("(") and t.endswith(")"):
        return None
    t2 = t.replace(",", "")
    if _NUM_RE.match(t2):
        try:
            return float(t2) if (("." in t2) or ("e" in t2.lower())) else int(t2)
        except ValueError:
            return None
    return None if not neg else None


def num_text(d):
    """CStr of a Double: up to 15 significant digits, E notation like VBA's."""
    if isinstance(d, bool):
        return "True" if d else "False"
    if isinstance(d, int):
        return str(d)
    if d == 0:
        return "0"
    s = "%.15g" % d
    if "e" in s:
        mant, exp = s.split("e")
        e = int(exp)
        return f"{mant}E{'+' if e >= 0 else '-'}{abs(e):02d}"
    return s


def date_text(v):
    d = v.to_datetime()
    if v.serial == int(v.serial):
        return f"{d.month}/{d.day}/{d.year}"
    t = d.strftime("%I:%M:%S %p").lstrip("0")
    return t if int(v.serial) == 0 else f"{d.month}/{d.day}/{d.year} {t}"


def to_str(v):
    if isinstance(v, str):
        return v
    if v is EMPTY:
        return ""
    if v is NULL:
        raise VbaError(94, "Invalid use of Null")
    if isinstance(v, bool):
        return "True" if v else "False"
    if is_num(v):
        return num_text(v)
    if isinstance(v, VDate):
        return date_text(v)
    if isinstance(v, VError):
        return f"Error {v.code}"
    raise VbaError(13, "Type mismatch")


def to_date(v):
    if isinstance(v, VDate):
        return v
    if is_num(v) or isinstance(v, bool):
        return VDate(to_double(v))
    if isinstance(v, str):
        m = re.match(r"^\s*(\d{4})-(\d{1,2})-(\d{1,2})\s*$", v)
        if m:
            return VDate.of(dt.date(int(m[1]), int(m[2]), int(m[3])))
        m = re.match(r"^\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*$", v)   # an English PC reads m/d/yyyy
        if m:
            return VDate.of(dt.date(int(m[3]), int(m[1]), int(m[2])))
    raise VbaError(13, "Type mismatch")


def vartype(v):
    if v is EMPTY:
        return 0
    if v is NULL:
        return 1
    if isinstance(v, bool):
        return 11
    if isinstance(v, int):
        return 3
    if isinstance(v, float):
        return 5
    if isinstance(v, VDate):
        return 7
    if isinstance(v, str):
        return 8
    if isinstance(v, VError):
        return 10
    if isinstance(v, VArray):
        return 8192 + 12
    return 9


def typename(v):
    return {0: "Empty", 1: "Null", 11: "Boolean", 3: "Long", 5: "Double", 7: "Date", 8: "String", 10: "Error",
            8204: "Variant()"}.get(vartype(v)) or getattr(v, "vba_type", "Object")


DEFAULTS = {"long": 0, "integer": 0, "byte": 0, "double": 0.0, "single": 0.0, "currency": 0.0, "string": "",
            "boolean": False, "date": VDate(0), "variant": EMPTY}
OBJECT_TYPES = {"object", "range", "worksheet", "workbook", "collection", "dictionary"}


def default_of(typ):
    t = (typ or "variant").lower()
    if t in DEFAULTS:
        return DEFAULTS[t]
    return NOTHING


def coerce(typ, v):
    """The value a variable of this type holds after an assignment (Let)."""
    t = (typ or "variant").lower()
    if t == "variant":
        return v.copy() if isinstance(v, VArray) else v
    if isinstance(v, VError) and t != "variant":
        raise VbaError(13, "Type mismatch")
    if t in ("long", "integer", "byte"):
        n = to_long(v)
        if t == "integer" and not -32768 <= n <= 32767:
            raise VbaError(6, "Overflow")
        return n
    if t in ("double", "single", "currency"):
        return to_double(v)
    if t == "string":
        return to_str(v)
    if t == "boolean":
        return to_bool(v)
    if t == "date":
        return to_date(v)
    return v


# ======================================================================= lexer

KEYWORDS = {"and", "or", "not", "xor", "eqv", "imp", "mod", "is", "like", "new", "to", "then", "else", "elseif",
            "end", "if", "for", "each", "in", "next", "do", "loop", "while", "wend", "until", "select", "case",
            "dim", "redim", "preserve", "set", "call", "exit", "sub", "function", "private", "public", "const",
            "as", "byval", "byref", "optional", "on", "error", "goto", "resume", "step", "with", "static",
            "option", "explicit", "true", "false", "nothing", "empty", "null"}

_TOKEN = re.compile(r"""
    (?P<ws>[ \t]+)
  | (?P<str>"(?:[^"]|"")*")
  | (?P<hex>&H[0-9A-Fa-f]+&?)
  | (?P<num>(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?[#&!@%]?)
  | (?P<id>[A-Za-z_][A-Za-z0-9_]*[$]?)
  | (?P<op>:=|<=|>=|<>|[-+*/\\^&=<>(),.:;#])
""", re.X)


class Tok:
    __slots__ = ("kind", "val", "low")

    def __init__(self, kind, val):
        self.kind, self.val = kind, val
        self.low = val.lower() if kind == "id" else val

    def __repr__(self):
        return f"{self.kind}:{self.val}"


def _strip_comment(line):
    out, i, q = [], 0, False
    while i < len(line):
        ch = line[i]
        if ch == '"':
            q = not q
        elif ch == "'" and not q:
            break
        out.append(ch)
        i += 1
    s = "".join(out)
    if re.match(r"^\s*Rem(\s|$)", s, re.I):
        return ""
    return s


def tokenize(text):
    toks, i = [], 0
    while i < len(text):
        m = _TOKEN.match(text, i)
        if not m:
            raise VbaUnsupported(f"can't read {text[i:i + 20]!r}")
        i = m.end()
        kind = m.lastgroup
        v = m.group()
        if kind == "ws":
            continue
        if kind == "str":
            toks.append(Tok("str", v[1:-1].replace('""', '"')))
        elif kind == "hex":
            toks.append(Tok("num", int(v[2:].rstrip("&"), 16) if v.endswith("&") else _hex16(v[2:])))
        elif kind == "num":
            suffix = v[-1] if v[-1] in "#&!@%" else ""
            body = v[:-1] if suffix else v
            if suffix in ("#", "!", "@") or "." in body or "e" in body.lower():
                toks.append(Tok("num", float(body)))
            else:
                toks.append(Tok("num", int(body)))
        elif kind == "id":
            toks.append(Tok("id", v[:-1] if v.endswith("$") else v))
        else:
            toks.append(Tok("op", v))
    return toks


def _hex16(digits):
    n = int(digits, 16)
    return n - 65536 if n <= 0xFFFF and n >= 0x8000 else n   # &H8000 to &HFFFF are negative Integers


def logical_lines(text):
    """(line number in the file, text) for every statement line: comments gone, continuations joined."""
    out, buf, start = [], "", 0
    for n, raw in enumerate(text.splitlines(), 1):
        line = _strip_comment(raw).rstrip()
        if not buf:
            start = n
        if line.endswith(" _") or line == "_":
            buf += line[:-1] + " "
            continue
        buf += line
        if buf.strip():
            out.append((start, buf))
        buf = ""
    return out


# ======================================================================= parser (expressions)

class Node:
    __slots__ = ("kind", "a", "b", "c")

    def __init__(self, kind, a=None, b=None, c=None):
        self.kind, self.a, self.b, self.c = kind, a, b, c

    def __repr__(self):
        return f"{self.kind}({self.a!r}, {self.b!r}, {self.c!r})"


BINARY = [("imp",), ("eqv",), ("xor",), ("or",), ("and",)]
COMPARE = {"=", "<>", "<", ">", "<=", ">=", "is", "like"}


class Parser:
    def __init__(self, toks):
        self.t, self.i = toks, 0

    def peek(self, k=0):
        j = self.i + k
        return self.t[j] if j < len(self.t) else None

    def at(self, *vals):
        p = self.peek()
        return p is not None and (p.low in vals if p.kind == "id" else p.val in vals) and p.kind != "str"

    def take(self):
        p = self.peek()
        self.i += 1
        return p

    def expect(self, val):
        p = self.take()
        if p is None or (p.low if p.kind == "id" else p.val) != val:
            raise VbaUnsupported(f"expected {val!r}, got {p!r}")
        return p

    def done(self):
        return self.i >= len(self.t)

    # expression grammar, loosest first
    def expr(self):
        return self._logic(0)

    def _logic(self, level):
        if level == len(BINARY):
            return self._not()
        left = self._logic(level + 1)
        while self.at(*BINARY[level]):
            op = self.take().low
            left = Node("bin", op, left, self._logic(level + 1))
        return left

    def _not(self):
        if self.at("not"):
            self.take()
            return Node("not", self._not())
        return self._cmp()

    def _cmp(self):
        left = self._concat()
        while self.peek() is not None and self.peek().kind != "str" and (
                (self.peek().kind == "op" and self.peek().val in COMPARE) or self.at("is", "like")):
            op = self.take()
            op = op.low if op.kind == "id" else op.val
            left = Node("bin", op, left, self._concat())
        return left

    def _concat(self):
        left = self._add()
        while self.at("&"):
            self.take()
            left = Node("bin", "&", left, self._add())
        return left

    def _add(self):
        left = self._mod()
        while self.at("+", "-"):
            op = self.take().val
            left = Node("bin", op, left, self._mod())
        return left

    def _mod(self):
        left = self._idiv()
        while self.at("mod"):
            self.take()
            left = Node("bin", "mod", left, self._idiv())
        return left

    def _idiv(self):
        left = self._mul()
        while self.at("\\"):
            self.take()
            left = Node("bin", "\\", left, self._mul())
        return left

    def _mul(self):
        left = self._neg()
        while self.at("*", "/"):
            op = self.take().val
            left = Node("bin", op, left, self._neg())
        return left

    def _neg(self):
        if self.at("-"):
            self.take()
            return Node("neg", self._neg())
        if self.at("+"):
            self.take()
            return self._neg()
        return self._pow()

    def _pow(self):
        left = self._postfix()
        while self.at("^"):
            self.take()
            left = Node("bin", "^", left, self._neg())
        return left

    def _postfix(self):
        node = self._primary()
        while True:
            if self.at("("):
                node = Node("call", node, self._args_paren())
            elif self.at(".") and self.peek(1) is not None and self.peek(1).kind == "id":
                self.take()
                node = Node("member", node, self.take().val)
            else:
                return node

    def _args_paren(self):
        self.expect("(")
        args = []
        if self.at(")"):
            self.take()
            return args
        while True:
            args.append(self._arg())
            if self.at(","):
                self.take()
                continue
            self.expect(")")
            return args

    def _arg(self):
        p = self.peek()
        if p is not None and p.kind == "op" and p.val in (",", ")"):
            return (None, Node("missing"))
        if p is not None and p.kind == "id" and self.peek(1) is not None and self.peek(1).val == ":=":
            self.take()
            self.take()
            return (p.val, self.expr())
        if self.at("byval"):   # ByVal in a call is ignored
            self.take()
        return (None, self.expr())

    def args_bare(self):
        """The arguments of a call statement without parentheses."""
        args = []
        if self.done():
            return args
        while True:
            if self.done():
                args.append((None, Node("missing")))
                return args
            args.append(self._arg())
            if self.at(","):
                self.take()
                continue
            if not self.done():
                raise VbaUnsupported(f"unexpected {self.peek()!r} after call arguments")
            return args

    def _primary(self):
        p = self.take()
        if p is None:
            raise VbaUnsupported("an expression is missing")
        if p.kind == "num":
            return Node("lit", p.val)
        if p.kind == "str":
            return Node("lit", p.val)
        if p.kind == "op" and p.val == "(":
            e = self.expr()
            self.expect(")")
            return Node("paren", e)
        if p.kind == "op" and p.val == ".":   # a member of the With object
            raise VbaUnsupported("With blocks are not supported")
        if p.kind == "id":
            low = p.low
            if low == "true":
                return Node("lit", True)
            if low == "false":
                return Node("lit", False)
            if low == "nothing":
                return Node("lit", NOTHING)
            if low == "empty":
                return Node("lit", EMPTY)
            if low == "null":
                return Node("lit", NULL)
            if low == "new":
                return Node("new", self.take().val)
            return Node("name", p.val)
        raise VbaUnsupported(f"can't start an expression with {p!r}")


# ======================================================================= parser (statements to flat code)

class Proc:
    def __init__(self, name, kind, params, ret_type, private):
        self.name, self.kind, self.params, self.ret_type, self.private = name, kind, params, ret_type, private
        self.code, self.labels, self.dims, self.consts = [], {}, [], {}


def _split_statements(text):
    """Split a logical line at ':' outside strings, keeping ':=' together."""
    parts, cur, q, i = [], [], False, 0
    while i < len(text):
        ch = text[i]
        if ch == '"':
            q = not q
        if ch == ":" and not q and not text.startswith(":=", i):
            parts.append("".join(cur))
            cur = []
        else:
            cur.append(ch)
        i += 1
    parts.append("".join(cur))
    return [p.strip() for p in parts if p.strip()]


_DECL = re.compile(r"^(?:(Public|Private|Global)\s+)?(Sub|Function)\s+(\w+)\s*\((.*)\)\s*(?:As\s+(\w+))?\s*$", re.I)
_END_PROC = re.compile(r"^End\s+(Sub|Function)\s*$", re.I)


class Module:
    """One parsed module: its variables, constants and procedures."""

    def __init__(self, text, name="Module1"):
        self.name = name
        self.procs, self.vars, self.consts = {}, [], []
        self._parse(text)

    def _parse(self, text):
        lines = logical_lines(text)
        i = 0
        while i < len(lines):
            n, line = lines[i]
            s = line.strip()
            low = s.lower()
            if low.startswith(("attribute ", "option ")):
                i += 1
                continue
            m = _DECL.match(s)
            if m:
                j = i + 1
                while j < len(lines) and not _END_PROC.match(lines[j][1].strip()):
                    j += 1
                if j == len(lines):
                    raise VbaUnsupported(f"{m[3]} has no End {m[2]}")
                proc = Proc(m[3], m[2].lower(), _params(m[4]), m[5], (m[1] or "").lower() == "private")
                for name in [proc.name] + [prm["name"] for prm in proc.params]:
                    _not_reserved(name, n)
                _Compiler(proc).build(lines[i + 1:j])
                self.procs[proc.name.lower()] = proc
                i = j + 1
                continue
            m = re.match(r"^(?:(Public|Private|Global|Dim)\s+)?Const\s+(\w+)\s*(?:As\s+\w+)?\s*=\s*(.+)$", s, re.I)
            if m:
                self.consts.append((m[2], Parser(tokenize(m[3])).expr()))
                i += 1
                continue
            m = re.match(r"^(Public|Private|Dim|Global)\s+(.+)$", s, re.I)
            if m:
                for decl in _dim_list(m[2]):
                    _not_reserved(decl[0], n)
                    self.vars.append(decl)
                i += 1
                continue
            raise VbaUnsupported(f"line {n}: module-level {s[:60]!r}")


def _not_reserved(name, line):
    if name.lower() in VBA_RESERVED:
        raise VbaUnsupported(f"line {line}: {name} is a name VBA reserves, so Excel would not compile the module")


def _params(text):
    out = []
    for part in _split_commas(text):
        part = part.strip()
        if not part:
            continue
        m = re.match(r"^(Optional\s+)?(ByVal\s+|ByRef\s+)?(ParamArray\s+)?(\w+)(\(\))?\s*(?:As\s+(\w+))?\s*(?:=\s*(.+))?$",
                     part, re.I)
        if not m:
            raise VbaUnsupported(f"parameter {part!r}")
        if m[3]:
            raise VbaUnsupported("ParamArray is not supported")
        default = Parser(tokenize(m[7])).expr() if m[7] else None
        out.append({"name": m[4], "byval": bool(m[2]) and m[2].strip().lower() == "byval", "type": m[6] or "Variant",
                    "optional": bool(m[1]), "default": default, "array": bool(m[5])})
    return out


def _split_commas(text):
    parts, cur, depth, q = [], [], 0, False
    for ch in text:
        if ch == '"':
            q = not q
        if not q:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif ch == "," and depth == 0:
                parts.append("".join(cur))
                cur = []
                continue
        cur.append(ch)
    parts.append("".join(cur))
    return parts


def _dim_list(text):
    """Dim a As Long, b(1 To 3) As Variant, c: [(name, type, bounds or None, is_new)]."""
    out = []
    for part in _split_commas(text):
        part = part.strip()
        m = re.match(r"^(\w+)\s*(?:\((.*)\))?\s*(?:As\s+(New\s+)?(\w+))?$", part, re.I)
        if not m:
            raise VbaUnsupported(f"declaration {part!r}")
        bounds = None
        if m[2] is not None:
            bounds = []
            for dim in _split_commas(m[2]):
                if not dim.strip():
                    continue
                mm = re.match(r"^(.+?)\s+To\s+(.+)$", dim.strip(), re.I)
                lo, hi = (mm[1], mm[2]) if mm else ("0", dim)
                bounds.append((Parser(tokenize(lo)).expr(), Parser(tokenize(hi)).expr()))
        out.append((m[1], m[4] or "Variant", bounds, bool(m[3])))
    return out


class _Compiler:
    """One procedure's statements as a flat list of instructions with jumps."""

    def __init__(self, proc):
        self.p = proc
        self.blocks = []   # open blocks: dicts with kind and jump lists

    def emit(self, *ins):
        self.p.code.append(list(ins))
        return len(self.p.code) - 1

    def here(self):
        return len(self.p.code)

    def build(self, lines):
        for n, text in lines:
            stripped = text.strip()
            m = re.match(r"^(\d+)\s+(.*)$", stripped)
            if m:   # a line number: Erl reports it
                self.emit("line", int(m[1]))
                stripped = m[2]
            m = re.match(r"^([A-Za-z_]\w*):(?!=)\s*(.*)$", stripped)
            if m and m[1].lower() not in ("else",):
                self.p.labels[m[1].lower()] = self.here()
                stripped = m[2]
                if not stripped:
                    continue
            try:
                self.line(stripped)
            except VbaUnsupported as e:
                raise VbaUnsupported(f"{self.p.name}, line {n}: {e}") from None
        if self.blocks:
            raise VbaUnsupported(f"{self.p.name}: an unclosed {self.blocks[-1]['kind']}")
        for ins in self.p.code:
            if ins[0] in ("goto", "onerror_goto", "resume_label") and ins[1].lower() not in self.p.labels:
                raise VbaUnsupported(f"{self.p.name}: no label {ins[1]}")

    def line(self, text):
        # a one-line If keeps its : parts together
        if re.match(r"^If\b", text, re.I) and not re.search(r"\bThen\s*$", text, re.I):
            self.single_if(text)
            return
        for part in _split_statements(text):
            self.stmt(part)

    def single_if(self, text):
        toks = tokenize(text)
        p = Parser(toks)
        p.expect("if")
        cond = p.expr()
        p.expect("then")
        rest = toks[p.i:]
        # split rest at a top-level Else (single-line If has no nesting to worry about here)
        depth_if, split = 0, None
        for k, t in enumerate(rest):
            if t.kind == "id" and t.low == "if":
                depth_if += 1
            if t.kind == "id" and t.low == "else":
                if depth_if == 0:
                    split = k
                    break
                depth_if -= 1
        then_toks, else_toks = (rest, []) if split is None else (rest[:split], rest[split + 1:])
        jf = self.emit("jf", cond, None)
        self._stmts_from_tokens(then_toks)
        if else_toks:
            j = self.emit("goto_idx", None)
            self.p.code[jf][2] = self.here()
            self._stmts_from_tokens(else_toks)
            self.p.code[j][1] = self.here()
        else:
            self.p.code[jf][2] = self.here()

    def _stmts_from_tokens(self, toks):
        parts, cur = [], []
        for t in toks:
            if t.kind == "op" and t.val == ":":
                parts.append(cur)
                cur = []
            else:
                cur.append(t)
        parts.append(cur)
        for part in parts:
            if part:
                if part[0].kind == "id" and part[0].low == "if":
                    raise VbaUnsupported("a nested one-line If")
                self.stmt_tokens(part)

    def stmt(self, text):
        self.stmt_tokens(tokenize(text), text)

    def stmt_tokens(self, toks, text=None):
        first = toks[0]
        low = first.low if first.kind == "id" else first.val
        second = toks[1].low if len(toks) > 1 and toks[1].kind == "id" else None
        P = Parser(toks)
        B = self.blocks

        if low == "if":   # block If
            P.take()
            cond = P.expr()
            P.expect("then")
            jf = self.emit("jf", cond, None)
            B.append({"kind": "if", "jf": jf, "ends": []})
            return
        if low == "elseif":
            b = B[-1]
            b["ends"].append(self.emit("goto_idx", None))
            self.p.code[b["jf"]][2] = self.here()
            P.take()
            cond = P.expr()
            P.expect("then")
            b["jf"] = self.emit("jf", cond, None)
            return
        if low == "else" and len(toks) == 1:
            b = B[-1]
            if b["kind"] == "if":
                b["ends"].append(self.emit("goto_idx", None))
                self.p.code[b["jf"]][2] = self.here()
                b["jf"] = None
                return
            raise VbaUnsupported("Else outside If")
        if low == "end" and second == "if":
            b = B.pop()
            if b["jf"] is not None:
                self.p.code[b["jf"]][2] = self.here()
            for e in b["ends"]:
                self.p.code[e][1] = self.here()
            return
        if low == "for" and second == "each":
            P.take()
            P.take()
            var = P.take().val
            P.expect("in")
            coll = P.expr()
            init = self.emit("foreach_init", var, coll)
            test = self.emit("foreach_next", var, None)
            B.append({"kind": "for", "test": test, "exits": [], "init": init})
            return
        if low == "for":
            P.take()
            var = P.take().val
            P.expect("=")
            start = P.expr()
            P.expect("to")
            stop = P.expr()
            step = Node("lit", 1)
            if P.at("step"):
                P.take()
                step = P.expr()
            self.emit("for_init", var, start, stop, step)
            test = self.emit("for_test", var, None)
            B.append({"kind": "for", "test": test, "exits": [], "var": var})
            return
        if low == "next":
            b = B.pop()
            if b["kind"] != "for":
                raise VbaUnsupported("Next without For")
            if "var" in b:
                self.emit("for_step", b["var"], b["test"])
            else:
                self.emit("goto_idx", b["test"])
            self.p.code[b["test"]][2] = self.here()
            for e in b["exits"]:
                self.p.code[e][1] = self.here()
            return
        if low == "do":
            P.take()
            top = self.here()
            jf = None
            if P.at("while", "until"):
                kind = P.take().low
                cond = P.expr()
                jf = self.emit("jf" if kind == "while" else "jt", cond, None)
            B.append({"kind": "do", "top": top, "jf": jf, "exits": []})
            return
        if low == "loop":
            b = B.pop()
            P.take()
            if P.at("while", "until"):
                kind = P.take().low
                cond = P.expr()
                self.emit("jt" if kind == "while" else "jf", cond, b["top"])
            else:
                self.emit("goto_idx", b["top"])
            if b["jf"] is not None:
                self.p.code[b["jf"]][2] = self.here()
            for e in b["exits"]:
                self.p.code[e][1] = self.here()
            return
        if low == "while":
            P.take()
            top = self.here()
            jf = self.emit("jf", P.expr(), None)
            B.append({"kind": "do", "top": top, "jf": jf, "exits": []})
            return
        if low == "wend":
            b = B.pop()
            self.emit("goto_idx", b["top"])
            self.p.code[b["jf"]][2] = self.here()
            return
        if low == "select":
            P.take()
            P.expect("case")
            sel = P.expr()
            slot = f"$select{len(self.p.code)}"
            self.emit("set_hidden", slot, sel)
            B.append({"kind": "select", "slot": slot, "jf": None, "ends": []})
            return
        if low == "case":
            b = B[-1]
            if b["jf"] is not None:
                b["ends"].append(self.emit("goto_idx", None))
                self.p.code[b["jf"]][-1] = self.here()
            P.take()
            if P.at("else"):
                b["jf"] = None
                return
            tests = []
            while True:
                if P.at("is"):
                    P.take()
                    op = P.take().val
                    tests.append(("is", op, P.expr()))
                else:
                    lo = P.expr()
                    if P.at("to"):
                        P.take()
                        tests.append(("to", lo, P.expr()))
                    else:
                        tests.append(("eq", lo))
                if P.at(","):
                    P.take()
                    continue
                break
            b["jf"] = self.emit("case_jf", b["slot"], tests, None)
            return
        if low == "end" and second == "select":
            b = B.pop()
            if b["jf"] is not None:
                self.p.code[b["jf"]][-1] = self.here()
            for e in b["ends"]:
                self.p.code[e][1] = self.here()
            return
        if low == "exit":
            what = second
            if what in ("sub", "function"):
                self.emit("return")
                return
            kind = {"for": "for", "do": "do"}[what]
            for b in reversed(B):
                if b["kind"] == kind:
                    b["exits"].append(self.emit("goto_idx", None))
                    return
            raise VbaUnsupported(f"Exit {what} outside a {what}")
        if low == "goto":
            self.emit("goto", toks[1].val)
            return
        if low == "on" and second == "error":
            rest = [t.low if t.kind == "id" else str(t.val) for t in toks[2:]]
            if rest == ["resume", "next"]:
                self.emit("onerror_next")
            elif rest == ["goto", "0"]:
                self.emit("onerror_off")
            elif rest[:1] == ["goto"]:
                self.emit("onerror_goto", toks[3].val)
            else:
                raise VbaUnsupported("On Error " + " ".join(rest))
            return
        if low == "resume":
            if len(toks) == 1:
                self.emit("resume_same")
            elif second == "next":
                self.emit("resume_next")
            else:
                self.emit("resume_label", toks[1].val)
            return
        if low in ("dim", "static"):
            for name, typ, bounds, is_new in _dim_list(_tokens_text(toks[1:])):
                if name.lower() in VBA_RESERVED:
                    raise VbaUnsupported(f"{name} is a name VBA reserves, so Excel would not compile the module")
                self.emit("dim", name, typ, bounds, is_new)
            return
        if low == "redim":
            j = 1
            preserve = toks[1].kind == "id" and toks[1].low == "preserve"
            if preserve:
                j = 2
            for name, typ, bounds, _ in _dim_list(_tokens_text(toks[j:])):
                self.emit("redim", name, bounds, preserve)
            return
        if low == "const":
            m = re.match(r"^Const\s+(\w+)\s*(?:As\s+\w+)?\s*=\s*(.+)$", _tokens_text(toks), re.I)
            self.emit("const", m[1], Parser(tokenize(m[2])).expr())
            return
        if low == "set":
            P.take()
            target = P._postfix()
            P.expect("=")
            self.emit("set", target, P.expr())
            return
        if low == "call":
            P.take()
            target = P._postfix()
            self.emit("call", target if target.kind == "call" else Node("call", target, []))
            return
        if low in ("with",):
            raise VbaUnsupported("With blocks are not supported")
        if low in ("open", "close", "get", "put", "print", "line", "input", "write", "kill", "mkdir", "rmdir",
                   "filecopy", "name", "shell", "declare"):
            if low != "print" and not (low == "name" and len(toks) > 1 and toks[1].val in ("=",)):
                self.file_stmt(low, toks)
                return
        # assignment, or a call statement
        target = P._postfix()
        if P.at("="):
            P.take()
            self.emit("let", target, P.expr())
            return
        if target.kind == "call" and P.done():
            self.emit("call", target)
            return
        if target.kind in ("name", "member"):
            self.emit("call", Node("call", target, P.args_bare(), "bare"))
            return
        if target.kind == "call" and not P.done():
            # Foo(a) , b  -> parenthesised first argument of a bare call
            rest = P.args_bare()
            self.emit("call", Node("call", target.a, [(None, Node("paren", _paren_arg(target)))] + rest, "bare"))
            return
        raise VbaUnsupported(f"statement {_tokens_text(toks)[:60]!r}")

    def file_stmt(self, low, toks):
        """Open, Get, Close and the like: kept as written, run by the simulator's file model (5c)."""
        if low in ("kill", "mkdir", "rmdir", "filecopy", "shell", "declare", "name", "put"):
            raise VbaUnsupported(f"{low.title()} is never allowed in generated VBA")
        self.emit("file", low, toks)


def _paren_arg(call):
    if len(call.b) != 1:
        raise VbaUnsupported("a call with a parenthesised list and more arguments")
    return call.b[0][1]


def _tokens_text(toks):
    out = []
    for t in toks:
        if t.kind == "str":
            out.append('"' + t.val.replace('"', '""') + '"')
        elif t.kind == "num":
            out.append(repr(t.val) + ("#" if isinstance(t.val, float) else ""))
        else:
            out.append(str(t.val))
    return " ".join(out).replace(" ( ", "(").replace(" )", ")").replace("( ", "(")


# ======================================================================= objects

class VbaObject:
    vba_type = "Object"

    def get(self, sim, name, args, named):
        raise VbaError(438, f"Object doesn't support this property or method ({self.vba_type}.{name})")

    def let(self, sim, name, args, value):
        raise VbaError(438, f"Object doesn't support this property or method ({self.vba_type}.{name})")

    def call(self, sim, name, args, named):
        return self.get(sim, name, args, named)

    def iterate(self, sim):
        raise VbaError(438, f"Object doesn't support For Each ({self.vba_type})")


class Collection(VbaObject):
    vba_type = "Collection"

    def __init__(self):
        self.items, self.keys = [], []

    def _index(self, k):
        if isinstance(k, str):
            for i, key in enumerate(self.keys):
                if key is not None and key.lower() == k.lower():
                    return i
            raise VbaError(5, "Invalid procedure call or argument")
        i = to_long(k) - 1
        if not 0 <= i < len(self.items):
            raise VbaError(9, "Subscript out of range")
        return i

    def get(self, sim, name, args, named):
        n = (name or "item").lower()
        if n == "item":
            return self.items[self._index(args[0])]
        if n == "count":
            return len(self.items)
        if n == "add":
            item = args[0] if args else named.get("item")
            key = args[1] if len(args) > 1 and args[1] is not MISSING else named.get("key")
            before = args[2] if len(args) > 2 and args[2] is not MISSING else named.get("before")
            after = args[3] if len(args) > 3 and args[3] is not MISSING else named.get("after")
            key = None if key is None or key is MISSING else to_str(key)
            if key is not None and any(k is not None and k.lower() == key.lower() for k in self.keys):
                raise VbaError(457, "This key is already associated with an element of this collection")
            if before is not None and before is not MISSING:
                i = self._index(before)
            elif after is not None and after is not MISSING:
                i = self._index(after) + 1
            else:
                i = len(self.items)
            self.items.insert(i, item)
            self.keys.insert(i, key)
            return EMPTY
        if n == "remove":
            i = self._index(args[0])
            del self.items[i]
            del self.keys[i]
            return EMPTY
        return super().get(sim, name, args, named)

    def iterate(self, sim):
        return list(self.items)


class Dictionary(VbaObject):
    vba_type = "Dictionary"

    def __init__(self):
        self.data = {}   # normalised key -> [key, value]
        self.compare = 0

    def _norm(self, k):
        if isinstance(k, str):
            return ("s", k.lower() if self.compare == 1 else k)
        if is_num(k) or isinstance(k, bool):
            return ("n", to_double(k))
        if isinstance(k, VDate):
            return ("d", k.serial)
        if k is EMPTY:
            return ("e",)
        return ("o", id(k))

    def get(self, sim, name, args, named):
        n = (name or "item").lower()
        if n == "item":
            k = self._norm(args[0])
            if k not in self.data:
                self.data[k] = [args[0], EMPTY]   # reading a missing key adds it, as Scripting.Dictionary does
            return self.data[k][1]
        if n == "exists":
            return self._norm(args[0]) in self.data
        if n == "add":
            k = self._norm(args[0])
            if k in self.data:
                raise VbaError(457, "This key is already associated with an element of this collection")
            self.data[k] = [args[0], args[1]]
            return EMPTY
        if n == "count":
            return len(self.data)
        if n == "keys":
            return VArray.from_list([kv[0] for kv in self.data.values()])
        if n == "items":
            return VArray.from_list([kv[1] for kv in self.data.values()])
        if n == "remove":
            k = self._norm(args[0])
            if k not in self.data:
                raise VbaError(32811, "Element not found")
            del self.data[k]
            return EMPTY
        if n == "removeall":
            self.data.clear()
            return EMPTY
        if n == "comparemode":
            return self.compare
        return super().get(sim, name, args, named)

    def let(self, sim, name, args, value):
        n = (name or "item").lower()
        if n == "item":
            k = self._norm(args[0])
            if k in self.data:
                self.data[k][1] = value
            else:
                self.data[k] = [args[0], value]
            return
        if n == "comparemode":
            if self.data:
                raise VbaError(5, "Invalid procedure call or argument")
            self.compare = to_long(value)
            return
        super().let(sim, name, args, value)

    def iterate(self, sim):
        return [kv[0] for kv in self.data.values()]


class ErrObject(VbaObject):
    vba_type = "ErrObject"

    def __init__(self):
        self.number, self.description, self.source = 0, "", ""

    def get(self, sim, name, args, named):
        n = (name or "number").lower()
        if n == "number":
            return self.number
        if n == "description":
            return self.description
        if n == "source":
            return self.source
        if n == "clear":
            self.number, self.description, self.source = 0, "", ""
            return EMPTY
        if n == "raise":
            number = args[0] if args else named.get("number")
            source = args[1] if len(args) > 1 and args[1] is not MISSING else named.get("source", "")
            desc = args[2] if len(args) > 2 and args[2] is not MISSING else named.get("description", "")
            raise VbaError(to_long(number), to_str(desc) if desc is not MISSING else "", to_str(source or ""))
        return super().get(sim, name, args, named)

    def let(self, sim, name, args, value):
        n = (name or "number").lower()
        if n == "number":
            self.number = to_long(value)
        elif n == "description":
            self.description = to_str(value)
        else:
            super().let(sim, name, args, value)


class Application(VbaObject):
    vba_type = "Application"

    def __init__(self):
        self.props = {"screenupdating": True, "calculation": -4105, "enableevents": True,
                      "displayalerts": True, "statusbar": False}

    def get(self, sim, name, args, named):
        n = name.lower()
        if n in self.props:
            return self.props[n]
        if n == "thisworkbook" or n == "activeworkbook":
            return sim.this_workbook
        return super().get(sim, name, args, named)

    def let(self, sim, name, args, value):
        n = name.lower()
        if n in self.props:
            self.props[n] = value
            return
        super().let(sim, name, args, value)


class Debug(VbaObject):
    vba_type = "Debug"

    def get(self, sim, name, args, named):
        if name.lower() == "print":
            sim.debug.append(" ".join(to_str(a) for a in args if a is not MISSING))
            return EMPTY
        return super().get(sim, name, args, named)


# ----------------------------------------------------------------------- the workbook model

def col_letters(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def col_number(letters):
    n = 0
    for ch in letters.upper():
        n = n * 26 + ord(ch) - 64
    return n


_A1 = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")


def parse_a1(ref):
    """A1, A1:F1, A:F or 3:3 as (r1, c1, r2, c2)."""
    ref = ref.replace("$", "").strip()
    parts = ref.split(":")
    if len(parts) == 1:
        parts = parts * 2
    m1, m2 = _A1.match(parts[0]), _A1.match(parts[1])
    if m1 and m2:
        r1, c1, r2, c2 = int(m1[2]), col_number(m1[1]), int(m2[2]), col_number(m2[1])
    elif re.match(r"^[A-Za-z]{1,3}$", parts[0]) and re.match(r"^[A-Za-z]{1,3}$", parts[1]):
        r1, c1, r2, c2 = 1, col_number(parts[0]), ROWS, col_number(parts[1])
    elif parts[0].isdigit() and parts[1].isdigit():
        r1, c1, r2, c2 = int(parts[0]), 1, int(parts[1]), COLS
    else:
        raise VbaError(1004, f"Method 'Range' of object '_Worksheet' failed ({ref})")
    return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


class Workbook(VbaObject):
    vba_type = "Workbook"

    def __init__(self, name="Book1.xlsx"):
        self.name = name
        self.sheets = []

    def add_sheet(self, name, after=None):
        ws = Worksheet(self, name)
        if after is None:
            self.sheets.append(ws)
        else:
            self.sheets.insert(self.sheets.index(after) + 1, ws)
        return ws

    def sheet(self, key):
        if isinstance(key, str):
            for ws in self.sheets:
                if ws.name.lower() == key.lower():
                    return ws
            raise VbaError(9, "Subscript out of range")
        i = to_long(key)
        if not 1 <= i <= len(self.sheets):
            raise VbaError(9, "Subscript out of range")
        return self.sheets[i - 1]

    def get(self, sim, name, args, named):
        n = name.lower()
        if n in ("worksheets", "sheets"):
            coll = SheetsCollection(self)
            return coll.get(sim, None, args, named) if args else coll
        if n == "name":
            return self.name
        return super().get(sim, name, args, named)


class SheetsCollection(VbaObject):
    vba_type = "Sheets"

    def __init__(self, book):
        self.book = book

    def get(self, sim, name, args, named):
        n = (name or "item").lower()
        if n == "item":
            return self.book.sheet(args[0])
        if n == "count":
            return len(self.book.sheets)
        if n == "add":
            after = named.get("after", args[1] if len(args) > 1 else None)
            before = named.get("before", args[0] if args and args[0] is not MISSING else None)
            k = len(self.book.sheets) + 1
            while any(s.name.lower() == f"sheet{k}" for s in self.book.sheets):
                k += 1
            if before not in (None, MISSING):
                ws = Worksheet(self.book, f"Sheet{k}")
                self.book.sheets.insert(self.book.sheets.index(before), ws)
                return ws
            return self.book.add_sheet(f"Sheet{k}", after if after not in (None, MISSING) else None)
        return super().get(sim, name, args, named)

    def iterate(self, sim):
        return list(self.book.sheets)


class Worksheet(VbaObject):
    vba_type = "Worksheet"

    def __init__(self, book, name):
        self.book, self.name = book, name
        self.cells = {}     # (r, c) -> value: str, float, bool, VDate, VError (never Empty)
        self.formulas = {}  # (r, c) -> formula text as written
        self.fills = {}     # (r, c) -> colour (BGR long, as Interior.Color)
        self.formats = {}   # (r, c) -> number format
        self.hidden = set()

    # -- storage
    def value(self, r, c):
        return self.cells.get((r, c), EMPTY)

    def store(self, r, c, v):
        """What Excel keeps when a value is written to a cell."""
        self.formulas.pop((r, c), None)
        if v is EMPTY or v is NULL:
            self.cells.pop((r, c), None)
            return
        if isinstance(v, str):
            if v.startswith("'"):
                v = v[1:]
            else:
                v = _excel_reads_text(v)
                if isinstance(v, tuple):   # a formula
                    self.formulas[(r, c)] = v[1]
                    self.cells.pop((r, c), None)
                    return
            if v == "":
                self.cells.pop((r, c), None)
                return
        elif isinstance(v, bool) or isinstance(v, (VDate, VError)):
            pass
        elif is_num(v):
            v = float(v)
        elif isinstance(v, VArray):
            raise VbaError(13, "Type mismatch")
        else:
            raise VbaError(13, "Type mismatch")
        self.cells[(r, c)] = v

    def shift_rows(self, row, step):
        """Delete (step -1) or insert (step 1) the sheet row `row`: the rows below move. An inserted row takes
        the formats of the row above it, as Excel's Insert does; ForgeInsertRow then clears them."""
        def moved(d):
            out = {}
            for (r, c), v in d.items():
                if step < 0 and r == row:
                    continue
                out[(r + step if r >= row + (1 if step < 0 else 0) else r, c)] = v
            return out
        above = {c: f for (r, c), f in self.formats.items() if r == row - 1} if step > 0 else {}
        fills_above = {c: f for (r, c), f in self.fills.items() if r == row - 1} if step > 0 else {}
        self.cells, self.formulas = moved(self.cells), moved(self.formulas)
        self.fills, self.formats = moved(self.fills), moved(self.formats)
        if step < 0:
            self.hidden = {r - 1 if r > row else r for r in self.hidden if r != row}
        else:
            self.hidden = {r + 1 if r >= row else r for r in self.hidden}
            for c, f in above.items():
                self.formats[(row, c)] = f
            for c, f in fills_above.items():
                self.fills[(row, c)] = f

    def used_box(self):
        keys = set(self.cells) | set(self.formulas) | set(self.fills) | set(self.formats)
        if not keys:
            return 1, 1, 1, 1
        rows = [k[0] for k in keys]
        cols = [k[1] for k in keys]
        return min(rows), min(cols), max(rows), max(cols)

    def rng(self, r1, c1, r2, c2):
        return Range(self, r1, c1, r2, c2)

    def get(self, sim, name, args, named):
        n = name.lower()
        if n == "name":
            return self.name
        if n == "cells":
            if not args:
                return self.rng(1, 1, ROWS, COLS)
            if len(args) == 2:
                return self.rng(to_long(args[0]), to_long(args[1]), to_long(args[0]), to_long(args[1]))
            return self.rng(1, 1, ROWS, COLS).get(sim, "item", args, named)
        if n == "range":
            if len(args) == 1:
                return self.rng(*parse_a1(to_str(args[0])))
            a, b = args[0], args[1]
            if not isinstance(a, Range) or not isinstance(b, Range):
                a = self.rng(*parse_a1(to_str(a))) if not isinstance(a, Range) else a
                b = self.rng(*parse_a1(to_str(b))) if not isinstance(b, Range) else b
            return self.rng(min(a.r1, b.r1), min(a.c1, b.c1), max(a.r2, b.r2), max(a.c2, b.c2))
        if n == "rows":
            if not args:
                return self.rng(1, 1, ROWS, COLS).get(sim, "rows", [], named)
            r = to_long(args[0])
            return self.rng(r, 1, r, COLS)
        if n == "columns":
            if not args:
                return self.rng(1, 1, ROWS, COLS).get(sim, "columns", [], named)
            if isinstance(args[0], str):
                r1, c1, r2, c2 = parse_a1(args[0] if ":" in args[0] else args[0] + ":" + args[0])
                return self.rng(1, c1, ROWS, c2)
            c = to_long(args[0])
            return self.rng(1, c, ROWS, c)
        if n == "usedrange":
            return self.rng(*self.used_box())
        if n == "parent":
            return self.book
        if n == "index":
            return self.book.sheets.index(self) + 1
        if n == "copy":
            after = named.get("after", args[1] if len(args) > 1 else None)
            copy = Worksheet(self.book, self.name + " (2)")
            copy.cells, copy.formulas = dict(self.cells), dict(self.formulas)
            copy.fills, copy.formats, copy.hidden = dict(self.fills), dict(self.formats), set(self.hidden)
            idx = self.book.sheets.index(after) + 1 if after not in (None, MISSING) else len(self.book.sheets)
            self.book.sheets.insert(idx, copy)
            return EMPTY
        return super().get(sim, name, args, named)

    def let(self, sim, name, args, value):
        if name.lower() == "name":
            self.name = to_str(value)
            return
        super().let(sim, name, args, value)


_EXCEL_ERRORS = {"#NULL!": 2000, "#DIV/0!": 2007, "#VALUE!": 2015, "#REF!": 2023, "#NAME?": 2029,
                 "#NUM!": 2036, "#N/A": 2042}


def _excel_reads_text(s):
    """What Excel 2016 (English) makes of text typed into a General cell: a number, TRUE/FALSE, an
    error, a formula, or the text. ForgeSetValue guards every such text with ', so this is the rare case."""
    t = s.strip()
    if t.startswith("=") and len(t) > 1:
        return ("formula", s)
    if t.upper() in ("TRUE", "FALSE"):
        return t.upper() == "TRUE"
    if t.upper() in _EXCEL_ERRORS:
        return VError(_EXCEL_ERRORS[t.upper()])
    n = _parse_number_text(t) if re.match(r"^[+-]?[\d.,]+([eE][+-]?\d+)?$", t) else None
    if n is not None:
        return float(n)
    return s


class Interior(VbaObject):
    vba_type = "Interior"

    def __init__(self, rng):
        self.rng = rng

    def get(self, sim, name, args, named):
        n = name.lower()
        ws, cells = self.rng.ws, list(self.rng.each())
        if n == "color":
            colours = {ws.fills.get(rc, 16777215) for rc in cells}
            return float(colours.pop()) if len(colours) == 1 else NULL
        if n in ("pattern", "colorindex"):
            fills = {rc in ws.fills for rc in cells}
            if fills == {False}:
                return -4142
            return 1 if fills == {True} else NULL
        return super().get(sim, name, args, named)

    def let(self, sim, name, args, value):
        n = name.lower()
        ws = self.rng.ws
        if n == "color":
            for rc in self.rng.each():
                ws.fills[rc] = int(to_double(value))
            return
        if n in ("pattern", "colorindex"):
            if to_long(value) in (-4142, 0):
                for rc in self.rng.each():
                    ws.fills.pop(rc, None)
                return
        super().let(sim, name, args, value)


class Range(VbaObject):
    vba_type = "Range"

    def __init__(self, ws, r1, c1, r2, c2):
        self.ws, self.r1, self.c1, self.r2, self.c2 = ws, r1, c1, r2, c2

    def __eq__(self, other):
        return isinstance(other, Range) and (self.ws, self.r1, self.c1, self.r2, self.c2) == (
            other.ws, other.r1, other.c1, other.r2, other.c2)

    def __hash__(self):
        return hash((id(self.ws), self.r1, self.c1, self.r2, self.c2))

    def each(self):
        """The cells, row by row, but never more than the used part of a whole-sheet range."""
        ur1, uc1, ur2, uc2 = self.ws.used_box()
        r2, c2 = min(self.r2, max(ur2, self.r1)), min(self.c2, max(uc2, self.c1))
        for r in range(self.r1, r2 + 1):
            for c in range(self.c1, c2 + 1):
                yield (r, c)

    def nrows(self):
        return self.r2 - self.r1 + 1

    def ncols(self):
        return self.c2 - self.c1 + 1

    def single(self):
        return self.r1 == self.r2 and self.c1 == self.c2

    def get_value(self):
        if self.single():
            return self.ws.value(self.r1, self.c1)
        a = VArray([(1, self.nrows()), (1, self.ncols())])
        for r in range(self.r1, self.r2 + 1):
            for c in range(self.c1, self.c2 + 1):
                a.set([r - self.r1 + 1, c - self.c1 + 1], self.ws.value(r, c))
        return a

    def set_value(self, v):
        ws = self.ws
        if isinstance(v, VArray):
            if len(v.bounds) == 1:
                lo, hi = v.bounds[0]
                for r in range(self.r1, self.r2 + 1):
                    for c in range(self.c1, self.c2 + 1):
                        i = lo + (c - self.c1)
                        ws.store(r, c, v.get([i]) if i <= hi else VError(2042))
                return
            (rlo, rhi), (clo, chi) = v.bounds
            for r in range(self.r1, self.r2 + 1):
                for c in range(self.c1, self.c2 + 1):
                    i, j = rlo + (r - self.r1), clo + (c - self.c1)
                    ws.store(r, c, v.get([i, j]) if i <= rhi and j <= chi else VError(2042))
            return
        for rc in self.each():
            ws.store(rc[0], rc[1], v)

    def end(self, direction):
        ws = self.ws
        r, c = self.r1, self.c1
        dr, dc = {-4162: (-1, 0), -4121: (1, 0), -4159: (0, -1), -4161: (0, 1)}[direction]
        rmax, cmax = ROWS, COLS

        def filled(rr, cc):
            return (rr, cc) in ws.cells or (rr, cc) in ws.formulas

        def inside(rr, cc):
            return 1 <= rr <= rmax and 1 <= cc <= cmax

        nr, nc = r + dr, c + dc
        if not inside(nr, nc):
            return r, c
        if filled(r, c) and filled(nr, nc):
            while inside(nr + dr, nc + dc) and filled(nr + dr, nc + dc):
                nr, nc = nr + dr, nc + dc
            return nr, nc
        # find the next filled cell in that direction
        if dr:
            rows = [rr for (rr, cc) in set(ws.cells) | set(ws.formulas) if cc == c and (rr - r) * dr > 0]
            if not rows:
                return (1 if dr < 0 else rmax), c
            return (max(rows) if dr < 0 else min(rows)), c
        cols = [cc for (rr, cc) in set(ws.cells) | set(ws.formulas) if rr == r and (cc - c) * dc > 0]
        if not cols:
            return r, (1 if dc < 0 else cmax)
        return r, (max(cols) if dc < 0 else min(cols))

    def get(self, sim, name, args, named):
        n = (name or "value").lower()
        ws = self.ws
        if n in ("value", "value2"):
            if args:
                raise VbaError(438, "Range.Value with arguments")
            return self.get_value()
        if n == "item" or (name is None and args):
            r = to_long(args[0])
            if len(args) > 1:
                c = to_long(args[1])
                return Range(ws, self.r1 + r - 1, self.c1 + c - 1, self.r1 + r - 1, self.c1 + c - 1)
            k = r - 1
            rr, cc = divmod(k, self.ncols())
            return Range(ws, self.r1 + rr, self.c1 + cc, self.r1 + rr, self.c1 + cc)
        if n == "cells":
            if args:
                return self.get(sim, "item", args, named)
            return self
        if n == "row":
            return self.r1
        if n == "column":
            return self.c1
        if n == "rows":
            if args:
                r = self.r1 + to_long(args[0]) - 1
                return Range(ws, r, self.c1, r, self.c2)
            return RangeRows(self)
        if n == "columns":
            if args:
                c = self.c1 + to_long(args[0]) - 1
                return Range(ws, self.r1, c, self.r2, c)
            return RangeCols(self)
        if n in ("count", "countlarge"):
            return self.nrows() * self.ncols()
        if n == "worksheet" or n == "parent":
            return ws
        if n == "end":
            r, c = self.end(to_long(args[0]))
            return Range(ws, r, c, r, c)
        if n == "interior":
            return Interior(self)
        if n == "numberformat":
            fmts = {ws.formats.get(rc, "General") for rc in self.each()}
            return fmts.pop() if len(fmts) == 1 else NULL
        if n in ("formula", "formular1c1"):
            if self.single():
                f = ws.formulas.get((self.r1, self.c1))
                return f if f is not None else to_str(ws.value(self.r1, self.c1)) if ws.value(self.r1, self.c1) is not EMPTY else ""
            raise VbaError(438, "Formula of several cells")
        if n == "address":
            absolute = not (args and args[0] is not MISSING and not to_bool(args[0]))
            d = "$" if absolute else ""
            a = f"{d}{col_letters(self.c1)}{d}{self.r1}"
            return a if self.single() else a + f":{d}{col_letters(self.c2)}{d}{self.r2}"
        if n == "resize":
            nr = to_long(args[0]) if args and args[0] is not MISSING else self.nrows()
            nc = to_long(args[1]) if len(args) > 1 and args[1] is not MISSING else self.ncols()
            return Range(ws, self.r1, self.c1, self.r1 + nr - 1, self.c1 + nc - 1)
        if n == "offset":
            dr = to_long(args[0]) if args and args[0] is not MISSING else 0
            dc = to_long(args[1]) if len(args) > 1 and args[1] is not MISSING else 0
            return Range(ws, self.r1 + dr, self.c1 + dc, self.r2 + dr, self.c2 + dc)
        if n == "entirerow":
            return Range(ws, self.r1, 1, self.r2, COLS)
        if n == "hidden":
            states = {r in ws.hidden for r in range(self.r1, self.r2 + 1)}
            return states.pop() if len(states) == 1 else NULL
        if n in ("clear", "clearcontents", "clearformats"):
            for rc in list(self.each()) if (self.r2 - self.r1 + 1) * (self.c2 - self.c1 + 1) < 10 ** 6 else self._used_cells():
                if n in ("clear", "clearcontents"):
                    ws.cells.pop(rc, None)
                    ws.formulas.pop(rc, None)
                if n in ("clear", "clearformats"):
                    ws.fills.pop(rc, None)
                    ws.formats.pop(rc, None)
            return EMPTY
        if n == "autofit":
            return EMPTY
        if n == "sort":
            self.sort(named, args)
            return EMPTY
        if n == "delete":
            if self.c1 != 1 or self.c2 != COLS:
                raise VbaUnsupported("Delete of part of a row")
            for _ in range(self.nrows()):
                ws.shift_rows(self.r1, -1)
            return EMPTY
        if n == "insert":
            if self.c1 != 1 or self.c2 != COLS:
                raise VbaUnsupported("Insert of part of a row")
            for _ in range(self.nrows()):
                ws.shift_rows(self.r1, 1)
            return EMPTY
        return super().get(sim, name, args, named)

    def sort(self, named, args):
        """Range.Sort as ForgeSort calls it: one key, MatchCase False, top to bottom. Excel's order: numbers and
        dates, then text (case ignored), then FALSE and TRUE, then errors; blanks always last; a stable sort."""
        ws = self.ws
        key = named.get("key1", args[0] if args else MISSING)
        if not isinstance(key, Range):
            raise VbaUnsupported("Range.Sort without a Key1 cell")
        order = to_long(named.get("order1", 1))
        header = to_long(named.get("header", 2)) == 1
        if to_long(named.get("orientation", 1)) != 1 or to_bool(named.get("matchcase", False)):
            raise VbaUnsupported("Range.Sort other than top to bottom, case ignored")
        first = self.r1 + (1 if header else 0)
        cols = range(self.c1, self.c2 + 1)
        rows = []
        for r in range(first, self.r2 + 1):
            rows.append({c: (ws.cells.get((r, c)), ws.formulas.get((r, c)), ws.fills.get((r, c)), ws.formats.get((r, c)))
                         for c in cols})

        def rank(v):
            if isinstance(v, bool):
                return (2, v)
            if isinstance(v, VError):
                return (3, 0)
            if isinstance(v, VDate):
                return (0, v.serial)
            if is_num(v):
                return (0, v)
            return (1, str(v).casefold())

        kc = key.c1
        filled = [i for i in range(len(rows)) if rows[i][kc][0] is not None]
        empty = [i for i in range(len(rows)) if rows[i][kc][0] is None]
        filled.sort(key=lambda i: rank(rows[i][kc][0]), reverse=(order == 2))
        for offset, i in enumerate(filled + empty):
            r = first + offset
            for c in cols:
                v, f, fill, fmt = rows[i][c]
                for store, x in ((ws.cells, v), (ws.formulas, f), (ws.fills, fill), (ws.formats, fmt)):
                    if x is None:
                        store.pop((r, c), None)
                    else:
                        store[(r, c)] = x

    def _used_cells(self):
        keys = set(self.ws.cells) | set(self.ws.formulas) | set(self.ws.fills) | set(self.ws.formats)
        return [k for k in keys if self.r1 <= k[0] <= self.r2 and self.c1 <= k[1] <= self.c2]

    def let(self, sim, name, args, value):
        n = (name or "value").lower()
        ws = self.ws
        if n in ("value", "value2"):
            self.set_value(value)
            return
        if n == "numberformat":
            for rc in self.each():
                if to_str(value) == "General":
                    ws.formats.pop(rc, None)
                else:
                    ws.formats[rc] = to_str(value)
            return
        if n in ("formula", "formular1c1"):
            for rc in self.each():
                ws.cells.pop(rc, None)
                ws.formulas[rc] = to_str(value)
            return
        if n == "hidden":
            for r in range(self.r1, self.r2 + 1):
                if to_bool(value):
                    ws.hidden.add(r)
                else:
                    ws.hidden.discard(r)
            return
        super().let(sim, name, args, value)

    def iterate(self, sim):
        return [Range(self.ws, r, c, r, c) for r, c in self.each()]


class RangeRows(VbaObject):
    vba_type = "Range"

    def __init__(self, rng):
        self.rng = rng

    def get(self, sim, name, args, named):
        n = (name or "item").lower()
        if n == "count":
            return self.rng.nrows()
        if n == "item":
            r = self.rng.r1 + to_long(args[0]) - 1
            return Range(self.rng.ws, r, self.rng.c1, r, self.rng.c2)
        return super().get(sim, name, args, named)


class RangeCols(RangeRows):
    def get(self, sim, name, args, named):
        n = (name or "item").lower()
        if n == "count":
            return self.rng.ncols()
        if n == "item":
            c = self.rng.c1 + to_long(args[0]) - 1
            return Range(self.rng.ws, self.rng.r1, c, self.rng.r2, c)
        return super().get(sim, name, args, named)


# ======================================================================= built-in functions

CONSTANTS = {
    "vbtextcompare": 1, "vbbinarycompare": 0, "vbempty": 0, "vbnull": 1, "vbinteger": 2, "vblong": 3, "vbsingle": 4,
    "vbdouble": 5, "vbcurrency": 6, "vbdate": 7, "vbstring": 8, "vbobject": 9, "vberror": 10, "vbboolean": 11,
    "vbvariant": 12, "vbdecimal": 14, "vbbyte": 17, "vbarray": 8192,
    "vbcrlf": "\r\n", "vbcr": "\r", "vblf": "\n", "vbnewline": "\r\n", "vbtab": "\t", "vbnullstring": "",
    "vbnullchar": "\0", "vbokonly": 0, "vbinformation": 64, "vbexclamation": 48, "vbcritical": 16,
    "xlerrna": 2042, "xlerrdiv0": 2007, "xlerrvalue": 2015, "xlerrref": 2023, "xlerrname": 2029,
    "xlerrnum": 2036, "xlerrnull": 2000,
    "xlup": -4162, "xldown": -4121, "xltoleft": -4159, "xltoright": -4161,
    "xlcalculationmanual": -4135, "xlcalculationautomatic": -4105, "xlcalculationsemiautomatic": 2,
    "xlascending": 1, "xldescending": 2, "xlyes": 1, "xlno": 2, "xlguess": 0, "xltoptobottom": 1,
    "xllefttoright": 2, "xlsortnormal": 0, "xlnone": -4142, "xlwhole": 1, "xlpart": 2, "xlvalues": -4163,
    "xlformulas": -4123, "xlbyrows": 1, "xlbycolumns": 2, "xlnext": 1, "xlprevious": 2, "xlcelltypelastcell": 11,
}


def _instr(args):
    if len(args) >= 3 and is_num(args[0]) or (len(args) >= 3 and isinstance(args[0], bool)):
        start, s1, s2 = to_long(args[0]), args[1], args[2]
        mode = to_long(args[3]) if len(args) > 3 and args[3] is not MISSING else 0
    else:
        start, s1, s2, mode = 1, args[0], args[1], 0
    if s1 is NULL or s2 is NULL:
        return NULL
    a, b = to_str(s1), to_str(s2)
    if start < 1:
        raise VbaError(5, "Invalid procedure call or argument")
    if mode == 1:
        a, b = a.lower(), b.lower()
    if b == "":
        return start if start <= len(a) + 1 else 0
    return a.find(b, start - 1) + 1


def _instrrev(args):
    a, b = to_str(args[0]), to_str(args[1])
    start = to_long(args[2]) if len(args) > 2 and args[2] is not MISSING else -1
    hay = a if start == -1 else a[:start]
    if b == "":
        return len(hay) if start == -1 else start
    return hay.rfind(b) + 1


def _mid(args):
    s = to_str(args[0])
    start = to_long(args[1])
    if start < 1:
        raise VbaError(5, "Invalid procedure call or argument")
    if len(args) > 2 and args[2] is not MISSING:
        n = to_long(args[2])
        if n < 0:
            raise VbaError(5, "Invalid procedure call or argument")
        return s[start - 1:start - 1 + n]
    return s[start - 1:]


def _ascw(args):
    s = to_str(args[0])
    if s == "":
        raise VbaError(5, "Invalid procedure call or argument")
    code = ord(s[0])
    return code - 65536 if code > 32767 else code


def _val(args):
    s = "".join(ch for ch in to_str(args[0]) if ch not in " \t\r\n")
    if s[:2].lower() == "&h":
        m = re.match(r"[0-9A-Fa-f]+", s[2:])
        return float(int(m.group(), 16)) if m else 0.0
    m = re.match(r"[+-]?(\d+\.?\d*|\.\d+)([eEdD][+-]?\d+)?", s)
    if not m:
        return 0.0
    return float(m.group().replace("d", "e").replace("D", "e"))


def _split(args):
    s = to_str(args[0])
    sep = to_str(args[1]) if len(args) > 1 and args[1] is not MISSING else " "
    if s == "":
        return VArray([(0, -1)])
    return VArray.from_list(s.split(sep) if sep else [s])


def _join(args):
    arr = args[0]
    sep = to_str(args[1]) if len(args) > 1 and args[1] is not MISSING else " "
    return sep.join(to_str(v) for v in arr.items())


def _replace(args):
    s, find, rep = to_str(args[0]), to_str(args[1]), to_str(args[2])
    if find == "":
        return s
    return s.replace(find, rep)


def _strcomp(args):
    if args[0] is NULL or args[1] is NULL:
        return NULL
    a, b = to_str(args[0]), to_str(args[1])
    mode = to_long(args[2]) if len(args) > 2 and args[2] is not MISSING else 0
    if mode == 1:
        a, b = a.lower(), b.lower()
    return (a > b) - (a < b)


def _bounds(args, upper):
    arr = args[0]
    if not isinstance(arr, VArray):
        raise VbaError(13, "Type mismatch")
    d = to_long(args[1]) if len(args) > 1 and args[1] is not MISSING else 1
    if not 1 <= d <= len(arr.bounds):
        raise VbaError(9, "Subscript out of range")
    return arr.bounds[d - 1][1 if upper else 0]


def _isnumeric(args):
    v = args[0]
    if isinstance(v, bool) or is_num(v):
        return True
    if isinstance(v, str):
        return _parse_number_text(v) is not None
    return False


def _date_parts(v):
    return to_date(v).to_datetime()


def _dateserial(args):
    y, m, d = to_long(args[0]), to_long(args[1]), to_long(args[2])
    y += (m - 1) // 12
    m = (m - 1) % 12 + 1
    base = dt.date(y, m, 1)
    return VDate.of(base + dt.timedelta(days=d - 1))


def _round(args):
    x = to_double(args[0])
    n = to_long(args[1]) if len(args) > 1 and args[1] is not MISSING else 0
    return round(x, n) if n else float(round(x))


def _format(args):
    v, fmt = args[0], (to_str(args[1]) if len(args) > 1 and args[1] is not MISSING else "")
    if fmt == "0" and (is_num(v) or isinstance(v, bool)):
        x = to_double(v)
        return str(int(math.floor(abs(x) + 0.5)) * (-1 if x < 0 else 1))
    if fmt.lower() in ("yyyy-mm-dd",) and isinstance(v, VDate):
        return v.to_datetime().strftime("%Y-%m-%d")
    if fmt == "" or fmt.lower() == "general":
        return to_str(v)
    raise VbaUnsupported(f"Format with {fmt!r}")


def _str(args):
    v = args[0]
    d = to_double(v) if not isinstance(v, str) else to_double(v)
    s = num_text(d if not (is_num(v) and isinstance(v, int)) else v)
    if s.startswith("0."):
        s = s[1:]            # Str(0.1) is " .1": VBA leaves out the 0 before the point
    elif s.startswith("-0."):
        s = "-" + s[2:]
    return s if s.startswith("-") else " " + s


def _like(s, pattern):
    regex, i = "", 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == "*":
            regex += ".*"
        elif ch == "?":
            regex += "."
        elif ch == "#":
            regex += r"\d"
        elif ch == "[":
            j = pattern.index("]", i)
            body = pattern[i + 1:j]
            regex += "[^" + re.escape(body[1:]) + "]" if body.startswith("!") else "[" + re.escape(body).replace(r"\-", "-") + "]"
            i = j
        else:
            regex += re.escape(ch)
        i += 1
    return re.fullmatch(regex, s, re.S) is not None


# ======================================================================= the simulator

class Box:
    __slots__ = ("value", "type")

    def __init__(self, value, typ="Variant"):
        self.value, self.type = value, typ


class Frame:
    def __init__(self, proc, module):
        self.proc, self.module = proc, module
        self.vars = {}
        self.hidden = {}
        self.erl = 0
        self.on_error = None      # None, "next", or a label
        self.in_handler = False
        self.err_at = None
        self.ret = None


class Simulator:
    """Runs procedures of one or more modules, as one VBA project in one workbook."""

    def __init__(self, modules, today=None, book=None, max_steps=20_000_000):
        if isinstance(modules, str):
            modules = [modules]
        self.modules = [m if isinstance(m, Module) else Module(m, f"Module{i + 1}") for i, m in enumerate(modules)]
        self.today = today or dt.date.today()
        self.this_workbook = book or Workbook("Book1.xlsm")
        if not self.this_workbook.sheets:
            self.this_workbook.add_sheet("Sheet1")
        self.err = ErrObject()
        self.app = Application()
        self.debug, self.messages = [], []
        self.steps, self.max_steps = 0, max_steps
        self.globals = {}
        self.procs = {}
        for mod in self.modules:
            for key, proc in mod.procs.items():
                if key in self.procs:
                    raise VbaUnsupported(f"two procedures named {proc.name}")
                self.procs[key] = (proc, mod)
        for mod in self.modules:
            for name, expr in mod.consts:
                self.globals[name.lower()] = Box(self.eval(expr, None), "Variant")
            for name, typ, bounds, is_new in mod.vars:
                self.globals[name.lower()] = Box(self._new_value(typ, bounds, is_new, None), typ)

    # ---------------------------------------------------------------- calling
    def call(self, name, *args):
        proc, mod = self.procs[name.lower()]
        boxes = [Box(a) for a in args]
        return self._invoke(proc, mod, [(None, b) for b in boxes])

    def _invoke(self, proc, mod, args):
        """args: (name or None, Box or value); a Box is passed ByRef when the parameter is ByRef."""
        f = Frame(proc, mod)
        named = {n.lower(): v for n, v in args if n}
        positional = [v for n, v in args if not n]
        for k, prm in enumerate(proc.params):
            a = positional[k] if k < len(positional) else named.get(prm["name"].lower(), MISSING)
            if a is MISSING or (isinstance(a, Box) and a.value is MISSING):
                if not prm["optional"]:
                    raise VbaError(449, f"Argument not optional ({proc.name}.{prm['name']})")
                v = self.eval(prm["default"], f) if prm["default"] is not None else (
                    MISSING if prm["type"].lower() == "variant" else default_of(prm["type"]))
                f.vars[prm["name"].lower()] = Box(v, prm["type"])
                continue
            if isinstance(a, Box) and not prm["byval"] and not prm["array"]:
                f.vars[prm["name"].lower()] = a     # ByRef: the caller's own variable
            else:
                v = a.value if isinstance(a, Box) else a
                f.vars[prm["name"].lower()] = Box(self._assign_value(prm["type"], v), prm["type"])
        if len(positional) > len(proc.params):
            raise VbaError(450, f"Wrong number of arguments ({proc.name})")
        if proc.kind == "function":
            f.ret = Box(default_of(proc.ret_type), proc.ret_type or "Variant")
        self._run(f)
        return f.ret.value if f.ret is not None else EMPTY

    def _assign_value(self, typ, v):
        if (typ or "").lower() in OBJECT_TYPES:
            return v
        if isinstance(v, VbaObject) and (typ or "variant").lower() == "variant":
            return v
        return coerce(typ, v)

    def _new_value(self, typ, bounds, is_new, frame):
        if bounds is not None:
            if not bounds:
                return VArray([(0, -1)])
            dims = [(to_long(self.eval(lo, frame)), to_long(self.eval(hi, frame))) for lo, hi in bounds]
            d = default_of(typ)
            return VArray(dims, d)
        if is_new:
            return self._new_object(typ)
        return default_of(typ)

    def _new_object(self, typ):
        t = typ.lower()
        if t == "collection":
            return Collection()
        if t == "dictionary" or t == "scripting.dictionary":
            return Dictionary()
        raise VbaUnsupported(f"New {typ}")

    # ---------------------------------------------------------------- running a procedure
    def _run(self, f):
        code = f.proc.code
        pc = 0
        while pc < len(code):
            self.steps += 1
            if self.steps > self.max_steps:
                raise VbaUnsupported("the run took too many steps (a loop that never ends?)")
            ins = code[pc]
            try:
                nxt = self._step(f, ins, pc)
            except _Exit:
                break
            except VbaError as e:
                nxt = self._on_error(f, e, pc)
            pc = pc + 1 if nxt is None else nxt
        if f.in_handler:   # leaving a procedure from its error handler resets Err
            self.err.number, self.err.description, self.err.source = 0, "", ""

    def _on_error(self, f, e, pc):
        self.err.number, self.err.description, self.err.source = e.number, e.description, e.source
        if f.in_handler or f.on_error is None:
            raise e
        if f.on_error == "next":
            return pc + 1
        f.in_handler = True
        f.err_at = pc
        return f.proc.labels[f.on_error.lower()]

    def _step(self, f, ins, pc):
        op = ins[0]
        if op == "line":
            f.erl = ins[1]
        elif op == "let":
            self.assign(ins[1], self.eval(ins[2], f), f, is_set=False)
        elif op == "set":
            self.assign(ins[1], self.eval(ins[2], f), f, is_set=True)
        elif op == "call":
            self.eval_call(ins[1], f, statement=True)
        elif op == "jf":
            if not to_bool(self._cond(ins[1], f)):
                return ins[2]
        elif op == "jt":
            if to_bool(self._cond(ins[1], f)):
                return ins[2]
        elif op in ("goto_idx",):
            return ins[1]
        elif op == "goto":
            return f.proc.labels[ins[1].lower()]
        elif op == "for_init":
            var = ins[1]
            start = self.eval(ins[2], f)
            f.hidden[var.lower()] = (self.eval(ins[3], f), self.eval(ins[4], f))
            self.assign(Node("name", var), start, f, is_set=False)
        elif op == "for_test":
            stop, step = f.hidden[ins[1].lower()]
            cur = self.eval(Node("name", ins[1]), f)
            s = to_double(step)
            if (s >= 0 and to_double(cur) > to_double(stop)) or (s < 0 and to_double(cur) < to_double(stop)):
                return ins[2]
        elif op == "for_step":
            stop, step = f.hidden[ins[1].lower()]
            cur = self.eval(Node("name", ins[1]), f)
            self.assign(Node("name", ins[1]), self.arith("+", cur, step), f, is_set=False)
            return ins[2]
        elif op == "foreach_init":
            coll = self.eval(ins[2], f)
            if isinstance(coll, VArray):
                items = coll.items()
            elif isinstance(coll, VbaObject):
                items = coll.iterate(self)
            else:
                raise VbaError(424, "Object required")
            f.hidden["$each" + ins[1].lower()] = [items, 0]
        elif op == "foreach_next":
            state = f.hidden["$each" + ins[1].lower()]
            items, i = state
            if i >= len(items):
                return ins[2]
            state[1] = i + 1
            v = items[i]
            self.assign(Node("name", ins[1]), v, f, is_set=isinstance(v, VbaObject))
        elif op == "set_hidden":
            f.hidden[ins[1]] = self.eval(ins[2], f)
        elif op == "case_jf":
            sel = f.hidden[ins[1]]
            for t in ins[2]:
                if t[0] == "eq" and to_bool(self.compare("=", sel, self.eval(t[1], f))):
                    return None
                if t[0] == "to" and to_bool(self.compare(">=", sel, self.eval(t[1], f))) and to_bool(
                        self.compare("<=", sel, self.eval(t[2], f))):
                    return None
                if t[0] == "is" and to_bool(self.compare(t[1], sel, self.eval(t[2], f))):
                    return None
            return ins[3]
        elif op == "return":
            raise _Exit()
        elif op == "onerror_goto":
            f.on_error = ins[1]
            if not f.in_handler:
                self.err.number, self.err.description = 0, ""
        elif op == "onerror_next":
            f.on_error = "next"
            f.in_handler = False
            self.err.number, self.err.description = 0, ""
        elif op == "onerror_off":
            f.on_error = None
            self.err.number, self.err.description = 0, ""
        elif op in ("resume_label", "resume_next", "resume_same"):
            if not f.in_handler:
                raise VbaError(20, "Resume without error")
            f.in_handler = False
            self.err.number, self.err.description, self.err.source = 0, "", ""
            if op == "resume_label":
                return f.proc.labels[ins[1].lower()]
            return f.err_at + 1 if op == "resume_next" else f.err_at
        elif op == "dim":
            name, typ, bounds, is_new = ins[1:]
            if name.lower() not in f.vars:   # a Dim is not run again when a loop passes it
                f.vars[name.lower()] = Box(self._new_value(typ, bounds, is_new, f), typ)
        elif op == "redim":
            name, bounds, preserve = ins[1:]
            box = self._lookup_box(name, f)
            dims = [(to_long(self.eval(lo, f)), to_long(self.eval(hi, f))) for lo, hi in bounds]
            old = box.value
            new = VArray(dims, default_of(box.type) if box.type.lower() != "variant" else EMPTY)
            if preserve and isinstance(old, VArray):
                for k in range(min(len(old.data), len(new.data))):
                    new.data[k] = old.data[k]
            box.value = new
        elif op == "file":
            self.file_statement(f, ins[1], ins[2])
        elif op == "const":
            f.vars[ins[1].lower()] = Box(self.eval(ins[2], f), "Variant")
        else:
            raise VbaUnsupported(f"instruction {op}")
        return None

    def file_statement(self, f, low, toks):
        raise VbaUnsupported(f"file statement {low.title()} needs the file model")

    def _cond(self, node, f):
        v = self.eval(node, f)
        if v is NULL:
            return False
        return v

    # ---------------------------------------------------------------- names
    def _lookup_box(self, name, f):
        low = name.lower()
        if f is not None:
            if low in f.vars:
                return f.vars[low]
            if f.ret is not None and low == f.proc.name.lower():
                return f.ret
        if low in self.globals:
            return self.globals[low]
        return None

    def assign(self, target, value, f, is_set):
        if is_set and not isinstance(value, VbaObject) and value is not NOTHING:
            raise VbaError(424, "Object required")
        if target.kind == "name":
            box = self._lookup_box(target.a, f)
            if box is None:
                raise VbaUnsupported(f"variable {target.a} is not declared (Option Explicit)")
            if is_set:
                box.value = value
            elif isinstance(box.value, VbaObject) and box.type.lower() not in ("variant",):
                box.value.let(self, None, [], value)
            else:
                if isinstance(value, VbaObject):
                    value = value.get(self, None, [], {})
                box.value = self._assign_value(box.type, value)
            return
        if target.kind == "call":
            callee, args = target.a, target.b
            vals = [self.eval(e, f) for n, e in args]
            if callee.kind == "name":
                box = self._lookup_box(callee.a, f)
                if box is not None and isinstance(box.value, VArray):
                    elem_type = box.type if box.type.lower() != "variant" else "Variant"
                    box.value.set(vals, self._assign_value(elem_type, value) if not is_set else value)
                    return
                if box is not None and isinstance(box.value, VbaObject):
                    box.value.let(self, None, vals, value)
                    return
                raise VbaUnsupported(f"assignment to {callee.a}(...)")
            if callee.kind == "member":
                obj = self.eval(callee.a, f)
                self._obj(obj).let(self, callee.b, vals, value)
                return
        if target.kind == "member":
            obj = self.eval(target.a, f)
            self._obj(obj).let(self, target.b, [], value)
            return
        raise VbaUnsupported(f"assignment to {target!r}")

    def _obj(self, v):
        if v is NOTHING or v is EMPTY:
            raise VbaError(91, "Object variable or With block variable not set")
        if not isinstance(v, VbaObject):
            raise VbaError(424, "Object required")
        return v

    # ---------------------------------------------------------------- expressions
    def eval(self, node, f):
        k = node.kind
        if k == "lit":
            return node.a
        if k == "name":
            return self.eval_name(node.a, f)
        if k == "paren":
            return self.eval(node.a, f)
        if k == "bin":
            op = node.a
            if op == "is":
                a, b = self.eval(node.b, f), self.eval(node.c, f)
                return a is b or (a == b and isinstance(a, Range))
            a, b = self.eval(node.b, f), self.eval(node.c, f)
            if op in ("and", "or", "xor", "eqv", "imp"):
                return self.logic(op, a, b)
            if op in ("=", "<>", "<", ">", "<=", ">=", "like"):
                return self.compare(op, a, b)
            return self.arith(op, a, b)
        if k == "not":
            v = self.eval(node.a, f)
            if v is NULL:
                return NULL
            if isinstance(v, bool):
                return not v
            return ~to_long(v)
        if k == "neg":
            v = self.eval(node.a, f)
            if v is EMPTY:
                return 0
            if isinstance(v, VDate):
                return VDate(-v.serial)
            if isinstance(v, bool):
                return 1 if v else 0
            if isinstance(v, int):
                return -v
            return -to_double(v)
        if k in ("call", "member"):
            return self.eval_call(node, f)
        if k == "new":
            return self._new_object(node.a)
        if k == "missing":
            return MISSING
        raise VbaUnsupported(f"expression {k}")

    def eval_name(self, name, f):
        box = self._lookup_box(name, f)
        if box is not None:
            v = box.value
            return v
        low = name.lower()
        if low in CONSTANTS:
            return CONSTANTS[low]
        if low == "erl":
            return f.erl if f is not None else 0
        if low in self.procs:
            return self._call_proc(low, [], f)
        if low == "err":
            return self.err
        if low == "application":
            return self.app
        if low == "thisworkbook" or low == "activeworkbook":
            return self.this_workbook
        if low == "debug":
            return Debug()
        if low in ("worksheets", "sheets"):
            return SheetsCollection(self.this_workbook)
        return self.builtin(low, [], {}, name)

    def _args(self, args, f, byref=False):
        out = []
        for n, e in args:
            if byref and e.kind == "name":
                box = self._lookup_box(e.a, f)
                if box is not None and not isinstance(box.value, VbaObject):
                    out.append((n, box))
                    continue
            out.append((n, self.eval(e, f)))
        return out

    def _call_proc(self, low, args, f):
        proc, mod = self.procs[low]
        return self._invoke(proc, mod, args)

    def eval_call(self, node, f, statement=False):
        if node.kind == "member":
            obj = self.eval(node.a, f)
            o = self._obj(obj)
            return o.call(self, node.b, [], {}) if statement else o.get(self, node.b, [], {})
        callee, args = node.a, node.b
        if callee.kind == "name":
            low = callee.a.lower()
            box = self._lookup_box(callee.a, f)
            if box is not None and not (f is not None and f.ret is box and args):
                vals = [self.eval(e, f) for n, e in args]
                v = box.value
                if isinstance(v, VArray):
                    return v.get(vals)
                if isinstance(v, VbaObject):
                    return v.get(self, None, vals, {})
                if not args:
                    return v
                raise VbaError(13, f"{callee.a} is not an array")
            if low in self.procs:
                return self._call_proc(low, self._args(args, f, byref=True), f)
            if low in ("msgbox",):
                vals = [self.eval(e, f) for n, e in args]
                self.messages.append(to_str(vals[0]))
                return 1
            vals, named = [], {}
            for n, e in args:
                if n:
                    named[n.lower()] = self.eval(e, f)
                else:
                    vals.append(self.eval(e, f))
            if low in ("createobject",):
                return self._create(to_str(vals[0]))
            return self.builtin(low, vals, named, callee.a)
        if callee.kind == "member":
            obj = self._obj(self.eval(callee.a, f))
            vals, named = [], {}
            for n, e in args:
                if n:
                    named[n.lower()] = self.eval(e, f)
                else:
                    vals.append(self.eval(e, f))
            if statement or node.c == "bare":
                return obj.call(self, callee.b, vals, named)
            return obj.get(self, callee.b, vals, named)
        if callee.kind == "call":   # a(1)(2), or Cells(1, 1).Value through a default member
            inner = self.eval_call(callee, f)
            vals = [self.eval(e, f) for n, e in args]
            if isinstance(inner, VArray):
                return inner.get(vals)
            return self._obj(inner).get(self, None, vals, {})
        raise VbaUnsupported(f"call of {callee.kind}")

    def _create(self, progid):
        p = progid.lower()
        if p == "scripting.dictionary":
            return Dictionary()
        if p == "scripting.filesystemobject":
            if getattr(self, "fso", None) is None:
                raise VbaUnsupported("Scripting.FileSystemObject needs a mock file system")
            return self.fso
        raise VbaUnsupported(f"CreateObject({progid!r})")

    # ---------------------------------------------------------------- operators
    def logic(self, op, a, b):
        if a is NULL or b is NULL:
            if op == "and" and (a is False or b is False):
                return False
            if op == "or" and (a is True or b is True):
                return True
            return NULL
        if isinstance(a, bool) and isinstance(b, bool):
            return {"and": a and b, "or": a or b, "xor": a != b, "eqv": a == b, "imp": (not a) or b}[op]
        x, y = to_long(a), to_long(b)
        return {"and": x & y, "or": x | y, "xor": x ^ y, "eqv": ~(x ^ y), "imp": (~x) | y}[op]

    def compare(self, op, a, b):
        if op == "like":
            if a is NULL or b is NULL:
                return NULL
            return _like(to_str(a), to_str(b))
        if a is NULL or b is NULL:
            return NULL
        if isinstance(a, VbaObject) or isinstance(b, VbaObject):
            raise VbaError(13, "Type mismatch")
        if isinstance(a, VError) or isinstance(b, VError):
            if isinstance(a, VError) and isinstance(b, VError):
                c = (a.code > b.code) - (a.code < b.code)
            else:
                raise VbaError(13, "Type mismatch")
        elif isinstance(a, str) and isinstance(b, str):
            c = (a > b) - (a < b)
        elif isinstance(a, str) or isinstance(b, str):
            s, other = (a, b) if isinstance(a, str) else (b, a)
            if other is EMPTY:
                x, y = (a, "") if isinstance(a, str) else ("", b)
                c = (x > y) - (x < y)
            else:
                n = _parse_number_text(s)
                if n is None:
                    # a number against text that is not one: VBA's Variant rule, the number is less
                    c = 1 if isinstance(a, str) else -1
                else:
                    x, y = (float(n), to_double(b)) if isinstance(a, str) else (to_double(a), float(n))
                    c = (x > y) - (x < y)
        else:
            x, y = to_double(a), to_double(b)
            c = (x > y) - (x < y)
        return {"=": c == 0, "<>": c != 0, "<": c < 0, ">": c > 0, "<=": c <= 0, ">=": c >= 0}[op]

    def arith(self, op, a, b):
        if a is NULL or b is NULL:
            return NULL
        if op == "&":
            return (to_str(a) if a is not NULL else "") + (to_str(b) if b is not NULL else "")
        if op == "+":
            if isinstance(a, str) and isinstance(b, str):
                return a + b
            if (isinstance(a, str) and b is EMPTY) or (isinstance(b, str) and a is EMPTY):
                return a if isinstance(a, str) else b
        if isinstance(a, VError) or isinstance(b, VError):
            raise VbaError(13, "Type mismatch")
        date_a, date_b = isinstance(a, VDate), isinstance(b, VDate)
        if op in ("+", "-", "*", "^") and _is_int(a) and _is_int(b):
            x, y = _int_of(a), _int_of(b)
            if op == "+":
                return x + y
            if op == "-":
                return x - y
            if op == "*":
                return x * y
            return float(x) ** y
        x, y = to_double(a), to_double(b)
        if op == "+":
            r = x + y
            return VDate(r) if (date_a or date_b) else r
        if op == "-":
            r = x - y
            return VDate(r) if (date_a and not date_b) else r
        if op == "*":
            return x * y
        if op == "/":
            if y == 0:
                raise VbaError(11, "Division by zero")
            return x / y
        if op == "\\":
            yi = _round_even(y)
            if yi == 0:
                raise VbaError(11, "Division by zero")
            xi = _round_even(x)
            q = abs(xi) // abs(yi)
            return q if (xi >= 0) == (yi >= 0) else -q
        if op == "mod":
            yi = _round_even(y)
            if yi == 0:
                raise VbaError(11, "Division by zero")
            xi = _round_even(x)
            r = abs(xi) % abs(yi)
            return r if xi >= 0 else -r
        if op == "^":
            return x ** y
        raise VbaUnsupported(f"operator {op}")

    # ---------------------------------------------------------------- built-in functions
    def builtin(self, low, a, named, shown):
        g = lambda i: a[i] if i < len(a) else MISSING   # noqa: E731
        if low in ("len",):
            return len(to_str(a[0])) if a[0] is not NULL else NULL
        if low == "mid":
            return _mid(a)
        if low == "left":
            return to_str(a[0])[:max(0, to_long(a[1]))]
        if low == "right":
            n = to_long(a[1])
            s = to_str(a[0])
            return s[len(s) - n:] if n > 0 else ""
        if low in ("lcase",):
            return to_str(a[0]).lower() if a[0] is not NULL else NULL
        if low in ("ucase",):
            return to_str(a[0]).upper() if a[0] is not NULL else NULL
        if low == "trim":
            return to_str(a[0]).strip(" ")
        if low == "ltrim":
            return to_str(a[0]).lstrip(" ")
        if low == "rtrim":
            return to_str(a[0]).rstrip(" ")
        if low == "instr":
            return _instr(a)
        if low == "instrrev":
            return _instrrev(a)
        if low == "replace":
            return _replace(a)
        if low == "split":
            return _split(a)
        if low == "join":
            return _join(a)
        if low == "space":
            return " " * to_long(a[0])
        if low == "string":
            return to_str(a[1])[:1] * to_long(a[0]) if isinstance(a[1], str) else chr(to_long(a[1])) * to_long(a[0])
        if low in ("ascw", "asc"):
            return _ascw(a)
        if low in ("chrw", "chr"):
            n = to_long(a[0])
            return chr(n + 65536 if n < 0 else n)
        if low == "strcomp":
            return _strcomp(a)
        if low == "val":
            return _val(a)
        if low == "cdbl":
            return to_double(a[0])
        if low in ("clng", "cint"):
            return to_long(a[0])
        if low == "cstr":
            return to_str(a[0])
        if low == "cbool":
            return to_bool(a[0])
        if low == "cdate":
            return to_date(a[0])
        if low == "cverr":
            return VError(to_long(a[0]))
        if low == "cvar":
            return a[0]
        if low == "isnumeric":
            return _isnumeric(a)
        if low == "isempty":
            return a[0] is EMPTY
        if low == "isnull":
            return a[0] is NULL
        if low == "iserror":
            return isinstance(a[0], VError)
        if low == "isarray":
            return isinstance(a[0], VArray)
        if low == "isobject":
            return isinstance(a[0], VbaObject) or a[0] is NOTHING
        if low == "ismissing":
            return a[0] is MISSING
        if low == "isdate":
            if isinstance(a[0], VDate):
                return True
            try:
                to_date(a[0]) if isinstance(a[0], str) else (_ for _ in ()).throw(VbaError(13))
                return True
            except VbaError:
                return False
        if low == "vartype":
            return vartype(a[0])
        if low == "typename":
            return typename(a[0])
        if low == "ubound":
            return _bounds(a, True)
        if low == "lbound":
            return _bounds(a, False)
        if low == "array":
            return VArray.from_list(list(a))
        if low == "iif":
            return a[1] if to_bool(a[0]) else a[2]
        if low == "abs":
            v = a[0]
            return abs(v) if is_num(v) else abs(to_double(v))
        if low == "int":
            v = a[0]
            if isinstance(v, VDate):
                return VDate(math.floor(v.serial))
            return math.floor(v) if isinstance(v, int) else float(math.floor(to_double(v)))
        if low == "fix":
            v = a[0]
            return int(v) if isinstance(v, int) else float(math.trunc(to_double(v)))
        if low == "sgn":
            d = to_double(a[0])
            return (d > 0) - (d < 0)
        if low == "round":
            return _round(a)
        if low == "sqr":
            return math.sqrt(to_double(a[0]))
        if low == "year":
            return _date_parts(a[0]).year
        if low == "month":
            return _date_parts(a[0]).month
        if low == "day":
            return _date_parts(a[0]).day
        if low == "hour":
            return _date_parts(a[0]).hour
        if low == "minute":
            return _date_parts(a[0]).minute
        if low == "second":
            return _date_parts(a[0]).second
        if low == "dateserial":
            return _dateserial(a)
        if low == "rgb":
            return to_long(a[0]) + to_long(a[1]) * 256 + to_long(a[2]) * 65536
        if low == "timeserial":
            return VDate((to_long(a[0]) * 3600 + to_long(a[1]) * 60 + to_long(a[2])) / 86400.0)
        if low == "date":
            return VDate.of(self.today)
        if low == "now":
            return VDate.of(self.today)
        if low == "timer":
            return 0.0
        if low == "format":
            return _format(a)
        if low == "str":
            return _str(a)
        if low == "doevents":
            return 0
        raise VbaUnsupported(f"function or name {shown}")


def _is_int(v):
    return isinstance(v, (int, bool)) and not isinstance(v, float) or v is EMPTY


def _int_of(v):
    if v is EMPTY:
        return 0
    if isinstance(v, bool):
        return -1 if v else 0
    return v
