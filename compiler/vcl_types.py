"""Prototype type checker for VCL 0.1 to 0.4 with flow-sensitive narrowing (milestone M2).

Types: Num, Text, Date, Bool, Cell (a raw cell value: empty, number, text, date or error),
Sheet, Book, Range, Dict, Coll, Arr, Iter.

Key rules
- A raw Cell can't be used in arithmetic or in < > <= >=, and can't be compared with = to a
  number. It must first be narrowed by IS_NUMBER or IS_DATE in an enclosing IF (or the
  condition of a one-line IF, or an earlier EXIT_IF).
- Both sides of AND / OR are evaluated, so a guard on the left does NOT protect the right.
  The same holds for both values of IIF: its condition never guards them.
- Facts are keyed by the exact text of an expression, for example CELL($bom,r,bomReq). A write
  to that exact cell, or a whole-sheet change (CLEAR of a range, COPY_ROW, SORT, DELETE_ROW),
  removes the facts it could affect. Assigning a variable removes facts that mention it.
- Paths (VCL 0.3, T4) are checked by their shape, not by flow: a path is a String $argument (a
  root) or PATH_JOIN($root, part, ...), where every part is SAFE_NAME(...) or a quoted safe name,
  written in place. So no path can climb out of its root, whatever the data holds.
"""
import sys
from pathlib import Path

from lark import Token, Tree

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "vcl"))
from semantics import safe_name  # noqa: E402
from vcl_parse import parse  # noqa: E402

VBA_TO_TYPE = {"Long": "Num", "Double": "Num", "String": "Text", "Date": "Date",
               "Boolean": "Bool", "Variant": "Variant", "Range": "Range", "Worksheet": "Sheet",
               "Workbook": "Book", "Dictionary": "Dict", "Collection": "Coll"}
SCALAR = {"Num", "Text", "Date", "Bool", "Cell", "Name", "Path"}
TEXTLIKE = {"Text", "Cell", "Num", "Name", "Path"}
STRINGS = {"Text", "Name", "Path"}   # a Name or a Path is text everywhere text is read
CONSTANTS = {"TODAY": "Date", "NOW": "Date", "TRUE": "Bool", "FALSE": "Bool",
             "THIS_BOOK": "Book", "ACTIVE_BOOK": "Book", "ACTIVE_SHEET": "Sheet",
             "NOTHING": "Range"}
# value functions: name -> (parameter types, return type); "S" = any scalar, "T" = text-like,
# "Num!" = a number and not a date; IIF's return type is worked out from its two values
FUNCS = {
    "LASTROW": (["Sheet", "Num"], "Num"), "LASTCOL": (["Sheet", "Num"], "Num"),
    "USED_RANGE": (["Sheet"], "Range"), "COL": (["Sheet", "Text"], "Num"),
    "CELL": (["Sheet", "Num", "Num"], "Cell"), "ROW_RANGE": (["Sheet", "Num", "Num", "Num"], "Range"),
    "RANGE": (["Sheet", "Text"], "Range"), "SHEET": (["Book", "Text"], "Sheet"),
    "FIND": (["Range", "S"], "Range"), "ROWS": (["Arr"], "Num"), "COLS": (["Arr"], "Num"),
    "ARR": (["Arr", "Num", "Num"], "Cell"), "COUNT": (["Dict"], "Num"),
    "EXISTS": (["Dict", "S"], "Bool"), "LEN": (["T"], "Num"), "TRIM": (["T"], "Text"),
    "UPPER": (["T"], "Text"), "LOWER": (["T"], "Text"), "LEFT": (["T", "Num"], "Text"),
    "RIGHT": (["T", "Num"], "Text"), "IS_EMPTY": (["S"], "Bool"), "IS_NUMBER": (["S"], "Bool"),
    "IS_DATE": (["S"], "Bool"), "IS_NOTHING": (["Range"], "Bool"), "FILE_NAME": (["Text"], "Text"),
    "KEYS": (["Dict"], "Iter"), "ITEMS": (["Coll"], "Iter"), "SHEETS": (["Book"], "Iter"),
    "CELLS": (["Range"], "Iter"),
    # VCL 0.2
    "SAME": (["S", "S"], "Bool"), "CLEAN_TEXT": (["T"], "Text"), "IS_ERROR": (["S"], "Bool"),
    "PARSE_DATE": (["S", "Text"], "Cell"), "IIF": (["Bool", "S", "S"], None),
    "MOD": (["Num!", "Num!"], "Num"),
    # VCL 0.3; PATH_JOIN and FOLDER_EXISTS are checked by shape in ty_call (T4)
    "SAFE_NAME": (["T"], "Name"), "PATH_JOIN": (None, "Path"), "FOLDER_EXISTS": (None, "Bool"),
    # VCL 0.4
    "SHEET_EXISTS": (["Book", "T"], "Bool"),
}
FILE_OPS = {"MAKE_FOLDER", "COPY_FILE"}   # VCL 0.3: every path argument is checked by shape (T4)


def tok(t):
    return isinstance(t, Token)


def canon(t):
    """Exact, whitespace-free text of an expression, used as the key for facts."""
    if tok(t):
        return str(t)
    if t.data == "call":
        args = t.children[1].children if len(t.children) > 1 else []
        return f"{t.children[0]}(" + ",".join(canon(a) for a in args) + ")"
    if t.data == "option":
        return f"{t.children[0]}={canon(t.children[1])}"
    if t.data == "not_op":
        return "NOT " + canon(t.children[0])
    if t.data == "neg":
        return "-" + canon(t.children[0])
    if t.data in ("and_expr", "or_expr", "concat"):
        sep = {"and_expr": " AND ", "or_expr": " OR ", "concat": "&"}[t.data]
        return "(" + sep.join(canon(c) for c in t.children) + ")"
    return "(" + "".join(canon(c) for c in t.children) + ")"


def mentions(text, name):
    import re
    return re.search(rf"(?<![\w$]){re.escape(name)}(?!\w)", text) is not None


class Checker:
    def __init__(self, spec):
        self.problems = []
        self.vars = {}       # name -> declared type
        self.args = {}       # name -> type
        self.dict_values = {}  # dict var -> set of value types
        self.ret = None
        for k in spec["keys"]:
            t = k["tree"]
            if t is None:
                continue
            if k["key"] == "ARGS":
                for p in t.children:
                    self.args[str(p.children[0])] = VBA_TO_TYPE[str(p.children[1])]
            elif k["key"] == "VAR":
                self.vars[str(t.children[0])] = VBA_TO_TYPE[str(t.children[1])]
            elif k["key"] == "RET":
                self.ret = VBA_TO_TYPE[str(t.children[0])]
        self.loopvars = {}
        self.selectors = []   # the SELECT_CASE selectors around the current line, innermost last
        self.spec = spec

    def err(self, line, msg):
        self.problems.append((line, msg))

    # ---------- expressions ----------
    def ty(self, t, facts, vflow, line):
        key = canon(t)
        if key in facts:
            return facts[key]
        if tok(t):
            kind = t.type
            if kind == "NUMBER":
                return "Num"
            if kind == "STRING":
                return "Text"
            if kind == "COLOR":
                return "Color"
            if kind == "ARG":
                name = str(t)[1:]
                if name not in self.args:
                    self.err(line, f"unknown argument {t}")
                    return "?"
                return self.args[name]
            if kind == "NAME":
                name = str(t)
                if name in CONSTANTS:
                    return CONSTANTS[name]
                if name in self.args:
                    self.err(line, f"argument {name} must be written ${name}")
                    return self.args[name]
                if name in self.loopvars:
                    return self.loopvars[name]
                if name in self.vars:
                    declared = self.vars[name]
                    return vflow.get(name, "Cell") if declared == "Variant" else declared
                return "Word"   # an option word such as up, sum, desc; the validator checks it
            return "?"
        d = t.data
        if d == "call":
            return self.ty_call(t, facts, vflow, line)
        if d == "not_op":
            self.need_bool(t.children[0], facts, vflow, line)
            return "Bool"
        if d in ("and_expr", "or_expr"):
            for c in t.children:          # no short-circuit: each side sees only outer facts
                self.need_bool(c, facts, vflow, line)
            return "Bool"
        if d == "neg":
            a = self.ty(t.children[0], facts, vflow, line)
            self.need_num(a, t.children[0], line)
            return "Num"
        if d == "concat":
            for c in t.children:
                self.ty(c, facts, vflow, line)
            return "Text"
        if d in ("sum", "product"):
            return self.ty_arith(t, facts, vflow, line)
        if d == "comparison":
            return self.ty_compare(t, facts, vflow, line)
        if d == "list_lit":
            return "List"
        return "?"

    def need_bool(self, t, facts, vflow, line):
        a = self.ty(t, facts, vflow, line)
        if a not in ("Bool", "?"):
            self.err(line, f"condition {canon(t)} is {a}, not Bool")

    def need_num(self, a, t, line):
        if a == "Cell":
            self.err(line, f"raw cell value {canon(t)} used as a number: guard it with IS_NUMBER first")
        elif a not in ("Num", "?", "Date"):
            self.err(line, f"{canon(t)} is {a}, not a number")

    def ty_arith(self, t, facts, vflow, line):
        kids = t.children
        cur = self.ty(kids[0], facts, vflow, line)
        if cur == "Cell":
            self.err(line, f"raw cell value {canon(kids[0])} used in arithmetic: guard it with IS_NUMBER or IS_DATE first")
            cur = "?"
        for i in range(1, len(kids), 2):
            op, rhs = str(kids[i]), kids[i + 1]
            r = self.ty(rhs, facts, vflow, line)
            if r == "Cell":
                self.err(line, f"raw cell value {canon(rhs)} used in arithmetic: guard it with IS_NUMBER or IS_DATE first")
                r = "?"
            if "?" in (cur, r):
                cur = "?"
            elif cur == "Date" and r == "Num" and op in "+-":
                cur = "Date"
            elif cur == "Num" and r == "Date" and op == "+":
                cur = "Date"
            elif cur == "Date" and r == "Date" and op == "-":
                cur = "Num"
            elif cur == "Num" and r == "Num":
                cur = "Num"
            else:
                self.err(line, f"can't apply {op} to {cur} and {r} in {canon(t)}")
                cur = "?"
        return cur

    def ty_compare(self, t, facts, vflow, line):
        a_t, op, b_t = t.children
        a = self.ty(a_t, facts, vflow, line)
        b = self.ty(b_t, facts, vflow, line)
        op = str(op)
        if "?" in (a, b):
            return "Bool"
        if op in ("<", ">", "<=", ">="):
            for side, x in ((a_t, a), (b_t, b)):
                if x == "Cell":
                    self.err(line, f"raw cell value {canon(side)} compared with {op}: guard it with IS_NUMBER or IS_DATE first")
                    return "Bool"
            if not ({a, b} <= {"Num"} or {a, b} <= {"Date"} or {a, b} <= STRINGS):
                self.err(line, f"can't compare {a} with {b} using {op}")
        else:  # = and <>
            if "Cell" in (a, b) and ("Num" in (a, b) or "Date" in (a, b)):
                side = a_t if a == "Cell" else b_t
                self.err(line, f"raw cell value {canon(side)} compared with a number or date: guard it first (a text cell would cause a type mismatch)")
            elif a != b and not ({a, b} <= STRINGS | {"Cell"}) and "Word" not in (a, b):
                self.err(line, f"can't compare {a} with {b}")
        return "Bool"

    def ty_call(self, t, facts, vflow, line):
        name = str(t.children[0])
        args = t.children[1].children if len(t.children) > 1 else []
        if name not in FUNCS:
            self.err(line, f"{name} is not a value function")
            return "?"
        params, ret = FUNCS[name]
        pos = [a for a in args if not (isinstance(a, Tree) and a.data == "option")]
        if params is None:   # PATH_JOIN, FOLDER_EXISTS: paths by shape
            self.check_path_call(name, pos, facts, vflow, line)
            return ret
        if len(pos) != len(params):
            self.err(line, f"{name} takes {len(params)} arguments, got {len(pos)}")
        types = []
        for p, a in zip(params, pos):
            at = self.ty(a, facts, vflow, line)
            types.append(at)
            if at == "?":
                continue
            if p in ("Num", "Num!") and at == "Cell":
                self.err(line, f"{name}: raw cell value {canon(a)} used as a number: guard it with IS_NUMBER first")
            elif p == "Num!" and at != "Num":
                self.err(line, f"{name}: {canon(a)} is {at}, expected a number")
            elif p == "Num!":
                pass
            elif p == "S" and at not in SCALAR:
                self.err(line, f"{name}: {canon(a)} is {at}, expected a value")
            elif p == "T" and at not in TEXTLIKE:
                self.err(line, f"{name}: {canon(a)} is {at}, expected text")
            elif (p not in ("S", "T") and at != p and not (p == "Num" and at == "Date")
                  and not (p == "Text" and at in STRINGS)):
                self.err(line, f"{name}: {canon(a)} is {at}, expected {p}")
        if name == "IIF":   # one type when both values have it, else a raw value that needs a guard
            values = types[1:3]
            return "?" if "?" in values or len(values) < 2 else values[0] if values[0] == values[1] else "Cell"
        return ret

    def is_root(self, t):
        """A String $argument: the only thing a path may start from."""
        return tok(t) and t.type == "ARG" and self.args.get(str(t)[1:]) == "Text"

    def path_problem(self, t):
        """None when t is a path by shape (a String $argument or PATH_JOIN(...)), else why not."""
        if self.is_root(t) or (isinstance(t, Tree) and t.data == "call" and str(t.children[0]) == "PATH_JOIN"):
            return None
        return f"{canon(t)} is not a path: write a String $argument or PATH_JOIN($root, ...)"

    def check_path_call(self, name, pos, facts, vflow, line):
        if name == "FOLDER_EXISTS":
            if len(pos) != 1:
                self.err(line, f"FOLDER_EXISTS takes 1 argument, got {len(pos)}")
                return
            problem = self.path_problem(pos[0])
            if problem:
                self.err(line, f"FOLDER_EXISTS: {problem}")
            else:
                self.ty(pos[0], facts, vflow, line)
            return
        if len(pos) < 2:   # PATH_JOIN
            self.err(line, f"PATH_JOIN takes 2 or more arguments, got {len(pos)}")
            return
        if not self.is_root(pos[0]):
            self.err(line, f"PATH_JOIN must start from a String $argument (the root), got {canon(pos[0])}")
        for part in pos[1:]:
            if isinstance(part, Tree) and part.data == "call" and str(part.children[0]) == "SAFE_NAME":
                self.ty(part, facts, vflow, line)
            elif tok(part) and part.type == "STRING":
                text = str(part)[1:-1].replace('""', '"')
                if safe_name(text) != text:
                    self.err(line, f"PATH_JOIN: {part} is not a safe name (see SAFE_NAME)")
            else:
                self.err(line, f"PATH_JOIN: {canon(part)} must be SAFE_NAME(...) or a \"quoted\" safe name, "
                               "written in place so every part is checked where it is used")

    # ---------- narrowing ----------
    def facts_if(self, t, truth):
        if isinstance(t, Tree) and t.data == "call" and str(t.children[0]) in ("IS_NUMBER", "IS_DATE"):
            if truth:
                target = canon(t.children[1].children[0])
                return {target: "Num" if str(t.children[0]) == "IS_NUMBER" else "Date"}
            return {}
        if isinstance(t, Tree) and t.data == "not_op":
            return self.facts_if(t.children[0], not truth)
        if isinstance(t, Tree) and t.data == "and_expr" and truth:
            out = {}
            for c in t.children:
                out.update(self.facts_if(c, True))
            return out
        if isinstance(t, Tree) and t.data == "or_expr" and not truth:
            out = {}
            for c in t.children:
                out.update(self.facts_if(c, False))
            return out
        return {}

    # ---------- statements ----------
    def invalidate(self, call, facts, vflow, line):
        """Remove facts that this action could make false; record variable flow types."""
        name = str(call.children[0])
        args = call.children[1].children if len(call.children) > 1 else []
        pos = [a for a in args if not (isinstance(a, Tree) and a.data == "option")]
        opts = {str(a.children[0]): a.children[1] for a in args if isinstance(a, Tree) and a.data == "option"}

        def drop_var(v):
            for k in [k for k in facts if mentions(k, v)]:
                del facts[k]

        def drop_sheet_of(rng, block=False):
            """block: the op writes a block that starts at rng, so a CELL(...) target loses the sheet."""
            text = canon(rng)
            if text.startswith("CELL(") and not block:
                facts.pop(text, None)
                return
            sheet = text[text.index("(") + 1:].split(",")[0].rstrip(")") if "(" in text else text
            for k in [k for k in facts if k.startswith("CELL(" + sheet + ",")]:
                del facts[k]

        if name == "LET":
            v = str(pos[0])
            drop_var(v)
            vt = self.ty(pos[1], facts, vflow, line)
            self.assign(v, vt, pos[1], vflow, line)
        elif name in ("DICT_GET", "READ_ARRAY") and "into" in opts:
            v = str(opts["into"])
            drop_var(v)
            if name == "READ_ARRAY":
                vflow[v] = "Arr"
            else:
                values = self.dict_values.get(str(pos[0]), {"Cell"})
                vt = "Num" if values == {"Num"} else ("Text" if values == {"Text"} else "Cell")
                self.assign(v, vt, opts["into"], vflow, line)
        elif name in FILE_OPS and "into" in opts:
            drop_var(str(opts["into"]))
        elif name in ("ADD_SHEET", "COPY_SHEET") and "into" in opts:   # the new sheet (D40, D41)
            drop_var(str(opts["into"]))
            self.assign(str(opts["into"]), "Sheet", opts["into"], vflow, line)
        elif name in ("SET_VALUE", "CLEAR"):
            drop_sheet_of(pos[0])
        elif name == "WRITE_ROW" and pos:
            for k in [k for k in facts if k.startswith("CELL(" + canon(pos[0]) + ",")]:
                del facts[k]
        elif name in ("COPY_ROW",) and len(pos) >= 3:
            drop_sheet_of(pos[2])
        elif name in ("DELETE_ROW", "INSERT_ROW", "SORT", "DEDUPE"):
            drop_sheet_of(pos[0])
        elif name == "WRITE_ARRAY" and "to" in opts:
            drop_sheet_of(opts["to"], block=True)

    def assign(self, v, vt, src, vflow, line):
        declared = self.vars.get(v)
        if declared is None:
            if v not in self.loopvars:
                self.err(line, f"{v} is not declared in VAR")
            return
        if declared == "Variant":
            vflow[v] = vt if vt in SCALAR | {"Arr"} else "Cell"
            return
        ok = (vt == declared or vt == "?" or (declared == "Text" and vt in ("Cell", "Num", "Name", "Path"))
              or (declared == "Num" and vt == "Num"))
        if vt == "Cell" and declared in ("Num", "Date"):
            self.err(line, f"assigning raw cell value {canon(src)} to {declared} variable {v}: guard it first")
        elif not ok:
            self.err(line, f"assigning {vt} to {declared} variable {v}")

    def check_action(self, call, facts, vflow, line):
        name = str(call.children[0])
        args = call.children[1].children if len(call.children) > 1 else []
        pos = [a for a in args if not (isinstance(a, Tree) and a.data == "option")]
        if name == "DICT_ADD" and len(pos) >= 3:
            opts = {str(a.children[0]): str(a.children[1]) for a in args
                    if isinstance(a, Tree) and a.data == "option"}
            vt = self.ty(pos[2], facts, vflow, line)
            if opts.get("mode") == "sum" and vt == "Cell":
                self.err(line, f"DICT_ADD mode=sum on raw cell value {canon(pos[2])}: guard it with IS_NUMBER first")
                vt = "Num"
            if opts.get("mode") in ("sum", "count"):
                vt = "Num"
            self.dict_values.setdefault(str(pos[0]), set()).add(vt if vt in SCALAR else "Cell")
            self.ty(pos[1], facts, vflow, line)
        elif name == "DEDUPE":   # cols= is a sheet column number or a list of them (D38)
            opts = {str(a.children[0]): a.children[1] for a in args if isinstance(a, Tree) and a.data == "option"}
            cols = opts.get("cols")
            items = cols.children if isinstance(cols, Tree) and cols.data == "list_lit" else [cols]
            for c in items:
                if c is not None:
                    self.need_num(self.ty(c, facts, vflow, line), c, line)
            for p in pos:
                self.ty(p, facts, vflow, line)
        elif name == "FILTER":   # FILTER(<range>, col=<sheet column>, crit=<value>) (D42)
            opts = {str(a.children[0]): a.children[1] for a in args if isinstance(a, Tree) and a.data == "option"}
            if pos and self.ty(pos[0], facts, vflow, line) not in ("Range", "?"):
                self.err(line, f"FILTER: {canon(pos[0])} is not a range")
            if "col" in opts:
                self.need_num(self.ty(opts["col"], facts, vflow, line), opts["col"], line)
            if "crit" in opts and self.ty(opts["crit"], facts, vflow, line) not in SCALAR | {"?"}:
                self.err(line, f"FILTER: crit={canon(opts['crit'])} is not a value")
        elif name in ("ADD_SHEET", "COPY_SHEET"):   # (<book> or <sheet>, name=<text>, after=<sheet>, into=<VAR>)
            opts = {str(a.children[0]): a.children[1] for a in args if isinstance(a, Tree) and a.data == "option"}
            first = "Book" if name == "ADD_SHEET" else "Sheet"
            if pos and self.ty(pos[0], facts, vflow, line) not in (first, "?"):
                self.err(line, f"{name}: {canon(pos[0])} is not a {'workbook' if first == 'Book' else 'sheet'}")
            if "name" in opts and self.ty(opts["name"], facts, vflow, line) not in TEXTLIKE | {"?"}:
                self.err(line, f"{name}: name={canon(opts['name'])} is not text")
            if "after" in opts and self.ty(opts["after"], facts, vflow, line) not in ("Sheet", "?"):
                self.err(line, f"{name}: after={canon(opts['after'])} is not a sheet")
        elif name == "WRITE_ROW" and len(pos) >= 3:
            st = self.ty(pos[0], facts, vflow, line)
            if st not in ("Sheet", "?"):
                self.err(line, f"WRITE_ROW: {canon(pos[0])} is {st}, expected Sheet")
            self.need_num(self.ty(pos[1], facts, vflow, line), pos[1], line)
            for a in pos[2:]:
                vt = self.ty(a, facts, vflow, line)
                if vt not in SCALAR and vt != "?":
                    self.err(line, f"WRITE_ROW: {canon(a)} is {vt}, expected a value")
            opts = {str(a.children[0]): a.children[1] for a in args if isinstance(a, Tree) and a.data == "option"}
            if "col" in opts:
                self.need_num(self.ty(opts["col"], facts, vflow, line), opts["col"], line)
        elif name in FILE_OPS:   # VCL 0.3: paths by shape (T4); the status goes into a String VAR
            for a in pos:
                problem = self.path_problem(a)
                if problem:
                    self.err(line, f"{name}: {problem}")
                else:
                    self.ty(a, facts, vflow, line)
            into = next((a.children[1] for a in args if isinstance(a, Tree) and a.data == "option"
                         and str(a.children[0]) == "into"), None)
            declared = self.vars.get(str(into)) if into is not None else None
            if declared not in (None, "Text"):
                self.err(line, f"{name}: into={into} is {declared}, but the status is text: declare it As String")
        elif name == "RETURN" and pos:
            rt = self.ty(pos[0], facts, vflow, line)
            if self.ret and rt not in (self.ret, "?"):
                self.err(line, f"RETURN gives {rt}, but RET is {self.ret}")
        elif name not in ("LET", "DICT_GET", "READ_ARRAY"):
            for a in pos:
                self.ty(a, facts, vflow, line)
        self.invalidate(call, facts, vflow, line)

    def block(self, nodes, facts, vflow):
        facts = dict(facts)
        for n in nodes:
            s, line, kind = n["stmt"], n["line"], n["kind"]
            if kind == "call":
                self.check_action(s, facts, vflow, line)
            elif kind == "if_inline":
                cond, act = s.children
                self.need_bool(cond, facts, vflow, line)
                inner = dict(facts)
                inner.update(self.facts_if(cond, True))
                if act.data == "call":            # THEN NOTE "..." has nothing to check
                    self.check_action(act, inner, vflow, line)
                for k in list(facts):
                    if k not in inner:
                        del facts[k]
            elif kind == "if_block":
                cond = s.children[0]
                self.need_bool(cond, facts, vflow, line)
                then_f = dict(facts); then_f.update(self.facts_if(cond, True))
                else_f = dict(facts); else_f.update(self.facts_if(cond, False))
                after_then = self.block(n["children"], then_f, vflow)
                after_else = self.block(n["else"], else_f, vflow)
                facts = {k: v for k, v in facts.items() if after_then.get(k) == v and after_else.get(k) == v}
            elif kind in ("exit_if", "skip_if", "exit_loop_if"):
                # past this line the condition was FALSE, for the rest of the block
                cond = s.children[0]
                self.need_bool(cond, facts, vflow, line)
                facts.update(self.facts_if(cond, False))
            elif kind in ("loop_rows", "loop_cols", "for_each"):
                opts = {str(o.children[0]): o.children[1] for o in s.children}
                var = str(opts.get("var", "_"))
                for k in ("from", "to"):
                    if k in opts:
                        self.need_num(self.ty(opts[k], facts, vflow, line), opts[k], line)
                if kind == "for_each":
                    src = opts.get("in")
                    if src is not None:
                        self.ty(src, facts, vflow, line)
                    over = canon(src).split("(")[0] if src is not None else ""
                    self.loopvars[var] = {"KEYS": "Text", "SHEETS": "Sheet"}.get(over, "Cell")   # D36
                else:
                    self.loopvars[var] = "Num"
                body = {k: v for k, v in facts.items() if not mentions(k, var)}
                after = self.block(n["children"], body, vflow)
                facts = {k: v for k, v in facts.items() if after.get(k) == v}
            elif kind == "select_case":   # each CASE value is compared with the selector as = does (D36)
                self.ty(s.children[0], facts, vflow, line)
                self.selectors.append(s.children[0])
                self.block(n["children"], facts, vflow)
                self.selectors.pop()
            elif kind == "case_line":
                for c in s.children:
                    if self.selectors:
                        self.ty_compare(Tree("comparison", [self.selectors[-1], Token("COMP_OP", "="), c]),
                                        facts, vflow, line)
                    else:
                        self.ty(c, facts, vflow, line)
                self.block(n["children"], facts, vflow)
            elif kind == "case_else":
                self.block(n["children"], facts, vflow)
        return facts

    def run(self):
        for k in self.spec["keys"]:
            if k["key"] == "PRE" and k["tree"] is not None:
                for cond in k["tree"].children:
                    self.need_bool(cond, {}, {}, k["line"])
        self.block(self.spec["ops"], {}, {})
        return self.problems


def check_text(text):
    spec, problems = parse(text)
    if problems or not spec or spec["kind"] != "TASK":
        return problems
    return Checker(spec).run()


if __name__ == "__main__":
    total = 0
    for p in sys.argv[1:]:
        found = check_text(Path(p).read_text(encoding="utf-8"))
        total += len(found)
        print(f"{'OK  ' if not found else 'FAIL'} {p}")
        for line, msg in found:
            print(f"     line {line}: {msg}")
    print(f"\n{len(sys.argv) - 1} files type-checked, {total} problems")
    sys.exit(1 if total else 0)
