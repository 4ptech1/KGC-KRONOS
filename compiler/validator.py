"""Validate VCL 0.1 specs against vcl/spec.md.

Usage:  python compiler/validator.py specs/*.vcl

Three stages, all reported the same way:
  1. the rules of spec.md (this file, on top of the hand-written parser in parser.py)
  2. "grammar:" problems from the formal grammar, vcl/grammar.lark (vcl_parse.py)
  3. "type:" problems from the type checker (vcl_types.py): unguarded calculations, date
     comparisons and assignments of raw cell values
A stage 2 or 3 problem on a line that already has a stage 1 error is left out, so each broken
line gets one clear message.

Prints one line per problem:  path:line: error|warning: message
Exits 1 if any spec has an error, or if no files were given. Warnings alone exit 0.
"""
import glob
import os
import re
import sys
from collections import Counter

from parser import Problem, Token, parse_spec, tokenize
from vcl_parse import OPTION_WORDS
from vcl_parse import parse as grammar_parse
from vcl_types import Checker as TypeChecker

TYPES = {"Long", "Double", "String", "Boolean", "Date", "Variant", "Range", "Worksheet", "Workbook",
         "Dictionary", "Collection"}
CONTROL_TYPES = {"TextBox", "CommandButton", "Label", "ListBox", "ComboBox", "CheckBox", "OptionButton"}
PERF_WORDS = {"screen_off", "calc_manual", "events_off"}
ALLOW_WORDS = {"file_delete", "file_overwrite", "shell", "http"}
FILL_NAMES = ("none", "black", "white", "red", "green", "blue", "yellow", "magenta", "cyan")
YES_NO = ("yes", "no")
MAX_TASK_OPS = 40

STRUCT = {"LOOP_ROWS", "LOOP_COLS", "FOR_EACH", "IF", "ELSE", "SELECT_CASE", "CASE", "CASE_ELSE", "EXIT_IF"}

# Call ops: (fewest positional args, most (None = any), {option: allowed values or kind}, required options)
# Option kinds: "expr", "var" (a declared VAR), "string", "fill", "width", "cols".
OPS = {
    "LET": (2, 2, {}, ()),
    "READ_ARRAY": (1, 1, {"into": "var"}, ("into",)),
    "WRITE_ARRAY": (1, 1, {"to": "expr"}, ("to",)),
    "DICT_ADD": (3, 3, {"mode": ("overwrite", "keep", "sum", "count")}, ("mode",)),
    "DICT_GET": (2, 2, {"into": "var"}, ("into",)),
    "SORT": (1, 1, {"key": "expr", "order": ("asc", "desc"), "header": YES_NO}, ("key",)),
    "FILTER": (1, 1, {"col": "expr", "crit": "expr"}, ("col", "crit")),
    "DEDUPE": (1, 1, {"cols": "cols", "header": YES_NO}, ("cols",)),
    "SET_VALUE": (2, 2, {}, ()),
    "SET_FORMULA": (2, 2, {"style": ("a1", "r1c1")}, ()),
    "FORMAT": (1, 1, {"numfmt": "string", "bold": YES_NO, "fill": "fill", "width": "width"}, ()),
    "CLEAR": (1, 1, {"what": ("all", "contents", "formats")}, ()),
    "COPY_ROW": (4, 4, {"cols": "expr"}, ("cols",)),
    "INSERT_ROW": (2, 2, {}, ()),
    "DELETE_ROW": (2, 2, {}, ()),
    "ADD_SHEET": (1, 1, {"name": "expr", "after": "expr", "into": "var"}, ()),
    "COPY_SHEET": (1, 1, {"after": "expr", "into": "var"}, ()),
    "OPEN_BOOK": (1, 1, {"readonly": YES_NO, "into": "var"}, ("into",)),
    "SAVE_AS": (2, 2, {"format": ("xlsx", "xlsm", "csv")}, ()),
    "CLOSE_BOOK": (1, 1, {"save": YES_NO}, ()),
    "LIST_FILES": (2, 2, {"into": "var"}, ("into",)),
    "READ_TEXT": (1, 1, {"into": "var"}, ("into",)),
    "WRITE_TEXT": (2, 2, {"mode": ("new", "append")}, ()),
    "CALL": (1, None, {}, ()),
    "LOG": (1, 1, {}, ()),
    "MSG": (1, 1, {}, ()),
    "RETURN": (0, 1, {}, ()),
}
# Value functions: (fewest args, most args, {option: allowed values})
VALUE_FUNCS = {
    "LASTROW": (2, 2, {}), "LASTCOL": (2, 2, {}), "USED_RANGE": (1, 1, {}), "COL": (2, 2, {}),
    "CELL": (3, 3, {}), "ROW_RANGE": (4, 4, {}), "RANGE": (2, 2, {}), "SHEET": (2, 2, {}),
    "FIND": (2, 2, {"match": ("whole", "part")}), "ROWS": (1, 1, {}), "COLS": (1, 1, {}),
    "ARR": (3, 3, {}), "COUNT": (1, 1, {}), "EXISTS": (2, 2, {}), "LEN": (1, 1, {}), "TRIM": (1, 1, {}),
    "UPPER": (1, 1, {}), "LOWER": (1, 1, {}), "LEFT": (2, 2, {}), "RIGHT": (2, 2, {}),
    "IS_EMPTY": (1, 1, {}), "IS_NUMBER": (1, 1, {}), "IS_DATE": (1, 1, {}), "IS_NOTHING": (1, 1, {}),
    "FILE_NAME": (1, 1, {}), "SHEETS": (1, 1, {}), "CELLS": (1, 1, {}), "KEYS": (1, 1, {}),
    "ITEMS": (1, 1, {}),
}
STRING_SECOND = {"COL", "RANGE", "SHEET"}  # second argument is a "quoted" header, ref or sheet name
CONSTANTS = {"THIS_BOOK", "ACTIVE_BOOK", "ACTIVE_SHEET", "TODAY", "NOW", "TRUE", "FALSE", "NOTHING"}
FOR_EACH_IN = {"SHEETS", "CELLS", "KEYS", "ITEMS"}
BINARY = {"=", "<>", "<", ">", "<=", ">=", "+", "-", "*", "/", "&"}
VBA_LEAKS = ["Dim ", "Set ", "End If", ".Value", "Cells(", "Range(", "Next ", "Wend"]

# TEST assertion functions and the kind of each argument
ASSERT_FUNCS = {"COUNTIF", "COUNTIFS", "COUNTA", "SUM", "SUMIFS", "VALUE", "ROWS_USED", "SHEET_EXISTS",
                "FILE_EXISTS", "FILL_COUNT"}
ASSERT_COMPARE = {"=", "<>", "<", ">", "<=", ">=", "+", "-", "*", "/"}

KEYWORD = re.compile(r"[A-Z][A-Z0-9_]+")
CAMEL = re.compile(r"[a-z][A-Za-z0-9]*")
PASCAL = re.compile(r"[A-Z][A-Za-z0-9]*")
FORM_EVENT = re.compile(r"[a-z][A-Za-z0-9]*_[A-Z][A-Za-z]*")
MODULE_NAME = re.compile(r"mod[A-Z][A-Za-z0-9]*")
FORM_NAME = re.compile(r"frm[A-Z][A-Za-z0-9]*")
FIXTURE_FILE = re.compile(r"[a-z][a-z0-9_]*_\d{2}\.xlsx")
SHEET_NAME = re.compile(r"[A-Za-z_]\w*")
HEX_COLOUR = re.compile(r"#[0-9A-Fa-f]{6}")


def norm(header):
    """Headers match ignoring case and extra spaces, like COL()."""
    return " ".join(header.split()).casefold()


def column_letter(index):
    letters = ""
    index += 1
    while index:
        index, rest = divmod(index - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


# ---------------------------------------------------------------- grouping tokens

class Call:
    """NAME(args): each arg is (key token or None, items)."""

    def __init__(self, tok, args):
        self.tok, self.name, self.args = tok, tok.text, args

    @property
    def positional(self):
        return [items for key, items in self.args if key is None]

    @property
    def options(self):
        return [(key, items) for key, items in self.args if key is not None]


class Group:
    """(items) used for grouping inside an expression."""

    def __init__(self, items):
        self.items = items


class Unbalanced(Exception):
    pass


def build(tokens):
    """Group tokens into Calls and Groups. Returns (items, error message)."""
    pos = 0

    def sequence(inside):
        nonlocal pos
        items = []
        while pos < len(tokens):
            tok = tokens[pos]
            if tok.kind == "op" and tok.text == ")":
                if inside:
                    return items
                raise Unbalanced("unbalanced brackets: ')' without '('")
            pos += 1
            if tok.kind == "op" and tok.text == "(":
                prev = items[-1] if items else None
                inner = sequence(True)
                pos += 1  # the closing ")"
                if isinstance(prev, Token) and prev.kind == "name" and not tok.spaced:
                    items[-1] = Call(prev, split_args(inner))
                else:
                    items.append(Group(inner))
            else:
                items.append(tok)
        if inside:
            raise Unbalanced("unbalanced brackets: '(' without ')'")
        return items

    try:
        return sequence(False), ""
    except Unbalanced as exc:
        return [], str(exc)


def split_on(items, text, kind="op"):
    parts = [[]]
    for item in items:
        if isinstance(item, Token) and item.kind == kind and item.text == text:
            parts.append([])
        else:
            parts[-1].append(item)
    return parts


def split_args(items):
    if not items:
        return []
    args = []
    for part in split_on(items, ","):
        # G2: only a reserved option word directly followed by "=" makes an option; any other
        # "name=value" is a comparison, exactly as in vcl/grammar.lark.
        if (len(part) >= 2 and is_tok(part[0], "name") and is_tok(part[1], "op", "=")
                and not part[1].spaced and part[0].text in OPTION_WORDS):
            args.append((part[0], part[2:]))
        else:
            args.append((None, part))
    return args


def is_tok(item, kind, text=None):
    return isinstance(item, Token) and item.kind == kind and (text is None or item.text == text)


def show(items):
    out = []
    for item in items:
        if isinstance(item, Call):
            out.append(f"{item.name}(...)")
        elif isinstance(item, Group):
            out.append("(...)")
        else:
            out.append(item.text)
    return " ".join(out)


def blank_strings(text):
    return re.sub(r'"(?:[^"]|"")*"', '""', text)


# ---------------------------------------------------------------- TASK checks

class TaskChecker:
    """Checks one TASK spec. in_build=True for a TASK section inside a BUILD."""

    def __init__(self, spec, in_build=False, build_fixtures=()):
        self.spec, self.in_build = spec, in_build
        self.fixtures = list(build_fixtures) + spec.fixtures
        self.problems = []
        self.line = spec.line
        self.mode = "op"
        self.kind = spec.value("KIND")
        self.args = {d.name: d.type for d in spec.args}
        self.vars = {d.name: d.type for d in spec.vars}
        self.loopvars = {}
        self.pre_cols = set()      # (arg, header) checked in PRE
        self.op_cols = []          # (line, arg, header) used in OP lines
        self.header_writes = []    # (arg, header) written to row 1
        self.header_copies = []    # (source arg, target arg) row 1 copied
        self.calls = []            # (line, "module.Proc", number of args)
        self.sheet_ops = False     # ADD_SHEET / COPY_SHEET can create sheets the fixture lacks
        self.pre_conditions = []   # [(arg, header)] when every PRE condition is COL($x, "h") > 0
        self.pre_evaluable = True

    def error(self, message, line=None):
        self.problems.append(Problem(line or self.line, "error", message))

    def warning(self, message, line=None):
        self.problems.append(Problem(line or self.line, "warning", message))

    def run(self):
        spec = self.spec
        self.check_keys()
        self.collect_loopvars()
        self.check_decls()
        for pre in spec.lines("PRE"):
            self.check_pre(pre)
        self.mode = "op"
        self.walk(spec.ops, None, [])
        for line, arg, header in self.op_cols:
            if (arg, norm(header)) not in self.pre_cols:
                self.error(f'header "{header}" on ${arg} is used but not checked in PRE (rule 14)', line)
        op_count = sum(1 for _ in spec.all_ops())
        if op_count > MAX_TASK_OPS:
            first = spec.lines("OP")[0].line
            self.warning(f"TASK has {op_count} OP lines (more than {MAX_TASK_OPS}): consider splitting it "
                         "into a BUILD (rule 9)", first)
        check_fixtures(self, spec.fixtures)
        self.check_tests()
        return self.problems

    # -- keys and declarations

    def check_keys(self):
        spec = self.spec
        counts = Counter(k.key for k in spec.keys)
        for key in ("PROC", "KIND", "MODULE", "OP", "ERR"):
            if not counts[key]:
                self.error(f"missing {key}")
        for key in ("PROC", "KIND", "MODULE", "ARGS", "RET", "PRE", "PERF", "ERR", "ALLOW"):
            for extra in spec.lines(key)[1:]:
                self.error(f"{key} appears more than once", extra.line)
        for k in spec.lines("KIND")[:1]:
            if k.value not in ("Sub", "Function"):
                self.error(f"KIND must be Sub or Function, got '{k.value}'", k.line)
        rets = spec.lines("RET")
        if self.kind == "Function" and not rets:
            self.error("a Function needs RET <Type>", spec.lines("KIND")[0].line)
        if self.kind == "Sub" and rets:
            self.error("RET is only for a Function", rets[0].line)
        for k in rets[:1]:
            if k.value not in TYPES:
                self.error(f"unknown type {k.value}", k.line)
        module = spec.value("MODULE")
        for k in spec.lines("MODULE")[:1]:
            if not (MODULE_NAME.fullmatch(module) or FORM_NAME.fullmatch(module)):
                self.error(f"MODULE must be mod + PascalCase, like modCleanup (or a form, frmName), "
                           f"got '{module}'", k.line)
        for k in spec.lines("PROC")[:1]:
            event = FORM_NAME.fullmatch(module) and FORM_EVENT.fullmatch(k.value)
            if not (PASCAL.fullmatch(k.value) or event):
                self.error(f"PROC must be PascalCase, like DeleteRowsByValue, got '{k.value}'", k.line)
        for k in spec.lines("ERR")[:1]:
            if k.value != "standard":
                self.error(f"ERR must be 'standard', got '{k.value}'", k.line)
        for k in spec.lines("PERF")[:1]:
            words = k.value.split()
            if not words:
                self.error("PERF needs at least one of screen_off, calc_manual, events_off", k.line)
            for w in words:
                if w not in PERF_WORDS:
                    self.error(f"unknown PERF value {w} (use screen_off, calc_manual, events_off)", k.line)
            for w, n in Counter(words).items():
                if n > 1:
                    self.error(f"PERF lists {w} twice", k.line)
        for k in spec.lines("ALLOW")[:1]:
            words = k.value.split()
            if not words:
                self.error("ALLOW needs at least one of file_delete, file_overwrite, shell, http", k.line)
            for w in words:
                if w not in ALLOW_WORDS:
                    self.error(f"unknown ALLOW value {w} (use file_delete, file_overwrite, shell, http)", k.line)
                else:
                    self.warning(f"ALLOW {w}: risky actions can't be tested on the work PC before March 2027",
                                 k.line)

    def collect_loopvars(self):
        for node in self.spec.all_ops():
            tokens, _ = tokenize(node.text)
            if tokens and tokens[0].text in ("LOOP_ROWS", "LOOP_COLS", "FOR_EACH"):
                for a, b, c in zip(tokens, tokens[1:], tokens[2:]):
                    if a.text == "var" and b.text == "=" and c.kind == "name":
                        self.loopvars.setdefault(c.text, node.line)

    def check_decls(self):
        seen = set()
        for decl in self.spec.args + self.spec.vars:
            if not CAMEL.fullmatch(decl.name):
                self.error(f"name {decl.name} must be camelCase (rule 12)", decl.line)
            if decl.type not in TYPES:
                self.error(f"unknown type {decl.type}", decl.line)
            if decl.name in seen:
                self.error(f"{decl.name} is declared twice", decl.line)
            seen.add(decl.name)
            if decl.name in self.loopvars:
                self.error(f"loop variable {decl.name} is declared automatically: don't declare it in "
                           "ARGS or VAR", decl.line)
            if decl.name in OPTION_WORDS:
                self.error(f"{decl.name} is a reserved option word: choose another name (G2)", decl.line)

    # -- PRE

    def check_pre(self, pre):
        self.mode, self.line = "pre", pre.line
        items = self.items_of(pre.value)
        if items is None:
            self.pre_evaluable = False
            return
        for condition in split_on(items, ";"):
            if not condition:
                self.error("empty PRE condition")
                self.pre_evaluable = False
                continue
            self.check_expr(condition, "PRE condition")
            if (len(condition) == 3 and isinstance(condition[0], Call) and condition[0].name == "COL"
                    and is_tok(condition[1], "op", ">") and is_tok(condition[2], "number", "0")
                    and len(condition[0].args) == 2):
                (_, first), (_, second) = condition[0].args
                if len(first) == 1 and is_tok(first[0], "arg") and len(second) == 1 \
                        and is_tok(second[0], "string"):
                    self.pre_conditions.append((first[0].text[1:], second[0].text[1:-1]))
                    continue
            self.pre_evaluable = False

    # -- OP lines

    def items_of(self, text):
        leak = next((leak for leak in VBA_LEAKS if leak in blank_strings(text)), None)
        if leak:
            self.error(f"VBA syntax in spec: {leak.strip()}")
            return None
        tokens, err = tokenize(text)
        if err:
            self.error(err)
            return None
        items, err = build(tokens)
        if err:
            self.error(err)
            return None
        return items

    def walk(self, nodes, parent, ancestors):
        prev = None
        for node in nodes:
            self.line = node.line
            info = self.check_op(node)
            name = info["name"]
            if node.children and not info["block"]:
                self.error(f"{name or 'this step'} can't have nested steps", node.children[0].line)
            if info["block"] and not node.children and name != "CASE":
                self.error(f"{name} has no nested steps under it", node.line)
            if name == "ELSE":
                if not (prev and prev["name"] == "IF" and prev["block"]):
                    self.error("ELSE must come right after a block IF at the same level", node.line)
            if name in ("CASE", "CASE_ELSE") and (parent or {}).get("name") != "SELECT_CASE":
                self.error(f"{name} must be nested under SELECT_CASE", node.line)
            if (parent or {}).get("name") == "SELECT_CASE" and name not in ("CASE", "CASE_ELSE"):
                self.error("SELECT_CASE can only contain CASE and CASE_ELSE", node.line)
            if name == "CASE" and prev and prev["name"] == "CASE_ELSE":
                self.error("CASE_ELSE must be the last case", node.line)
            if "DELETE_ROW" in (name, info.get("action")) and not any(
                    a["name"] == "LOOP_ROWS" and a.get("dir") == "up" for a in ancestors):
                self.error("DELETE_ROW is only allowed inside LOOP_ROWS dir=up (rule 10)", node.line)
            self.walk(node.children, info, ancestors + [info])
            prev = info

    def check_op(self, node):
        info = {"name": "", "block": False}
        items = self.items_of(node.text)
        if not items:
            if items is not None:
                self.error("OP needs a step")
            info["block"] = bool(node.children)  # don't pile up nesting errors on a bad line
            return info
        head = items[0]
        if isinstance(head, Call):
            name = head.name
        elif is_tok(head, "name"):
            name = head.text
        else:
            self.error(f"an OP step starts with an op name, got '{show(items[:1])}'")
            return info
        info["name"] = name
        if name == "NOTE":
            if not (len(items) == 2 and is_tok(items[1], "string")):
                self.error('NOTE takes one "<terse technical text>"')
        elif name in STRUCT:
            self.check_struct(name, items[1:], info)
        elif name in OPS:
            if not isinstance(head, Call):
                self.error(f"{name} needs its brackets right after the name: {name}(...)")
            else:
                if len(items) > 1:
                    self.error(f"unexpected text after {name}(...): {show(items[1:])}")
                self.check_call_op(head)
        elif name in VALUE_FUNCS or name in CONSTANTS:
            self.error(f"{name} is a value, not an op")
        else:
            self.error(f"unknown op {name}")
            info["block"] = bool(node.children)
        return info

    def check_struct(self, name, rest, info):
        if name in ("LOOP_ROWS", "LOOP_COLS"):
            info["block"] = True
            dirs = ("up", "down") if name == "LOOP_ROWS" else ("left", "right")
            opts = self.struct_options(name, rest, ("dir", "from", "to", "var"))
            for key in ("dir", "from", "to", "var"):
                if key not in opts:
                    self.error(f"{name} needs {key}=")
            if "dir" in opts:
                info["dir"] = self.check_choice(name, "dir", opts["dir"], dirs)
            for key in ("from", "to"):
                if key in opts:
                    self.check_expr(opts[key], f"{key}= value")
            if "var" in opts:
                self.check_loopvar(name, opts["var"])
        elif name == "FOR_EACH":
            info["block"] = True
            opts = self.struct_options(name, rest, ("var", "in"))
            for key in ("var", "in"):
                if key not in opts:
                    self.error(f"FOR_EACH needs {key}=")
            if "var" in opts:
                self.check_loopvar(name, opts["var"])
            if "in" in opts:
                value = opts["in"]
                if len(value) == 1 and isinstance(value[0], Call) and value[0].name in FOR_EACH_IN:
                    self.check_value_call(value[0])
                else:
                    self.error("FOR_EACH in= must be SHEETS(...), CELLS(...), KEYS(...) or ITEMS(...)")
        elif name == "IF":
            then = next((i for i, item in enumerate(rest) if is_tok(item, "name", "THEN")), None)
            if then is None:
                info["block"] = True
                self.check_expr(rest, "IF condition")
                return
            self.check_expr(rest[:then], "IF condition")
            action = rest[then + 1:]
            if len(action) == 1 and isinstance(action[0], Call) and action[0].name in OPS:
                info["action"] = action[0].name
                self.check_call_op(action[0])
            elif len(action) == 2 and is_tok(action[0], "name", "NOTE") and is_tok(action[1], "string"):
                info["action"] = "NOTE"
            else:
                self.error("IF ... THEN needs exactly one op after THEN, like THEN LET(x, 1)")
        elif name in ("ELSE", "CASE_ELSE"):
            info["block"] = True
            if rest:
                self.error(f"{name} takes nothing after it")
        elif name == "SELECT_CASE":
            info["block"] = True
            self.check_expr(rest, "SELECT_CASE value")
        elif name == "CASE":
            info["block"] = True
            for value in split_on(rest, ","):
                self.check_expr(value, "CASE value")
        elif name == "EXIT_IF":
            self.check_expr(rest, "EXIT_IF condition")

    def struct_options(self, name, rest, allowed):
        opts, key, i = {}, None, 0
        while i < len(rest):
            item = rest[i]
            nxt = rest[i + 1] if i + 1 < len(rest) else None
            if is_tok(item, "name") and item.text in allowed and is_tok(nxt, "op", "=") and not nxt.spaced:
                key = item.text
                if key in opts:
                    self.error(f"{name} has {key}= twice")
                opts[key] = []
                i += 2
                continue
            if key is None:
                self.error(f"{name} expects key=value options ({', '.join(a + '=' for a in allowed)}), "
                           f"got '{show([item])}'")
                return opts
            opts[key].append(item)
            i += 1
        return opts

    def check_loopvar(self, name, value):
        if not (len(value) == 1 and is_tok(value[0], "name") and CAMEL.fullmatch(value[0].text)):
            self.error(f"{name} var= needs a camelCase loop variable name, got '{show(value)}'")
        elif value[0].text in OPTION_WORDS:
            self.error(f"{value[0].text} is a reserved option word: choose another loop variable name (G2)")

    def check_choice(self, name, key, value, allowed):
        text = show(value)
        if not (len(value) == 1 and is_tok(value[0], "name") and value[0].text in allowed):
            self.error(f"{name} {key}= must be {'|'.join(allowed)}, got '{text}'")
            return None
        return text

    def call_parts(self, call):
        """(positional, options) of a call. Under G2 only reserved words make options, so to the
        grammar "colour=red" is a comparison. When its name isn't declared it is almost always a
        misspelt option, so it is reported as one ("FORMAT has no option colour=")."""
        positional, options = [], []
        for key, items in call.args:
            if key is not None:
                options.append((key, items))
            elif (len(items) >= 2 and is_tok(items[0], "name") and is_tok(items[1], "op", "=")
                  and not items[1].spaced and CAMEL.fullmatch(items[0].text)
                  and items[0].text not in self.vars and items[0].text not in self.loopvars
                  and items[0].text not in self.args):
                options.append((items[0], items[2:]))
            else:
                positional.append(items)
        return positional, options

    def check_call_op(self, call):
        name = call.name
        fewest, most, allowed, required = OPS[name]
        positional, options = self.call_parts(call)
        if name == "DEDUPE":  # cols=<n,n>: the numbers after cols= belong to it
            merged, positional, last = [], [], None
            for key, items in call.args:
                if key is None and last == "cols" and len(items) == 1 and is_tok(items[0], "number"):
                    merged[-1][1].extend([Token("op", ",", False)] + items)
                    continue
                if key is None:
                    positional.append(items)
                else:
                    merged.append((key, list(items)))
                    last = key.text
            options = merged
        count = len(positional)
        if count < fewest or (most is not None and count > most):
            want = str(fewest) if fewest == most else (f"{fewest} or more" if most is None
                                                       else f"{fewest} to {most}")
            self.error(f"{name} takes {want} argument{'s' if want != '1' else ''} before its options, "
                       f"got {count}")
        seen = set()
        for key, items in options:
            if key.text in seen:
                self.error(f"{name} has {key.text}= twice")
            seen.add(key.text)
            if key.text not in allowed:
                self.error(f"{name} has no option {key.text}=")
                continue
            self.check_option(name, key.text, allowed[key.text], items)
        for key in required:
            if key not in seen:
                self.error(f"{name} needs {key}=")
        if name == "FORMAT" and not options:
            self.error("FORMAT needs at least one of numfmt=, bold=, fill=, width=")
        for i, items in enumerate(positional):
            if name == "LET" and i == 0:
                if not (len(items) == 1 and is_tok(items[0], "name") and items[0].text in self.vars):
                    self.error(f"LET needs a declared VAR first, got '{show(items)}'")
            elif name == "CALL" and i == 0:
                target = items[0].text if len(items) == 1 and is_tok(items[0], "name") else ""
                parts = target.split(".")
                if len(parts) != 2 or not (MODULE_NAME.fullmatch(parts[0]) or FORM_NAME.fullmatch(parts[0])) \
                        or not PASCAL.fullmatch(parts[1]):
                    self.error(f"CALL needs <module>.<Proc> first, like modCleanup.RemoveRows, got "
                               f"'{show(items)}'")
                else:
                    self.calls.append((self.line, target, count - 1))
            elif name in ("SET_FORMULA", "LIST_FILES") and i == 1:
                if not (len(items) == 1 and is_tok(items[0], "string")):
                    self.error(f'{name} needs a "quoted" {"formula" if name == "SET_FORMULA" else "pattern"} '
                               "as its second argument")
            else:
                self.check_expr(items, f"{name} argument")
        if name == "RETURN":
            if self.kind == "Function" and count != 1:
                self.error("a Function's RETURN needs a value: RETURN(<expr>)")
            if self.kind == "Sub" and count:
                self.error("a Sub can't RETURN a value")
        if name in ("ADD_SHEET", "COPY_SHEET"):
            self.sheet_ops = True
        if name == "SET_VALUE" and len(positional) == 2:
            target, value = positional
            arg = self.row1_arg(target)
            if arg and len(value) == 1 and is_tok(value[0], "string"):
                self.header_writes.append((arg, value[0].text[1:-1]))
        if name == "COPY_ROW" and len(positional) == 4:
            src, src_row, dst, dst_row = positional
            if all(len(p) == 1 for p in positional) and is_tok(src[0], "arg") and is_tok(dst[0], "arg") \
                    and is_tok(src_row[0], "number", "1") and is_tok(dst_row[0], "number", "1"):
                self.header_copies.append((src[0].text[1:], dst[0].text[1:]))

    @staticmethod
    def row1_arg(target):
        """'x' if target is CELL($x, 1, <col>), else None."""
        if len(target) == 1 and isinstance(target[0], Call) and target[0].name == "CELL":
            pos = target[0].positional
            if len(pos) == 3 and len(pos[0]) == 1 and is_tok(pos[0][0], "arg") \
                    and len(pos[1]) == 1 and is_tok(pos[1][0], "number", "1"):
                return pos[0][0].text[1:]
        return None

    def check_option(self, name, key, kind, items):
        text = show(items)
        if isinstance(kind, tuple):
            self.check_choice(name, key, items, kind)
        elif kind == "var":
            if not (len(items) == 1 and is_tok(items[0], "name") and items[0].text in self.vars):
                self.error(f"{name} {key}= needs a declared VAR, got '{text}'")
        elif kind == "string":
            if not (len(items) == 1 and is_tok(items[0], "string")):
                self.error(f'{name} {key}= needs a "quoted" value, got \'{text}\'')
        elif kind == "fill":
            ok = len(items) == 1 and ((is_tok(items[0], "name") and items[0].text in FILL_NAMES)
                                      or (is_tok(items[0], "hex") and HEX_COLOUR.fullmatch(items[0].text)))
            if not ok:
                self.error(f"{name} fill= must be {', '.join(FILL_NAMES)} or #RRGGBB, got '{text}'")
        elif kind == "width":
            if not (len(items) == 1 and (is_tok(items[0], "name", "auto") or is_tok(items[0], "number"))):
                self.error(f"{name} width= must be auto or a number, got '{text}'")
        elif kind == "cols":
            for part in split_on(items, ","):
                self.check_expr(part, f"{name} cols= value")
        else:
            self.check_expr(items, f"{name} {key}= value")

    # -- expressions

    def check_expr(self, items, what="expression"):
        if not items:
            self.error(f"missing {what}")
            return
        want_operand = True
        for item in items:
            if want_operand:
                if is_tok(item, "name", "NOT") or is_tok(item, "op", "-"):
                    continue
                self.check_operand(item)
                want_operand = False
            elif (is_tok(item, "op") and item.text in BINARY) or is_tok(item, "name", "AND") \
                    or is_tok(item, "name", "OR"):
                want_operand = True
            else:
                self.error(f"missing operator before '{show([item])}' in {what}")
        if want_operand:
            self.error(f"{what} ends with an operator")

    def check_operand(self, item):
        if isinstance(item, Group):
            self.check_expr(item.items)
        elif isinstance(item, Call):
            self.check_value_call(item)
        elif item.kind in ("string", "number"):
            pass
        elif item.kind == "arg":
            self.check_arg(item.text[1:])
        elif item.kind == "name":
            self.check_name(item.text)
        elif item.kind == "hex":
            self.error(f"{item.text}: #RRGGBB colours are only for fill=")
        elif item.kind == "range":
            self.error(f"{item.text}: Sheet!range references are only for TEST assertions")
        else:
            self.error(f"unexpected '{item.text}'")

    def check_name(self, word):
        if KEYWORD.fullmatch(word):
            if word in CONSTANTS:
                return
            if word in VALUE_FUNCS:
                self.error(f"{word} needs its brackets: {word}(...)")
            elif word in OPS or word in STRUCT or word == "NOTE":
                self.error(f"{word} is an op; it can't be used inside an expression")
            elif word in ("AND", "OR", "NOT", "THEN"):
                self.error(f"misplaced {word}")
            else:
                self.error(f"unknown word {word}")
        elif "." in word:
            self.error(f"'{word}': VCL has no .Property syntax; dots only appear in a CALL target (module.Proc)")
        elif word in self.vars or word in self.loopvars:
            pass
        elif word in self.args:
            self.error(f"argument {word} must be written ${word}")
        else:
            self.error(f"undeclared name {word}")

    def check_arg(self, name):
        if name in self.args:
            return
        if name in self.vars or name in self.loopvars:
            self.error(f"{name} is a variable, not an argument: write it without $")
        else:
            self.error(f"undeclared argument ${name}")

    def check_value_call(self, call):
        name = call.name
        if name in OPS or name in STRUCT:
            self.error(f"{name} is an op; it can't be used inside an expression")
            return
        if name not in VALUE_FUNCS:
            self.error(f"unknown function {name}")
            for items in call.positional:
                self.check_expr(items, f"{name} argument")
            return
        fewest, most, allowed = VALUE_FUNCS[name]
        positional, options = self.call_parts(call)
        if not fewest <= len(positional) <= most:
            self.error(f"{name} takes {fewest} argument{'s' if fewest > 1 else ''}, got {len(positional)}")
        for key, items in options:
            if key.text not in allowed:
                self.error(f"{name} has no option {key.text}=")
            else:
                self.check_choice(name, key.text, items, allowed[key.text])
        for i, items in enumerate(positional):
            if name in STRING_SECOND and i == 1:
                if not (len(items) == 1 and is_tok(items[0], "string")):
                    self.error(f'{name} needs a "quoted" text as its second argument')
            else:
                self.check_expr(items, f"{name} argument")
        if name == "COL" and len(positional) == 2:
            first, second = positional
            if len(first) == 1 and is_tok(first[0], "arg") and len(second) == 1 and is_tok(second[0], "string"):
                arg, header = first[0].text[1:], second[0].text[1:-1]
                if self.mode == "pre":
                    self.pre_cols.add((arg, norm(header)))
                else:
                    self.op_cols.append((self.line, arg, header))

    # -- TESTs

    def check_tests(self):
        spec = self.spec
        tests = spec.tests
        files = {}
        for fx in self.fixtures:
            files.setdefault(fx.file, {})[fx.sheet] = fx
        flags = []
        for test in tests:
            self.line = test.line
            flags.append(self.check_test(test, files))
        if self.in_build:
            return
        first = tests[0].line if tests else spec.line
        if len(tests) < 2:
            self.error("a spec needs at least two TEST lines (rule 5)", first)
        if tests and not any(f["error_count_0"] for f in flags):
            self.error("no TEST checks ERROR_COUNT = 0 (rule 5)", first)
        if tests and not any(f["pre"] for f in flags):
            self.error("no TEST guards against collateral changes with a PRE_ measure (rule 5)", first)
        if tests and not any(t.runs >= 2 for t in tests):
            self.error("no TEST uses runs=2 to prove re-runs are safe (rule 16)", first)
        used = {t.fixture for t in tests}
        many_sheets = any(len(files.get(f, {})) > 1 for f in used)
        if spec.lines("PRE") and many_sheets:
            want = "RESULT = -1 AND ERROR_COUNT = 1" if self.kind == "Function" else "ERROR_COUNT = 1"
            if not any(f["wrong_sheet"] for f in flags):
                self.error(f"PRE is used and the fixture has another sheet, but no wrong-sheet TEST expects "
                           f"{want} (rule 5)", first)
        if spec.lines("PRE") and self.pre_evaluable and self.pre_conditions and flags \
                and not any(f["passes_pre"] for f in flags if f["passes_pre"] is not None):
            self.error("no TEST passes PRE: every TEST's sheets miss a header that PRE checks", first)

    def check_test(self, test, files):
        flags = {"error_count_0": False, "wrong_sheet": False, "pre": False, "passes_pre": None}
        spec = self.spec
        if test.fixture == "target":
            self.error("fixture=target is not allowed: use a named synthetic FIXTURE (rule 18)")
            return flags
        sheets = files.get(test.fixture)
        if sheets is None:
            self.error(f"TEST uses fixture {test.fixture}, but no FIXTURE line builds it")
        proc = spec.value("PROC")
        if not self.in_build and test.procedure != proc:
            self.error(f"TEST calls {test.procedure}, but PROC is {proc}")
        argmap = {}
        if len(test.args) != len(spec.args):
            self.error(f"TEST passes {len(test.args)} argument{'s' if len(test.args) != 1 else ''}, "
                       f"but ARGS has {len(spec.args)}")
        elif sheets is not None:
            for decl, value in zip(spec.args, test.args):
                if decl.type != "Worksheet":
                    continue
                if value in sheets:
                    argmap[decl.name] = value
                else:
                    self.error(f"TEST passes {value} for ${decl.name}, but {test.fixture} has no sheet {value}")
        headers = None
        if sheets is not None:
            headers = {name: {norm(h) for _, h in fx.headers} for name, fx in sheets.items()}
            for _ in range(len(self.header_copies) + 1):
                for arg, header in self.header_writes:
                    if arg in argmap:
                        headers[argmap[arg]].add(norm(header))
                for src, dst in self.header_copies:
                    if src in argmap and dst in argmap:
                        headers[argmap[dst]] |= headers[argmap[src]]
        result = check_assertion(self, test, sheets, headers)
        flags.update(result)
        if self.pre_evaluable and self.pre_conditions and sheets is not None:
            missing = [(arg, h) for arg, h in self.pre_conditions
                       if arg in argmap and norm(h) not in {norm(x) for _, x in sheets[argmap[arg]].headers}]
            if all(arg in argmap for arg, _ in self.pre_conditions):
                flags["passes_pre"] = not missing
                if missing and self.kind == "Function" and not result["result_minus_1"]:
                    arg, h = missing[0]
                    self.error(f'TEST fails PRE (sheet {argmap[arg]} has no "{h}" header) but doesn\'t expect '
                               "RESULT = -1")
        return flags


def check_assertion(checker, test, sheets, headers):
    """Check one TEST assertion. Returns flags used by the spec-level TEST rules."""
    flags = {"error_count_0": False, "wrong_sheet": False, "pre": False, "result_minus_1": False}
    tokens, err = tokenize(test.assertion)
    if err:
        checker.error(f"assertion: {err}")
        return flags
    conjuncts, depth = [[]], 0
    for tok in tokens:
        depth += tok.text == "(" and tok.kind == "op"
        depth -= tok.text == ")" and tok.kind == "op"
        if depth == 0 and tok.kind == "name" and tok.text == "AND":
            conjuncts.append([])
        else:
            conjuncts[-1].append(tok)
    texts = {"".join(t.text for t in c) for c in conjuncts}
    flags["error_count_0"] = "ERROR_COUNT=0" in texts
    flags["result_minus_1"] = "RESULT=-1" in texts
    if checker.kind == "Function":
        flags["wrong_sheet"] = "RESULT=-1" in texts and "ERROR_COUNT=1" in texts
    else:
        flags["wrong_sheet"] = "ERROR_COUNT=1" in texts
    flags["pre"] = any(t.kind == "name" and t.text.startswith("PRE_") for t in tokens)
    for conjunct in conjuncts:
        items, err = build(conjunct)
        if err:
            checker.error(f"assertion: {err}")
            continue
        if not items:
            checker.error("assertion: empty condition around AND")
            continue
        AssertionChecker(checker, test, sheets, headers).expr(items)
    return flags


class AssertionChecker:
    def __init__(self, checker, test, sheets, headers):
        self.c, self.test, self.sheets, self.headers = checker, test, sheets, headers

    def error(self, message):
        self.c.error(f"assertion: {message}")

    def expr(self, items):
        want_operand = True
        for item in items:
            if want_operand:
                if is_tok(item, "op", "-"):
                    continue
                self.operand(item)
                want_operand = False
            elif is_tok(item, "op") and item.text in ASSERT_COMPARE:
                want_operand = True
            elif is_tok(item, "name", "OR"):
                self.error("combine assertions with AND, not OR")
                want_operand = True
            else:
                self.error(f"missing operator before '{show([item])}'")
        if want_operand:
            self.error("ends with an operator")

    def operand(self, item):
        if isinstance(item, Group):
            self.expr(item.items)
        elif isinstance(item, Call):
            self.function(item)
        elif item.kind in ("number", "string"):
            pass
        elif item.kind == "name" and item.text in ("RESULT", "ERROR_COUNT"):
            if item.text == "RESULT" and self.c.kind == "Sub":
                self.error("RESULT needs a Function: a Sub returns nothing")
        elif item.kind == "range":
            self.error(f"{item.text}: put ranges inside a function, like COUNTA({item.text})")
        elif item.kind == "name" and item.text.startswith("PRE_"):
            self.error(f"{item.text}: PRE_ goes before a measure, like PRE_COUNTA(...)")
        else:
            self.error(f"unknown word {item.text}")

    def function(self, call):
        name = call.name[4:] if call.name.startswith("PRE_") else call.name
        if name not in ASSERT_FUNCS:
            self.error(f"unknown function {call.name}")
            return
        if call.options:
            self.error(f"{call.name} has no options")
        args = call.positional
        n = len(args)
        if name == "COUNTIFS":
            kinds = ["range", "crit"] * (n // 2)
            ok = n >= 2 and n % 2 == 0
            want = "range, criteria pairs"
        elif name == "SUMIFS":
            kinds = ["range"] + ["range", "crit"] * ((n - 1) // 2)
            ok = n >= 3 and n % 2 == 1
            want = "a sum range, then range, criteria pairs"
        else:
            kinds = {"COUNTIF": ["range", "crit"], "COUNTA": ["range"], "SUM": ["range"], "VALUE": ["cell"],
                     "ROWS_USED": ["sheet"], "SHEET_EXISTS": ["string"], "FILE_EXISTS": ["string"],
                     "FILL_COUNT": ["range", "hex"]}[name]
            ok = n == len(kinds)
            want = f"{len(kinds)} argument{'s' if len(kinds) > 1 else ''}"
        if not ok:
            self.error(f"{call.name} takes {want}, got {n}")
            return
        for kind, items in zip(kinds, args):
            self.argument(call.name, kind, items)

    def argument(self, fname, kind, items):
        text = show(items)
        single = items[0] if len(items) == 1 else None
        if kind in ("range", "cell"):
            if not is_tok(single, "range"):
                self.error(f"{fname} needs a <Sheet>!<range> here, got '{text}'")
                return
            self.range(single)
            if kind == "cell":
                ref = single.text.split("!", 1)[1]
                if (single.header is not None and not single.row) or \
                        (single.header is None and (":" in ref or not re.search(r"\d", ref))):
                    self.error(f"{fname} needs one cell, like Sheet![Header]2 or Sheet!B2, got '{text}'")
        elif kind == "crit":
            if not (is_tok(single, "string") or is_tok(single, "number")
                    or (len(items) == 2 and is_tok(items[0], "op", "-") and is_tok(items[1], "number"))):
                self.error(f"{fname} criteria must be a \"quoted\" text or a number, got '{text}'")
        elif kind == "sheet":
            if not is_tok(single, "name"):
                self.error(f"{fname} needs a sheet name, got '{text}'")
            else:
                self.sheet(single.text)
        elif kind == "string":
            if not is_tok(single, "string"):
                self.error(f'{fname} needs a "quoted" name, got \'{text}\'')
        elif kind == "hex":
            if not (is_tok(single, "hex") and HEX_COLOUR.fullmatch(single.text)):
                self.error(f"{fname} needs a #RRGGBB colour, got '{text}'")

    def sheet(self, name):
        if self.sheets is None or name in self.sheets or self.c.sheet_ops:
            return False
        self.error(f"sheet {name} is not in {self.test.fixture}")
        return True

    def range(self, tok):
        if self.sheet(tok.sheet) or tok.header is None or self.headers is None:
            return
        known = self.headers.get(tok.sheet)
        if known is not None and norm(tok.header) not in known:
            self.error(f"sheet {tok.sheet} has no [{tok.header}] header in this TEST (not in its FIXTURE cols, "
                       "and not written or copied into row 1)")


def check_fixtures(checker, fixtures):
    seen = set()
    for fx in fixtures:
        line = fx.line
        if not FIXTURE_FILE.fullmatch(fx.file):
            checker.error(f"fixture file {fx.file} must be named like snake_case_01.xlsx (rule 12)", line)
        if not SHEET_NAME.fullmatch(fx.sheet):
            checker.error(f"fixture sheet name {fx.sheet} must be one word, so TESTs can pass it bare", line)
        if (fx.file, fx.sheet) in seen:
            checker.error(f"{fx.file} has two FIXTURE lines for sheet {fx.sheet}", line)
        seen.add((fx.file, fx.sheet))
        letters = [col for col, _ in fx.headers]
        expected = [column_letter(i) for i in range(len(letters))]
        if letters != expected:
            checker.error(f"FIXTURE cols must run {', '.join(expected)} in order, got {', '.join(letters)}", line)
        names = Counter(norm(h) for _, h in fx.headers)
        for header, n in names.items():
            if n > 1:
                checker.error(f'FIXTURE sheet {fx.sheet} has the header "{header}" twice', line)
        if fx.rows == 0 and fx.rules:
            checker.error("rows=0 creates a header-only sheet, so it can't have data rules", line)
        ruled = Counter(rule.col for rule in fx.rules)
        for col, n in ruled.items():
            if n > 1:
                checker.error(f"FIXTURE data has two rules for column {col}", line)
        for rule in fx.rules:
            if rule.col not in letters:
                checker.error(f"FIXTURE data uses column {rule.col}, which is not in cols=", line)
            if rule.blank > fx.rows:
                checker.error(f"{rule.col}: +blank({rule.blank}) is more than rows={fx.rows}", line)
            if rule.kind == "values":
                total = sum(count for _, count in rule.args)
                if total != fx.rows:
                    checker.error(f"{rule.col}: values(...) counts add up to {total}, but rows={fx.rows}", line)
            elif rule.kind in ("int", "num", "dates"):
                lo, hi = rule.args
                if (lo > hi) if rule.kind == "dates" else (float(lo) > float(hi)):
                    checker.error(f"{rule.col}: {rule.kind}() starts after it ends", line)
            elif rule.kind == "text" and int(rule.args[0]) < 1:
                checker.error(f"{rule.col}: text() needs a length of 1 or more", line)


# ---------------------------------------------------------------- BUILD checks

def check_build(spec):
    problems = []

    def error(message, line):
        problems.append(Problem(line, "error", message))

    counts = Counter(k.key for k in spec.keys)
    for key in ("NAME", "TARGET", "ENTRY", "MODULE", "SIG", "FLOW"):
        if not counts[key]:
            error(f"BUILD is missing {key}", spec.line)
    for key in ("NAME", "TARGET", "ENTRY"):
        for extra in spec.lines(key)[1:]:
            error(f"{key} appears more than once", extra.line)
    for k in spec.lines("NAME")[:1]:
        if not PASCAL.fullmatch(k.value):
            error(f"NAME must be PascalCase, got '{k.value}'", k.line)
    for k in spec.lines("TARGET")[:1]:
        if k.value not in ("xlsm", "xlam"):
            error(f"TARGET must be xlsm or xlam, got '{k.value}'", k.line)
    for k in spec.lines("ENTRY")[:1]:
        if k.value != "modMain.Run":
            error(f"ENTRY must be modMain.Run, got '{k.value}'", k.line)
    modules = {k.value for k in spec.lines("MODULE")}
    for k in spec.lines("MODULE"):
        if not MODULE_NAME.fullmatch(k.value):
            error(f"MODULE must be mod + PascalCase, got '{k.value}'", k.line)
    if spec.lines("MODULE") and "modMain" not in modules:
        error("a BUILD always has MODULE modMain", spec.lines("MODULE")[0].line)
    forms = {f.name for f in spec.forms}
    for form in spec.forms:
        if not FORM_NAME.fullmatch(form.name):
            error(f"FORM name must be frm + PascalCase, got '{form.name}'", form.line)
        for control, ctype in form.controls:
            if ctype not in CONTROL_TYPES:
                error(f"unknown control type {ctype} (use {', '.join(sorted(CONTROL_TYPES))})", form.line)
    sigs = {}
    for sig in spec.sigs:
        full = f"{sig.module}.{sig.proc}"
        if sig.module not in modules | forms:
            error(f"SIG {full}: {sig.module} is not a MODULE or FORM", sig.line)
        if full in sigs:
            error(f"SIG {full} appears twice", sig.line)
        sigs[full] = sig
        for decl in sig.params:
            if decl.type not in TYPES:
                error(f"unknown type {decl.type}", sig.line)
        if sig.ret and sig.ret not in TYPES:
            error(f"unknown type {sig.ret}", sig.line)
    for needed in ("modMain.Run", "modMain.RunHeadless"):
        if spec.sigs and needed not in sigs:
            error(f"a BUILD needs SIG {needed}(...)", spec.sigs[0].line)
    for line, names in spec.flows:
        for name in names:
            if name not in sigs and not any(s.proc == name for s in spec.sigs):
                error(f"FLOW names {name or '(nothing)'}, which has no SIG", line)

    # the BUILD's own TESTs call modMain.RunHeadless
    headless = sigs.get("modMain.RunHeadless")
    files = {}
    for fx in spec.fixtures:
        files.setdefault(fx.file, {})[fx.sheet] = fx
    written = set()
    for task in spec.tasks:
        for node in task.all_ops():
            for header in re.findall(r'SET_VALUE\(CELL\(\$\w+, 1, [^)]*\), "([^"]+)"\)', node.text):
                written.add(norm(header))
    shim = TaskChecker(spec)
    shim.kind = "Function" if headless and headless.ret else "Sub"
    shim.sheet_ops = True  # BUILD headers come from several procedures: only check sheets loosely
    flags = []
    for test in spec.tests:
        shim.line = test.line
        if test.fixture == "target":
            shim.error("fixture=target is not allowed: use a named synthetic FIXTURE (rule 18)")
            continue
        if test.procedure != "modMain.RunHeadless":
            shim.error(f"a BUILD TEST calls modMain.RunHeadless, not {test.procedure}")
        if headless and len(test.args) != len(headless.params):
            shim.error(f"TEST passes {len(test.args)} arguments, but modMain.RunHeadless takes "
                       f"{len(headless.params)}")
        sheets = files.get(test.fixture)
        if sheets is None:
            shim.error(f"TEST uses fixture {test.fixture}, but no FIXTURE line builds it")
        headers = None
        if sheets is not None:
            all_headers = {norm(h) for fx in sheets.values() for _, h in fx.headers} | written
            headers = {name: all_headers for name in sheets}
        flags.append(check_assertion(shim, test, sheets, headers))
    first = spec.tests[0].line if spec.tests else spec.line
    if len(spec.tests) < 2:
        shim.error("a BUILD needs at least two TEST lines (rule 5)", first)
    if spec.tests and not any(f["error_count_0"] for f in flags):
        shim.error("no TEST checks ERROR_COUNT = 0 (rule 5)", first)
    if spec.tests and not any(f["pre"] for f in flags):
        shim.error("no TEST guards against collateral changes with a PRE_ measure (rule 5)", first)
    if spec.tests and not any(t.runs >= 2 for t in spec.tests):
        shim.error("no TEST uses runs=2 to prove re-runs are safe (rule 16)", first)
    check_fixtures(shim, spec.fixtures)
    problems += shim.problems

    # one TASK per SIG, matching it
    matched = set()
    for task in spec.tasks:
        problems += task.problems
        full = f"{task.value('MODULE')}.{task.value('PROC')}"
        sig = sigs.get(full)
        if sig is None:
            error(f"TASK {full} has no matching SIG", task.line)
        elif full in matched:
            error(f"two TASK sections for {full}", task.line)
        else:
            matched.add(full)
            kind = "Function" if sig.ret else "Sub"
            if task.value("KIND") != kind:
                error(f"TASK {full} must be KIND {kind} to match its SIG", task.line)
            if sig.ret and task.value("RET") != sig.ret:
                error(f"TASK {full} must have RET {sig.ret} to match its SIG", task.line)
            if [(d.name, d.type) for d in task.args] != [(d.name, d.type) for d in sig.params]:
                error(f"TASK {full} ARGS don't match its SIG", task.line)
        checker = TaskChecker(task, in_build=True, build_fixtures=spec.fixtures)
        problems += checker.run()
        for line, target, nargs in checker.calls:
            if target not in sigs:
                problems.append(Problem(line, "error", f"CALL {target} has no SIG in this BUILD"))
            elif nargs != len(sigs[target].params):
                problems.append(Problem(line, "error", f"CALL {target} passes {nargs} arguments, but its SIG "
                                                       f"takes {len(sigs[target].params)}"))
    for full, sig in sigs.items():
        if full not in matched:
            error(f"SIG {full} has no TASK section after ---", sig.line)
    return problems


# ---------------------------------------------------------------- entry points

def _sorted_unique(problems):
    unique = []
    for p in sorted(problems, key=lambda p: (p.line, p.severity != "error")):
        if p not in unique:
            unique.append(p)
    return unique


def validate(spec):
    """Stage 1 problems for a parsed spec, sorted by line."""
    problems = list(spec.problems)
    if spec.kind == "TASK":
        problems += TaskChecker(spec).run()
    elif spec.kind == "BUILD":
        problems += check_build(spec)
    return _sorted_unique(problems)


def grammar_and_type_problems(text, known):
    """Stages 2 and 3: the formal grammar, then the type checker on every TASK it parsed.
    Problems on lines that already have a stage 1 error are left out."""
    bad_lines = {p.line for p in known if p.severity == "error"}
    spec, found = grammar_parse(text)
    out = [Problem(line, "error", f"grammar: {msg}") for line, msg in found if line not in bad_lines]
    if found or spec is None:
        return out
    tasks = [spec] if spec["kind"] == "TASK" else spec.get("tasks", []) if spec["kind"] == "BUILD" else []
    for task in tasks:
        if not task or task.get("kind") != "TASK":
            continue
        try:
            typed = TypeChecker(task).run()
        except Exception as exc:  # a crash here is a bug in the type checker, never in the spec
            if not bad_lines:
                out.append(Problem(task["keys"][0]["line"] if task["keys"] else 1, "error",
                                   f"the type checker could not read this spec ({type(exc).__name__}: {exc})"))
            continue
        out += [Problem(line, "error", f"type: {msg}") for line, msg in typed if line not in bad_lines]
    return out


def validate_text(text):
    spec = parse_spec(text)
    problems = validate(spec)
    if spec.kind in ("TASK", "BUILD"):
        problems = _sorted_unique(problems + grammar_and_type_problems(text, problems))
    return problems


def validate_file(path):
    try:
        with open(path, encoding="utf-8") as f:
            text = f.read()
    except (OSError, UnicodeDecodeError) as exc:
        return [Problem(1, "error", f"cannot read file ({type(exc).__name__}: {exc})")]
    return validate_text(text)


def expand(paths):
    """Expand wildcards ourselves too, for shells (like Windows cmd) that don't."""
    out = []
    for path in paths:
        if any(ch in path for ch in "*?[") and not os.path.exists(path):
            matches = sorted(glob.glob(path))
            out += matches or [path]
        else:
            out.append(path)
    return out


def main(argv):
    if not argv:
        print("usage: python compiler/validator.py specs/*.vcl", file=sys.stderr)
        print("error: no spec files given", file=sys.stderr)
        return 1
    errors = warnings = 0
    paths = expand(argv)
    for path in paths:
        for p in validate_file(path):
            print(f"{path}:{p.line}: {p.severity}: {p.message}")
            errors += p.severity == "error"
            warnings += p.severity == "warning"
    plural = lambda n, word: f"{n} {word}{'' if n == 1 else 's'}"
    print(f"{plural(len(paths), 'file')} checked: {plural(errors, 'error')}, {plural(warnings, 'warning')}")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
