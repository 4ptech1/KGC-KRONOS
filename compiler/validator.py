"""Validate VCL 0.1 to 0.5 specs against vcl/spec.md.

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

from parser import Problem, Token, parse_spec, split_top, tokenize
from vcl_fixtures import parse_fixture, sheet_rows, stale_rows
from vcl_parse import OPTION_WORDS
from vcl_parse import parse as grammar_parse
from vcl_types import Checker as TypeChecker

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "vcl"))
from semantics import safe_name, sheet_name_ok  # noqa: E402
from vcl_formula import formula_problems, sheet_names  # noqa: E402

TYPES = {"Long", "Double", "String", "Boolean", "Date", "Variant", "Range", "Worksheet", "Workbook",
         "Dictionary", "Collection"}
CONTROL_TYPES = {"TextBox", "CommandButton", "Label", "ListBox", "ComboBox", "CheckBox", "OptionButton"}
PERF_WORDS = {"screen_off", "calc_manual", "events_off"}
ALLOW_WORDS = {"file_delete", "file_overwrite", "shell", "http"}
FILL_NAMES = ("none", "black", "white", "red", "green", "blue", "yellow", "magenta", "cyan")
YES_NO = ("yes", "no")
MAX_TASK_OPS = 40

STRUCT = {"LOOP_ROWS", "LOOP_COLS", "FOR_EACH", "IF", "ELSE", "SELECT_CASE", "CASE", "CASE_ELSE", "EXIT_IF",
          "SKIP_IF", "EXIT_LOOP_IF"}
LOOPS = {"LOOP_ROWS", "LOOP_COLS", "FOR_EACH"}

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
    "DEDUPE": (1, 1, {"cols": "cols", "header": YES_NO, "into": "var"}, ("cols",)),   # into= (VCL 0.5, D55)
    "SET_VALUE": (2, 2, {}, ()),
    "SET_FORMULA": (2, 2, {"style": ("a1", "r1c1")}, ()),
    "FORMAT": (1, 1, {"numfmt": "string", "bold": YES_NO, "fill": "fill", "width": "width"}, ()),
    "CLEAR": (1, 1, {"what": ("all", "contents", "formats")}, ()),
    "COPY_ROW": (4, 4, {"cols": "expr"}, ("cols",)),
    "INSERT_ROW": (2, 2, {}, ()),
    "DELETE_ROW": (2, 2, {}, ()),
    "ADD_SHEET": (1, 1, {"name": "expr", "after": "expr", "into": "var"}, ()),
    "COPY_SHEET": (1, 1, {"name": "expr", "after": "expr", "into": "var"}, ()),
    "OPEN_BOOK": (1, 1, {"readonly": YES_NO, "into": "var"}, ("into",)),
    "SAVE_AS": (2, 2, {"format": ("xlsx", "xlsm", "csv")}, ()),
    "CLOSE_BOOK": (1, 1, {"save": YES_NO}, ()),
    "LIST_FILES": (2, 2, {"into": "var"}, ("into",)),
    "READ_TEXT": (2, 2, {"into": "var"}, ("into",)),    # D45: (path, String VAR, into=status)
    "WRITE_TEXT": (2, 2, {"into": "var"}, ("into",)),   # D45: new files only, so no mode=
    "CALL": (1, None, {"into": "var"}, ()),   # into= (VCL 0.5): a Function's result (D49)
    "LOG": (1, 1, {}, ()),
    "MSG": (1, 1, {}, ()),
    "RETURN": (0, 1, {}, ()),
    "WRITE_ROW": (3, None, {"col": "expr"}, ()),   # VCL 0.2
    "MAKE_FOLDER": (1, 1, {"into": "var"}, ("into",)),   # VCL 0.3 (E9)
    "COPY_FILE": (2, 2, {"into": "var"}, ("into",)),
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
    # VCL 0.2 (docs/vep-0.2-core.md, PR A)
    "SAME": (2, 2, {}), "CLEAN_TEXT": (1, 1, {}), "IS_ERROR": (1, 1, {}), "PARSE_DATE": (2, 2, {}),
    "IIF": (3, 3, {}), "MOD": (2, 2, {}),
    # VCL 0.3 (docs/vep-0.3-file-ops.md, Phase 1)
    "SAFE_NAME": (1, 1, {}), "PATH_JOIN": (2, None, {}), "FOLDER_EXISTS": (1, 1, {}),
    # VCL 0.4 (docs/vep-0.4-rev3-rest.md, Phase 4)
    "SHEET_EXISTS": (2, 2, {}),
    # VCL 0.5 (docs/vep-0.5-build.md, Phase 4)
    "ROW_OF": (1, 1, {}),
}
VCL_02_FUNCS = {"SAME", "CLEAN_TEXT", "IS_ERROR", "PARSE_DATE", "IIF", "MOD"}
VCL_02_OPS = {"SKIP_IF", "EXIT_LOOP_IF", "WRITE_ROW"}
VCL_03_FUNCS = {"SAFE_NAME", "PATH_JOIN", "FOLDER_EXISTS"}
VCL_03_OPS = {"MAKE_FOLDER", "COPY_FILE"}
VCL_04_FUNCS = {"SHEET_EXISTS"}
VCL_05_FUNCS = {"ROW_OF"}   # D55
VCL_04_OPS = {"LIST_FILES", "READ_TEXT", "WRITE_TEXT"}   # D44, D45: S24, S25 and the path rules of VCL 0.3
VERSION_RANK = {"VCL 0.1": 1, "VCL 0.2": 2, "VCL 0.3": 3, "VCL 0.4": 4, "VCL 0.5": 5}
# MODIFIES (VCL 0.2): what each write op changes, and so which MODIFIES line it needs
SHEET_FUNCS = {"CELL", "ROW_RANGE", "USED_RANGE", "RANGE"}
ROW_OPS = {"INSERT_ROW", "DELETE_ROW", "SORT", "DEDUPE", "FILTER"}   # these move or hide rows
MODIFIES_RE = re.compile(r"\$(?P<arg>[A-Za-z_]\w*)\s*(?:\[(?P<cols>.*)\]|\s(?P<fills>fills)|\s(?P<files>files)"
                         r"|\s(?P<sheets>sheets))?")
DATE_FORMATS = ("dd/mm/yyyy", "dd-mm-yyyy", "dd.mm.yyyy", "mm/dd/yyyy", "yyyy-mm-dd")  # as vcl/semantics.py
# second argument is a "quoted" header, ref, sheet name or date format
STRING_SECOND = {"COL", "RANGE", "SHEET", "PARSE_DATE"}
CONSTANTS = {"THIS_BOOK", "ACTIVE_BOOK", "ACTIVE_SHEET", "TODAY", "NOW", "TRUE", "FALSE", "NOTHING"}
FOR_EACH_IN = {"SHEETS", "CELLS", "KEYS", "ITEMS"}
BINARY = {"=", "<>", "<", ">", "<=", ">=", "+", "-", "*", "/", "&"}
VBA_LEAKS = ["Dim ", "Set ", "End If", ".Value", "Cells(", "Range(", "Next ", "Wend"]

# TEST assertion functions and the kind of each argument
ASSERT_FUNCS = {"COUNTIF", "COUNTIFS", "COUNTA", "SUM", "SUMIFS", "VALUE", "ROWS_USED", "SHEET_EXISTS",
                "FILE_EXISTS", "FILL_COUNT", "FILE_COUNT", "VISIBLE_COUNT", "FORMULA", "FILE_TEXT"}
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


class ListLit:
    """[a, b, ...]: a list (VCL 0.4), so far only as DEDUPE cols= (D38). elements: one items list each."""

    def __init__(self, items):
        self.elements = split_on(items, ",")


class Unbalanced(Exception):
    pass


OPENER = {")": "(", "]": "["}


def build(tokens):
    """Group tokens into Calls and Groups. Returns (items, error message)."""
    pos = 0

    def sequence(closer):
        nonlocal pos
        items = []
        while pos < len(tokens):
            tok = tokens[pos]
            if tok.kind == "op" and tok.text in (")", "]"):
                if tok.text == closer:
                    return items
                if closer:
                    raise Unbalanced(f"unbalanced brackets: '{OPENER[closer]}' without '{closer}'")
                raise Unbalanced(f"unbalanced brackets: '{tok.text}' without '{OPENER[tok.text]}'")
            pos += 1
            if tok.kind == "op" and tok.text == "(":
                prev = items[-1] if items else None
                inner = sequence(")")
                pos += 1  # the closing ")"
                if isinstance(prev, Token) and prev.kind == "name" and not tok.spaced:
                    items[-1] = Call(prev, split_args(inner))
                else:
                    items.append(Group(inner))
            elif tok.kind == "op" and tok.text == "[":
                items.append(ListLit(sequence("]")))
                pos += 1  # the closing "]"
            else:
                items.append(tok)
        if closer:
            raise Unbalanced(f"unbalanced brackets: '{OPENER[closer]}' without '{closer}'")
        return items

    try:
        return sequence(None), ""
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


def target_sheet(items):
    """The $sheet argument a write target is on: $ws, CELL/ROW_RANGE/USED_RANGE/RANGE($ws, ...), or
    FIND(<one of those>, ...) written in place (D39). None when it can't be told."""
    if len(items) == 1 and is_tok(items[0], "arg"):
        return items[0].text[1:]
    if len(items) == 1 and isinstance(items[0], Call) and items[0].name in SHEET_FUNCS | {"FIND"}:
        first = items[0].positional[:1]
        if first and items[0].name == "FIND":
            return target_sheet(first[0]) if isinstance(first[0][0], Call) else None
        if first and len(first[0]) == 1 and is_tok(first[0][0], "arg"):
            return first[0][0].text[1:]
    return None


def is_tok(item, kind, text=None):
    return isinstance(item, Token) and item.kind == kind and (text is None or item.text == text)


def show(items):
    out = []
    for item in items:
        if isinstance(item, Call):
            out.append(f"{item.name}(...)")
        elif isinstance(item, Group):
            out.append("(...)")
        elif isinstance(item, ListLit):
            out.append("[...]")
        else:
            out.append(item.text)
    return " ".join(out)


def blank_strings(text):
    return re.sub(r'"(?:[^"]|"")*"', '""', text)


# ---------------------------------------------------------------- TASK checks

class TaskChecker:
    """Checks one TASK spec. in_build=True for a TASK section inside a BUILD."""

    def __init__(self, spec, in_build=False, build_fixtures=(), build_folders=()):
        self.spec, self.in_build = spec, in_build
        self.fixtures = list(build_fixtures) + spec.fixtures
        self.folders = list(build_folders) + spec.folders
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
        self.modifies_col_lines = {}   # (arg, header) -> the MODIFIES line that names it
        self.header_copies = []    # (source arg, target arg) row 1 copied
        self.calls = []            # (line, "module.Proc", number of args, [arg items], into= items or None)
        self.sheet_ops = False     # ADD_SHEET / COPY_SHEET can create sheets the fixture lacks
        self.formula_sheets = []   # (line, sheet name) named in a SET_FORMULA formula (D51)
        self.named_adds = set()    # casefolded name= of every ADD_SHEET / COPY_SHEET with a quoted name
        self.pre_conditions = []   # [(arg, header)] when every PRE condition is COL($x, "h") > 0
        self.pre_evaluable = True
        self.modifies = {}         # arg -> {"whole": line, "cols": line, "fills": line}
        self.modifies_cols = {}    # arg -> {normalised header: as written} of its MODIFIES $arg["..."] line (D49)
        self.writes = []           # (line, op, sheet arg or None, what it changes)
        self.pre_roots = []        # [arg] checked with FOLDER_EXISTS($arg) in PRE (VCL 0.3)
        self.modifies_files = {}   # arg -> line of MODIFIES $arg files (VCL 0.3)
        self.modifies_sheets = {}  # arg -> line of MODIFIES $arg sheets (VCL 0.4)
        self.sheet_adds = []       # (line, op, $book, None or "*" for COPY_SHEET's own workbook) (VCL 0.4)
        self.new_sheets = {}       # Worksheet VAR -> line: set by ADD_SHEET or COPY_SHEET into= (D40, D41)
        self.var_sets = []         # (line, VAR) set by LET or any other into=
        self.row_order_ops = {}    # "FILTER" / "SORT" -> first line: a spec may not use both (D42)
        self.row_moves = {}        # SET_FORMULA and the ops that move rows -> first line (D43)
        self.file_paths = []       # (line, op, "from" or "to", root arg or None, bare root?)

    def error(self, message, line=None):
        self.problems.append(Problem(line or self.line, "error", message))

    def at_least(self, version):
        """True when the spec's version line is this version or a later one."""
        return VERSION_RANK.get(self.spec.version, 1) >= VERSION_RANK[version]

    def warning(self, message, line=None):
        self.problems.append(Problem(line or self.line, "warning", message))

    def run(self):
        spec = self.spec
        self.check_keys()
        self.collect_loopvars()
        self.check_decls()
        for pre in spec.lines("PRE"):
            self.check_pre(pre)
        for mod in spec.lines("MODIFIES"):
            self.check_modifies(mod)
        self.mode = "op"
        self.walk(spec.ops, None, [])
        if spec.lines("MODIFIES"):
            self.check_writes_are_declared()
        elif self.at_least("VCL 0.2") and self.writes and not self.in_build:
            self.warning("the spec writes to sheets but has no MODIFIES lines, so no TEST checks that nothing "
                         "else changed: declare each sheet it changes with MODIFIES", self.writes[0][0])
        elif self.at_least("VCL 0.5") and self.writes:   # D49: a BUILD's TESTs check them all
            self.error("a TASK of a BUILD declares every sheet it changes with MODIFIES, so the BUILD's TESTs can "
                       "check that nothing else changed (D49)", self.writes[0][0])
        if self.calls and not self.in_build and self.at_least("VCL 0.5"):
            self.error("CALL runs another procedure of the same BUILD: a TASK spec has none to call (D49)",
                       self.calls[0][0])
        self.check_file_ops()
        self.check_sheet_adds()
        moves = sorted(op for op in self.row_moves if op != "SET_FORMULA")
        if "SET_FORMULA" in self.row_moves and moves:
            self.error(f"a spec may not both SET_FORMULA and {' or '.join(moves)}: Excel rewrites formula references "
                       "when rows move, which a TEST can't prove (D43)", self.row_moves["SET_FORMULA"])
        if len(self.row_order_ops) == 2:
            self.error("a spec may not both FILTER and SORT: how Excel sorts rows that FILTER hid can't be proven "
                       "by a TEST (D42); sort in one spec and filter in another", self.row_order_ops["FILTER"])
        made = {(arg, norm(header)) for arg, header in self.header_writes}   # the spec writes them to row 1
        for line, arg, header in self.op_cols:
            if (arg, norm(header)) in self.pre_cols or ((arg, norm(header)) in made and self.at_least("VCL 0.5")):
                continue   # D56: a header the spec adds itself can't be in PRE; its TESTs prove it
            hint = ": a header the spec writes to row 1 itself needs VCL 0.5" if (arg, norm(header)) in made else ""
            self.error(f'header "{header}" on ${arg} is used but not checked in PRE (rule 14){hint}', line)
        for (arg, key), line in self.modifies_col_lines.items():
            if (arg, key) in made and (arg, key) not in self.pre_cols and self.at_least("VCL 0.5"):
                self.error(f'MODIFIES ${arg}["{self.modifies_cols[arg][key]}"]: the spec adds that header itself, so '
                           f'the column is not there before the run and no TEST can check it; declare the whole '
                           f'sheet with MODIFIES ${arg} (D56)', line)
        op_count = sum(1 for _ in spec.all_ops())
        if op_count > MAX_TASK_OPS:
            first = spec.lines("OP")[0].line
            self.warning(f"TASK has {op_count} OP lines (more than {MAX_TASK_OPS}): keep a TASK small, or ask "
                         "which part to do first (rule 9)", first)
        check_fixtures(self, spec.fixtures)
        check_folders(self, spec.folders, self.fixtures)
        self.check_tests()
        if not self.in_build:
            self.check_fixtures_can_catch_mistakes()
            self.check_folders_can_catch_mistakes()
        return self.problems

    def check_folders_can_catch_mistakes(self):
        """Rule 19 (warning, VCL 0.3): COPY_FILE, and READ_TEXT (VCL 0.4, D46), need a locked file in a
        FIXTURE folder, or no TEST shows what the macro does with a file that is open in another program."""
        reads = [(line, name) for line, name, role, _, _ in self.file_paths if role == "from"]
        if reads and self.spec.tests and not any(ff.locked for ff in self.folders):
            line, name = reads[0]
            self.warning(f"the spec {'copies' if name == 'COPY_FILE' else 'reads'} files, but no FIXTURE folder has "
                         "a locked= file, so no TEST shows what happens to a file that is open elsewhere (rule 19)",
                         line)

    def check_fixtures_can_catch_mistakes(self):
        """Rule 19 (warnings): fixtures that let the TESTs tell a right macro from a wrong one."""
        spec = self.spec
        if not spec.tests or not spec.fixtures:
            return
        first = spec.tests[0].line
        texts = [blank_strings(node.text) for node in spec.all_ops()]
        rules = [rule for fx in spec.fixtures for rule in fx.rules]
        # stale rows: sheets the macro rebuilds (CLEAR of the used range) or whose fill it owns
        owned = {}
        for text in texts:
            for arg in re.findall(r"CLEAR\(USED_RANGE\(\$(\w+)\)", text):
                owned.setdefault(arg, "rows")
            if re.search(r"\bfill=none\b", text):
                for arg in re.findall(r"FORMAT\((?:ROW_RANGE|CELL|USED_RANGE)\(\$(\w+)", text):
                    owned.setdefault(arg, "fills")
        params = [d.name for d in spec.args]
        warned = set()
        for test in spec.tests:
            if not self.passes_pre(test):
                continue  # a wrong-sheet TEST changes nothing, so its sheets need no stale rows
            for name, value in zip(params, test.args):
                fx = next((f for f in spec.fixtures if f.file == test.fixture and f.sheet == value), None)
                if name in owned and fx is not None and not fx.stale and fx.sheet not in warned:
                    warned.add(fx.sheet)
                    what = "old rows are cleared" if owned[name] == "rows" else "old fills are cleared"
                    self.warning(f"FIXTURE sheet {fx.sheet} has no stale=<n> rows, so no TEST proves {what} "
                                 "(rule 19)", fx.line)
        if any(re.search(r"\bTODAY\b", t) for t in texts) and not any(r.kind == "reldates" for r in rules):
            self.warning("the spec uses TODAY, but no FIXTURE column uses reldates(lo, hi): fixed dates drift "
                         "as TODAY moves (rule 19)", first)
        repeats = any(r.dupes for r in rules) or any(
            count >= 2 for r in rules if r.kind == "values" for text, count in r.args if text)
        if any("DICT_ADD(" in t for t in texts) and not repeats:
            self.warning("the spec adds to a dictionary, but no FIXTURE column repeats a key: add +dupes(n) "
                         "so the TESTs see repeated keys (rule 19)", first)
        if any(re.search(r"<=|>=|(?<![<>])[<>](?![<>=])", t) for t in texts) and not any(r.boundary for r in rules):
            self.warning("the spec compares with <, <=, > or >=, but no FIXTURE column has +boundary(...): "
                         "put the values that sit exactly on each comparison in the first rows (rule 19)", first)
        self.check_skip_if_has_blanks()

    def check_skip_if_has_blanks(self):
        """Rule 19 (warning): SKIP_IF IS_EMPTY(CELL($x, r, c)) needs a blank cell in that FIXTURE column,
        or no TEST shows the skip matters (D22)."""
        spec = self.spec
        raw = [node.text for node in spec.all_ops()]
        headers = {}  # column variable -> (arg, header), from LET(c, COL($x, "Header"))
        for text in raw:
            m = re.fullmatch(r'LET\((\w+),\s*COL\(\$(\w+),\s*"((?:[^"]|"")*)"\)\)', text)
            if m:
                headers[m.group(1)] = (m.group(2), m.group(3).replace('""', '"'))
        params = [d.name for d in spec.args]
        warned = set()
        for text in raw:
            m = re.fullmatch(r"SKIP_IF IS_EMPTY\(CELL\(\$(\w+),\s*\w+,\s*(\w+)\)\)", text)
            if not m or headers.get(m.group(2), (None,))[0] != m.group(1):
                continue
            arg, header = headers[m.group(2)]
            for test in spec.tests:
                if not self.passes_pre(test):
                    continue
                sheet = dict(zip(params, test.args)).get(arg)
                fx = next((f for f in spec.fixtures if f.file == test.fixture and f.sheet == sheet), None)
                letter = next((c for c, h in fx.headers if norm(h) == norm(header)), None) if fx else None
                if letter is None or (fx.sheet, letter) in warned:
                    continue
                rule = next((r for r in fx.rules if r.col == letter), None)
                has_blank = (rule is None or rule.blank or any(v == "" for _, v in rule.boundary)
                             or (rule.kind == "values" and any(t == "" for t, _ in rule.args)))
                if not has_blank and fx.rows > 0:
                    warned.add((fx.sheet, letter))
                    self.warning(f'SKIP_IF skips blank "{header}" cells, but FIXTURE sheet {fx.sheet} has none, '
                                 "so no TEST shows the skip matters: add +blank(<n>) (rule 19)", fx.line)

    def passes_pre(self, test):
        """False only when PRE is a list of COL(...) > 0 and FOLDER_EXISTS($root) checks and this TEST's
        sheets miss a header or its roots are not folders of the fixture."""
        if not (self.pre_evaluable and (self.pre_conditions or self.pre_roots)):
            return True
        argmap = dict(zip([d.name for d in self.spec.args], test.args))
        folders = {f.folder for f in self.folders if f.file == test.fixture}
        if any(argmap.get(root) not in folders for root in self.pre_roots):
            return False
        for arg, header in self.pre_conditions:
            fx = next((f for f in self.fixtures if f.file == test.fixture and f.sheet == argmap.get(arg)), None)
            if fx is not None and norm(header) not in {norm(h) for _, h in fx.headers}:
                return False
        return True

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

    # -- MODIFIES (VCL 0.2)

    def check_modifies(self, mod):
        self.line = mod.line
        if not self.at_least("VCL 0.2"):
            self.error("MODIFIES is VCL 0.2: start the spec with 'VCL 0.2' to use it")
        m = MODIFIES_RE.fullmatch(mod.value)
        if not m:
            forms = ('$<sheet>, $<sheet>["<header>", ...], $<sheet> fills, $<root> files or $<book> sheets'
                     if self.at_least("VCL 0.4") else
                     '$<sheet>, $<sheet>["<header>", ...], $<sheet> fills or $<root> files' if self.at_least("VCL 0.3")
                     else '$<sheet>, $<sheet>["<header>", ...] or $<sheet> fills')
            self.error(f"MODIFIES needs {forms}, got '{mod.value}'")
            return
        arg = m.group("arg")
        if m.group("files"):   # VCL 0.3 (E10)
            if not self.at_least("VCL 0.3"):
                self.error("MODIFIES ... files is VCL 0.3: start the spec with 'VCL 0.3' to use it")
            if self.args.get(arg) != "String":
                self.error(f"MODIFIES ${arg} files: ${arg} must be a String argument (a root folder)")
            elif arg in self.modifies_files:
                self.error(f"MODIFIES ${arg} files appears twice: use one line")
            else:
                self.modifies_files[arg] = mod.line
            return
        if m.group("sheets"):   # VCL 0.4 (D40)
            if not self.at_least("VCL 0.4"):
                self.error("MODIFIES ... sheets is VCL 0.4: start the spec with 'VCL 0.4' to use it")
            if self.args.get(arg) != "Workbook":
                self.error(f"MODIFIES ${arg} sheets: ${arg} must be a Workbook argument")
            elif arg in self.modifies_sheets:
                self.error(f"MODIFIES ${arg} sheets appears twice: use one line")
            else:
                self.modifies_sheets[arg] = mod.line
            return
        if self.args.get(arg) == "String":
            self.error(f"MODIFIES ${arg}: a String argument is a root folder, so write MODIFIES ${arg} files")
            return
        if self.args.get(arg) == "Workbook":
            self.error(f"MODIFIES ${arg}: a Workbook argument gets new sheets only, so write MODIFIES ${arg} sheets")
            return
        if self.args.get(arg) != "Worksheet":
            self.error(f"MODIFIES ${arg}: ${arg} is not a Worksheet argument")
            return
        kind = "cols" if m.group("cols") is not None else "fills" if m.group("fills") else "whole"
        if kind == "cols":
            headers = []
            for part in split_top(m.group("cols"), ","):
                tokens, _ = tokenize(part.strip())
                if len(tokens) != 1 or tokens[0].kind != "string" or not tokens[0].text[1:-1].strip():
                    self.error(f'MODIFIES ${arg}[...] needs "quoted" header names, got \'{part.strip()}\'')
                    continue
                header = tokens[0].text[1:-1].replace('""', '"')
                if norm(header) in headers:
                    self.error(f'MODIFIES ${arg} lists "{header}" twice')
                headers.append(norm(header))
                self.modifies_cols.setdefault(arg, {}).setdefault(norm(header), header)
                self.modifies_col_lines.setdefault((arg, norm(header)), mod.line)
        seen = self.modifies.setdefault(arg, {})
        if kind in seen:
            what = {"whole": "the whole sheet", "cols": "columns", "fills": "fills"}[kind]
            self.error(f"MODIFIES ${arg} declares {what} twice: use one line")
        elif "whole" in seen or (kind == "whole" and seen):
            self.error(f"MODIFIES ${arg} (the whole sheet) already covers its other MODIFIES lines")
        seen[kind] = mod.line

    def record_write(self, name, positional, options):
        """Remember which sheet argument a write op changes, and how (MODIFIES checks it)."""
        if name not in ("SET_VALUE", "WRITE_ROW", "FORMAT", "CLEAR", "COPY_ROW", "INSERT_ROW", "DELETE_ROW",
                        "SORT", "WRITE_ARRAY", "SET_FORMULA", "DEDUPE", "FILTER"):
            return
        if name == "WRITE_ARRAY":
            target = dict((k.text, v) for k, v in options).get("to")
        elif name == "COPY_ROW":
            target = positional[2] if len(positional) > 2 else None
        else:
            target = positional[0] if positional else None
        if target is None:
            return
        if name == "SET_VALUE" and len(target) == 1 and isinstance(target[0], Call) and target[0].name == "ARR":
            return   # an array element, not a cell
        base = target[0].positional[0] if (len(target) == 1 and isinstance(target[0], Call)
                                           and target[0].name in SHEET_FUNCS and target[0].positional) else target
        if len(base) == 1 and is_tok(base[0], "name") and base[0].text in self.new_sheets:
            return   # a sheet ADD_SHEET made: MODIFIES $book sheets covers all of it (D40)
        arg = target_sheet(target)
        keys = {k.text for k, _ in options}
        what = dict((k.text, show(v)) for k, v in options).get("what", "all")
        if name in ROW_OPS:
            change = "rows"
        elif name == "FORMAT":
            change = "fills" if keys == {"fill"} else "formats"
        elif name == "CLEAR":
            change = "values" if what == "contents" else "formats"
        else:
            change = "values"
        self.writes.append((self.line, name, arg, change))

    def check_sheet_adds(self):
        """VCL 0.4 (D40, D41): ADD_SHEET and COPY_SHEET add to a Workbook argument declared with
        MODIFIES $book sheets."""
        added = set()
        for line, name, book in self.sheet_adds:
            if book == "*":   # COPY_SHEET: the copy goes into the TEST's workbook, every Workbook argument
                if not self.modifies_sheets:
                    why = "" if self.at_least("VCL 0.4") else ", which is VCL 0.4: start the spec with 'VCL 0.4'"
                    self.error(f"{name} adds a sheet to the workbook of the sheet it copies: give the spec a Workbook "
                               f"argument $book and declare MODIFIES $book sheets{why}", line)
                added |= set(self.modifies_sheets)
            elif book is None or self.args.get(book) != "Workbook":
                self.error(f"{name} needs a Workbook argument ($book) as its first argument, so MODIFIES can "
                           "name the workbook it changes", line)
            elif book not in self.modifies_sheets:
                why = "" if self.at_least("VCL 0.4") else ", which is VCL 0.4: start the spec with 'VCL 0.4'"
                self.error(f"{name} adds a sheet to ${book}: declare MODIFIES ${book} sheets{why}", line)
            else:
                added.add(book)
        for book, line in self.modifies_sheets.items():
            if book not in added:
                self.warning(f"MODIFIES ${book} sheets, but no op adds a sheet to ${book}", line)
        for line, var in self.var_sets:
            if var in self.new_sheets:
                self.error(f"{var} holds the sheet made on line {self.new_sheets[var]}, so only ADD_SHEET or "
                           "COPY_SHEET may set it: writes through it are checked as writes to a new sheet", line)

    def check_writes_are_declared(self):
        """MODIFIES: every write names a declared sheet in a way its lines allow; every declared
        sheet is written somewhere."""
        needs = {"values": ("cols",), "fills": ("fills",), "formats": (), "rows": ()}
        line_for = {"values": '${0}["<header>", ...]', "fills": "${0} fills", "formats": "${0}", "rows": "${0}"}
        what_for = {"values": "values", "fills": "fills", "formats": "formats or fills", "rows": "rows"}
        written = set()
        for line, name, arg, change in self.writes:
            if arg is None:
                self.error(f"{name}: MODIFIES can't tell which sheet this changes; write through CELL, "
                           "ROW_RANGE, USED_RANGE, RANGE or FIND of a $sheet argument", line)
                continue
            written.add(arg)
            declared = self.modifies.get(arg)
            if declared is None:
                self.error(f"{name} changes ${arg}, but no MODIFIES line declares ${arg}", line)
            elif "whole" not in declared and not any(k in declared for k in needs[change]):
                self.error(f"{name} changes {what_for[change]} on ${arg}: declare "
                           f"MODIFIES {line_for[change].format(arg)}", line)
        for arg, kinds in self.modifies.items():
            if arg not in written:
                self.warning(f"MODIFIES ${arg}, but no op writes to ${arg}", min(kinds.values()))

    def check_file_ops(self):
        """VCL 0.3 (E9, E10): every file op path is PATH_JOIN($root, ...); every root it adds to is
        declared with MODIFIES $root files; every root it uses is checked with FOLDER_EXISTS in PRE."""
        used, written = {}, set()
        for line, name, role, root, bare in self.file_paths:
            if root is None or self.args.get(root) != "String":
                continue  # not a path: the type checker says why (T4)
            used.setdefault(root, line)
            if bare and role != "list":
                what = ("the folder to make" if name == "MAKE_FOLDER" else "the file to read" if name == "READ_TEXT"
                        else "the file to write" if name == "WRITE_TEXT" else f"the file to copy {role}")
                self.error(f"{name}: {what} must be PATH_JOIN(${root}, ...), not the root ${root} itself, "
                           "which must already exist", line)
            if role == "to":
                written.add(root)
                if root not in self.modifies_files:
                    self.error(f"{name} adds to ${root}, but no line declares MODIFIES ${root} files", line)
        for root, line in used.items():
            if root not in self.pre_roots:
                self.error(f"file ops use ${root}, but PRE doesn't check FOLDER_EXISTS(${root}) (rule 14)", line)
        for root, line in self.modifies_files.items():
            if root not in written:
                self.warning(f"MODIFIES ${root} files, but no op adds to ${root}", line)

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
            if (len(condition) == 1 and isinstance(condition[0], Call) and condition[0].name == "FOLDER_EXISTS"
                    and len(condition[0].args) == 1 and condition[0].args[0][0] is None):
                root = condition[0].args[0][1]
                if len(root) == 1 and is_tok(root[0], "arg"):
                    self.pre_roots.append(root[0].text[1:])
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
            if name in ("SKIP_IF", "EXIT_LOOP_IF") and not any(a["name"] in LOOPS for a in ancestors):
                self.error(f"{name} is only allowed inside LOOP_ROWS, LOOP_COLS or FOR_EACH", node.line)
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
        if name in VCL_02_OPS and not self.at_least("VCL 0.2"):
            self.error(f"{name} is VCL 0.2: start the spec with 'VCL 0.2' to use it")
        if name in VCL_03_OPS and not self.at_least("VCL 0.3"):
            self.error(f"{name} is VCL 0.3: start the spec with 'VCL 0.3' to use it")
        if name in VCL_04_OPS and not self.at_least("VCL 0.4"):
            self.error(f"{name} is VCL 0.4: start the spec with 'VCL 0.4' to use it")
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
        elif name in ("EXIT_IF", "SKIP_IF", "EXIT_LOOP_IF"):
            self.check_expr(rest, f"{name} condition")

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
        if name in ("CALL", "DEDUPE") and "into" in seen and not self.at_least("VCL 0.5"):
            self.error(f"{name} into= is VCL 0.5: start the spec with 'VCL 0.5' to use it")
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
                    into = dict((k.text, v) for k, v in options).get("into")
                    self.calls.append((self.line, target, count - 1, positional[1:], into))
            elif name in ("SET_FORMULA", "LIST_FILES") and i == 1:
                if not (len(items) == 1 and is_tok(items[0], "string")):
                    self.error(f'{name} needs a "quoted" {"formula" if name == "SET_FORMULA" else "pattern"} '
                               "as its second argument")
                elif name == "LIST_FILES":   # D44: a name pattern with * and ? only
                    pattern = items[0].text[1:-1].replace('""', '"')
                    bad = sorted(set(ch for ch in pattern if ch in '\\/:"<>|'))
                    if not pattern or bad:
                        self.error(f"LIST_FILES pattern \"{pattern}\" must be a file name pattern with * and ? "
                                   f"(not empty, none of \\ / : \" < > |)")
                elif name == "SET_FORMULA":   # D43: Excel 2016 only, this workbook only
                    r1c1 = any(k.text == "style" and show(v) == "r1c1" for k, v in call.options)
                    formula = items[0].text[1:-1].replace('""', '"')
                    for problem in formula_problems(formula, r1c1):
                        self.error(f"SET_FORMULA: {problem}")
                    self.formula_sheets += [(self.line, sheet) for sheet in sheet_names(formula)]
            else:
                self.check_expr(items, f"{name} argument")
        if name == "RETURN":
            if self.kind == "Function" and count != 1:
                self.error("a Function's RETURN needs a value: RETURN(<expr>)")
            if self.kind == "Sub" and count:
                self.error("a Sub can't RETURN a value")
        if name in ("ADD_SHEET", "COPY_SHEET"):
            self.sheet_ops = True
        if name in ("FILTER", "SORT"):
            self.row_order_ops.setdefault(name, self.line)
        if name in ("SET_FORMULA", "INSERT_ROW", "DELETE_ROW", "DEDUPE", "SORT"):
            self.row_moves.setdefault(name, self.line)
        if name == "LET" and positional and len(positional[0]) == 1 and is_tok(positional[0][0], "name"):
            self.var_sets.append((self.line, positional[0][0].text))
        for key, items in options:
            if (key.text == "into" and name not in ("ADD_SHEET", "COPY_SHEET") and len(items) == 1
                    and is_tok(items[0], "name")):
                self.var_sets.append((self.line, items[0].text))
        if name in ("ADD_SHEET", "COPY_SHEET") and self.at_least("VCL 0.2"):   # they change the sheet list (D40, D41)
            book = positional[0] if positional else []
            if name == "COPY_SHEET":
                self.sheet_adds.append((self.line, name, "*"))   # the workbook of the sheet it copies
            else:
                self.sheet_adds.append((self.line, name, book[0].text[1:] if len(book) == 1 and is_tok(book[0], "arg")
                                        else None))
            given = dict((k.text, v) for k, v in options)
            if len(given.get("name", [])) == 1 and is_tok(given["name"][0], "string"):
                self.named_adds.add(given["name"][0].text[1:-1].replace('""', '"').casefold())
            into = given.get("into")
            if into and len(into) == 1 and is_tok(into[0], "name"):
                self.new_sheets.setdefault(into[0].text, self.line)
            if "name" not in given:
                why = ("Excel's own names (Sheet4) depend on the PC's language" if name == "ADD_SHEET"
                       else "Excel's own name for a copy (Template (2)) is not one a TEST can rely on")
                self.error(f"{name} needs name=: {why}")
            elif len(given["name"]) == 1 and is_tok(given["name"][0], "string") \
                    and not sheet_name_ok(given["name"][0].text[1:-1].replace('""', '"')):
                self.error(f"{name} name={given['name'][0].text} is not a valid sheet name (S23: 1 to 31 "
                           "characters, no spaces at the ends, none of \\ / ? * [ ] :, not History)")
        self.record_write(name, positional, options)
        if name in VCL_03_OPS:
            roles = ["to"] if name == "MAKE_FOLDER" else ["from", "to"]
            for role, items in zip(roles, positional):
                self.file_paths.append((self.line, name, role) + path_root(items))
        if name == "LIST_FILES" and positional:   # D44: it reads a folder, the root itself or one below it
            self.file_paths.append((self.line, name, "list") + path_root(positional[0]))
        if name in ("READ_TEXT", "WRITE_TEXT") and positional:   # D45: a file below a root
            self.file_paths.append((self.line, name, "from" if name == "READ_TEXT" else "to")
                                   + path_root(positional[0]))
        if name == "READ_TEXT" and len(positional) > 1:   # D45: the text goes into a declared VAR
            given = positional[1]
            if not (len(given) == 1 and is_tok(given[0], "name") and given[0].text in self.vars):
                self.error(f"READ_TEXT needs a declared String VAR second, which gets the file's text, "
                           f"got '{show(given)}'")
            else:
                self.var_sets.append((self.line, given[0].text))
        if name == "SET_VALUE" and len(positional) == 2:
            target, value = positional
            arg = self.row1_arg(target)
            if arg and len(value) == 1 and is_tok(value[0], "string"):
                self.header_writes.append((arg, value[0].text[1:-1]))
        if name == "WRITE_ROW" and len(positional) >= 3:
            ws, row = positional[:2]
            if len(ws) == 1 and is_tok(ws[0], "arg") and len(row) == 1 and is_tok(row[0], "number", "1"):
                for value in positional[2:]:
                    if len(value) == 1 and is_tok(value[0], "string"):
                        self.header_writes.append((ws[0].text[1:], value[0].text[1:-1]))
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
        elif kind == "cols" and len(items) == 1 and isinstance(items[0], ListLit):   # D38
            if not self.at_least("VCL 0.4"):
                self.error(f"{name} cols=[...] is VCL 0.4: start the spec with 'VCL 0.4' to use it")
            for part in items[0].elements:
                self.check_expr(part, f"{name} cols= list value")
        elif kind == "cols":
            parts = split_on(items, ",")
            if len(parts) > 1 and self.at_least("VCL 0.4"):
                self.error(f"{name} cols={text.replace(' , ', ',')} is ambiguous (G3): write "
                           f"cols=[{', '.join(show(p) for p in parts)}]")
            for part in parts:
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
        if isinstance(item, ListLit):
            self.error("a list [...] is only allowed as DEDUPE cols=")
        elif isinstance(item, Group):
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
        elif item.kind == "qsheet":
            self.error(f"{item.text}: text goes in \"double quotes\"; 'single quotes' are only for a sheet name "
                       "in a TEST assertion")
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
        if name in VCL_02_FUNCS and not self.at_least("VCL 0.2"):
            self.error(f"{name} is VCL 0.2: start the spec with 'VCL 0.2' to use it")
        if name in VCL_03_FUNCS and not self.at_least("VCL 0.3"):
            self.error(f"{name} is VCL 0.3: start the spec with 'VCL 0.3' to use it")
        if name in VCL_04_FUNCS and not self.at_least("VCL 0.4"):
            self.error(f"{name} is VCL 0.4: start the spec with 'VCL 0.4' to use it")
        if name in VCL_05_FUNCS and not self.at_least("VCL 0.5"):
            self.error(f"{name} is VCL 0.5: start the spec with 'VCL 0.5' to use it")
        fewest, most, allowed = VALUE_FUNCS[name]
        positional, options = self.call_parts(call)
        if len(positional) < fewest or (most is not None and len(positional) > most):
            want = f"{fewest} or more" if most is None else str(fewest)
            self.error(f"{name} takes {want} argument{'s' if fewest > 1 or most is None else ''}, got {len(positional)}")
        for key, items in options:
            if key.text not in allowed:
                self.error(f"{name} has no option {key.text}=")
            else:
                self.check_choice(name, key.text, items, allowed[key.text])
        for i, items in enumerate(positional):
            if name in STRING_SECOND and i == 1:
                if not (len(items) == 1 and is_tok(items[0], "string")):
                    self.error(f'{name} needs a "quoted" text as its second argument')
                elif name == "PARSE_DATE" and items[0].text[1:-1] not in DATE_FORMATS:
                    self.error(f"PARSE_DATE format {items[0].text} is not one of "
                               + ", ".join(f'"{f}"' for f in DATE_FORMATS))
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
        files, folders = {}, {}
        for fx in self.fixtures:
            files.setdefault(fx.file, {})[fx.sheet] = fx
        for ff in self.folders:
            folders.setdefault(ff.file, {})[ff.folder] = ff
        flags = []
        for test in tests:
            self.line = test.line
            flags.append(self.check_test(test, files, folders))
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
        if spec.lines("PRE") and self.pre_evaluable and (self.pre_conditions or self.pre_roots) and flags \
                and not any(f["passes_pre"] for f in flags if f["passes_pre"] is not None):
            what = "TEST misses a header or a folder" if self.pre_roots else "TEST's sheets miss a header"
            self.error(f"no TEST passes PRE: every {what} that PRE checks", first)

    def check_test(self, test, files, folders):
        flags = {"error_count_0": False, "wrong_sheet": False, "pre": False, "passes_pre": None}
        spec = self.spec
        if test.fixture == "target":
            self.error("fixture=target is not allowed: use a named synthetic FIXTURE (rule 18)")
            return flags
        sheets = files.get(test.fixture)
        dirs = folders.get(test.fixture, {})
        if sheets is None and dirs:
            sheets = {}   # a fixture of folders only
        if sheets is None:
            self.error(f"TEST uses fixture {test.fixture}, but no FIXTURE line builds it")
        else:   # D51: in Excel a formula naming a missing sheet opens a file dialog and links to a file
            have = {name.casefold() for name in sheets} | self.named_adds
            for line, sheet in self.formula_sheets:
                if sheet.casefold() not in have:
                    self.error(f"SET_FORMULA names the sheet {sheet}, which the TEST workbook {test.fixture} "
                               "doesn't have: in Excel the formula would ask for a file (D51)")
        proc = spec.value("PROC")
        if not self.in_build and test.procedure != proc:
            self.error(f"TEST calls {test.procedure}, but PROC is {proc}")
        argmap, rootmap = {}, {}
        if len(test.args) != len(spec.args):
            self.error(f"TEST passes {len(test.args)} argument{'s' if len(test.args) != 1 else ''}, "
                       f"but ARGS has {len(spec.args)}")
        elif sheets is not None:
            for decl, value in zip(spec.args, test.args):
                if decl.type == "Workbook":   # VCL 0.4 (D40): the workbook is the TEST's fixture file
                    if value != test.fixture:
                        self.error(f"TEST passes {value} for ${decl.name}, a Workbook: pass this TEST's "
                                   f"fixture file, {test.fixture}")
                    continue
                if decl.type == "String" and SHEET_NAME.fullmatch(value):   # a bare name is a FIXTURE folder (F6)
                    if value in dirs:
                        rootmap[decl.name] = value
                    else:
                        self.error(f"TEST passes {value} for ${decl.name}, but {test.fixture} has no folder {value}")
                    continue
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
        if self.pre_evaluable and (self.pre_conditions or self.pre_roots) and sheets is not None:
            missing = [(arg, h) for arg, h in self.pre_conditions
                       if arg in argmap and norm(h) not in {norm(x) for _, x in sheets[argmap[arg]].headers}]
            no_folder = [root for root in self.pre_roots if root not in rootmap]
            if all(arg in argmap for arg, _ in self.pre_conditions):
                flags["passes_pre"] = not missing and not no_folder
                if (missing or no_folder) and self.kind == "Function" and not result["result_minus_1"]:
                    if missing:
                        arg, h = missing[0]
                        why = f'sheet {argmap[arg]} has no "{h}" header'
                    else:
                        given = dict(zip([d.name for d in spec.args], test.args)).get(no_folder[0])
                        why = f"${no_folder[0]} is {given}, which is not a folder of {test.fixture}"
                    self.error(f"TEST fails PRE ({why}) but doesn't expect RESULT = -1")
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
        if isinstance(item, ListLit):
            self.error("a list [...] can't be used in an assertion")
        elif isinstance(item, Group):
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
        elif item.kind == "qsheet":
            self.error(f"{item.text}: a sheet name goes before ! or inside ROWS_USED, like {item.text}!A1")
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
        if name == "FILE_COUNT" and not self.c.at_least("VCL 0.3"):
            self.error("FILE_COUNT is VCL 0.3: start the spec with 'VCL 0.3' to use it")
        if name in ("VISIBLE_COUNT", "FORMULA", "FILE_TEXT") and not self.c.at_least("VCL 0.4"):
            self.error(f"{name} is VCL 0.4: start the spec with 'VCL 0.4' to use it")
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
                     "ROWS_USED": ["sheet"], "SHEET_EXISTS": ["string"], "FILE_EXISTS": ["path"],
                     "FILL_COUNT": ["range", "hex"], "FILE_COUNT": ["path"], "VISIBLE_COUNT": ["range"],
                     "FORMULA": ["cell"], "FILE_TEXT": ["path"]}[name]
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
            if single.quoted and self.quoted(single):
                return
            self.range(single)
            if kind == "cell":
                ref = single.ref
                if (single.header is not None and not single.row) or \
                        (single.header is None and (":" in ref or not re.search(r"\d", ref))):
                    self.error(f"{fname} needs one cell, like Sheet![Header]2 or Sheet!B2, got '{text}'")
        elif kind == "crit":
            if not (is_tok(single, "string") or is_tok(single, "number")
                    or (len(items) == 2 and is_tok(items[0], "op", "-") and is_tok(items[1], "number"))):
                self.error(f"{fname} criteria must be a \"quoted\" text or a number, got '{text}'")
        elif kind == "sheet":
            if is_tok(single, "qsheet"):
                if not self.quoted(single):
                    self.sheet(single.sheet)
            elif not is_tok(single, "name"):
                self.error(f"{fname} needs a sheet name, got '{text}'")
            else:
                self.sheet(single.text)
        elif kind == "string" or (kind == "path" and not self.c.at_least("VCL 0.3")):
            if not is_tok(single, "string"):
                self.error(f'{fname} needs a "quoted" name, got \'{text}\'')
        elif kind == "path":   # VCL 0.3: below one of this TEST's FIXTURE folders (F7)
            folders = {f.folder.casefold() for f in self.c.folders if f.file == self.test.fixture}
            parts = single.text[1:-1].replace('""', '"').split("\\") if is_tok(single, "string") else []
            if not parts or parts[0].casefold() not in folders:
                self.error(f'{fname} needs a "quoted" path that starts with a FIXTURE folder of {self.test.fixture}, '
                           f'like "Jobs\\J-101\\r1.pdf", got \'{text}\'')
            elif any(safe_name(part) != part for part in parts):
                self.error(f"{fname}: {text} is not a safe path (every part must be a SAFE_NAME)")
        elif kind == "hex":
            if not (is_tok(single, "hex") and HEX_COLOUR.fullmatch(single.text)):
                self.error(f"{fname} needs a #RRGGBB colour, got '{text}'")

    def quoted(self, tok):
        """D57: a sheet name in 'quotes' is VCL 0.5, only for a name that isn't one word, and one Excel allows.
        True when it is refused."""
        plain = f"{tok.sheet}!{tok.ref}" if tok.kind == "range" else tok.sheet
        if not self.c.at_least("VCL 0.5"):
            self.error(f"{tok.text}: a sheet name in quotes is VCL 0.5: start the spec with 'VCL 0.5' to use it")
        elif SHEET_NAME.fullmatch(tok.sheet):
            self.error(f"{tok.text}: {tok.sheet} is one word, so write {plain} without quotes")
        elif not sheet_name_ok(tok.sheet):
            self.error(f"{tok.text}: '{tok.sheet}' can't be a sheet name (S23)")
        else:
            return False
        return True

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


def path_root(items):
    """(arg, True) for a bare $arg, (arg, False) for PATH_JOIN($arg, ...), else (None, False)."""
    if len(items) == 1 and is_tok(items[0], "arg"):
        return items[0].text[1:], True
    if len(items) == 1 and isinstance(items[0], Call) and items[0].name == "PATH_JOIN":
        first = items[0].positional[:1]
        if first and len(first[0]) == 1 and is_tok(first[0][0], "arg"):
            return first[0][0].text[1:], False
    return None, False


def check_folders(checker, folders, fixtures):
    """VCL 0.3 folder FIXTUREs (F6): one-word folder names, safe file paths, locked files listed."""
    seen = set()
    sheets = {(fx.file, fx.sheet) for fx in fixtures}
    for ff in folders:
        line = ff.line
        if not checker.at_least("VCL 0.3"):
            checker.error("folder FIXTUREs are VCL 0.3: start the spec with 'VCL 0.3' to use them", line)
        if not FIXTURE_FILE.fullmatch(ff.file):
            checker.error(f"fixture file {ff.file} must be named like snake_case_01.xlsx (rule 12)", line)
        if not SHEET_NAME.fullmatch(ff.folder):
            checker.error(f"fixture folder name {ff.folder} must be one word, so TESTs can pass it bare", line)
        if (ff.file, ff.folder) in seen:
            checker.error(f"{ff.file} has two FIXTURE lines for folder {ff.folder}", line)
        if (ff.file, ff.folder) in sheets:
            checker.error(f"{ff.file} has a sheet and a folder named {ff.folder}: a TEST could not tell them apart",
                          line)
        seen.add((ff.file, ff.folder))
        paths, shown = [], {}
        for path, _, text in ff.files:
            if text is not None and not checker.at_least("VCL 0.4"):
                checker.error(f'FIXTURE files: {path}:"..." (a text file) is VCL 0.4: start the spec with '
                              "'VCL 0.4' to use it", line)
            elif text is not None and any(not 32 <= ord(ch) <= 126 for ch in text):
                checker.error(f"FIXTURE files: the text of {path} must be one line of printable ASCII (D45)", line)
            parts = path.split("\\")
            bad = next((part for part in parts if safe_name(part) != part), None)
            if bad is not None:
                checker.error(f"FIXTURE files: {path} is not a safe path: {bad or 'an empty part'} is not "
                              "a safe name (see SAFE_NAME)", line)
            if path.casefold() in paths:
                checker.error(f"FIXTURE folder {ff.folder} lists {path} twice", line)
            paths.append(path.casefold())
            shown.setdefault(path.casefold(), path)
        for path in paths:
            if any(other.startswith(path + "\\") for other in paths):
                checker.error(f"FIXTURE folder {ff.folder}: {shown[path]} is a file, so nothing can be below it",
                              line)
        for path in ff.locked:
            if path.casefold() not in paths:
                checker.error(f"FIXTURE locked= {path} is not one of the folder's files=", line)
        if len({p.casefold() for p in ff.locked}) != len(ff.locked):
            checker.error(f"FIXTURE folder {ff.folder} lists a locked file twice", line)


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
            check_modifiers(checker, rule, fx.rows, line)
            if rule.kind == "values":
                total = sum(count for _, count in rule.args)
                if total != fx.rows:
                    checker.error(f"{rule.col}: values(...) counts add up to {total}, but rows={fx.rows}", line)
            elif rule.kind in ("int", "num", "dates", "reldates"):
                lo, hi = rule.args
                if (lo > hi) if rule.kind == "dates" else (float(lo) > float(hi)):
                    checker.error(f"{rule.col}: {rule.kind}() starts after it ends", line)
            elif rule.kind == "text" and int(rule.args[0]) < 1:
                checker.error(f"{rule.col}: text() needs a length of 1 or more", line)
        if fx.stale and fx.rows and fx.stale > fx.rows:
            checker.error(f"stale={fx.stale} is more than rows={fx.rows}", line)


MODIFIER_ORDER = ["boundary", "blank", "dupes"]
BOUNDARY_KIND = {"values": "text", "text": "text", "seq": "number", "int": "number", "num": "number",
                 "reldates": "number", "dates": "date"}


def check_modifiers(checker, rule, rows, line):
    """+boundary, +blank and +dupes: each at most once, in that order, with values the rule can hold."""
    col = rule.col
    for name in sorted(set(rule.mods)):
        if rule.mods.count(name) > 1:
            checker.error(f"{col}: +{name}() is written twice", line)
    positions = [MODIFIER_ORDER.index(name) for name in rule.mods]
    if positions != sorted(positions):
        checker.error(f"{col}: write the modifiers in this order: +boundary, +blank, +dupes", line)
    top = len(rule.boundary)
    if top > rows:
        checker.error(f"{col}: +boundary has {top} values, but rows={rows}", line)
    if rule.blank > max(rows - top, 0):
        checker.error(f"{col}: +blank({rule.blank}) is more than the {max(rows - top, 0)} rows below the "
                      "boundary rows", line)
    if "dupes" in rule.mods and rule.dupes < 1:
        checker.error(f"{col}: +dupes() needs 1 or more", line)
    want = BOUNDARY_KIND.get(rule.kind)
    for kind, value in rule.boundary:
        if kind == "text" and value == "":
            continue  # "" is a blank cell, allowed for every rule
        if want and kind != want:
            checker.error(f"{col}: +boundary value {value!r} doesn't fit {rule.kind}() (use "
                          f"{ {'text': 'quoted text', 'number': 'numbers', 'date': 'yyyy-mm-dd dates'}[want] })",
                          line)
        elif rule.kind == "reldates" and not re.fullmatch(r"-?\d+", value):
            checker.error(f"{col}: reldates() +boundary values are whole numbers of days, got {value}", line)
    if rule.kind == "values":
        listed = Counter()
        for text, count in rule.args:
            listed[text] += count
        wanted = Counter(value for kind, value in rule.boundary if kind == "text")
        for value, n in wanted.items():
            if n > listed[value]:
                checker.error(f'{col}: +boundary uses "{value}" {n} times, but values(...) has it '
                              f"{listed[value]} times", line)


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
    for ff in spec.folders:
        files.setdefault(ff.file, {})
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
    check_folders(shim, spec.folders, spec.fixtures)
    problems += shim.problems

    # one TASK per SIG, matching it
    matched = set()
    checkers = {}
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
        checker = TaskChecker(task, in_build=True, build_fixtures=spec.fixtures, build_folders=spec.folders)
        problems += checker.run()
        checkers.setdefault(full, checker)
        for line, target, nargs, _, _ in checker.calls:
            if target not in sigs:
                problems.append(Problem(line, "error", f"CALL {target} has no SIG in this BUILD"))
            elif nargs != len(sigs[target].params):
                problems.append(Problem(line, "error", f"CALL {target} passes {nargs} arguments, but its SIG "
                                                       f"takes {len(sigs[target].params)}"))
    for full, sig in sigs.items():
        if full not in matched:
            error(f"SIG {full} has no TASK section after ---", sig.line)
    if VERSION_RANK.get(spec.version, 1) >= VERSION_RANK["VCL 0.5"]:
        found, passed_on = check_build_05(spec, sigs, checkers)
        # a MODIFIES line a CALL uses is not unused: the callee makes that change (D49)
        unused = {f"MODIFIES ${arg}{kind}, but no op " for checker in checkers.values()
                  for arg in passed_on.get(id(checker), ()) for kind in ("", " files", " sheets")}
        problems = [p for p in problems
                    if not (p.severity == "warning" and any(p.message.startswith(u) for u in unused))] + found
    return problems


OBJECT_TYPES = ("Worksheet", "Workbook", "Range")
RUN_STEPS = {"CALL", "LET", "MSG", "IF", "ELSE"}


def declared_changes(checker, arg):
    """What a procedure's MODIFIES lines let it change through its argument arg: {(kind, header)}."""
    kinds = checker.modifies.get(arg, {})
    out = {(kind, None) for kind in ("whole", "fills") if kind in kinds}
    if "cols" in kinds:
        out |= {("cols", h) for h in checker.modifies_cols.get(arg, {}).values()}
    if arg in checker.modifies_files:
        out.add(("files", None))
    if arg in checker.modifies_sheets:
        out.add(("sheets", None))
    return out


def allows(checker, arg, change):
    """True when the caller's MODIFIES lines on arg allow a change a callee declares (D49)."""
    kind, header = change
    kinds = checker.modifies.get(arg, {})
    if kind in ("whole", "cols", "fills") and "whole" in kinds:
        return True
    if kind == "cols":
        return "cols" in kinds and norm(header) in checker.modifies_cols.get(arg, {})
    if kind == "fills":
        return "fills" in kinds
    if kind == "files":
        return arg in checker.modifies_files
    if kind == "sheets":
        return arg in checker.modifies_sheets
    return False


def modifies_form(arg, change):
    kind, header = change
    return {"whole": f"MODIFIES ${arg}", "cols": f'MODIFIES ${arg}["{header}"]', "fills": f"MODIFIES ${arg} fills",
            "files": f"MODIFIES ${arg} files", "sheets": f"MODIFIES ${arg} sheets"}[kind]


def check_build_05(spec, sigs, checkers):
    """VCL 0.5 (D49): the rules that let a BUILD's TESTs prove the whole tool. One module per BUILD,
    procedures called by name, sheets passed as $arguments so MODIFIES follows them, no recursion,
    FLOW as the CALL edges, and a modMain.Run that has nothing left to test."""
    problems = []

    def error(message, line):
        problems.append(Problem(line, "error", message))

    for form in spec.forms:
        error(f"FORM {form.name}: a UserForm can't be pasted in as text, so FORM waits until after March 2027 "
              "(D49); modMain.Run starts the tool", form.line)
    for k in spec.lines("TARGET")[:1]:
        if k.value == "xlam":
            error("TARGET xlam waits (D49): an add-in's THIS_BOOK is the add-in, not the data; use TARGET xlsm",
                  k.line)
    homes = {}
    for sig in spec.sigs:
        if homes.setdefault(sig.proc, sig.module) != sig.module:
            error(f"SIG {sig.module}.{sig.proc}: {homes[sig.proc]}.{sig.proc} has the same name; a BUILD calls its "
                  "procedures by name, so each name is used once (D49)", sig.line)

    edges = {}   # caller -> {callee: first line}
    passed_on = {}   # id(caller's checker) -> its $arguments a callee may change
    for caller, checker in checkers.items():
        from_run = caller == "modMain.Run"
        for line, target, nargs, args, into in checker.calls:
            sig = sigs.get(target)
            if sig is None or nargs != len(sig.params):
                continue   # reported with the BUILD's other CALL checks
            edges.setdefault(caller, {}).setdefault(target, line)
            if from_run and target != "modMain.RunHeadless":
                error(f"modMain.Run only CALLs modMain.RunHeadless, got {target} (D49)", line)
            if into is not None:
                v = into[0].text if len(into) == 1 and is_tok(into[0], "name") else None
                if not sig.ret:
                    error(f"CALL {target} into=: {target} is a Sub, so it gives no result", line)
                elif v in checker.vars and checker.vars[v] != sig.ret:
                    error(f"CALL {target} gives a {sig.ret}, but VAR {v} is {checker.vars[v]}", line)
            callee = checkers.get(target)
            for decl, items in zip(sig.params, args):
                own = items[0].text[1:] if len(items) == 1 and is_tok(items[0], "arg") else None
                changes = declared_changes(callee, decl.name) if callee else set()
                if from_run:
                    by_name = (len(items) == 1 and isinstance(items[0], Call) and items[0].name == "SHEET"
                               and len(items[0].positional) == 2 and len(items[0].positional[0]) == 1
                               and is_tok(items[0].positional[0][0], "name", "THIS_BOOK")
                               and len(items[0].positional[1]) == 1 and is_tok(items[0].positional[1][0], "string"))
                    if decl.type == "Worksheet" and not by_name:
                        error(f'CALL {target}: modMain.Run passes ${decl.name} as SHEET(THIS_BOOK, "<real name>") (D49)',
                              line)
                    elif decl.type == "Workbook" and not (len(items) == 1 and is_tok(items[0], "name", "THIS_BOOK")):
                        error(f"CALL {target}: modMain.Run passes ${decl.name} as THIS_BOOK (D49)", line)
                    elif decl.type == "Range":
                        error(f"CALL {target}: modMain.Run can't pass a Range; let RunHeadless take a sheet (D49)",
                              line)
                    continue
                if (decl.type in OBJECT_TYPES or changes) and own is None:
                    error(f"CALL {target}: pass ${decl.name} as one of this procedure's own $arguments, written as "
                          "$name, so MODIFIES can follow it (D49)", line)
                    continue
                if own is not None and decl.type in OBJECT_TYPES and checker.args.get(own) != decl.type:
                    error(f"CALL {target}: ${decl.name} is a {decl.type}, but ${own} is a {checker.args.get(own)}", line)
                    continue
                if changes:
                    passed_on.setdefault(id(checker), set()).add(own)
                for change in sorted(changes, key=str):
                    if not allows(checker, own, change):
                        error(f"CALL {target} may change ${decl.name} ({modifies_form(decl.name, change)}), so this "
                              f"procedure needs {modifies_form(own, change)} too (D49)", line)

    # no recursion: every loop has a bound (CLAUDE.md)
    state = {}

    def visit(node, path):
        state[node] = "open"
        for nxt, line in edges.get(node, {}).items():
            if state.get(nxt) == "open":
                cycle = path[path.index(nxt):] + [nxt]
                error(f"CALL makes a cycle ({' -> '.join(cycle)}): a BUILD has no recursion (D49)", line)
            elif nxt not in state:
                visit(nxt, path + [nxt])
        state[node] = "done"

    for node in sorted(edges):
        if node not in state:
            visit(node, [node])

    # FLOW lists exactly the CALL edges
    def full_name(name):
        if name in sigs:
            return name
        found = [f for f, s in sigs.items() if s.proc == name]
        return found[0] if len(found) == 1 else name

    flowed = set()
    for line, names in spec.flows:
        for a, b in zip(names, names[1:]):
            a, b = full_name(a), full_name(b)
            flowed.add((a, b))
            if b not in edges.get(a, {}):
                error(f"FLOW says {a} -> {b}, but {a} has no CALL to {b} (D49)", line)
    for caller, targets in edges.items():
        for target, line in targets.items():
            if (caller, target) not in flowed:
                error(f"{caller} CALLs {target}, but no FLOW line shows {caller} -> {target} (D49)", line)

    # modMain.Run starts the tool and has nothing left to test
    run_sig = sigs.get("modMain.Run")
    if run_sig is not None and (run_sig.params or run_sig.ret):
        error("SIG modMain.Run() takes no arguments and gives no result: it starts the tool (D49)", run_sig.line)
    for task in spec.tasks:
        if f"{task.value('MODULE')}.{task.value('PROC')}" != "modMain.Run":
            continue
        for node in task.all_ops():
            words = re.findall(r"[A-Z_]+", node.text)
            word = words[0] if words else ""
            if word == "IF" and " THEN " in node.text:
                after = re.findall(r"[A-Z_]+", node.text.split(" THEN ", 1)[1])
                word = after[0] if after else ""
            if word not in RUN_STEPS:
                error(f"modMain.Run only CALLs modMain.RunHeadless, LETs, IFs and MSGs, so nothing in it goes "
                      f"untested (D49); move {word or 'this step'} into RunHeadless", node.line)
    return problems, passed_on


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


def grammar_and_type_problems(text, known, sigs=None):
    """Stages 2 and 3: the formal grammar (and building every FIXTURE from it), then the type
    checker on every TASK it parsed. Problems on lines that already have a stage 1 error are left out."""
    bad_lines = {p.line for p in known if p.severity == "error"}
    spec, found = grammar_parse(text)
    out = [Problem(line, "error", f"grammar: {msg}") for line, msg in found if line not in bad_lines]
    if found or spec is None:
        return out
    tasks = [spec] if spec["kind"] == "TASK" else spec.get("tasks", []) if spec["kind"] == "BUILD" else []
    for k in spec.get("keys", []) + [k for t in tasks if t and t is not spec for k in t.get("keys", [])]:
        if k["key"] == "FIXTURE" and k["line"] not in bad_lines and k["tree"].data != "folder_fixture":
            try:
                fx = parse_fixture(k["tree"])
                sheet_rows(fx)
                stale_rows(fx)
            except ValueError as exc:
                out.append(Problem(k["line"], "error", f"fixture: {exc}"))
    for task in tasks:
        if not task or task.get("kind") != "TASK":
            continue
        try:
            typed = TypeChecker(task, sigs).run()
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
        sigs = {f"{s.module}.{s.proc}": ([(d.name, d.type) for d in s.params], s.ret) for s in spec.sigs}
        problems = _sorted_unique(problems + grammar_and_type_problems(text, problems, sigs))
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
