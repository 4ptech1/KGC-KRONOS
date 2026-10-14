"""Starter checker for VCL 0.1 (revision 3) TASK specs.

Usage:  python compiler/vcl_check.py specs/*.vcl
Checks structure and the spec-writing rules. It is a starting point for validator.py,
not a full parser: it does not type-check expressions.
"""
import re
import sys

ORDER = ["PROC", "KIND", "MODULE", "ARGS", "RET", "VAR", "PRE", "OP", "PERF", "ERR", "ALLOW", "FIXTURE", "TEST"]
STRUCT = {"LOOP_ROWS", "LOOP_COLS", "FOR_EACH", "IF", "ELSE", "SELECT_CASE", "CASE", "CASE_ELSE", "EXIT_IF"}
ACTIONS = {"LET", "READ_ARRAY", "WRITE_ARRAY", "DICT_ADD", "DICT_GET", "SORT", "FILTER", "DEDUPE", "SET_VALUE",
           "SET_FORMULA", "FORMAT", "CLEAR", "COPY_ROW", "INSERT_ROW", "DELETE_ROW", "ADD_SHEET", "COPY_SHEET",
           "OPEN_BOOK", "SAVE_AS", "CLOSE_BOOK", "LIST_FILES", "READ_TEXT", "WRITE_TEXT", "CALL", "LOG", "MSG",
           "RETURN", "NOTE"}
VALUES = {"LASTROW", "LASTCOL", "USED_RANGE", "COL", "CELL", "ROW_RANGE", "RANGE", "SHEET", "FIND", "ROWS", "COLS",
          "ARR", "COUNT", "EXISTS", "LEN", "TRIM", "UPPER", "LOWER", "LEFT", "RIGHT", "IS_EMPTY", "IS_NUMBER",
          "IS_DATE", "IS_NOTHING", "FILE_NAME", "SHEETS", "CELLS", "KEYS", "ITEMS"}
WORDS = {"THIS_BOOK", "ACTIVE_BOOK", "ACTIVE_SHEET", "TODAY", "NOW", "TRUE", "FALSE", "NOTHING", "AND", "OR",
         "NOT", "THEN"}
TYPES = {"Long", "Double", "String", "Boolean", "Date", "Variant", "Range", "Worksheet", "Workbook",
         "Dictionary", "Collection"}
OPTION_WORDS = {"up", "down", "right", "left", "asc", "desc", "yes", "no", "all", "contents", "formats", "none",
                "sum", "count", "keep", "overwrite", "whole", "part", "new", "append"}
VBA_LEAKS = ["Dim ", "Set ", "End If", ".Value", "Cells(", "Range(", "Next ", "Wend"]


def check(path):
    problems = []
    err = problems.append
    lines = open(path, encoding="utf-8").read().splitlines()
    if lines[:2] != ["VCL 0.1", "TASK"]:
        return ["first two lines must be 'VCL 0.1' and 'TASK'"]
    body = [l for l in lines[2:] if l.strip()]
    keys = [l.split()[0] for l in body if not l.startswith("#")]
    for k in keys:
        if k not in ORDER:
            err(f"unknown key {k}")
    idx = [ORDER.index(k) for k in keys if k in ORDER]
    if idx != sorted(idx):
        err("keys are out of order")
    val = lambda l: l[len(l.split()[0]):].strip()
    get = lambda k: [val(l) for l in body if l.split()[0] == k]

    proc = (get("PROC") or [""])[0]
    kind = (get("KIND") or [""])[0]
    if kind not in ("Sub", "Function"):
        err("KIND must be Sub or Function")
    if (kind == "Function") != bool(get("RET")):
        err("RET is required for a Function and only for a Function")

    names = set()
    for a in ",".join(get("ARGS")).split(","):
        if a.strip():
            n, t = [x.strip() for x in a.split(" As ")]
            names.add(n)
            if t not in TYPES:
                err(f"unknown type {t}")
    for v in get("VAR"):
        n, t = [x.strip() for x in v.split(" As ")]
        names.add(n)
        if t not in TYPES:
            err(f"unknown type {t}")
    text = "\n".join(body)
    loopvars = set(re.findall(r"var=(\w+)", text))
    pre = " ".join(get("PRE"))

    # OP lines
    prev = None
    op_lines = [l for l in body if l.split()[0] == "OP"]
    for l in op_lines:
        raw = l[2:]
        step = raw.lstrip(" ")
        indent = len(raw) - len(step) - (len(op_lines[0][2:]) - len(op_lines[0][2:].lstrip(" ")))
        if indent % 2 or (prev is not None and indent > prev + 2) or indent < 0:
            err(f"bad nesting: {step}")
        prev = indent
        first = re.match(r"[A-Z_]+", step).group(0)
        if first not in STRUCT | ACTIONS:
            err(f"unknown op {first}")
        if step.count("(") != step.count(")") or step.count('"') % 2:
            err(f"unbalanced brackets or quotes: {step}")
        code = re.sub(r'"[^"]*"', '""', step)
        for w in re.findall(r"\b([A-Z_]{2,})\b", code):
            if w not in STRUCT | ACTIONS | VALUES | WORDS:
                err(f"unknown word {w}: {step}")
        for a in re.findall(r"\$(\w+)", code):
            if a not in names:
                err(f"undeclared argument ${a}")
        bare = re.sub(r"\b\w+=", "", code)
        for w in re.findall(r"(?<![\$\w#])([a-z]\w*)\b", bare):
            if w not in names | loopvars | OPTION_WORDS:
                err(f"undeclared name {w}: {step}")
        for leak in VBA_LEAKS:
            if leak in code:
                err(f"VBA syntax in spec: {leak.strip()}")
        for ws, header in re.findall(r'COL\(\$(\w+), "([^"]+)"\)', step):
            if f'COL(${ws}, "{header}")' not in pre:
                err(f'header "{header}" on ${ws} is used but not checked in PRE')

    # fixtures and tests
    fixture_headers = set()
    for f in get("FIXTURE"):
        m = re.search(r"cols=(.*?) rows=", f)
        if m:
            fixture_headers |= {c.split(":", 1)[1] for c in m.group(1).split(",")}
    written = set(re.findall(r'SET_VALUE\(CELL\(\$\w+, 1, [^)]*\), "([^"]+)"\)', text))
    tests = get("TEST")
    if len(tests) < 2:
        err("fewer than two TEST lines")
    if not any("ERROR_COUNT = 0" in t for t in tests):
        err("no TEST checks ERROR_COUNT = 0")
    if not any("PRE_" in t for t in tests):
        err("no TEST guards against collateral changes (PRE_...)")
    if not any("runs=2" in t for t in tests):
        err("no TEST uses runs=2 to prove re-runs are safe")
    if pre and len({re.search(r"sheet=(\w+)", f).group(1) for f in get("FIXTURE")}) > 1 \
            and not any("RESULT = -1" in t for t in tests):
        err("PRE is used but no wrong-sheet TEST expects RESULT = -1")
    for t in tests:
        if t.startswith("fixture=target"):
            err("fixture=target is not allowed: use a named synthetic FIXTURE")
        m = re.search(r"call=(\w+)\(", t)
        if not m or m.group(1) != proc:
            err(f"TEST calls a different procedure: {t[:60]}")
        for h in re.findall(r"!\[([^\]]+)\]", t):
            if h not in fixture_headers | written:
                err(f"TEST uses header [{h}] that no fixture or output row creates")
    for h in re.findall(r'COL\(\$\w+, "([^"]+)"\)', pre):
        if get("FIXTURE") and h not in fixture_headers:
            err(f'PRE header "{h}" is not in any fixture')
    return problems


if __name__ == "__main__":
    total = 0
    for p in sys.argv[1:]:
        found = check(p)
        total += len(found)
        print(f"{'OK  ' if not found else 'FAIL'} {p}")
        for f in found:
            print(f"     - {f}")
    print(f"\n{len(sys.argv) - 1} files checked, {total} problems")
    sys.exit(1 if total else 0)
