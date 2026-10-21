"""Rewrite a VCL 0.1 TASK spec into VCL 0.2 (docs/vep-0.2-core.md, PR D).

Usage:  python compiler/vcl_migrate.py specs/*.vcl            show what would change
        python compiler/vcl_migrate.py --write specs/*.vcl    rewrite the files
        python compiler/vcl_migrate.py --prove specs/*.vcl    also prove each rewrite (slow)

Only rewrites that can't change what a spec does:
  UPPER(TRIM(x)) = "LIT"   ->  SAME(x, "LIT")       LIT already upper case, trimmed and ASCII
  UPPER(TRIM(x)) <> "LIT"  ->  NOT SAME(x, "LIT")   (the literal keeps its text)
  a loop whose whole body is one block IF NOT IS_EMPTY(x), with no ELSE
                           ->  SKIP_IF IS_EMPTY(x), then the body one level up
  2 or more SET_VALUE(CELL($o, r, k), v) in a row, k = n, n+1, ...
                           ->  WRITE_ROW($o, r, v1, v2, ...) (col=n when n > 1), only when no
                               value reads $o and none can raise an error (so evaluating them all
                               first, as WRITE_ROW does, gives the same cells)
  MODIFIES lines from the sheets the write ops change: the whole sheet when it is cleared, its
  formats change, rows move or its columns can't be named; "fills" for fill-only FORMATs;
  otherwise the headers of the columns written (from LET(c, COL($x, "Header")))
  and the version line becomes VCL 0.2.

--prove checks, for every spec: the validator gives 0 errors and the same warnings; every TEST
passes before and after with the same RESULT and the same final workbook (UNCHANGED included);
and the TESTs miss no more mutants than before (D20). Mutants of the new SKIP_IF lines have no
VCL 0.1 counterpart, so they are listed as findings instead. Exits 1 if any spec fails.
"""
import datetime as dt
import re
import sys
from collections import Counter
from pathlib import Path

from lark import Token, Tree

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from parser import split_top  # noqa: E402
from vcl_parse import parse  # noqa: E402

LOOPS = ("loop_rows", "loop_cols", "for_each")
SHEET_FUNCS = ("CELL", "ROW_RANGE", "USED_RANGE", "RANGE")
WRITE_OPS = ("SET_VALUE", "WRITE_ROW", "FORMAT", "CLEAR", "COPY_ROW", "INSERT_ROW", "DELETE_ROW", "SORT",
             "WRITE_ARRAY", "SET_FORMULA", "DEDUPE", "FILTER")
ROW_OPS = ("INSERT_ROW", "DELETE_ROW", "SORT", "DEDUPE", "FILTER")
OP_LINE = re.compile(r"^(OP\s+)(.*)$")


class MigrateError(Exception):
    pass


def _args(call):
    args = call.children[1].children if len(call.children) > 1 else []
    pos = [a for a in args if not (isinstance(a, Tree) and a.data == "option")]
    opts = {str(a.children[0]): a.children[1] for a in args if isinstance(a, Tree) and a.data == "option"}
    return pos, opts


def _is_call(t, name=None):
    return isinstance(t, Tree) and t.data == "call" and (name is None or str(t.children[0]) == name)


def _walk(nodes):
    for n in nodes:
        yield n
        yield from _walk(n["children"])
        yield from _walk(n.get("else", []))


def _actions(spec):
    """Every op call in the spec, with its line: plain call lines and the action of IF ... THEN."""
    for n in _walk(spec["ops"]):
        if n["kind"] == "call":
            yield n["line"], n["stmt"]
        elif n["kind"] == "if_inline" and _is_call(n["stmt"].children[1]):
            yield n["line"], n["stmt"].children[1]


def _parse(text):
    spec, problems = parse(text)
    if problems or not spec or spec["kind"] != "TASK":
        raise MigrateError(f"only a TASK spec that parses can be migrated: {problems[:2]}")
    return spec


# ---------------------------------------------------------------- SAME

def _close(text, start):
    """Index just past the bracket that closes the one opened at text[start - 1]."""
    depth, quoted, i = 1, False, start
    while i < len(text):
        ch = text[i]
        if ch == '"':
            quoted = not quoted
        elif not quoted and ch == "(":
            depth += 1
        elif not quoted and ch == ")":
            depth -= 1
            if depth == 0:
                return i + 1
        i += 1
    return -1


SAME_BEFORE = re.compile(r"(?:^|[(,]|\b(?:IF|NOT|AND|OR|EXIT_IF|SKIP_IF|EXIT_LOOP_IF))\s*$")
SAME_AFTER = re.compile(r"^\s*(?:$|[),]|(?:AND|OR|THEN)\b)")
SAME_TAIL = re.compile(r'\s*(=|<>)\s*("(?:[^"]|"")*")')


def rewrite_same(value):
    """value is one OP line's text; returns it with every provable UPPER(TRIM(x)) = "LIT" rewritten."""
    out, i, count = "", 0, 0
    while True:
        j = value.find("UPPER(TRIM(", i)
        if j < 0:
            return out + value[i:], count
        inner_start = j + len("UPPER(TRIM(")
        inner_end = _close(value, inner_start)              # closes TRIM(
        ok = inner_end > 0 and value[inner_end:inner_end + 1] == ")"
        tail = SAME_TAIL.match(value, inner_end + 1) if ok else None
        if ok and tail and SAME_BEFORE.search(value[:j]) and SAME_AFTER.match(value[tail.end():]):
            lit = tail.group(2)
            text = lit[1:-1].replace('""', '"')
            if text and text.isascii() and text == text.upper() and text == text.strip(" \t\r\n "):
                x = value[inner_start:inner_end - 1]
                same = f"SAME({x}, {lit})"
                out += value[i:j] + (same if tail.group(1) == "=" else "NOT " + same)
                i = tail.end()
                count += 1
                continue
        out += value[i:inner_start]
        i = inner_start


# ---------------------------------------------------------------- SKIP_IF

def skip_if_edits(spec):
    """{line: new value} and {line: levels to dedent} for every loop whose whole body is one
    block IF NOT IS_EMPTY(x) without ELSE."""
    replace, dedent = {}, Counter()
    for n in _walk(spec["ops"]):
        if n["kind"] not in LOOPS or len(n["children"]) != 1:
            continue
        body = n["children"][0]
        if body["kind"] != "if_block" or body.get("else") or "has_else" in body or not body["children"]:
            continue
        cond = body["stmt"].children[0]
        if not (isinstance(cond, Tree) and cond.data == "not_op" and _is_call(cond.children[0], "IS_EMPTY")):
            continue
        replace[body["line"]] = cond.children[0]
        for inner in _walk(body["children"]):
            dedent[inner["line"]] += 1
            if "has_else" in inner:      # ELSE lines are not nodes of their own
                dedent[inner["has_else"]] += 1
    return replace, dedent


# ---------------------------------------------------------------- WRITE_ROW

def _safe(t, sheet):
    """True when evaluating t can't raise an error and doesn't read sheet."""
    if isinstance(t, Token):
        return t.type in ("STRING", "NUMBER", "NAME") or (t.type == "ARG" and str(t) != sheet)
    if _is_call(t, "CELL"):
        pos, _ = _args(t)
        return (len(pos) == 3 and isinstance(pos[0], Token) and pos[0].type == "ARG" and str(pos[0]) != sheet
                and all(isinstance(p, Token) and p.type in ("NAME", "NUMBER") for p in pos[1:]))
    if _is_call(t, "TRIM"):
        pos, _ = _args(t)
        return len(pos) == 1 and _safe(pos[0], sheet)
    if isinstance(t, Tree) and t.data == "sum":
        return all(_safe(c, sheet) for c in t.children[::2])
    return False


def _cell_write(n):
    """(sheet, row text key, column number, value tree) for SET_VALUE(CELL($o, r, <number>), v)."""
    if n["kind"] != "call" or not _is_call(n["stmt"], "SET_VALUE"):
        return None
    pos, _ = _args(n["stmt"])
    if len(pos) != 2 or not _is_call(pos[0], "CELL"):
        return None
    cpos, _ = _args(pos[0])
    if len(cpos) != 3 or not (isinstance(cpos[0], Token) and cpos[0].type == "ARG"):
        return None
    if not (isinstance(cpos[2], Token) and cpos[2].type == "NUMBER" and str(cpos[2]).isdigit()):
        return None
    return str(cpos[0]), cpos[1], int(str(cpos[2])), pos[1]


def _same_tree(a, b):
    return a == b if isinstance(a, Token) or isinstance(b, Token) else (
        isinstance(a, Tree) and isinstance(b, Tree) and a.data == b.data and len(a.children) == len(b.children)
        and all(_same_tree(x, y) for x, y in zip(a.children, b.children)))


def write_row_runs(spec):
    """Lists of sibling SET_VALUE nodes that WRITE_ROW can replace."""
    runs = []

    def scan(nodes):
        run = []
        for n in nodes + [None]:
            w = _cell_write(n) if n is not None else None
            if run and w is not None:
                first, last = _cell_write(run[0]), _cell_write(run[-1])
                if (w[0] == first[0] and _same_tree(w[1], first[1]) and w[2] == last[2] + 1
                        and _safe(w[3], w[0])):
                    run.append(n)
                    continue
            if len(run) >= 2:
                runs.append(run)
            run = [n] if w is not None and _safe(w[3], w[0]) else []
        for n in nodes:
            scan(n["children"])
            scan(n.get("else", []))

    scan(spec["ops"])
    return runs


def _call_parts(value, name):
    """The top-level argument texts of '<name>(...)' (the whole OP value)."""
    if not (value.startswith(name + "(") and value.endswith(")")):
        raise MigrateError(f"can't read {value!r}")
    return [p.strip() for p in split_top(value[len(name) + 1:-1], ",")]


# ---------------------------------------------------------------- MODIFIES

def modifies_lines(spec):
    """MODIFIES values for the sheets the write ops change, in ARGS order."""
    sheets = [str(p.children[0]) for k in spec["keys"] if k["key"] == "ARGS" for p in k["tree"].children
              if str(p.children[1]) == "Worksheet"]
    headers = {}   # (sheet arg, column var) -> header, from LET(var, COL($x, "Header"))
    lets = Counter()
    for _, call in _actions(spec):
        if _is_call(call, "LET"):
            pos, _ = _args(call)
            lets[str(pos[0])] += 1
            if _is_call(pos[1], "COL"):
                cpos, _ = _args(pos[1])
                if isinstance(cpos[0], Token) and cpos[0].type == "ARG":
                    headers[(str(cpos[0])[1:], str(pos[0]))] = str(cpos[1])[1:-1].replace('""', '"')
    need = {}      # arg -> {"whole": bool, "cols": [header], "fills": bool}
    for _, call in _actions(spec):
        name = str(call.children[0])
        if name not in WRITE_OPS:
            continue
        pos, opts = _args(call)
        target = opts.get("to") if name == "WRITE_ARRAY" else pos[2] if name == "COPY_ROW" else pos[0]
        if name == "SET_VALUE" and _is_call(target, "ARR"):
            continue
        if isinstance(target, Token) and target.type == "ARG":
            arg, col = str(target)[1:], None
        elif isinstance(target, Tree) and _is_call(target) and str(target.children[0]) in SHEET_FUNCS:
            tpos, _ = _args(target)
            if not (isinstance(tpos[0], Token) and tpos[0].type == "ARG"):
                raise MigrateError(f"{name} writes through {tpos[0]}, not a $sheet argument")
            arg = str(tpos[0])[1:]
            col = tpos[2] if str(target.children[0]) == "CELL" and len(tpos) == 3 else None
        else:
            raise MigrateError(f"{name} writes through something that isn't a $sheet argument")
        entry = need.setdefault(arg, {"whole": False, "cols": [], "fills": False})
        if name in ROW_OPS:
            entry["whole"] = True
        elif name == "FORMAT":
            if set(opts) == {"fill"}:
                entry["fills"] = True
            else:
                entry["whole"] = True
        elif name == "CLEAR" and str(opts.get("what", "all")) != "contents":
            entry["whole"] = True
        else:   # values
            header = None
            if name in ("SET_VALUE", "CLEAR") and isinstance(col, Token) and col.type == "NAME" \
                    and lets[str(col)] == 1:
                header = headers.get((arg, str(col)))
            if header is None:
                entry["whole"] = True
            elif header not in entry["cols"]:
                entry["cols"].append(header)
    lines = []
    for arg in sheets:
        entry = need.get(arg)
        if entry is None:
            continue
        if entry["whole"]:
            lines.append(f"${arg}")
            continue
        if entry["cols"]:
            lines.append(f"${arg}[" + ", ".join('"' + h.replace('"', '""') + '"' for h in entry["cols"]) + "]")
        if entry["fills"]:
            lines.append(f"${arg} fills")
    return lines


# ---------------------------------------------------------------- the whole rewrite

def migrate(text):
    """(VCL 0.2 text, list of what changed). Raises MigrateError for specs it can't migrate."""
    lines = text.splitlines()
    first = next((i for i, l in enumerate(lines) if l.strip()), None)
    if first is None or lines[first].strip() != "VCL 0.1":
        raise MigrateError("not a VCL 0.1 spec")
    spec = _parse(text)
    notes = []

    # 1. SAME
    for i, line in enumerate(lines):
        m = OP_LINE.match(line)
        if m and "UPPER(TRIM(" in m.group(2):
            new, count = rewrite_same(m.group(2))
            if count:
                lines[i] = m.group(1) + new
                notes.append(f"line {i + 1}: SAME x{count}")
    spec = _parse("\n".join(lines))

    # 2. SKIP_IF
    replace, dedent = skip_if_edits(spec)
    for line_no, cond in replace.items():
        m = OP_LINE.match(lines[line_no - 1])
        value = m.group(2)
        x = value[len("IF NOT "):] if value.startswith("IF NOT ") else None
        if x is None or not x.startswith("IS_EMPTY("):
            raise MigrateError(f"line {line_no}: can't read {value!r}")
        lines[line_no - 1] = m.group(1) + "SKIP_IF " + x
        notes.append(f"line {line_no}: SKIP_IF")
    for line_no, levels in dedent.items():
        m = OP_LINE.match(lines[line_no - 1])
        pad = m.group(1)
        if len(pad) - 2 * levels < len("OP "):
            raise MigrateError(f"line {line_no}: can't dedent")
        lines[line_no - 1] = pad[:len(pad) - 2 * levels] + m.group(2)
    spec = _parse("\n".join(lines))

    # 3. WRITE_ROW
    drop = set()
    for run in write_row_runs(spec):
        m = OP_LINE.match(lines[run[0]["line"] - 1])
        target, _ = _call_parts(m.group(2), "SET_VALUE")
        sheet, row, col = _call_parts(target, "CELL")
        values = [_call_parts(OP_LINE.match(lines[n["line"] - 1]).group(2), "SET_VALUE")[1] for n in run]
        tail = f", col={col}" if col != "1" else ""
        lines[run[0]["line"] - 1] = m.group(1) + f"WRITE_ROW({sheet}, {row}, {', '.join(values)}{tail})"
        drop |= {n["line"] for n in run[1:]}
        notes.append(f"line {run[0]['line']}: WRITE_ROW of {len(run)} values")
    lines = [l for i, l in enumerate(lines, 1) if i not in drop]
    spec = _parse("\n".join(lines))

    # 4. MODIFIES, after PRE (or before the first OP)
    mods = modifies_lines(spec)
    at = next(i for i, l in enumerate(lines) if l.startswith("OP "))
    lines[at:at] = [f"MODIFIES {m}" for m in mods]
    notes += [f"MODIFIES {m}" for m in mods]

    # 5. the version line
    lines[first] = "VCL 0.2"
    return "\n".join(lines) + ("\n" if text.endswith("\n") else ""), notes


# ---------------------------------------------------------------- the proof

def prove(old, new, today=None, mutation=True):
    """Problems (empty when the rewrite is proven) for one spec, as the VEP defines the proof."""
    from validator import validate_text
    from vcl_interp import run_tests
    from vcl_mutate import mutate_report
    today = today or dt.date.today()
    problems = []
    before, after = validate_text(old), validate_text(new)
    problems += [f"validator: {p.message}" for p in after if p.severity == "error"]
    old_warn = Counter(re.sub(r"\d+", "#", p.message) for p in before if p.severity == "warning")
    new_warn = Counter(re.sub(r"\d+", "#", p.message) for p in after if p.severity == "warning")
    problems += [f"new warning: {w}" for w in new_warn - old_warn]
    r_old, r_new = run_tests(old, today), run_tests(new, today)
    if len(r_old) != len(r_new):
        problems.append("a different number of TESTs")
    for a, b in zip(r_old, r_new):
        if not b.passed:
            problems.append(f"TEST (line {b.line}) fails: {b.failed[:3]}")
        if (a.result, a.state) != (b.result, b.state):
            problems.append(f"TEST (line {b.line}) gives a different RESULT or workbook")
    scores = None
    if mutation and not problems:
        k1, t1, old_survivors = mutate_report(old, snapshot=True, today=today)
        k2, t2, new_survivors = mutate_report(new, snapshot=True, today=today)
        new_lines = new.splitlines()
        # Mutants of SKIP_IF lines have no counterpart in VCL 0.1, so they are reported, not compared
        probes = [s for s in new_survivors if new_lines[s[0] - 1].lstrip("OP ").startswith(("SKIP_IF ", "EXIT_LOOP_IF "))]
        comparable = len(new_survivors) - len(probes)
        if comparable > len(old_survivors):
            problems.append(f"the TESTs miss {comparable} mutants, {len(old_survivors)} before")
        scores = (k1, t1, k2, t2, len(old_survivors), comparable, probes)
    return problems, scores


def main(argv):
    write, do_prove = "--write" in argv, "--prove" in argv
    files = [a for a in argv if not a.startswith("--")]
    if not files:
        print("usage: python compiler/vcl_migrate.py [--write] [--prove] specs/*.vcl", file=sys.stderr)
        return 1
    bad = 0
    for path in files:
        old = Path(path).read_text(encoding="utf-8")
        try:
            new, notes = migrate(old)
        except MigrateError as exc:
            print(f"SKIP {path}: {exc}")
            continue
        print(f"{path}: " + "; ".join(notes))
        problems = []
        if do_prove:
            problems, scores = prove(old, new)
            for p in problems:
                print(f"     FAIL {p}")
            if problems:
                bad += 1
            else:
                k1, t1, k2, t2, before, after, probes = scores
                print(f"     proven: same results; {before} -> {after} missed mutants "
                      f"(score {100 * k1 / t1:.1f}% -> {100 * k2 / t2:.1f}% of {t1} -> {t2})")
                for line, what in probes:
                    print(f"     finding: no TEST notices line {line} ({what}): the FIXTURE has no row it skips")
        if write and not problems:
            Path(path).write_text(new, encoding="utf-8")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
