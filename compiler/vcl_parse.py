"""Grammar-based parser for VCL 0.1, 0.2 and 0.3 (prototype for milestone M2).

parse(text) -> (spec, problems)
  spec: dict with kind (TASK, BUILD, QUESTION), version ("VCL 0.1", "VCL 0.2" or "VCL 0.3"), assumes,
        keys, ops (a nested tree) and the parsed Lark trees for every line.
  problems: list of (line_no, message). Parsing never raises on bad input.
"""
import re
from pathlib import Path

from lark import Lark, Token, Tree
from lark.exceptions import LarkError, UnexpectedInput

GRAMMAR = Path(__file__).resolve().parent.parent / "vcl" / "grammar.lark"

TASK_ORDER = ["PROC", "KIND", "MODULE", "ARGS", "RET", "VAR", "PRE", "MODIFIES", "OP", "PERF", "ERR",
              "ALLOW", "FIXTURE", "TEST"]
BUILD_ORDER = ["NAME", "TARGET", "ENTRY", "MODULE", "FORM", "SIG", "FLOW", "FIXTURE", "TEST"]
START = {"PROC": "proc_value", "KIND": "kind_value", "MODULE": "module_value",
         "ARGS": "args_value", "RET": "ret_value", "VAR": "var_value", "PRE": "pre_value",
         "MODIFIES": "modifies_value",
         "OP": "op_value", "PERF": "perf_value", "ERR": "err_value", "ALLOW": "allow_value",
         "FIXTURE": "fixture_value", "TEST": "test_value", "NAME": "name_value",
         "TARGET": "target_value", "ENTRY": "entry_value", "FORM": "form_value",
         "SIG": "sig_value", "FLOW": "flow_value"}
VERSIONS = ("VCL 0.1", "VCL 0.2", "VCL 0.3")
OPENERS = {"loop_rows", "loop_cols", "for_each", "if_block", "select_case", "case_line",
           "case_else"}
# G2: option names are reserved words, read from the grammar's OPTKEY terminal so there is one list.
OPTION_WORDS = tuple(re.search(r"OPTKEY\.2: /\(([a-z|]+)\)\(\?==\)/",
                               GRAMMAR.read_text(encoding="utf-8")).group(1).split("|"))

_parser = None


def parser():
    global _parser
    if _parser is None:
        _parser = Lark(GRAMMAR.read_text(encoding="utf-8"), parser="lalr",
                       lexer="contextual", start=sorted(set(START.values())),
                       maybe_placeholders=False)
    return _parser


def _parse_value(start, value, line_no, problems):
    try:
        return parser().parse(value, start=start)
    except UnexpectedInput as e:
        col = getattr(e, "column", "?")
        got = getattr(e, "token", None) or getattr(e, "char", "")
        problems.append((line_no, f"syntax error at column {col} near {str(got)!r}"))
    except LarkError as e:
        problems.append((line_no, f"syntax error: {e.__class__.__name__}"))
    return None


def _parse_task(lines, first_no, problems, version):
    spec = {"kind": "TASK", "version": version, "assumes": [], "keys": [], "ops": []}
    seen_key = False
    order_idx = -1
    op_base = None
    stack = []  # (depth, node, list that children go into)
    for offset, raw in enumerate(lines):
        line_no = first_no + offset
        if not raw.strip():
            continue
        if raw.startswith("#"):
            if seen_key:
                problems.append((line_no, "comments are only allowed before the first key"))
            elif raw.startswith("# ASSUME:"):
                spec["assumes"].append(raw[len("# ASSUME:"):].strip())
            else:
                problems.append((line_no, "only '# ASSUME:' comments are allowed in a TASK"))
            continue
        key = raw.split()[0]
        if key not in TASK_ORDER:
            problems.append((line_no, f"unknown key {key!r}"))
            continue
        seen_key = True
        idx = TASK_ORDER.index(key)
        if idx < order_idx:
            problems.append((line_no, f"key {key} is out of order"))
        order_idx = max(order_idx, idx)
        after_key = raw[len(key):]
        value = after_key.lstrip(" ")
        tree = _parse_value(START[key], value, line_no, problems)
        entry = {"key": key, "line": line_no, "text": value, "tree": tree}
        spec["keys"].append(entry)
        if key != "OP" or tree is None:
            continue
        col = len(raw) - len(value)
        if op_base is None:
            op_base = col
        extra = col - op_base
        if extra < 0 or extra % 2:
            problems.append((line_no, "OP indentation must be the first OP's column plus 2 spaces per level"))
            continue
        depth = extra // 2
        stmt = tree.children[0]
        kind = stmt.data if isinstance(stmt, Tree) else "call"
        if kind == "else_line":
            while stack and stack[-1][0] > depth:
                stack.pop()
            top = stack[-1] if stack else None
            if top and top[0] == depth and top[1]["kind"] == "if_block" and "has_else" not in top[1]:
                top[1]["has_else"] = line_no
                stack[-1] = (depth, top[1], top[1]["else"])
            else:
                problems.append((line_no, "ELSE without a matching block IF at the same depth"))
            continue
        node = {"line": line_no, "stmt": stmt, "kind": kind, "children": [], "else": []}
        while stack and stack[-1][0] >= depth:
            stack.pop()
        allowed = stack[-1][0] + 1 if stack else 0
        if depth > allowed:
            problems.append((line_no, "OP is indented deeper than its parent allows"))
            continue
        (stack[-1][2] if stack else spec["ops"]).append(node)
        if kind in OPENERS:
            stack.append((depth, node, node["children"]))
    return spec


def _parse_lines(lines, first_no, problems, allow_build=True, want=None):
    """Parse one spec held in lines; first_no is the file line number of lines[0]. want is the
    version line a TASK inside a BUILD must repeat (the BUILD's own)."""
    nonblank = [(i, l) for i, l in enumerate(lines) if l.strip()]
    if not nonblank:
        problems.append((first_no, "empty spec"))
        return None
    head_i, head = nonblank[0]
    if head.startswith("# QUESTION:"):
        if len(nonblank) > 1:
            problems.append((first_no + nonblank[1][0], "a QUESTION reply must be the only line"))
        return {"kind": "QUESTION", "text": head[len("# QUESTION:"):].strip()}
    version = head.strip()
    if want and version != want:
        problems.append((first_no + head_i, f"line 1 must be '{want}', as in the BUILD"))
    elif version not in VERSIONS:
        problems.append((first_no + head_i, "line 1 must be 'VCL 0.1', 'VCL 0.2' or 'VCL 0.3'"))
        version = VERSIONS[0]
    kind_i = nonblank[1][0] if len(nonblank) > 1 else head_i + 1
    kind = lines[kind_i].strip() if kind_i < len(lines) else ""
    if kind == "TASK":
        return _parse_task(lines[kind_i + 1:], first_no + kind_i + 1, problems, want or version)
    if kind == "BUILD" and allow_build:
        head = {"kind": "BUILD", "version": version, "keys": [], "tasks": []}
        seps = [i for i in range(kind_i + 1, len(lines)) if lines[i].strip() == "---"]
        for i in range(kind_i + 1, seps[0] if seps else len(lines)):
            raw, line_no = lines[i], first_no + i
            if raw.strip() and not raw.startswith("#"):
                key = raw.split()[0]
                if key not in BUILD_ORDER:
                    problems.append((line_no, f"unknown BUILD key {key!r}"))
                else:
                    value = raw[len(key):].lstrip(" ")
                    tree = _parse_value(START[key], value, line_no, problems)
                    head["keys"].append({"key": key, "line": line_no, "tree": tree})
        for a, b in zip(seps, seps[1:] + [len(lines)]):
            head["tasks"].append(_parse_lines(lines[a + 1:b], first_no + a + 1, problems,
                                              allow_build=False, want=version))
        return head
    problems.append((first_no + kind_i, "line 2 must be TASK" + (" or BUILD" if allow_build else "")))
    return None


def parse(text):
    """parse(text) -> (spec, problems). Line numbers are the file's, also inside a BUILD."""
    problems = []
    if text.startswith("\ufeff"):  # a UTF-8 byte order mark, as Windows Notepad saves it
        text = text[1:]
    spec = _parse_lines(text.splitlines(), 1, problems)
    return spec, problems


if __name__ == "__main__":
    import sys
    bad = 0
    for p in sys.argv[1:]:
        spec, probs = parse(Path(p).read_text(encoding="utf-8"))
        ops = 0
        def count(nodes):
            n = 0
            for x in nodes:
                n += 1 + count(x["children"]) + count(x.get("else", []))
            return n
        if spec and spec.get("kind") == "TASK":
            ops = count(spec["ops"])
        status = "OK  " if not probs else "FAIL"
        bad += bool(probs)
        print(f"{status} {p}  ({len(spec['keys']) if spec and 'keys' in spec else 0} key lines, {ops} ops)")
        for ln, msg in probs:
            print(f"     line {ln}: {msg}")
    sys.exit(1 if bad else 0)
