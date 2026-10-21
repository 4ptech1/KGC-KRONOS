"""Grammar-based parser for VCL 0.1 to 0.5 (prototype for milestone M2).

parse(text) -> (spec, problems)
  spec: dict with kind (TASK, BUILD, QUESTION), version ("VCL 0.1" to "VCL 0.6"), assumes,
        keys, ops (a nested tree) and the parsed Lark trees for every line.
  problems: list of (line_no, message). Parsing never raises on bad input.
"""
import re
from pathlib import Path

from lark import Lark, Token, Tree
from lark.exceptions import LarkError, UnexpectedInput

from errcodes import coded

GRAMMAR = Path(__file__).resolve().parent.parent / "vcl" / "grammar.lark"

TASK_ORDER = ["GOAL", "PROC", "KIND", "MODULE", "ARGS", "RET", "VAR", "PRE", "MODIFIES", "OP", "PERF", "ERR",
              "ALLOW", "FIXTURE", "TEST"]
BUILD_ORDER = ["GOAL", "NAME", "TARGET", "ENTRY", "MODULE", "FORM", "SIG", "FLOW", "FIXTURE", "TEST"]
START = {"PROC": "proc_value", "KIND": "kind_value", "MODULE": "module_value",
         "ARGS": "args_value", "RET": "ret_value", "VAR": "var_value", "PRE": "pre_value",
         "MODIFIES": "modifies_value",
         "OP": "op_value", "PERF": "perf_value", "ERR": "err_value", "ALLOW": "allow_value",
         "FIXTURE": "fixture_value", "TEST": "test_value", "NAME": "name_value",
         "TARGET": "target_value", "ENTRY": "entry_value", "FORM": "form_value",
         "SIG": "sig_value", "FLOW": "flow_value"}
VERSIONS = ("VCL 0.1", "VCL 0.2", "VCL 0.3", "VCL 0.4", "VCL 0.5", "VCL 0.6")
OPENERS = {"loop_rows", "loop_cols", "for_each", "if_block", "select_case", "case_line",
           "case_else"}
# G2: option names are reserved words, read from the grammar's OPTKEY terminal so there is one list.
OPTION_WORDS = tuple(re.search(r"OPTKEY\.2: /\(([a-z|]+)\)\(\?==\)/",
                               GRAMMAR.read_text(encoding="utf-8")).group(1).split("|"))

_parser = None


def fold_dedupe_cols(tree):
    """Before VCL 0.4, DEDUPE wrote two key columns as cols=1,2, which the grammar reads as cols=1
    and a positional 2 (G3). Fold those numbers into a list, so every stage sees cols=[1, 2] (D38);
    the validator refuses the old form in a VCL 0.4 spec."""
    calls = [t for t in tree.iter_subtrees() if t.data == "call" and str(t.children[0]) == "DEDUPE"
             and len(t.children) > 1]
    for call in calls:
        out, made, cols = [], [], None
        for item in call.children[1].children:
            if cols is not None and isinstance(item, Token) and item.type == "NUMBER":
                cols.children.append(item)   # cols=1,2: the 2 belongs to cols
                continue
            cols = None
            out.append(item)
            if (isinstance(item, Tree) and item.data == "option" and str(item.children[0]) == "cols"
                    and isinstance(item.children[1], Token) and item.children[1].type == "NUMBER"):
                cols = Tree("list_lit", [item.children[1]])
                item.children[1] = cols
                made.append(item)
        for option in made:   # a single number stays a number
            if len(option.children[1].children) == 1:
                option.children[1] = option.children[1].children[0]
        call.children[1].children = out
    return tree


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
        problems.append((line_no, coded(f"syntax error at column {col} near {str(got)!r}", "G001")))
    except LarkError as e:
        problems.append((line_no, coded(f"syntax error: {e.__class__.__name__}", "G002")))
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
                problems.append((line_no, coded("comments are only allowed before the first key", "G003")))
            elif raw.startswith("# ASSUME:"):
                spec["assumes"].append(raw[len("# ASSUME:"):].strip())
            else:
                problems.append((line_no, coded("only '# ASSUME:' comments are allowed in a TASK", "G004")))
            continue
        key = raw.split()[0]
        if key not in TASK_ORDER:
            problems.append((line_no, coded(f"unknown key {key!r}", "G005")))
            continue
        seen_key = True
        idx = TASK_ORDER.index(key)
        if idx < order_idx:
            problems.append((line_no, coded(f"key {key} is out of order", "G006")))
        order_idx = max(order_idx, idx)
        after_key = raw[len(key):]
        value = after_key.lstrip(" ")
        tree = None if key == "GOAL" else _parse_value(START[key], value, line_no, problems)   # GOAL: free text (D61)
        if key == "OP" and tree is not None:
            fold_dedupe_cols(tree)
        entry = {"key": key, "line": line_no, "text": value, "tree": tree}
        spec["keys"].append(entry)
        if key != "OP" or tree is None:
            continue
        col = len(raw) - len(value)
        if op_base is None:
            op_base = col
        extra = col - op_base
        if extra < 0 or extra % 2:
            problems.append((line_no, coded("OP indentation must be the first OP's column plus 2 spaces per level", "G007")))
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
                problems.append((line_no, coded("ELSE without a matching block IF at the same depth", "G008")))
            continue
        node = {"line": line_no, "stmt": stmt, "kind": kind, "children": [], "else": []}
        while stack and stack[-1][0] >= depth:
            stack.pop()
        allowed = stack[-1][0] + 1 if stack else 0
        if depth > allowed:
            problems.append((line_no, coded("OP is indented deeper than its parent allows", "G009")))
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
        problems.append((first_no, coded("empty spec", "G010")))
        return None
    head_i, head = nonblank[0]
    if head.startswith("# QUESTION:"):
        if len(nonblank) > 1:
            problems.append((first_no + nonblank[1][0], coded("a QUESTION reply must be the only line", "G011")))
        return {"kind": "QUESTION", "text": head[len("# QUESTION:"):].strip()}
    version = head.strip()
    if want and version != want:
        problems.append((first_no + head_i, coded(f"line 1 must be '{want}', as in the BUILD", "G012")))
    elif version not in VERSIONS:
        problems.append((first_no + head_i, coded("line 1 must be " + ", ".join(f"'{v}'" for v in VERSIONS[:-1])
                         + f" or '{VERSIONS[-1]}'", "G013")))
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
                    problems.append((line_no, coded(f"unknown BUILD key {key!r}", "G014")))
                else:
                    value = raw[len(key):].lstrip(" ")
                    tree = None if key == "GOAL" else _parse_value(START[key], value, line_no, problems)
                    head["keys"].append({"key": key, "line": line_no, "tree": tree})
        for a, b in zip(seps, seps[1:] + [len(lines)]):
            head["tasks"].append(_parse_lines(lines[a + 1:b], first_no + a + 1, problems,
                                              allow_build=False, want=version))
        return head
    problems.append((first_no + kind_i, coded("line 2 must be TASK" + (" or BUILD" if allow_build else ""), "G015")))
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
