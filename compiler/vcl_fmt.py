"""vclc fmt: one canonical text for every spec (VEP 0.6, D62).

    python compiler/vclc.py fmt FILE...            print the canonical text (a header line per file when
                                                   there is more than one)
    python compiler/vclc.py fmt --check FILE...    exit 1 and name each file that is not canonical
    python compiler/vclc.py fmt --write FILE...    rewrite the files in place

The same meaning always gives the same text, so the model (and a reader) meets one way to write each
spec. Only layout changes: key names padded to column 9, OP steps indented 2 spaces per level, one space
after a comma or a semicolon and around operators and keywords, none inside brackets or before a call's
bracket, no spaces in `name=value` options, no blank lines, one final newline. It never reorders keys and
never changes a quoted text, a name, a number or a FIXTURE rule.

It refuses a spec the validator rejects, and it proves its own work: every line must give exactly the same
tokens, the OP steps the same nesting, and the validator the same messages. If not, it raises FormatError
instead of writing a different spec.
"""
import re
import sys
from pathlib import Path

from lark import Lark

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import vcl_parse  # noqa: E402
from parser import parse_spec  # noqa: E402
from validator import validate_text  # noqa: E402

KEY_WIDTH = 9
# the `name=` words that start a field of a TEST or a FIXTURE line
FIELDS = {"fixture=", "call=", "runs=", "assert=", "sheet=", "cols=", "rows=", "stale=", "data=", "folder=",
          "files=", "locked="}
NO_BEFORE = {"RPAR", "RSQB", "COMMA", "SEMICOLON", "BANG", "EQUAL", "COLON"}
NO_AFTER = {"LPAR", "LSQB", "BANG", "EQUAL", "COLON", "MINUS"}   # MINUS is the unary minus (a binary one is ADD_OP)


class FormatError(Exception):
    """The spec can't be formatted: it has errors, or formatting would have changed its meaning."""

    def __init__(self, problems):
        super().__init__("; ".join(problems))
        self.problems = problems


_parser = None


def token_parser():
    """The grammar's parser with every token kept (even punctuation), so a line comes back as its tokens."""
    global _parser
    if _parser is None:
        _parser = Lark(vcl_parse.GRAMMAR.read_text(encoding="utf-8"), parser="lalr", lexer="contextual",
                       start=sorted(set(vcl_parse.START.values())), maybe_placeholders=False, keep_all_tokens=True)
    return _parser


def leaves(tree):
    if hasattr(tree, "children"):
        for child in tree.children:
            yield from leaves(child)
    else:
        yield tree


def tokens_of(key, value):
    return list(leaves(token_parser().parse(value, start=vcl_parse.START[key])))


def is_field(token):
    return str(token) in FIELDS and token.type.startswith("__ANON")


def gap(prev, nxt, mode):
    """The space between two tokens of a line: "" or " ". mode is the kind of key: fixture, or anything else."""
    if is_field(nxt):
        return " "
    if is_field(prev):
        return ""
    if mode == "fixture":
        return ""
    if nxt.type in NO_BEFORE or prev.type in NO_AFTER:
        return ""
    if nxt.type == "LPAR" and prev.type in ("NAME", "QUALNAME"):   # a call
        return ""
    if nxt.type == "LSQB" and prev.type == "ARG":   # $ws["Status"]
        return ""
    return " "


def format_value(key, value):
    """The canonical text of one key line's value."""
    if key == "GOAL":
        return " ".join(value.split())
    tokens = tokens_of(key, value)
    mode = "fixture" if key == "FIXTURE" else "other"
    out = []
    for i, token in enumerate(tokens):
        if i:
            out.append(gap(tokens[i - 1], token, mode))
        out.append(str(token))
    text = "".join(out)
    if [(t.type, str(t)) for t in tokens_of(key, text)] != [(t.type, str(t)) for t in tokens]:
        raise FormatError([f"{key} {value!r}: formatting would change its tokens"])
    return text


def op_depths(spec):
    """{file line: nesting depth} of every OP step, in a TASK or in the TASKs of a BUILD."""
    found = {}

    def walk(nodes, depth):
        for node in nodes:
            found[node.line] = depth
            walk(node.children, depth + 1)

    for part in [spec] + list(spec.tasks):
        walk(part.ops, 0)
    return found


def format_text(text):
    """The spec in its canonical form, or FormatError."""
    text = text.lstrip("﻿").replace("\r\n", "\n").replace("\r", "\n")
    errors = [p.message for p in validate_text(text) if p.severity == "error"]
    if errors:
        raise FormatError(errors)
    depths = op_depths(parse_spec(text))
    out = []
    for number, raw in enumerate(text.split("\n"), 1):
        line = raw.strip()
        if not line:
            continue
        if line == "---" or line in ("TASK", "BUILD") or re.fullmatch(r"VCL 0\.\d", line):
            out.append(line)
        elif line.startswith("# QUESTION:"):
            out.append("# QUESTION: " + line[len("# QUESTION:"):].strip())
        elif line.startswith("# ASSUME:"):
            out.append("# ASSUME: " + line[len("# ASSUME:"):].strip())
        else:
            key, value = split_key(line)
            indent = "  " * depths[number] if key == "OP" else ""
            out.append(f"{key:<{KEY_WIDTH}}{indent}{format_value(key, value)}")
    result = "\n".join(out) + "\n"
    check_same_spec(text, result)
    return result


def messages(text):
    return [(p.severity, p.message) for p in validate_text(text)]


def split_key(line):
    """(key, value) of a key line; the key and its value may be parted by spaces or tabs."""
    parts = line.split(None, 1)
    return parts[0], parts[1].strip() if len(parts) > 1 else ""


def key_lines(text):
    """(key, tokens) of every key line, GOAL as its words."""
    for raw in text.split("\n"):
        line = raw.strip()
        if not line or line in ("---", "TASK", "BUILD") or line.startswith("#") or re.fullmatch(r"VCL 0\.\d", line):
            continue
        key, value = split_key(line)
        yield key, value.split() if key == "GOAL" else [(t.type, str(t)) for t in tokens_of(key, value)]


def check_same_spec(before, after):
    """Formatting only moves whitespace: the same tokens on every line, the same OP nesting, the same validator
    messages."""
    if list(key_lines(before)) != list(key_lines(after)):
        raise FormatError(["formatting would change the tokens of a line"])
    a, b = parse_spec(before), parse_spec(after)
    if ([k.key for k in a.keys], list(op_depths(a).values())) != ([k.key for k in b.keys], list(op_depths(b).values())):
        raise FormatError(["formatting would change the keys or the nesting of the steps"])
    if messages(before) != messages(after):
        raise FormatError(["formatting would change what the validator says"])


def main(argv):
    check = "--check" in argv
    write = "--write" in argv
    files = [a for a in argv if not a.startswith("--")]
    if not files:
        print("usage: python compiler/vclc.py fmt [--check | --write] FILE...")
        return 1
    status = 0
    for name in files:
        path = Path(name)
        text = path.read_text(encoding="utf-8")
        try:
            result = format_text(text)
        except FormatError as err:
            print(f"{name}: can't format: {err.problems[0]}" + (f" (and {len(err.problems) - 1} more)" if len(err.problems) > 1 else ""))
            status = 1
            continue
        if check:
            if result != text:
                print(f"{name}: not in canonical form (run: python compiler/vclc.py fmt --write {name})")
                status = 1
        elif write:
            if result != text:
                path.write_text(result, encoding="utf-8")
                print(f"{name}: formatted")
        else:
            if len(files) > 1:
                print(f"==> {name} <==")
            sys.stdout.write(result)
    if check and not status:
        print(f"{len(files)} file{'s' if len(files) != 1 else ''} in canonical form")
    return status


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
