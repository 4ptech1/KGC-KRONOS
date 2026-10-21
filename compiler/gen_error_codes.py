"""Gives every message of the checks a stable code, and writes the catalogue (VEP 0.6, D60).

    python compiler/gen_error_codes.py            give new messages their codes, write the files
    python compiler/gen_error_codes.py --check    exit 1 if a message has no code, a code is used
                                                  twice or isn't registered, or the catalogue is stale

A message site is a call the checks use to report a problem (see SITES). The tool wraps its
message as coded("...", "V012"); a site written by hand as coded("...") gets its code the same
way. Codes are never reused or renumbered: compiler/error_codes.json keeps every code ever given
(with the first wording), and docs/error-codes.md lists them with their current wording. A code
whose message was deleted is shown as retired.
"""
import ast
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
REGISTRY = ROOT / "compiler" / "error_codes.json"
CATALOGUE = ROOT / "docs" / "error-codes.md"

# file -> (code prefix, {callee name: index of the message argument},
#          receivers whose .append(message) or .append((..., message)) is a message,
#          {function name: index of the message in a returned tuple})
SITES = {
    "parser.py": ("V", {"error": 1, "Problem": 2}, set(), {"tokenize": 1}),
    "validator.py": ("V", {"error": 0, "warning": 0, "Problem": 2, "Unbalanced": 0}, set(), {}),
    "vcl_formula.py": ("V", {}, {"problems"}, {}),
    "vcl_parse.py": ("G", {}, {"problems"}, {}),
    "vcl_types.py": ("T", {"err": 1}, set(), {}),
    "vcl_fixtures.py": ("X", {"ValueError": 0}, set(), {}),
    "vcl_interp.py": ("R", {"VclError": 0}, {"found", "failed"}, {}),
}
PREFIX_NAMES = {"V": "validator and parser", "G": "grammar", "T": "type checker", "X": "fixture builder",
                "R": "TEST run"}
MESSAGE_KINDS = (ast.Constant, ast.JoinedStr, ast.BinOp, ast.IfExp)
CODE_RE = re.compile(r"[VGTXR]\d{3}")


def callee(call):
    func = call.func
    return func.attr if isinstance(func, ast.Attribute) else getattr(func, "id", "")


def is_message(node):
    if isinstance(node, ast.Constant):
        return isinstance(node.value, str) and node.value != ""   # "" is no message (tokenize: no error)
    return isinstance(node, MESSAGE_KINDS)


def parents_of(tree):
    parents = {}
    for node in ast.walk(tree):
        for child in ast.iter_child_nodes(node):
            parents[child] = node
    return parents


def enclosing_function(node, parents):
    while node in parents:
        node = parents[node]
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            return node.name
    return ""


class Site:
    def __init__(self, file, node, wrapped, severity):
        self.file, self.node, self.wrapped, self.severity = file, node, wrapped, severity
        self.code = ""

    def template(self):
        node = self.node
        if isinstance(node, ast.JoinedStr):
            parts = []
            for piece in node.values:
                if isinstance(piece, ast.Constant):
                    parts.append(str(piece.value))
                else:
                    inner = piece.value
                    parts.append("{" + ast.unparse(inner) + "}")
            return "".join(parts)
        if isinstance(node, ast.Constant):
            return str(node.value)
        return ast.unparse(node)


def find_sites(path, source):
    """Every message site in one file: those already written coded(...), and those to be wrapped."""
    prefix, calls, receivers, returns = SITES[path.name]
    tree = ast.parse(source)
    parents = parents_of(tree)
    sites, seen = [], set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Call) and callee(node) == "coded" and node.args:
            message = node.args[0]
            literal = node.args[1] if len(node.args) > 1 else None
            if literal is None or (isinstance(literal, ast.Constant) and literal.value in ("", "NEW")):
                site = Site(path.name, message, node, "error")
            elif isinstance(literal, ast.Constant):
                site = Site(path.name, message, node, "error")
                site.code = literal.value
            else:
                continue   # a wrapper that keeps the code of the message it rewords
            sites.append(site)
            seen.add(id(message))
    for node in ast.walk(tree):
        found = []
        if isinstance(node, ast.Call):
            name = callee(node)
            if name in calls and len(node.args) > calls[name]:
                found.append((node.args[calls[name]], "warning" if name == "warning" else "error"))
                if name == "Problem" and len(node.args) > 2 and isinstance(node.args[1], ast.Constant):
                    found[-1] = (node.args[2], node.args[1].value)
            elif name == "append" and isinstance(node.func, ast.Attribute) and isinstance(node.func.value, ast.Name) \
                    and node.func.value.id in receivers and node.args:
                arg = node.args[0]
                found.append((arg.elts[-1] if isinstance(arg, ast.Tuple) else arg, "error"))
        elif isinstance(node, ast.Return) and isinstance(node.value, ast.Tuple):
            index = returns.get(enclosing_function(node, parents))
            if index is not None and len(node.value.elts) > index:
                found.append((node.value.elts[index], "error"))
        for message, severity in found:
            if not is_message(message) or id(message) in seen:
                continue
            inside = parents.get(message)
            if isinstance(inside, ast.Call) and callee(inside) == "coded":
                continue
            seen.add(id(message))
            site = Site(path.name, message, None, severity)
            sites.append(site)
    sites.sort(key=lambda s: (s.node.lineno, s.node.col_offset))
    return sites


def byte_offsets(source):
    starts, total = [], 0
    for line in source.encode("utf-8").splitlines(keepends=True):
        starts.append(total)
        total += len(line)
    return starts


def load_registry():
    if REGISTRY.exists():
        return json.loads(REGISTRY.read_text(encoding="utf-8"))
    return {}


def next_code(registry, prefix):
    used = [int(code[1:]) for code in registry if code.startswith(prefix)]
    return f"{prefix}{max(used, default=0) + 1:03d}"


def scan(directory):
    result = {}
    for name in SITES:
        path = directory / name
        result[name] = (path, path.read_text(encoding="utf-8"), find_sites(path, path.read_text(encoding="utf-8")))
    return result


def problems_with(scanned, registry):
    """What --check reports: uncoded sites, reused or unregistered codes, wrong prefixes."""
    found, owner = [], {}
    for name, (path, source, sites) in scanned.items():
        prefix = SITES[name][0]
        for site in sites:
            where = f"{name}:{site.node.lineno}"
            if not site.code:
                found.append(f"{where}: this message has no code: {site.template()[:70]!r}")
                continue
            if not CODE_RE.fullmatch(site.code):
                found.append(f"{where}: {site.code!r} is not a code like {prefix}012")
            elif site.code[0] != prefix:
                found.append(f"{where}: {site.code} must start with {prefix} in {name}")
            if site.code in owner:
                found.append(f"{where}: {site.code} is already used at {owner[site.code]}")
            owner[site.code] = where
            if site.code not in registry:
                found.append(f"{where}: {site.code} is not in compiler/error_codes.json: run gen_error_codes.py")
    return found, owner


def catalogue_text(scanned, registry):
    rows = {}
    for name, (path, source, sites) in scanned.items():
        for site in sites:
            if site.code:
                rows[site.code] = (site.severity, site.template(), name)
    out = ["# Error codes", "",
           "Generated by `python compiler/gen_error_codes.py`: do not edit by hand. Every message the checks",
           "give carries one of these codes (`Problem.code`, shown as `[V012]` on the command line). A code",
           "never changes meaning and is never reused; a message may be reworded and keep its code. In a",
           "message, `{...}` stands for the part that varies.", ""]
    for prefix, label in PREFIX_NAMES.items():
        out += [f"## {prefix}: {label}", "", "| Code | Kind | Message | File |", "| --- | --- | --- | --- |"]
        for code in sorted(c for c in registry if c.startswith(prefix)):
            if code in rows:
                severity, template, name = rows[code]
                text = template.replace("|", "\\|").replace("\n", " ")
                out.append(f"| {code} | {severity} | {text} | {name} |")
            else:
                first = registry[code].replace("|", "\\|").replace("\n", " ")
                out.append(f"| {code} | retired | {first} | |")
        out.append("")
    return "\n".join(out)


def assign(scanned, registry):
    """Give every uncoded site its code; returns the new sources {name: text} and the registry."""
    new_sources = {}
    for name, (path, source, sites) in scanned.items():
        prefix = SITES[name][0]
        pending = [s for s in sites if not s.code]
        if not pending:
            continue
        data = source.encode("utf-8")
        starts = byte_offsets(source)

        def offset(lineno, col):
            return starts[lineno - 1] + col

        edits = []
        for site in pending:
            code = next_code(registry, prefix)
            registry[code] = site.template()
            site.code = code
            node = site.node
            begin, end = offset(node.lineno, node.col_offset), offset(node.end_lineno, node.end_col_offset)
            if site.wrapped is None:
                edits.append((begin, end, b"coded(" + data[begin:end] + f', "{code}")'.encode("utf-8")))
            else:   # coded(<message>) written by hand
                call = site.wrapped
                cb, ce = offset(call.lineno, call.col_offset), offset(call.end_lineno, call.end_col_offset)
                edits.append((cb, ce, b"coded(" + data[begin:end] + f', "{code}")'.encode("utf-8")))
        for begin, end, text in sorted(edits, reverse=True):
            data = data[:begin] + text + data[end:]
        new_sources[name] = data.decode("utf-8")
    return new_sources


def check():
    registry = load_registry()
    scanned = scan(ROOT / "compiler")
    found, _ = problems_with(scanned, registry)
    if not CATALOGUE.exists() or CATALOGUE.read_text(encoding="utf-8") != catalogue_text(scanned, registry):
        found.append("docs/error-codes.md is out of date: run python compiler/gen_error_codes.py")
    return found


def main(argv):
    if "--list" in argv:   # what the tool sees, before it changes anything
        for name, (path, source, sites) in scan(ROOT / "compiler").items():
            print(f"== {name}: {len(sites)} messages")
            for site in sites:
                print(f"  {site.node.lineno:5} {site.code or '-':5} {site.template()[:110]}")
        return 0
    if "--check" in argv:
        found = check()
        for line in found:
            print(line)
        if not found:
            print("error codes: every message has a stable code, and docs/error-codes.md is up to date")
        return 1 if found else 0
    registry = load_registry()
    scanned = scan(ROOT / "compiler")
    new_sources = assign(scanned, registry)
    for name, text in new_sources.items():
        (ROOT / "compiler" / name).write_text(text, encoding="utf-8")
    scanned = scan(ROOT / "compiler")
    found, _ = problems_with(scanned, registry)
    REGISTRY.write_text(json.dumps(dict(sorted(registry.items())), indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    CATALOGUE.write_text(catalogue_text(scanned, registry), encoding="utf-8")
    for line in found:
        print(line)
    print(f"gave {sum(1 for _ in registry)} codes in all; wrote {REGISTRY.name} and docs/{CATALOGUE.name}")
    return 1 if found else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
