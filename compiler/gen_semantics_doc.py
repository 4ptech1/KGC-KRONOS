"""Generate docs/vcl-semantics.md from vcl/semantics.py docstrings and tests/conformance_cases.py.

Usage: python compiler/gen_semantics_doc.py [--check]   (--check fails if the doc is out of date)
"""
import datetime as dt
import inspect
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "vcl"))
sys.path.insert(0, str(ROOT / "tests"))
import semantics  # noqa: E402
from conformance_cases import CASES  # noqa: E402

NAMES = {"is_empty": "IS_EMPTY", "is_number": "IS_NUMBER", "is_date": "IS_DATE",
         "is_error": "IS_ERROR (VCL 0.2)", "trim": "TRIM",
         "text_of": "Text conversion (TRIM, UPPER, &, keys)", "upper": "UPPER", "length": "LEN",
         "left": "LEFT", "right": "RIGHT", "equals": "= with text", "col": "COL",
         "dict_key": "Dictionary keys", "last_row": "LASTROW", "same": "SAME (VCL 0.2)",
         "clean_text": "CLEAN_TEXT (VCL 0.2)", "parse_date": "PARSE_DATE (VCL 0.2)",
         "iif": "IIF (VCL 0.2)", "mod": "MOD (VCL 0.2)",
         "needs_text_guard": "Text written by SET_VALUE, WRITE_ROW and COPY_ROW (VCL 0.2)"}
DECISION_FUNCS = ["is_empty", "trim", "is_number", "is_date", "is_error", "text_of", "equals",
                  "left", "header_key", "dict_key", "last_row", "same", "clean_text", "parse_date",
                  "iif", "mod", "needs_text_guard"]


def show(v):
    if isinstance(v, str):
        v = v.replace("\u00a0", "⍽").replace("\n", "\\n").replace("\t", "\\t")
        return '"' + "".join(ch if ord(ch) >= 32 else f"\\x{ord(ch):02x}" for ch in v) + '"'
    if v is None:
        return "(empty)"
    if isinstance(v, list):
        return "[" + ", ".join(show(x) for x in v) + "]"
    if isinstance(v, dt.datetime):
        return "date and time " + v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, dt.date):
        return "date " + v.strftime("%Y-%m-%d")
    return repr(v)


def render():
    decisions = []
    for fn in DECISION_FUNCS:
        doc = " ".join(inspect.getdoc(getattr(semantics, fn)).split())
        m = re.match(r"(S\d+(?: \([^)]*\))?): (.*)", doc)
        if m:
            decisions.append((m.group(1), fn, m.group(2)))
    decisions.sort(key=lambda d: int(re.match(r"S(\d+)", d[0]).group(1)))
    out = ["# VCL core semantics (reference)", "",
           "Generated from `vcl/semantics.py` and `tests/conformance_cases.py`. Do not edit by hand.",
           "ForgeLib (VBA) must give exactly these results; the same table becomes the Excel conformance batch.",
           "In the tables, ⍽ is a non-breaking space (character 160).", "", "## Decisions", "",
           "| Id | Function | Meaning |", "| --- | --- | --- |"]
    out += [f"| {d} | `{fn}` | {text.replace('|', '/')} |" for d, fn, text in decisions]
    out += ["", "## Conformance cases", ""]
    current = None
    for fn, args, expected, why in CASES:
        if fn != current:
            out += ["", f"### {NAMES.get(fn, fn)}", "", "| Input | Result | Why |", "| --- | --- | --- |"]
            current = fn
        out.append(f"| {', '.join(show(a) for a in args)} | {show(expected)} | {why} |")
    return "\n".join(out) + "\n"


if __name__ == "__main__":
    target = ROOT / "docs" / "vcl-semantics.md"
    text = render()
    if "--check" in sys.argv:
        ok = target.exists() and target.read_text(encoding="utf-8") == text
        print("docs/vcl-semantics.md is up to date" if ok else "docs/vcl-semantics.md is OUT OF DATE")
        sys.exit(0 if ok else 1)
    target.write_text(text, encoding="utf-8")
    print(f"wrote {target.relative_to(ROOT)}")
