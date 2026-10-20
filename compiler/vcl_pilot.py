"""The Excel pilot of a spec's TESTs, and the check of its results (router Phase 5b, D35).

  python compiler/vcl_pilot.py check RESULTS.txt SPEC...   compare the lines pasted back from Excel

`vclc build` writes dist/forge_pilot.bas for the specs in PILOT_SPECS. Pasted into Excel next to
forge_module.bas, its ForgePilotRun runs every TEST of those specs as the interpreter does: fresh
fixture sheets and, for folder FIXTUREs, a new sandbox below %TEMP% (D30); the procedure called
runs=<n> times, with ForgeLog cleared before each run. After each TEST it writes plain lines to
the sheet ForgePilot:

  PILOT <n> spec(s)                 then DATE yyyy-mm-dd (Excel's TODAY)
  SPEC <file>
  TEST <spec line> RESULT=<text> ERRORS=<n>     RESULT is (none) for a Sub
  CELL <sheet> <row> <col> [<text>]             every non-empty cell of the fixture's sheets
  FILL <sheet> <row> <col> #RRGGBB              every filled cell
  FOLDER <path below the root>                  every folder and file in the sandbox
  FILE <path below the root> <bytes>
  END                                           or ABORTED error <n>: <text>

Text is the canonical text of S6 (ForgeText), with line breaks and tabs written as {LF}, {CR} and
{TAB}. `check` runs the same TESTs in the interpreter, with TODAY set to the pasted DATE, and
prints one line per difference; it exits 0 when Excel and the interpreter agree.
"""
import datetime as dt
import re
import sys
from pathlib import Path

from lark import Token, Tree

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "vcl"))
import semantics as S  # noqa: E402
from vcl_codegen import vba_name  # noqa: E402
from vcl_fixtures import (EXCEL_SANDBOX_MAX, FIXTURE_HELPERS, SANDBOX, parse_fixture, parse_folder,  # noqa: E402
                          proc_name, sandbox_name, vba_for_file, vba_for_folders, vba_literal)
from vcl_interp import run_tests  # noqa: E402
from vcl_parse import parse  # noqa: E402
from vcl_selftest import vba_text  # noqa: E402

PILOT_SPECS = ("task_12.vcl",)   # the router MVP (docs/vep-0.3-file-ops.md, Phase 5)
ESCAPES = (("\n", "{LF}"), ("\r", "{CR}"), ("\t", "{TAB}"))


def privately(vba):
    """A fixture builder's procedures as Private ones: the pilot module keeps its own copies."""
    return re.sub(r"^Public (Sub|Function) ", r"Private \1 ", vba, flags=re.M)


def test_calls(spec):
    """Each TEST of a parsed spec: (spec line, fixture file, [call value tokens], runs)."""
    found = []
    for k in spec["keys"]:
        if k["key"] != "TEST":
            continue
        kids = k["tree"].children
        rest, i, vals = kids[1:], 1, []
        while i < len(rest) and not (isinstance(rest[i], Token) and rest[i].type == "INT") \
                and not (isinstance(rest[i], Tree) and rest[i].data == "assertion"):
            vals.append(rest[i])
            i += 1
        runs = int(rest[i]) if i < len(rest) and isinstance(rest[i], Token) and rest[i].type == "INT" else 1
        found.append((k["line"], str(kids[0]), vals, runs))
    return found


def spec_procedure(name, text):
    """The VBA that runs every TEST of one spec, and the fixture files it needs."""
    spec, problems = parse(text)
    if problems or spec["kind"] != "TASK":
        raise ValueError(f"{name}: only a TASK spec that parses can be piloted")
    keys = {k["key"]: k["tree"] for k in spec["keys"] if k["key"] in ("PROC", "KIND")}
    proc, kind = str(keys["PROC"].children[0]), str(keys["KIND"].children[0])
    trees = [k["tree"] for k in spec["keys"] if k["key"] == "FIXTURE"]
    sheets = {}
    for t in trees:
        if t.data != "folder_fixture":
            fx = parse_fixture(t)
            sheets.setdefault(fx.file, []).append(fx.sheet)
    folders = {str(t.children[0]) for t in trees if t.data == "folder_fixture"}
    stem = name.rsplit(".", 1)[0]
    out = [f"Private Sub PilotSpec_{stem}()", "    Dim root As String", "    Dim result As Variant",
           "    Dim i As Long", f'    PilotLine "SPEC {name}"']
    for line, fixture, vals, runs in test_calls(spec):
        args = []
        for v in vals:
            if isinstance(v, Token) and v.type == "NAME" and str(v) in sheets.get(fixture, []):
                args.append(f'ThisWorkbook.Worksheets("{v}")')
            elif isinstance(v, Token) and v.type == "NAME":   # a FIXTURE folder (F6)
                args.append(f'root & "\\{v}"')
            elif isinstance(v, Token) and v.type == "STRING":   # passed as it is: no ' prefix, not a cell
                args.append(vba_text(str(v)[1:-1].replace('""', '"')))
            else:
                args.append(vba_literal(S.to_number(str(v))))
        call = f"{vba_name(proc)}({', '.join(args)})" if args else vba_name(proc)
        out += [f"    ' TEST on spec line {line}"]
        out += [f"    {proc_name(fixture)} ThisWorkbook"] if fixture in sheets else []
        out += [f"    root = {sandbox_name(fixture)}()" if fixture in folders else '    root = ""',
                "    result = Empty",
                f"    For i = 1 To {runs}",
                "        ForgeClearLog",
                f"        result = {call}" if kind == "Function" else f"        Call {call}",
                "    Next i"]
        out += ["    ForgeSandboxUnlock"] if fixture in folders else []
        names = ", ".join(f'"{s}"' for s in sheets.get(fixture, []))
        out += [f"    PilotReport {line}, result, {'True' if kind == 'Function' else 'False'}, root, Array({names})"]
    out += ["End Sub", ""]
    return "\n".join(out), {str(t.children[0]) for t in trees}


RUNNER = r'''Private gOut As Worksheet
Private gLine As Long

Public Sub ForgePilotRun()
    ' Runs every TEST of the piloted specs. Results: the sheet ForgePilot, one line per row.
    On Error GoTo PilotFailed
    Set gOut = FixtureSheet(ThisWorkbook, "ForgePilot")
    gLine = 0
    PilotLine "PILOT {count} spec(s)"
    PilotLine "DATE " & ForgeText(Date)
{specs}
    PilotLine "END"
    gOut.Columns("A").AutoFit
    MsgBox "Pilot done: " & gLine & " lines on the sheet ForgePilot. Copy column A and paste it to Claude."
    Exit Sub
PilotFailed:
    PilotLine "ABORTED error " & Err.Number & ": " & Err.Description
    ForgeSandboxUnlock
    MsgBox "Pilot stopped by error " & Err.Number & ". Copy column A of the sheet ForgePilot to Claude."
End Sub

Private Sub PilotLine(ByVal s As String)
    gLine = gLine + 1
    ForgeSetValue gOut.Cells(gLine, 1), s
End Sub

Private Function PilotEscape(ByVal s As String) As String
    PilotEscape = Replace(Replace(Replace(s, vbLf, "{LF}"), vbCr, "{CR}"), vbTab, "{TAB}")
End Function

Private Function PilotColor(ByVal c As Long) As String
    ' Interior.Color is BGR: #RRGGBB as the interpreter writes fills.
    PilotColor = "#" & Right$("0" & Hex(c And &HFF), 2) & Right$("0" & Hex((c \ &H100) And &HFF), 2) & _
        Right$("0" & Hex((c \ &H10000) And &HFF), 2)
End Function

Private Sub PilotReport(ByVal specLine As Long, ByVal result As Variant, ByVal isFunction As Boolean, _
        ByVal root As String, ByVal sheetNames As Variant)
    ' One TEST's lines: RESULT and ERROR_COUNT, then every value and fill of its sheets, then its files.
    Dim i As Long
    Dim ws As Worksheet
    Dim c As Range
    PilotLine "TEST " & specLine & " RESULT=" & IIf(isFunction, ForgeText(result), "(none)") & " ERRORS=" & ForgeErrorCount()
    For i = LBound(sheetNames) To UBound(sheetNames)
        Set ws = ThisWorkbook.Worksheets(sheetNames(i))
        For Each c In ws.UsedRange.Cells
            If Not IsEmpty(c.Value) Then
                PilotLine "CELL " & ws.Name & " " & c.Row & " " & c.Column & " [" & PilotEscape(ForgeText(c.Value)) & "]"
            End If
            If c.Interior.ColorIndex <> xlColorIndexNone Then
                PilotLine "FILL " & ws.Name & " " & c.Row & " " & c.Column & " " & PilotColor(c.Interior.Color)
            End If
        Next c
    Next i
    If Len(root) > 0 Then PilotFiles root, ""
End Sub

Private Sub PilotFiles(ByVal folderPath As String, ByVal rel As String)
    ' Every folder and file below folderPath, read only. The depth is the sandbox's, a few levels.
    Dim item As Object
    For Each item In SandboxFso().GetFolder(folderPath).SubFolders
        PilotLine "FOLDER " & rel & item.Name
        PilotFiles item.Path, rel & item.Name & "\"
    Next item
    For Each item In SandboxFso().GetFolder(folderPath).Files
        PilotLine "FILE " & rel & item.Name & " " & item.Size
    Next item
End Sub
'''


def pilot_module(specs):
    """One paste-ready module: ForgePilotRun for [(spec file name, spec text)], with its own
    private copies of the fixture builders, so nothing else but forge_module.bas is needed."""
    procs, calls, builders, seen = [], [], [], {}
    for name, text in specs:
        proc, files = spec_procedure(name, text)
        procs.append(proc)
        calls.append(f"    PilotSpec_{name.rsplit('.', 1)[0]}")
        spec, _ = parse(text)
        trees = [k["tree"] for k in spec["keys"] if k["key"] == "FIXTURE"]
        for file in sorted(files):
            mine = [t for t in trees if str(t.children[0]) == file]
            shape = ([parse_fixture(t) for t in mine if t.data != "folder_fixture"],
                     [parse_folder(t) for t in mine if t.data == "folder_fixture"])
            if file in seen:
                if seen[file] != shape:
                    raise ValueError(f"fixture {file} is defined differently in two piloted specs")
                continue
            seen[file] = shape
            if shape[0]:
                builders.append(privately(vba_for_file(file, shape[0])))
            if shape[1]:
                builders.append(privately(vba_for_folders(file, [t for t in mine if t.data == "folder_fixture"])))
    head = ["' VBA Forge pilot: runs the TESTs of " + ", ".join(n for n, _ in specs) + " in Excel and writes",
            "' the results to the sheet ForgePilot. Generated by compiler/vcl_pilot.py; do not edit by hand.",
            "' Paste into a NEW module next to forge_module.bas, then run ForgePilotRun (docs/excel-pilot.md).",
            "Option Explicit", "",
            f"Private Const SANDBOX_MAX_ROOT As Long = {EXCEL_SANDBOX_MAX}",
            f"Private Const SANDBOX_TEST_ROOT As Long = {len(SANDBOX)}   ' the interpreter's {SANDBOX}"]
    runner = RUNNER.replace("{count}", str(len(specs))).replace("{specs}", "\n".join(calls))
    return ("\n".join(head) + "\n" + runner + "\n" + privately(FIXTURE_HELPERS) + "\n"
            + "\n".join(procs) + "\n" + "\n".join(builders) + "\n")


# ---------------------------------------------------------------- what Excel should show, and the check

def escape(text):
    for raw, code in ESCAPES:
        text = text.replace(raw, code)
    return text


def expected_lines(specs, today):
    """The lines a correct Excel run writes, from the interpreter, for [(name, text)] on day today."""
    lines = [f"PILOT {len(specs)} spec(s)", f"DATE {today.isoformat()}"]
    for name, text in specs:
        lines.append(f"SPEC {name}")
        spec, _ = parse(text)
        is_function = any(k["key"] == "KIND" and str(k["tree"].children[0]) == "Function" for k in spec["keys"])
        for r in run_tests(text, today):
            shown = ("" if r.result is None else S.text_of(r.result)) if is_function else "(none)"
            lines.append(f"TEST {r.line} RESULT={shown} ERRORS={r.errors}")
            for sheet, (cells, fills, _) in r.state.items():
                lines += [f"CELL {sheet} {row} {col} [{escape(v)}]" for (row, col), v in cells if v != ""]
                lines += [f"FILL {sheet} {row} {col} {f}" for (row, col), f in fills]
            for path, kind, size, _ in r.files:
                lines.append(f"FOLDER {path}" if kind == "folder" else f"FILE {path} {size}")
    return lines + ["END"]


def read_results(lines):
    """Pasted lines -> (date, {(spec, test line): {"head": ..., "rows": set}}, problems)."""
    date, spec, test, tests, problems = None, None, None, {}, []
    for raw in lines:
        line = raw.strip("\r\n").strip()
        if not line:
            continue
        word = line.split(" ", 1)[0]
        if word == "DATE":
            date = dt.date.fromisoformat(line.split(" ", 1)[1].strip())
        elif word == "SPEC":
            spec, test = line.split(" ", 1)[1].strip(), None
        elif word == "TEST":
            test = (spec, int(line.split()[1]))
            tests[test] = {"head": line, "rows": set()}
        elif word in ("CELL", "FILL", "FILE", "FOLDER"):
            if test is None:
                problems.append(f"line outside a TEST: {line}")
            else:
                tests[test]["rows"].add(line)
        elif word == "ABORTED":
            problems.append(f"Excel stopped: {line}")
    return date, tests, problems


def check(pasted, specs):
    """Differences between the pasted Excel lines and the interpreter, as one line each."""
    date, got, problems = read_results(pasted)
    if date is None:
        return problems + ["no DATE line: paste the whole of column A, from PILOT to END"]
    _, want, _ = read_results(expected_lines(specs, date))
    if not any(line.strip() == "END" for line in pasted):
        problems.append("no END line: the run did not finish, or the paste was cut short")
    for key in sorted(want):
        spec, line = key
        if key not in got:
            problems.append(f"{spec} TEST line {line}: missing from the Excel results")
            continue
        if got[key]["head"] != want[key]["head"]:
            problems.append(f"{spec} TEST line {line}: Excel has {got[key]['head']!r}, "
                            f"the interpreter {want[key]['head']!r}")
        for row in sorted(want[key]["rows"] - got[key]["rows"]):
            problems.append(f"{spec} TEST line {line}: missing in Excel: {row}")
        for row in sorted(got[key]["rows"] - want[key]["rows"]):
            problems.append(f"{spec} TEST line {line}: only in Excel: {row}")
    for key in sorted(set(got) - set(want)):
        problems.append(f"{key[0]} TEST line {key[1]}: in the Excel results but not in the spec")
    return problems


def main(argv):
    if len(argv) < 4 or argv[1] != "check":
        print(__doc__)
        return 1
    pasted = Path(argv[2]).read_text(encoding="utf-8").splitlines()
    specs = [(Path(p).name, Path(p).read_text(encoding="utf-8")) for p in argv[3:]]
    problems = check(pasted, specs)
    for p in problems:
        print(p)
    print("PILOT OK: Excel and the interpreter agree" if not problems else f"PILOT: {len(problems)} differences")
    return 0 if not problems else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
