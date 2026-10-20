"""Build the ForgeLib self-test: every conformance case, run through ForgeLib in Excel (router Phase 5).

  python compiler/vcl_selftest.py [-o dist]   writes dist/forge_selftest.bas

tests/conformance_cases.py is the contract between vcl/semantics.py and ForgeLib. The unit tests
run it against semantics.py; this module runs the same cases against ForgeLib, so Excel 2016 is
the only thing needed to prove ForgeLib means what the interpreter means. Paste it into a new
module next to forge_module.bas and run ForgeSelfTest: the sheet ForgeSelfTest gets one row per
case (PASS or FAIL, expected and got), and cell H4 a line to paste back (D34).

Each case calls the same ForgeLib expression the compiler emits for that VCL meaning. A case whose
call raises an error is a FAIL with the error, and the run goes on. The cases are fixed (no TODAY),
so the module never changes unless the cases or the mapping below do; CI checks it is up to date.
"""
import datetime as dt
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "vcl"))
sys.path.insert(0, str(ROOT / "tests"))
from semantics import XlError  # noqa: E402
from conformance_cases import CASES  # noqa: E402

CASES_PER_PROC = 40   # keeps each VBA procedure small
MAX_LINE = 1000       # VBA allows 1023 characters per line

# The ForgeLib expression for each semantics function, as the compiler emits it (vcl_codegen.py).
CALLS = {
    "is_empty": "ForgeIsEmpty({0})", "is_number": "ForgeIsNumber({0})", "is_date": "ForgeIsDate({0})",
    "is_error": "ForgeIsError({0})", "trim": "ForgeTrim({0})", "text_of": "ForgeText({0})",
    "upper": "UCase$(ForgeText({0}))", "length": "Len(ForgeText({0}))",
    "left": "ForgeLeft({0}, {1})", "right": "ForgeRight({0}, {1})", "equals": "ForgeEq({0}, {1})",
    "col": "SelfTestCol({0}, {1})", "last_row": "SelfTestLastRow({0})",
    "same": "ForgeSame({0}, {1})", "clean_text": "ForgeCleanText({0})",
    "parse_date": "ForgeParseDate({0}, {1})", "iif": "ForgeIif({0}, {1}, {2})", "mod": "ForgeMod({0}, {1})",
    "needs_text_guard": "ForgeNeedsTextGuard({0})", "safe_name": "ForgeSafeName({0})",
    "file_name": "ForgeFileName({0})",
    "same_row_key": "ForgeSameRowKey({0}, {1})", "find": "SelfTestFind({0}, {1}, {2})",
    "sheet_name_ok": "ForgeSheetNameOk({0})",
    "text_ok": "ForgeIsTextOk({0})", "file_text": "SelfTestFileText({0})",
}
XL_ERRORS = {"#N/A": "xlErrNA", "#DIV/0!": "xlErrDiv0", "#VALUE!": "xlErrValue", "#REF!": "xlErrRef",
             "#NAME?": "xlErrName", "#NUM!": "xlErrNum", "#NULL!": "xlErrNull"}


def vba_text(s):
    """A VBA string expression for s: printable ASCII in quotes, anything else as ChrW(code)."""
    parts, run = [], ""
    for ch in s:
        if 32 <= ord(ch) <= 126:
            run += '""' if ch == '"' else ch
        else:
            if run:
                parts.append(f'"{run}"')
                run = ""
            parts.append(f"ChrW({ord(ch)})")
    if run or not parts:
        parts.append(f'"{run}"')
    return " & ".join(parts)


def vba_value(v):
    """A VBA expression for a conformance value. Numbers are Doubles, as Excel cells hold them."""
    if v is None:
        return "Empty"
    if isinstance(v, bool):
        return "True" if v else "False"
    if isinstance(v, int):
        return f"{v}#"
    if isinstance(v, float):
        if v.is_integer():
            return f"{int(v)}#"
        text = repr(v)
        if "e" in text or "inf" in text or "nan" in text:
            raise ValueError(f"number {v!r} can't be written as a plain VBA literal")
        return text
    if isinstance(v, str):
        return vba_text(v)
    if isinstance(v, dt.datetime):
        return f"(DateSerial({v.year}, {v.month}, {v.day}) + TimeSerial({v.hour}, {v.minute}, {v.second}))"
    if isinstance(v, dt.date):
        return f"DateSerial({v.year}, {v.month}, {v.day})"
    if isinstance(v, XlError):
        return f"CVErr({XL_ERRORS[v.text]})"
    if isinstance(v, list):
        return "Array(" + ", ".join(vba_value(x) for x in v) + ")"
    raise ValueError(f"no VBA form for {v!r}")


def case_lines(n, fn, args, expected, why):
    """The two VBA lines of case n: name it, then check the ForgeLib call against the expected value."""
    if fn == "dict_key":   # S10: a key added once is found again under the expected text, case ignored
        got, want = f"SelfTestDictHas({vba_value(args[0])}, {vba_value(expected)})", "True"
    elif fn == "list_files":   # S24: the names that match, joined by | in their order
        got = f"SelfTestJoin(ForgeMatchNames({vba_value(args[0])}, {vba_value(args[1])}))"
        want = vba_text("|".join(expected))
    elif fn == "path_join":
        got = f"ForgePathJoin({vba_value(args[0])}, Array({', '.join(vba_value(a) for a in args[1:])}))"
        want = vba_value(expected)
    else:
        got, want = CALLS[fn].format(*[vba_value(a) for a in args]), vba_value(expected)
    lines = [f"    SelfTestCase {n}, {vba_text(fn)}, {vba_text(why)}", f"    SelfTestCheck {got}, {want}"]
    for line in lines:
        if len(line) > MAX_LINE:
            raise ValueError(f"case {n} makes a VBA line of {len(line)} characters (limit {MAX_LINE})")
    return lines


HELPERS = '''Private gWs As Worksheet
Private gScratch As Worksheet
Private gCase As Long

Public Sub ForgeSelfTest()
    ' Runs every conformance case through ForgeLib. Results: the sheet ForgeSelfTest.
    Dim total As Long
    Dim passed As Long
    Dim failedList As String
    Dim r As Long
    Set gWs = SelfTestSheet("ForgeSelfTest")
    Set gScratch = SelfTestSheet("ForgeSelfTestScratch")
    gWs.Range("A1:F1").Value = Array("No", "Function", "Case", "Expected", "Got", "Result")
    gCase = 0
{parts}
    SelfTestFlush
    total = {total}
    For r = 2 To total + 1
        If gWs.Cells(r, 6).Value = "PASS" Then
            passed = passed + 1
        Else
            failedList = failedList & " " & gWs.Cells(r, 1).Value
        End If
    Next r
    gWs.Range("H1").Value = "Passed"
    gWs.Range("I1").Value = passed
    gWs.Range("H2").Value = "Cases"
    gWs.Range("I2").Value = total
    gWs.Range("H3").Value = "Paste this line back to Claude:"
    ForgeSetValue gWs.Range("H4"), "SELFTEST " & passed & "/" & total & " FAILED:" & IIf(failedList = "", " none", failedList)
    gWs.Columns("A:F").AutoFit
    MsgBox "ForgeLib self-test: " & passed & " of " & total & " cases pass. Details on the sheet ForgeSelfTest."
End Sub

Private Function SelfTestSheet(ByVal sheetName As String) As Worksheet
    ' The self-test's own sheet in this workbook, added if missing and emptied.
    Dim ws As Worksheet
    Dim found As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then Set found = ws
    Next ws
    If found Is Nothing Then
        Set found = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        found.Name = sheetName
    End If
    found.Cells.Clear
    Set SelfTestSheet = found
End Function

Private Sub SelfTestCase(ByVal n As Long, ByVal fn As String, ByVal why As String)
    ' Starts case n: a FAIL row until SelfTestCheck proves otherwise.
    SelfTestFlush
    gCase = n
    gWs.Cells(n + 1, 1).Value = n
    ForgeSetValue gWs.Cells(n + 1, 2), fn
    ForgeSetValue gWs.Cells(n + 1, 3), why
    ForgeSetValue gWs.Cells(n + 1, 5), "(not run)"
    gWs.Cells(n + 1, 6).Value = "FAIL"
End Sub

Private Sub SelfTestFlush()
    ' A case whose call raised an error: its row shows the error. Err is the caller's.
    If Err.Number <> 0 And gCase > 0 Then
        ForgeSetValue gWs.Cells(gCase + 1, 5), "error " & Err.Number & ": " & Err.Description
    End If
    Err.Clear
End Sub

Private Sub SelfTestCheck(ByVal got As Variant, ByVal want As Variant)
    ForgeSetValue gWs.Cells(gCase + 1, 4), SelfTestShow(want)
    ForgeSetValue gWs.Cells(gCase + 1, 5), SelfTestShow(got)
    If SelfTestSame(got, want) Then gWs.Cells(gCase + 1, 6).Value = "PASS"
End Sub

Private Function SelfTestShow(ByVal v As Variant) As String
    ' How a value looks in the results: its type, then its canonical text.
    If IsError(v) Then
        SelfTestShow = "error " & ForgeText(v)
    ElseIf IsEmpty(v) Then
        SelfTestShow = "(empty)"
    ElseIf VarType(v) = vbString Then
        SelfTestShow = "text [" & v & "]"
    ElseIf VarType(v) = vbBoolean Then
        SelfTestShow = ForgeText(v)
    ElseIf VarType(v) = vbDate Then
        SelfTestShow = "date " & ForgeText(v)
    Else
        SelfTestShow = "number " & ForgeText(v)
    End If
End Function

Private Function SelfTestSame(ByVal got As Variant, ByVal want As Variant) As Boolean
    ' The same value and the same kind of value, as Python's == on the reference results.
    If IsError(got) Or IsError(want) Then
        SelfTestSame = False
    ElseIf IsEmpty(got) Or IsEmpty(want) Then
        SelfTestSame = (IsEmpty(got) And IsEmpty(want))
    ElseIf VarType(got) = vbString Or VarType(want) = vbString Then
        If VarType(got) = vbString And VarType(want) = vbString Then
            SelfTestSame = (StrComp(got, want, vbBinaryCompare) = 0)
        End If
    ElseIf VarType(got) = vbBoolean Or VarType(want) = vbBoolean Then
        If VarType(got) = vbBoolean And VarType(want) = vbBoolean Then SelfTestSame = (got = want)
    ElseIf VarType(got) = vbDate Or VarType(want) = vbDate Then
        If VarType(got) = vbDate And VarType(want) = vbDate Then SelfTestSame = (CDbl(got) = CDbl(want))
    Else
        SelfTestSame = (CDbl(got) = CDbl(want))
    End If
End Function

Private Function SelfTestCol(ByVal headers As Variant, ByVal header As Variant) As Long
    ' COL on a scratch sheet whose row 1 holds these headers.
    Dim i As Long
    gScratch.Cells.Clear
    For i = LBound(headers) To UBound(headers)
        ForgeSetValue gScratch.Cells(1, i - LBound(headers) + 1), headers(i)
    Next i
    SelfTestCol = ForgeCol(gScratch, header)
End Function

Private Function SelfTestLastRow(ByVal values As Variant) As Long
    ' LASTROW of column A on a scratch sheet holding these values from row 1 (Empty: a blank cell).
    Dim i As Long
    gScratch.Cells.Clear
    For i = LBound(values) To UBound(values)
        If Not IsEmpty(values(i)) Then ForgeSetValue gScratch.Cells(i - LBound(values) + 1, 1), values(i)
    Next i
    SelfTestLastRow = ForgeLastRow(gScratch, 1)
End Function

Private Function SelfTestFind(ByVal values As Variant, ByVal x As Variant, ByVal match As String) As Long
    ' FIND on row 1 of a scratch sheet holding these values (Empty: a blank cell): the column found, or 0.
    Dim i As Long
    Dim found As Range
    gScratch.Cells.Clear
    For i = LBound(values) To UBound(values)
        If Not IsEmpty(values(i)) Then ForgeSetValue gScratch.Cells(1, i - LBound(values) + 1), values(i)
    Next i
    Set found = ForgeFind(gScratch.Range(gScratch.Cells(1, 1), gScratch.Cells(1, UBound(values) - LBound(values) + 1)), _
        x, match = "whole")
    If Not found Is Nothing Then SelfTestFind = found.Column
End Function

Private Function SelfTestFileText(ByVal raw As Variant) As Variant
    ' S25: ForgeFileText gives Null for what is not text; the cases write that as Empty.
    SelfTestFileText = ForgeFileText(raw)
    If IsNull(SelfTestFileText) Then SelfTestFileText = Empty
End Function

Private Function SelfTestJoin(ByVal names As Collection) As String
    ' A list's names joined by |, in their order.
    Dim i As Long
    For i = 1 To names.Count
        If i > 1 Then SelfTestJoin = SelfTestJoin & "|"
        SelfTestJoin = SelfTestJoin & names(i)
    Next i
End Function

Private Function SelfTestDictHas(ByVal key As Variant, ByVal want As String) As Boolean
    ' S10: after one DICT_ADD of key, the dictionary holds one key, found again as want.
    Dim d As Object
    Set d = ForgeNewDict()
    ForgeDictAdd d, key, 1, "overwrite"
    SelfTestDictHas = (d.Count = 1 And ForgeDictExists(d, want))
End Function
'''


def selftest_module(cases=CASES):
    """The whole self-test module, ready to paste into Excel."""
    chunks = [cases[i:i + CASES_PER_PROC] for i in range(0, len(cases), CASES_PER_PROC)]
    calls, procs, n = [], [], 0
    for part, chunk in enumerate(chunks, start=1):
        calls.append(f"    SelfTestPart{part}")
        body = [f"Private Sub SelfTestPart{part}()",
                "    ' A call that raises an error goes to the next line; SelfTestFlush records it.",
                "    On Error Resume Next"]
        for fn, args, expected, why in chunk:
            n += 1
            body += case_lines(n, fn, args, expected, why)
        body += ["    SelfTestFlush", "End Sub", ""]
        procs.append("\n".join(body))
    head = ["' VBA Forge: the ForgeLib self-test. Every conformance case (tests/conformance_cases.py) runs",
            "' through ForgeLib in Excel. Generated by compiler/vcl_selftest.py; do not edit by hand.",
            "' Paste into a NEW module next to forge_module.bas, then run ForgeSelfTest.",
            "Option Explicit", ""]
    helpers = HELPERS.replace("{parts}", "\n".join(calls)).replace("{total}", str(n))
    return "\n".join(head) + "\n" + helpers + "\n" + "\n".join(procs)


def main(argv):
    out = Path(argv[argv.index("-o") + 1]) if "-o" in argv else ROOT / "dist"
    out.mkdir(parents=True, exist_ok=True)
    (out / "forge_selftest.bas").write_text(selftest_module(), encoding="utf-8")
    print(f"wrote {out / 'forge_selftest.bas'} ({len(CASES)} cases)")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
