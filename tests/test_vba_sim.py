"""The VBA simulator (compiler/vba_sim.py, D90 Phase 1, D98).

Two kinds of proof. First, VBA's own rules, case by case, as Microsoft documents them (banker's rounding,
integer division, Str's missing 0, On Error, ByRef...). Second, the whole ForgeLib self-test
(dist/forge_selftest.bas), which Excel runs too: every conformance case must pass in the simulator,
and a bug planted in ForgeLib must make some case fail, so the simulator is not just agreeing.
"""
import datetime as dt
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
sys.path.insert(0, str(ROOT / "tests"))

import vba_sim as V  # noqa: E402
from conformance_cases import CASES  # noqa: E402

TODAY = dt.date(2026, 11, 19)


def run(body, *args, ret="Variant", params=""):
    """Run one Function T with this body; returns T's value."""
    sim = V.Simulator([f"Option Explicit\nPublic Function T({params}) As {ret}\n{body}\nEnd Function\n"], today=TODAY)
    return sim.call("T", *args)


def value(expr):
    return run(f"    T = {expr}")


class VbaRules(unittest.TestCase):
    def test_conversions_round_half_to_even(self):
        self.assertEqual([value("CLng(2.5)"), value("CLng(3.5)"), value("CLng(-2.5)"), value("CInt(0.5)")], [2, 4, -2, 0])
        self.assertEqual([value("Int(-2.5)"), value("Fix(-2.5)"), value("Int(2.5)")], [-3.0, -2.0, 2.0])

    def test_integer_division_and_mod_round_first(self):
        self.assertEqual([value("7 \\ 2"), value("-7 \\ 2"), value("7.6 \\ 2"), value("-7 Mod 3"), value("7.5 Mod 2")],
                         [3, -3, 4, -1, 0])

    def test_text_of_numbers(self):
        self.assertEqual([value("CStr(0.1)"), value("Str$(0.1)"), value("Str$(25)"), value("CStr(1E+16)"),
                          value("CStr(2.5)"), value("Str$(-0.25)")], ["0.1", " .1", " 25", "1E+16", "2.5", "-.25"])

    def test_val_ignores_spaces_and_stops_at_the_first_non_number(self):
        self.assertEqual([value('Val(" 12 3")'), value('Val("1.5e3x")'), value('Val("abc")'), value('Val("&H10")')],
                         [123.0, 1500.0, 0.0, 16.0])

    def test_strings(self):
        self.assertEqual([value('Mid$("abcdef", 2, 3)'), value('Mid$("abc", 5)'), value('InStr(3, "abcabc", "b")'),
                          value('InStrRev("a\\b\\c", "\\")'), value('Left$("ab", 5)'), value('Right$("abc", 0)'),
                          value('Replace("a  b", "  ", " ")'), value('LCase$("ÄB")'), value('Trim$(" x ")')],
                         ["bcd", "", 5, 4, "ab", "", "a b", "äb", "x"])
        self.assertEqual([value('AscW(ChrW(40000))'), value('UBound(Split("", ","))'), value('UBound(Split("a,b", ","))')],
                         [40000 - 65536, -1, 1])

    def test_comparisons(self):
        self.assertEqual([value('"10" < "9"'), value('10 < 9'), value('"a" < "B"'), value('StrComp("a", "B", 1)'),
                          value('"abc" Like "a*"'), value('"a1" Like "a#"'), value('Empty = 0'), value('Empty = ""')],
                         [True, False, False, -1, True, True, True, True])
        self.assertTrue(V.Simulator([""]).compare("=", V.NULL, 1) is V.NULL)

    def test_types(self):
        self.assertEqual([value("VarType(25)"), value("VarType(25#)"), value('VarType("x")'), value("VarType(Empty)"),
                          value("VarType(True)"), value("VarType(DateSerial(2026, 1, 1))"), value("VarType(CVErr(2042))")],
                         [3, 5, 8, 0, 11, 7, 10])

    def test_dates(self):
        d = value("DateSerial(2026, 2, 31)")   # rolls over into March, as VBA's DateSerial does
        self.assertEqual(d.to_datetime().date(), dt.date(2026, 3, 3))
        self.assertEqual([value("Day(DateSerial(2026, 13, 1))"), value("Month(DateSerial(2026, 13, 1))"),
                          value("DateSerial(2026, 3, 1) - DateSerial(2026, 2, 1)")], [1, 1, 28.0])
        self.assertEqual(value("Date").to_datetime().date(), TODAY)

    def test_logic_on_numbers_is_bitwise(self):
        self.assertEqual([value("AscW(ChrW(40000)) And &HFFFF&"), value("6 And 3"), value("Not 0"), value("True And 6")],
                         [40000, 2, -1, 6])

    def test_typed_variables_convert_on_assignment(self):
        self.assertEqual(run("    Dim n As Long\n    n = 2.5\n    T = n"), 2)
        self.assertEqual(run("    Dim s As String\n    s = 2.5\n    T = s"), "2.5")
        with self.assertRaises(V.VbaError) as caught:
            run('    Dim n As Long\n    n = "abc"\n    T = n')
        self.assertEqual(caught.exception.number, 13)

    def test_flow(self):
        body = """    Dim i As Long, s As String
    For i = 5 To 1 Step -2
        s = s & i
    Next i
    Do Until Len(s) > 5
        s = s & "x"
    Loop
    For i = 1 To 10
        If i = 3 Then Exit For
    Next i
    Select Case i
        Case 1, 2
            s = s & "a"
        Case 3 To 4
            s = s & "b"
        Case Is > 4
            s = s & "c"
        Case Else
            s = s & "d"
    End Select
    T = s"""
        self.assertEqual(run(body), "531xxxb")

    def test_on_error_goto_with_erl_and_resume(self):
        body = """    On Error GoTo Handler
10  T = "start"
20  T = 1 / 0
30  T = "not here"
    Exit Function
Handler:
    T = "error " & Err.Number & " at " & Erl
    Resume Done
Done:
    T = T & ", Err is " & Err.Number"""
        self.assertEqual(run(body), "error 11 at 20, Err is 0")

    def test_on_error_resume_next_and_a_raise_from_a_callee(self):
        code = """Public Function T() As String
    On Error Resume Next
    Boom
    T = "after " & Err.Number & " " & Err.Description
End Function
Private Sub Boom()
    Err.Raise 1004, "Boom", "bad"
    MsgBox "never"
End Sub
"""
        sim = V.Simulator([code])
        self.assertEqual(sim.call("T"), "after 1004 bad")
        self.assertEqual(sim.messages, [])

    def test_an_error_without_a_handler_stops_the_run(self):
        with self.assertRaises(V.VbaError) as caught:
            value("1 / 0")
        self.assertEqual(caught.exception.number, 11)

    def test_byref_and_byval(self):
        code = """Public Function T() As String
    Dim a As String, b As String
    a = "a": b = "b"
    Change a, b
    T = a & b
End Function
Private Sub Change(ByRef x As String, ByVal y As String)
    x = x & "!": y = y & "!"
End Sub
"""
        self.assertEqual(V.Simulator([code]).call("T"), "a!b")

    def test_collections_and_dictionaries(self):
        body = """    Dim c As Collection, d As Object, k As Variant, s As String
    Set c = New Collection
    c.Add "b"
    c.Add "a", Before:=1
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    d("Key") = 1
    d("KEY") = d("key") + 1
    For Each k In d.Keys
        s = s & k & "=" & d(k)
    Next k
    T = c(1) & c(2) & c.Count & " " & s & " " & d.Exists("kEy")"""
        self.assertEqual(run(body), "ab2 Key=2 True")

    def test_excel_keeps_numbers_as_doubles_and_text_after_an_apostrophe(self):
        body = """    ThisWorkbook.Worksheets(1).Cells(1, 1).Value = 5
    ThisWorkbook.Worksheets(1).Cells(2, 1).Value = "'0123"
    ThisWorkbook.Worksheets(1).Cells(3, 1).Value = "12"
    T = VarType(ThisWorkbook.Worksheets(1).Cells(1, 1).Value) & " " & ThisWorkbook.Worksheets(1).Cells(2, 1).Value & " " & VarType(ThisWorkbook.Worksheets(1).Cells(3, 1).Value)"""
        self.assertEqual(run(body), "5 0123 5")

    def test_end_and_used_range(self):
        body = """    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Cells(3, 2).Value = "x"
    ws.Cells(7, 2).Value = "y"
    T = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row & " " & ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column & " " & ws.UsedRange.Row & ":" & ws.UsedRange.Rows.Count"""
        self.assertEqual(run(body), "7 2 3:5")

    def test_unsupported_vba_is_refused_not_guessed(self):
        with self.assertRaises(V.VbaUnsupported):
            V.Simulator(["Public Sub S()\n    With ThisWorkbook\n    End With\nEnd Sub\n"])
        with self.assertRaises(V.VbaUnsupported):
            V.Simulator(["Public Sub S()\n    Kill \"C:\\x\"\nEnd Sub\n"])
        with self.assertRaises(V.VbaUnsupported):
            value("Shell(\"x\")")


class SelfTest(unittest.TestCase):
    """dist/forge_selftest.bas in the simulator, as Excel would run it."""

    def run_selftest(self, module):
        sim = V.Simulator([module, (ROOT / "dist" / "forge_selftest.bas").read_text(encoding="utf-8")], today=TODAY)
        sim.call("ForgeSelfTest")
        return sim, sim.this_workbook.sheet("ForgeSelfTest")

    def test_every_conformance_case_passes(self):
        from vcl_selftest import NO_VBA   # a DOCUMENT's functions have no VBA (D115; pinned in test_selftest.py)
        n = len([c for c in CASES if c[0] not in NO_VBA])
        sim, ws = self.run_selftest((ROOT / "dist" / "forge_module.bas").read_text(encoding="utf-8"))
        failed = [(ws.value(r, 1), ws.value(r, 2), ws.value(r, 3), ws.value(r, 4), ws.value(r, 5))
                  for r in range(2, n + 2) if ws.value(r, 6) != "PASS"]
        self.assertEqual(failed, [])
        self.assertEqual(ws.value(4, 8), f"SELFTEST {n}/{n} FAILED: none")
        self.assertEqual(sim.messages, [f"ForgeLib self-test: {n} of {n} cases pass. "
                                        "Details on the sheet ForgeSelfTest."])

    def test_a_bug_planted_in_forgelib_is_caught(self):
        module = (ROOT / "dist" / "forge_module.bas").read_text(encoding="utf-8")
        for old, new in [("code = 32 Or code = 160 Or", "code = 32 Or"),
                         ('If Left$(s, 1) = "." Then s = "0" & s', "s = s"),
                         ("If Day(result) <> CLng(dd) Then Exit Function", ""),
                         ("ForgeMod = a - b * Int(a / b)", "ForgeMod = a Mod b"),
                         ("ElseIf Right$(s, 1) = vbLf Then", "ElseIf Right$(s, 1) = vbCr Then")]:
            with self.subTest(bug=old):
                self.assertEqual(module.count(old), 1)
                _, ws = self.run_selftest(module.replace(old, new))
                self.assertNotIn("FAILED: none", ws.value(4, 8))


if __name__ == "__main__":
    unittest.main()
