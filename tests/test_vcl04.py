"""VEP 0.4 (docs/vep-0.4-rev3-rest.md). Phase 1: the interpreter runs what the compiler already
compiled: INSERT_ROW, READ_ARRAY, WRITE_ARRAY, ARR, ROWS, COLS, SELECT_CASE, LOG, MSG, RANGE, SHEET,
FILE_NAME and FOR_EACH over SHEETS and CELLS (D36, D37). Phase 2: the VCL 0.4 line, lists in
DEDUPE cols=[...] and DEDUPE itself (S21, D38)."""
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
import validator  # noqa: E402
from vba_lint import lint  # noqa: E402
from vcl_codegen import assemble, compile_text  # noqa: E402
from vcl_interp import run_tests  # noqa: E402
from vcl_mutate import mutate_report  # noqa: E402
from vcl_types import check_text  # noqa: E402

FORGELIB = (ROOT / "vba" / "ForgeLib.bas").read_text(encoding="utf-8")

# Arrays, SELECT_CASE, INSERT_ROW and LOG: grade each weld from its Result, on a copy in Report.
GRADE = r"""VCL 0.3
TASK
# ASSUME: Welds headers (row 1) include "Result" and "Grade"
# ASSUME: Result is matched as typed, after TRIM; anything not listed gets "Unknown"
PROC     GradeWelds
KIND     Function
MODULE   modGrades
ARGS     ws As Worksheet, out As Worksheet
RET      Long
VAR      cResult As Long
VAR      cGrade As Long
VAR      data As Variant
VAR      n As Long
PRE      COL($ws, "Result") > 0; COL($ws, "Grade") > 0
MODIFIES $out
OP       LET(cResult, COL($ws, "Result"))
OP       LET(cGrade, COL($ws, "Grade"))
OP       CLEAR(USED_RANGE($out))
OP       READ_ARRAY(USED_RANGE($ws), into=data)
OP       LOOP_ROWS dir=down from=2 to=ROWS(data) var=r
OP         SELECT_CASE TRIM(ARR(data, r, cResult))
OP           CASE "Accepted"
OP             SET_VALUE(ARR(data, r, cGrade), "A")
OP           CASE "Rejected", "Repair"
OP             SET_VALUE(ARR(data, r, cGrade), "R-1")
OP             LET(n, n + 1)
OP           CASE_ELSE
OP             SET_VALUE(ARR(data, r, cGrade), "Unknown")
OP       WRITE_ARRAY(data, to=CELL($out, 1, 1))
OP       INSERT_ROW($out, 1)
OP       SET_VALUE(CELL($out, 1, 1), "Graded")
OP       LOG("graded " & n)
OP       RETURN(n)
ERR      standard
FIXTURE  grades_01.xlsx sheet=Welds cols=A:Seam,B:Result,C:Grade rows=8 data=A:seq(1,1);B:values("Accepted"x3,"Rejected"x2,"Repair"x1,"accepted"x1,""x1)+boundary("Accepted","Rejected","Repair","accepted","")
FIXTURE  grades_01.xlsx sheet=Report cols=A:Old rows=0 stale=3
TEST     fixture=grades_01.xlsx call=GradeWelds(Welds, Report) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(Report!C:C, "R-1") = 3 AND COUNTIF(Report!C:C, "A") = 3 AND COUNTIF(Report!C:C, "Unknown") = 2 AND VALUE(Report!A1) = "Graded" AND VALUE(Report!A2) = "Seam"
TEST     fixture=grades_01.xlsx call=GradeWelds(Welds, Report) assert=COUNTA(Welds![Result]) = PRE_COUNTA(Welds![Result]) AND COUNTA(Welds![Grade]) = PRE_COUNTA(Welds![Grade])
TEST     fixture=grades_01.xlsx call=GradeWelds(Report, Report) assert=RESULT = -1 AND ERROR_COUNT = 1
"""

# FOR_EACH over SHEETS and CELLS, RANGE, SHEET, FILE_NAME and MSG.
MARKS = r"""VCL 0.3
TASK
# ASSUME: a sheet "has marks" when its row 1 includes the header "Mark"
# ASSUME: the marks to count are in Checks!B2:B6, matched with SAME; D2 gets the file name of the path in C2
PROC     CountMarks
KIND     Function
MODULE   modMarks
ARGS     ws As Worksheet
RET      Long
VAR      sheetsWithMarks As Long
VAR      ticks As Long
VAR      notes As Long
PRE      COL($ws, "Mark") > 0
MODIFIES $ws
OP       FOR_EACH var=sh in=SHEETS(THIS_BOOK)
OP         IF COL(sh, "Mark") > 0 THEN LET(sheetsWithMarks, sheetsWithMarks + 1)
OP       FOR_EACH var=v in=CELLS(RANGE($ws, "B2:B6"))
OP         IF SAME(v, "x") THEN LET(ticks, ticks + 1)
OP       LET(notes, LASTROW(SHEET(THIS_BOOK, "Notes"), 1))
OP       SET_VALUE(RANGE($ws, "D2"), FILE_NAME(TRIM(CELL($ws, 2, 3))))
OP       MSG("ticks: " & ticks)
OP       RETURN(sheetsWithMarks * 100 + ticks * 10 + notes)
ERR      standard
FIXTURE  marks_01.xlsx sheet=Checks cols=A:Item,B:Mark,C:Path,D:File rows=6 data=A:seq(1,1);C:values("C:\Docs\WPS-7.pdf"x1,"D:\a\b.pdf"x5)+boundary("C:\Docs\WPS-7.pdf");B:values("x"x2," X "x1,"o"x2,""x1)+boundary("x"," X ","o","x","o","")
FIXTURE  marks_01.xlsx sheet=Notes cols=A:Note rows=3 data=A:text(5)
FIXTURE  marks_01.xlsx sheet=Spare cols=A:Mark rows=2 data=A:values("x"x2)
TEST     fixture=marks_01.xlsx call=CountMarks(Checks) runs=2 assert=RESULT = 234 AND ERROR_COUNT = 0 AND VALUE(Checks!D2) = "WPS-7.pdf"
TEST     fixture=marks_01.xlsx call=CountMarks(Checks) assert=COUNTA(Checks![Mark]) = PRE_COUNTA(Checks![Mark]) AND COUNTA(Checks![File]) = PRE_COUNTA(Checks![File]) + 1
TEST     fixture=marks_01.xlsx call=CountMarks(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
"""

# DEDUPE with a list of key columns (VCL 0.4): drop repeated welds, never a row with a blank key.
DEDUPE = r"""VCL 0.4
TASK
# ASSUME: a weld is listed twice when Seam and Welder repeat (case ignored, as typed); the first listing stays
# ASSUME: a row with a blank Seam or Welder is never removed
PROC     DropRepeatedWelds
KIND     Function
MODULE   modDedupe
ARGS     ws As Worksheet
RET      Long
VAR      before As Long
PRE      COL($ws, "Seam") > 0; COL($ws, "Welder") > 0
MODIFIES $ws
OP       LET(before, LASTROW($ws, 3))
OP       DEDUPE(USED_RANGE($ws), cols=[COL($ws, "Seam"), COL($ws, "Welder")], header=yes)
OP       RETURN(before - LASTROW($ws, 3))
ERR      standard
FIXTURE  welds_01.xlsx sheet=Welds cols=A:Seam,B:Welder,C:Note rows=8 data=C:seq(1,1);A:values("S1"x2,"s1"x1,"S2"x1,"S1 "x1,""x2,"S3"x1)+boundary("S1","S1","s1","S2","S1 ","","","S3");B:values("W1"x7,""x1)+boundary("W1","W1","W1","W1","W1","W1","W1","")
FIXTURE  welds_01.xlsx sheet=Other cols=A:Item rows=2 data=A:seq(1,1)
TEST     fixture=welds_01.xlsx call=DropRepeatedWelds(Welds) assert=RESULT = 2 AND ERROR_COUNT = 0 AND VALUE(Welds!C2) = 1 AND VALUE(Welds!C3) = 4 AND VALUE(Welds!A4) = "S1 " AND VALUE(Welds!C7) = 8
TEST     fixture=welds_01.xlsx call=DropRepeatedWelds(Welds) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTA(Welds!C:C) = PRE_COUNTA(Welds!C:C) - 2
TEST     fixture=welds_01.xlsx call=DropRepeatedWelds(Other) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(Other!A:A) = PRE_COUNTA(Other!A:A)
"""

FIXTURE = 'FIXTURE  t_01.xlsx sheet=Data cols=A:Key,B:Qty rows=4 data=A:values("a"x2,"b"x2);B:int(1,9)'


def ops(*lines, head=(), version="VCL 0.3"):
    """A small spec around OP lines, with one TEST, to watch the interpreter."""
    return "\n".join([version, "TASK", "PROC     P", "KIND     Function", "MODULE   modP",
                      "ARGS     ws As Worksheet", "RET      Long", "VAR      n As Long", *head, "MODIFIES $ws"]
                     + ["OP       " + o for o in lines]
                     + ["OP       RETURN(n)", "ERR      standard", FIXTURE,
                        "TEST     fixture=t_01.xlsx call=P(Data) assert=ERROR_COUNT = 0"])


def run(text):
    (result,) = run_tests(text)
    return result


def cells(result, sheet="Data"):
    return dict(result.state[sheet][0])


class Samples(unittest.TestCase):
    def test_they_validate_pass_and_catch_every_mutant(self):
        for name, text in (("GRADE", GRADE), ("MARKS", MARKS), ("DEDUPE", DEDUPE)):
            with self.subTest(name):
                self.assertEqual([p.message for p in validator.validate_text(text)], [])
                results = run_tests(text)
                self.assertTrue(all(r.passed for r in results), [(r.line, r.failed) for r in results])
                killed, total, survivors = mutate_report(text, snapshot=True)
                self.assertEqual((killed, survivors), (total, []))

    def test_they_compile_lint_clean(self):
        vba = compile_text(GRADE) + compile_text(MARKS)
        self.assertEqual(lint(assemble([compile_text(GRADE), compile_text(MARKS)])), [])
        self.assertIn("data = ForgeReadArray(ForgeUsedRange(ws))", vba)
        self.assertIn("forgeSel1 = ForgeTrim(data(r, cResult))", vba)     # the selector, evaluated once
        self.assertIn("Dim forgeSel1 As Variant", vba)
        self.assertIn("Select Case True", vba)
        self.assertIn('Case (forgeSel1 = "Rejected"), (forgeSel1 = "Repair")', vba)
        self.assertIn("Case Else", vba)
        self.assertIn("ForgeInsertRow out, 1", vba)
        self.assertIn("ForgeWriteArray data, out.Cells(1, 1)", vba)
        self.assertIn("For Each sh In ThisWorkbook.Worksheets", vba)
        self.assertIn('For Each forgeCell_v In ws.Range("B2:B6").Cells', vba)
        self.assertIn("v = forgeCell_v.Value", vba)                       # the loop variable is the value
        self.assertIn('ForgeLastRow(ThisWorkbook.Worksheets("Notes"), 1)', vba)
        self.assertIn("ForgeFileName(ForgeTrim(ws.Cells(2, 3).Value))", vba)
        vba = compile_text(DEDUPE)
        self.assertEqual(lint(assemble([vba])), [])
        self.assertIn('ForgeDedupe ForgeUsedRange(ws), Array(ForgeCol(ws, "Seam"), ForgeCol(ws, "Welder")), True', vba)

    def test_log_and_msg_are_never_errors(self):
        grade, marks = run_tests(GRADE)[0], run_tests(MARKS)[0]
        self.assertEqual((grade.errors, grade.messages), (0, [("LOG", "graded 3")]))
        self.assertEqual((marks.errors, marks.messages), (0, [("MSG", "ticks: 3")]))


class Interpreter(unittest.TestCase):
    def test_insert_row_is_blank_and_moves_fills(self):
        r = run(ops("FORMAT(ROW_RANGE($ws, 2, 1, 2), fill=red)", "INSERT_ROW($ws, 2)"))
        got = cells(r)
        self.assertEqual((got[(1, 1)], got.get((2, 1)), got[(3, 1)]), ("Key", None, "a"))
        self.assertEqual(sorted(dict(r.state["Data"][1])), [(3, 1), (3, 2)])   # the red row moved; the new one has none

    def test_read_array_is_2d_even_for_one_cell(self):
        text = ops("READ_ARRAY(CELL($ws, 1, 1), into=a)", "LET(n, ROWS(a) * 10 + COLS(a))", head=["VAR      a As Variant"])
        self.assertEqual(run(text).result, 11)

    def test_arr_outside_the_array_is_an_error(self):
        text = ops("READ_ARRAY(CELL($ws, 1, 1), into=a)", "SET_VALUE(ARR(a, 2, 1), 0)", head=["VAR      a As Variant"])
        r = run(text)
        self.assertEqual((r.result, r.errors), (-2, 1))

    def test_write_array_keeps_text_as_text(self):
        text = ops("READ_ARRAY(CELL($ws, 1, 1), into=a)", 'SET_VALUE(ARR(a, 1, 1), "1001")',
                   "WRITE_ARRAY(a, to=CELL($ws, 1, 3))", head=["VAR      a As Variant"])
        self.assertEqual(cells(run(text))[(1, 3)], "1001")
        body = FORGELIB[FORGELIB.index("Public Sub ForgeWriteArray("):]
        self.assertIn("If ForgeNeedsTextGuard(out(r, c)) Then out(r, c) = \"'\" & out(r, c)", body[:body.index("End Sub")])

    def test_select_case_compares_as_equals(self):
        def chosen(selector, *cases):
            lines = ["SELECT_CASE " + selector]
            for i, case in enumerate(cases, start=1):
                lines += ["  " + case, f"    LET(n, {i})"]
            return run(ops(*lines)).result
        self.assertEqual(chosen('"25"', 'CASE "25.0"', "CASE_ELSE"), 2)        # text, not numbers (S7)
        self.assertEqual(chosen('"b"', 'CASE "a", "b"', 'CASE "b"'), 1)        # the first match wins
        self.assertEqual(chosen('"B"', 'CASE "b"'), 0)                         # = is case-sensitive; no CASE_ELSE
        self.assertEqual(chosen("2 + 3", "CASE 4, 5", "CASE_ELSE"), 1)

    def test_select_case_values_are_typed_as_equals(self):
        found = check_text(ops("SELECT_CASE CELL($ws, 2, 2)", "  CASE 5", "    LET(n, 1)"))
        self.assertTrue(any("raw cell value" in m for _, m in found), found)
        self.assertEqual(check_text(ops("SELECT_CASE CELL($ws, 2, 1)", '  CASE "a"', "    LET(n, 1)")), [])

    def test_ranges_inside_the_sheet_only(self):
        self.assertEqual(cells(run(ops('SET_VALUE(RANGE($ws, "C2:$C$3"), 7)')))[(3, 3)], "7")
        for ref in ("A:A", "1:1", "XFE1", "A1048577", "Totals"):
            with self.subTest(ref):
                r = run(ops(f'SET_VALUE(RANGE($ws, "{ref}"), 7)'))
                self.assertEqual((r.result, r.errors), (-2, 1))

    def test_books_and_sheets(self):
        self.assertEqual(run(ops('LET(n, LASTROW(SHEET(THIS_BOOK, "data"), 1))')).result, 5)   # name, any case
        self.assertEqual(run(ops('LET(n, LASTROW(SHEET(THIS_BOOK, "Missing"), 1))')).result, -2)
        self.assertEqual(run(ops("LET(n, LASTROW(ACTIVE_SHEET, 1))")).result, -2)

    def test_a_sheet_loop_variable_is_a_sheet(self):
        text = ops("FOR_EACH var=sh in=SHEETS(THIS_BOOK)", '  IF COL(sh, "Qty") > 0 THEN LET(n, n + 1)')
        self.assertEqual(check_text(text), [])
        self.assertEqual(run(text).result, 1)


def errors(text):
    """The errors on the OP lines (the small specs of ops() break the spec-wide TEST rules on purpose)."""
    first_op = text.splitlines().index(next(line for line in text.splitlines() if line.startswith("OP "))) + 1
    return [p.message for p in validator.validate_text(text) if p.severity == "error" and p.line >= first_op
            and "TEST" not in p.message]


class Dedupe(unittest.TestCase):
    KEYS = 'FIXTURE  t_01.xlsx sheet=Data cols=A:Key,B:Qty rows=6 data=A:values("a"x2,"A"x1,"b"x2,""x1)+boundary("a","A","b","a","","b");B:values("1"x5,"2"x1)+boundary("1","1","1","2","1","1")'

    def run_keys(self, op, version="VCL 0.4"):
        text = ops(op, head=[], version=version).replace(FIXTURE, self.KEYS)
        return run(text)

    def test_the_first_row_of_each_key_stays(self):
        r = self.run_keys("DEDUPE(USED_RANGE($ws), cols=[1, 2], header=yes)")
        self.assertEqual(r.errors, 0)
        rows = [(cells(r).get((i, 1)), cells(r).get((i, 2))) for i in range(1, 7)]
        self.assertEqual(rows, [("Key", "Qty"), ("a", "1"), ("b", "1"), ("a", "2"), (None, "1"), (None, None)])

    def test_one_key_column_and_no_header(self):
        r = self.run_keys("DEDUPE(USED_RANGE($ws), cols=1, header=no)")
        self.assertEqual([cells(r).get((i, 1)) for i in range(1, 6)], ["Key", "a", "b", None, None])

    def test_the_old_form_means_the_same_before_0_4(self):
        old = self.run_keys("DEDUPE(USED_RANGE($ws), cols=1,2, header=yes)", version="VCL 0.3")
        new = self.run_keys("DEDUPE(USED_RANGE($ws), cols=[1, 2], header=yes)")
        self.assertEqual(old.state, new.state)

    def test_a_key_column_beyond_the_sheet_is_an_error(self):
        for cols in ("[1, 16385]", "[0, 1]"):
            with self.subTest(cols):
                r = self.run_keys(f"DEDUPE(USED_RANGE($ws), cols={cols}, header=yes)")
                self.assertEqual((r.result, r.errors), (-2, 1))
        r = self.run_keys("DEDUPE(CELL($ws, 1, 1), cols=[0], header=yes)")   # no row to check: nothing read
        self.assertEqual(r.errors, 0)

    def test_the_validator_and_the_version_line(self):
        dedupe = "DEDUPE(USED_RANGE($ws), {}, header=yes)"
        self.assertIn("DEDUPE cols=[...] is VCL 0.4: start the spec with 'VCL 0.4' to use it",
                      errors(ops(dedupe.format("cols=[1, 2]"))))
        self.assertIn("DEDUPE cols=1,2 is ambiguous (G3): write cols=[1, 2]",
                      errors(ops(dedupe.format("cols=1,2"), version="VCL 0.4")))
        self.assertIn("unbalanced brackets: '[' without ']'", errors(ops(dedupe.format("cols=[1, 2"), version="VCL 0.4")))
        self.assertIn("a list [...] is only allowed as DEDUPE cols=", errors(ops("LET(n, [1, 2])", version="VCL 0.4")))
        for form in ("cols=1", "cols=[1, n + 1]"):
            with self.subTest(form):
                self.assertEqual(errors(ops(dedupe.format(form), version="VCL 0.4")), [])

    def test_key_columns_are_numbers(self):
        found = check_text(ops("DEDUPE(USED_RANGE($ws), cols=[1, CELL($ws, 1, 2)], header=yes)", version="VCL 0.4"))
        self.assertTrue(any("raw cell value" in m for _, m in found), found)


class ForgeLib(unittest.TestCase):
    def body(self, name):
        for kind in ("Function", "Sub"):
            head = f"Public {kind} {name}("
            if head in FORGELIB:
                start = FORGELIB.index(head)
                return FORGELIB[start:FORGELIB.index("\nEnd ", start)]
        raise AssertionError(f"{name} is not in ForgeLib")

    def test_the_new_helpers(self):
        self.assertIn("ws.Rows(rowNo).ClearFormats", self.body("ForgeInsertRow"))
        self.assertIn("If rng.Cells.CountLarge = 1 Then", self.body("ForgeReadArray"))
        self.assertIn('ForgeFileName = Mid$(s, InStrRev(s, "\\") + 1)', self.body("ForgeFileName"))
        dedupe = self.body("ForgeDedupe")
        self.assertLess(dedupe.index("key = ForgeRowKey(parts)"), dedupe.index("Step -1"))   # marked first, deleted after
        self.assertIn("If doomed(r) Then rng.Worksheet.Rows(rng.Row + r - 1).Delete", dedupe)
        self.assertIn("If ForgeIsEmpty(values(i)) Or ForgeIsError(values(i)) Then", self.body("ForgeRowKey"))


if __name__ == "__main__":
    unittest.main()
