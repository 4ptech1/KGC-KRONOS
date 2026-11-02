"""compiler/vcl_pilot.py: the Excel pilot of task_12 and the check of its pasted results (D35)."""
import datetime as dt
import re
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
from vba_lint import lint, public_helpers  # noqa: E402
from vcl_parse import parse  # noqa: E402
from vcl_pilot import PILOT_SPECS, check, expected_lines, not_piloted, pilot_module, unquote  # noqa: E402

FORGELIB = (ROOT / "vba" / "ForgeLib.bas").read_text(encoding="utf-8")
SPECS = [(name, (ROOT / "specs" / name).read_text(encoding="utf-8")) for name in PILOT_SPECS]
DAY = dt.date(2026, 10, 20)


class Module(unittest.TestCase):
    MODULE = pilot_module(SPECS)

    def test_the_router_is_piloted(self):
        self.assertEqual(PILOT_SPECS, ("task_12.vcl",))

    def test_lint_clean_and_self_contained(self):
        self.assertEqual(lint(self.MODULE, public_helpers(FORGELIB)), [])
        self.assertEqual(re.findall(r"^Public (?:Sub|Function) (\w+)", self.MODULE, re.M), ["ForgePilotRun"])
        for name in ("ForgeFixture_jobpacks_01", "ForgeSandbox_jobpacks_01", "ForgeSandboxUnlock", "FixtureSheet"):
            self.assertRegex(self.MODULE, rf"(?m)^Private (Sub|Function) {name}\(")

    def test_each_test_runs_as_the_interpreter_runs_it(self):
        calls = re.findall(r"^        result = (RouteJobPacks\(.*\))$", self.MODULE, re.M)
        self.assertEqual(calls, [
            'RouteJobPacks(ThisWorkbook.Worksheets("Tracker"), root & "\\Dump", root & "\\Jobs")',
            'RouteJobPacks(ThisWorkbook.Worksheets("Tracker"), root & "\\Dump", root & "\\Jobs")',
            'RouteJobPacks(ThisWorkbook.Worksheets("Tracker"), root & "\\Dump", "X:\\KRONOS_NO_SUCH_FOLDER")',
            'RouteJobPacks(ThisWorkbook.Worksheets("Notes"), root & "\\Dump", root & "\\Jobs")'])
        self.assertEqual(re.findall(r"^    For i = 1 To (\d+)$", self.MODULE, re.M), ["1", "2", "1", "1"])
        spec = self.MODULE[self.MODULE.index("Private Sub PilotSpec_task_12()"):]
        spec = spec[:spec.index("End Sub")]
        self.assertEqual(spec.count("    root = ForgeSandbox_jobpacks_01()\n"), 4)   # a new sandbox each TEST
        self.assertEqual(spec.count("    ForgeSandboxUnlock\n"), 4)                 # locks released each TEST
        self.assertEqual(spec.count("        ForgeClearLog\n"), 4)                  # ERROR_COUNT of the last run
        self.assertEqual(re.findall(r"^    PilotReport (\d+),", self.MODULE, re.M), ["43", "44", "45", "46"])

    def test_it_runs_only_in_a_workbook_with_nothing_else(self):   # FixtureSheet empties sheets (D81)
        entry = self.MODULE[self.MODULE.index("Public Sub ForgePilotRun()"):]
        entry = entry[:entry.index("End Sub")]
        guard = entry.index('If Not PilotBookIsSafe(Array("Tracker", "Notes")) Then')
        self.assertLess(guard, entry.index('Set gOut = FixtureSheet(ThisWorkbook, "ForgePilot")'))
        self.assertLess(guard, entry.index("PilotSpec_task_12"))
        self.assertIn("Exit Sub", entry[guard:entry.index("End If", guard)])
        safe = self.MODULE[self.MODULE.index("Private Function PilotBookIsSafe("):]
        safe = safe[:safe.index("End Function")]
        self.assertIn('Application.WorksheetFunction.CountA(ws.Cells) > 0 Or ws.UsedRange.Address(False, False) <> "A1"',
                      safe)   # anything on a sheet, a value or only formatting, makes it someone's sheet
        self.assertIn("If ranBefore Then", safe)   # practice sheets count as the pilot's own only after a run
        self.assertIn('If Left$(ws.Cells(1, 1).Text, 6) = "PILOT " Then ranBefore = True', safe)   # a real marker
        self.assertIn('If StrComp(ws.Name, "ForgePilot", vbTextCompare) = 0 Then ours = ranBefore', safe)
        self.assertNotRegex(safe, r"(?i)\b(Delete|Clear|Remove)\b")   # the check itself changes nothing

    def test_other_words_for_kronos(self):   # KRONOS's practice run (D81)
        words = {"entry": "TryOnPracticeData", "done_words": "Practice run done", "stopped_words": "It stopped",
                 "paste_to": "into KRONOS", "head": ["' a KRONOS practice run"]}
        module = pilot_module(SPECS, words)
        self.assertEqual(lint(module, public_helpers(FORGELIB)), [])
        self.assertEqual(re.findall(r"^Public (?:Sub|Function) (\w+)", module, re.M), ["TryOnPracticeData"])
        self.assertIn('MsgBox "Practice run done: " & gLine & " lines on the sheet ForgePilot. Copy column A and '
                      'paste it into KRONOS."', module)
        self.assertIn("' a KRONOS practice run\nOption Explicit", module)
        self.assertNotIn("{", module.replace("{LF}", "").replace("{CR}", "").replace("{TAB}", ""))

    def test_what_cannot_be_piloted_is_said_in_plain_words(self):   # D81
        spec = (ROOT / "vcl" / "examples" / "example1_rev3.vcl").read_text(encoding="utf-8")
        parsed, _ = parse(spec)
        self.assertEqual(not_piloted(parsed, spec), "")
        book = (ROOT / "data" / "verified" / "c028" / "spec.vcl").read_text(encoding="utf-8")   # a Workbook argument
        whole = spec.replace("OP       LET(removed, 0)", "OP       LET(removed, SHEETS(THIS_BOOK))")
        for text, words in ((book, "give the macro a whole workbook"), (whole, "looks at the whole workbook")):
            parsed, problems = parse(text)
            self.assertEqual(problems, [])
            self.assertIn(words, not_piloted(parsed, text))
            with self.assertRaisesRegex(ValueError, words):
                pilot_module([("a.vcl", text)])

    def test_excel_quotes_a_pasted_cell_that_holds_a_quote(self):   # D81
        self.assertEqual(unquote('"CELL Orders 3 3 [5"" pipe]"'), 'CELL Orders 3 3 [5" pipe]')
        self.assertEqual(unquote("CELL Orders 3 3 [5 pipe]"), "CELL Orders 3 3 [5 pipe]")
        self.assertEqual(unquote('"'), '"')
        lines = list(Check.GOOD)
        lines[lines.index("CELL Tracker 2 3 [Copied]")] = '"CELL Tracker 2 3 [Copied]"'   # as pasted, wrapped
        self.assertEqual(check(lines, SPECS), [])
        lines[lines.index('"CELL Tracker 2 3 [Copied]"')] = '"CELL Tracker 2 3 [Co""pied]"'
        self.assertEqual(check(lines, SPECS), ["task_12.vcl TEST line 43: missing in Excel: CELL Tracker 2 3 [Copied]",
                                               'task_12.vcl TEST line 43: only in Excel: CELL Tracker 2 3 [Co"pied]'])

    def test_the_file_walk_only_reads(self):
        body = self.MODULE[self.MODULE.index("Private Sub PilotFiles("):]
        body = body[:body.index("End Sub")]
        self.assertNotRegex(body, r"(?i)\b(Create|Copy|Delete|Move|Open|Write)")

    def test_the_committed_module_is_up_to_date(self):
        committed = (ROOT / "dist" / "forge_pilot.bas").read_text(encoding="utf-8")
        self.assertTrue(committed == self.MODULE, "dist/forge_pilot.bas is out of date: run "
                                                  "python compiler/vclc.py all specs/*.vcl -o dist and commit it")


class Check(unittest.TestCase):
    GOOD = expected_lines(SPECS, DAY)

    def test_what_a_correct_run_shows(self):
        self.assertEqual(self.GOOD[:4], ["PILOT 1 spec(s)", "DATE 2026-10-20", "SPEC task_12.vcl",
                                         "TEST 43 RESULT=3 ERRORS=0"])
        self.assertIn("CELL Tracker 2 3 [Copied]", self.GOOD)
        self.assertIn("FILE Jobs\\J-101\\r1.pdf 120", self.GOOD)
        self.assertIn("FOLDER Jobs\\J-103", self.GOOD)
        self.assertIn("TEST 45 RESULT=-1 ERRORS=1", self.GOOD)
        self.assertEqual(check(self.GOOD, SPECS), [])

    def test_a_line_belongs_to_its_test(self):
        lines = list(self.GOOD)
        lines.remove("CELL Tracker 2 3 [Copied]")
        lines.insert(lines.index("TEST 44 RESULT=3 ERRORS=0") + 1, "CELL Tracker 2 3 [Copied]")
        self.assertEqual(check(lines, SPECS), [
            "task_12.vcl TEST line 43: missing in Excel: CELL Tracker 2 3 [Copied]",
            "task_12.vcl TEST line 44: only in Excel: CELL Tracker 2 3 [Copied]"])

    def test_order_and_extra_spaces_dont_matter(self):
        head, body = self.GOOD[:4], self.GOOD[4:]
        test_43 = body[:body.index(next(x for x in body if x.startswith("TEST 44")))]
        shuffled = head + list(reversed(test_43)) + body[len(test_43):]
        self.assertEqual(check(["  " + x + " \r" for x in shuffled], SPECS), [])

    def test_every_kind_of_difference_is_reported(self):
        def edit(old, new):
            lines = list(self.GOOD)
            lines[lines.index(old)] = new
            return check(lines, SPECS)
        self.assertEqual(edit("CELL Tracker 2 3 [Copied]", "CELL Tracker 2 3 [Failed]"),
                         ["task_12.vcl TEST line 43: missing in Excel: CELL Tracker 2 3 [Copied]",
                          "task_12.vcl TEST line 43: only in Excel: CELL Tracker 2 3 [Failed]"])
        self.assertEqual(edit("TEST 45 RESULT=-1 ERRORS=1", "TEST 45 RESULT=-1 ERRORS=2"),
                         ["task_12.vcl TEST line 45: Excel has 'TEST 45 RESULT=-1 ERRORS=2', the interpreter "
                          "'TEST 45 RESULT=-1 ERRORS=1'"])
        self.assertEqual(len(edit("FILE Jobs\\J-101\\r1.pdf 120", "FILE Jobs\\J-101\\r1.pdf 0")), 2)
        self.assertIn("Excel stopped: ABORTED error 70: Permission denied",
                      check(self.GOOD[:-1] + ["ABORTED error 70: Permission denied"], SPECS))
        self.assertIn("no END line: the run did not finish, or the paste was cut short", check(self.GOOD[:-1], SPECS))
        self.assertEqual(check([x for x in self.GOOD if not x.startswith("DATE")], SPECS)[-1],
                         "no DATE line: paste the whole of column A, from PILOT to END")

    def test_the_command_line(self):
        with tempfile.TemporaryDirectory() as tmp:
            pasted = Path(tmp) / "results.txt"
            pasted.write_text("\r\n".join(self.GOOD) + "\r\n", encoding="utf-8")
            run = subprocess.run([sys.executable, "compiler/vclc.py", "pilot-check", str(pasted), "specs/task_12.vcl"],
                                 cwd=ROOT, capture_output=True, text=True)
            self.assertEqual(run.returncode, 0, run.stdout + run.stderr)
            self.assertIn("PILOT OK", run.stdout)
            pasted.write_text("\n".join(x.replace("[Copied]", "[Copy]") for x in self.GOOD), encoding="utf-8")
            run = subprocess.run([sys.executable, "compiler/vclc.py", "pilot-check", str(pasted), "specs/task_12.vcl"],
                                 cwd=ROOT, capture_output=True, text=True)
            self.assertEqual(run.returncode, 1)


if __name__ == "__main__":
    unittest.main()
