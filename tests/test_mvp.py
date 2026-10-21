"""Regression tests for the MVP toolchain: interpreter, compiler, VBA lint, mutation scoring."""
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
from vba_lint import lint  # noqa: E402
from vcl_codegen import CompileError, assemble, compile_text  # noqa: E402
from vcl_interp import run_tests  # noqa: E402
from vcl_mutate import mutate_report  # noqa: E402
from vcl_parse import parse  # noqa: E402
from vcl_types import Checker  # noqa: E402

EXAMPLES = ROOT / "vcl" / "examples" if (ROOT / "vcl" / "examples").exists() else ROOT / "examples"
HEAD = ["VCL 0.1", "TASK", "PROC     P", "KIND     Function", "MODULE   m", "ARGS     ws As Worksheet",
        "RET      Long", "VAR      n As Long", "PRE      COL($ws, \"Status\") > 0"]
TAIL = ["PERF     screen_off", "ERR      standard",
        "FIXTURE  t_01.xlsx sheet=S cols=A:Status,B:Qty rows=6 data=A:values(\"Open\"x3,\"Closed\"x3);B:int(1,9)",
        "FIXTURE  t_01.xlsx sheet=Other cols=A:Note rows=1 data=A:text(3)"]


def mini(ops, tests):
    return "\n".join(HEAD + ["OP       " + o for o in ops] + TAIL + ["TEST     " + t for t in tests])


def all_specs():
    return sorted((ROOT / "specs").glob("*.vcl")) + sorted(EXAMPLES.glob("*.vcl"))


class Interpreter(unittest.TestCase):
    def test_every_spec_passes_its_tests(self):
        for p in all_specs():
            if p.name == "build_example.vcl":
                continue
            with self.subTest(spec=p.name):
                self.assertTrue(all(r.passed for r in run_tests(p.read_text())))

    def test_pre_failure_returns_minus_one_and_changes_nothing(self):
        text = mini(["LET(n, 1)", "RETURN(n)"],
                    ["fixture=t_01.xlsx call=P(Other) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(Other!A:A) = PRE_COUNTA(Other!A:A)"])
        self.assertTrue(run_tests(text)[0].passed)

    def test_row_or_column_below_one_is_an_error_like_excel(self):
        text = mini(["LET(n, 0)", "SET_VALUE(CELL($ws, 1, n), 5)", "RETURN(1)"],
                    ["fixture=t_01.xlsx call=P(S) assert=RESULT = -2 AND ERROR_COUNT = 1"])
        self.assertTrue(run_tests(text)[0].passed)


    def test_exit_if_leaves_the_procedure_not_just_the_loop(self):
        # RESULT = 0, not 5: RETURN(n) is never reached, and a Function then returns its default.
        text = mini(["LET(n, 5)", "LOOP_ROWS dir=down from=2 to=4 var=r", "  EXIT_IF r = 3",
                     "  SET_VALUE(CELL($ws, r, 3), 1)", "RETURN(n)"],
                    ["fixture=t_01.xlsx call=P(S) assert=RESULT = 0 AND COUNTA(S!C:C) = 1 AND ERROR_COUNT = 0"])
        self.assertTrue(run_tests(text)[0].passed)
        self.assertIn("If (r = 3) Then GoTo CleanExit", compile_text(text))

    def test_fill_names_are_their_colours(self):
        # fill=red is the same fill as fill=#FF0000, so FILL_COUNT(..., #FF0000) counts it.
        text = mini(["FORMAT(CELL($ws, 2, 1), fill=red)", "FORMAT(CELL($ws, 3, 1), fill=#FF0000)",
                     "FORMAT(CELL($ws, 4, 1), fill=yellow)", "RETURN(1)"],
                    ["fixture=t_01.xlsx call=P(S) assert=FILL_COUNT(S!A:A, #FF0000) = 2 AND FILL_COUNT(S!A4, #FFFF00) = 1"])
        self.assertTrue(run_tests(text)[0].passed)


class Compiler(unittest.TestCase):
    def test_compiles_what_type_checks_and_refuses_the_rest(self):
        for p in all_specs():
            if p.name in ("build_example.vcl", "example4_build.vcl"):
                continue   # BUILDs compile with compile_build (tests/test_vcl05.py)
            text = p.read_text()
            spec, _ = parse(text)
            with self.subTest(spec=p.name):
                if Checker(spec).run():
                    with self.assertRaises(CompileError):
                        compile_text(text)
                else:
                    self.assertIn("End Function", compile_text(text))

    def test_fill_names_compile_to_their_colours(self):
        vba = compile_text(mini(["FORMAT(CELL($ws, 2, 1), fill=red)", "FORMAT(CELL($ws, 3, 1), fill=cyan)",
                                 "RETURN(1)"], ["fixture=t_01.xlsx call=P(S) assert=RESULT = 1"]))
        self.assertIn("ws.Cells(2, 1).Interior.Color = RGB(255, 0, 0)", vba)
        self.assertIn("ws.Cells(3, 1).Interior.Color = RGB(0, 255, 255)", vba)

    def test_every_fill_name_has_a_colour(self):
        sys.path.insert(0, str(ROOT / "vcl"))
        import semantics
        from validator import FILL_NAMES
        self.assertEqual(FILL_NAMES, ("none",) + tuple(semantics.FILL_COLORS))

    def test_module_lints_clean(self):
        procs = []
        for p in all_specs():
            try:
                procs.append(compile_text(p.read_text(), p.name))
            except CompileError:
                pass
        self.assertGreaterEqual(len(procs), 10)
        self.assertEqual(lint(assemble(procs)), [])

    def test_house_style(self):
        vba = compile_text(mini(['LOOP_ROWS dir=up from=LASTROW($ws, 1) to=2 var=r',
                                 '  IF CELL($ws, r, 1) = "Open" THEN LET(n, n + 1)',
                                 '  IF IS_NUMBER(CELL($ws, r, 2)) THEN LET(n, n + CELL($ws, r, 2))',
                                 'RETURN(n)'],
                                ["fixture=t_01.xlsx call=P(S) assert=RESULT >= 3 AND ERROR_COUNT = 0"]))
        self.assertIn("Step -1", vba)                                   # dir=up
        self.assertIn('ForgeEq(ws.Cells(r, 1).Value, "Open")', vba)     # S7 text equality
        self.assertIn("ForgeNum(ws.Cells(r, 2).Value)", vba)            # guarded number
        self.assertIn("ForgeRestore", vba)                              # state always restored
        self.assertIn("P = -1", vba)                                    # PRE failure code
        self.assertIn("P = -2", vba)                                    # runtime error code


class Lint(unittest.TestCase):
    def test_catches_planted_compile_errors(self):
        procs = [compile_text(p.read_text(), p.name) for p in all_specs()
                 if p.name in ("task_01.vcl", "example1_rev3.vcl")]
        m = assemble(procs)
        faults = [m.replace("End If\n", "", 1), m.replace("ForgeTrim(", "ForgeTirm(", 1),
                  m.replace("outRow = 2", "outRw = 2", 1), m.replace("Next r", "Next k", 1),
                  m.replace("Option Explicit\n", "", 1)]
        for bad in faults:
            self.assertNotEqual(bad, m)
            self.assertTrue(lint(bad))


class Mutation(unittest.TestCase):
    def test_snapshot_checking_is_at_least_as_strong(self):
        text = (ROOT / "specs" / "task_05.vcl").read_text()
        k1, t1, _ = mutate_report(text)
        k2, t2, _ = mutate_report(text, snapshot=True)
        self.assertEqual(t1, t2)
        self.assertGreaterEqual(k2, k1)
        self.assertGreater(t1, 10)


if __name__ == "__main__":
    unittest.main()
