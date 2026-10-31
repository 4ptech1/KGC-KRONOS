"""VEP 0.6 Phase 1 (D60): every message the checks give carries a stable code."""
import os
import subprocess
import sys
import unittest
from pathlib import Path

# From here on every Problem made while the tests run must have a code (errcodes.strict).
os.environ["KGC_STRICT_CODES"] = "1"

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
import gen_error_codes  # noqa: E402
import validator  # noqa: E402
from errcodes import Coded, code_of, coded, tag  # noqa: E402
from parser import Problem  # noqa: E402
from test_mvp import mini  # noqa: E402
from vcl_fixtures import parse_fixture, sheet_rows  # noqa: E402
from vcl_interp import VclError, run_tests  # noqa: E402
from vcl_parse import parse  # noqa: E402

GOAL_TEST = ["fixture=t_01.xlsx call=P(S) assert=RESULT >= 0", "fixture=t_01.xlsx call=P(S) assert=ERROR_COUNT = 0"]


class Catalogue(unittest.TestCase):
    def test_every_message_has_a_code_and_the_catalogue_is_current(self):
        self.assertEqual(gen_error_codes.check(), [])

    def test_codes_have_the_shape_of_their_stage(self):
        registry = gen_error_codes.load_registry()
        self.assertGreater(len(registry), 400)
        for code in registry:
            self.assertRegex(code, r"^[VGTXRK]\d{3}$")

    def test_a_code_in_a_source_must_be_in_the_registry(self):
        scanned = gen_error_codes.scan(ROOT / "compiler")
        found, _ = gen_error_codes.problems_with(scanned, {})
        self.assertTrue(found and all("is not in compiler/error_codes.json" in line for line in found))

    def test_a_reused_code_is_reported(self):
        source = 'from errcodes import coded\nself.error(coded("a", "V001"))\nself.error(coded("b", "V001"))\n'
        path = ROOT / "compiler" / "validator.py"
        sites = gen_error_codes.find_sites(path, source)
        found, _ = gen_error_codes.problems_with({"validator.py": (path, source, sites)}, {"V001": "a"})
        self.assertEqual(found, ["validator.py:3: V001 is already used at validator.py:2"])   # the second use is the one reported


class Messages(unittest.TestCase):
    def test_a_coded_message_is_a_str(self):
        message = coded("unknown key X", "V005")
        self.assertIsInstance(message, str)
        self.assertEqual(message, "unknown key X")
        self.assertEqual(code_of(message), "V005")
        self.assertIsInstance(str(message), Coded)   # str() keeps the code
        self.assertEqual(code_of(f"{message}!"), "")   # text built from it is a new text
        self.assertEqual(tag(message), " [V005]")
        self.assertEqual(tag("plain"), "")

    def test_a_problem_takes_its_code_from_its_message(self):
        problem = Problem(3, "error", coded("unknown key X", "V005"))
        self.assertEqual(problem.code, "V005")
        self.assertEqual(problem, Problem(3, "error", "unknown key X", "V999"))   # the code isn't compared

    def test_a_message_without_a_code_is_a_bug_in_the_tests(self):
        with self.assertRaises(ValueError):
            Problem(1, "error", "no code here")


class Stages(unittest.TestCase):
    """One example per stage: the validator V, the grammar G, the type checker T, the fixture builder X,
    and a TEST run R."""

    def codes(self, text):
        return [(p.code, p.message) for p in validator.validate_text(text)]

    def test_validator_rules_are_v(self):
        found = self.codes("VCL 0.4\nTASK\nPROC     x\n")
        self.assertTrue(found and all(code.startswith("V") for code, _ in found), found)
        self.assertIn("missing KIND", [message for _, message in found])

    def test_the_grammar_is_g(self):
        _, problems = parse("VCL 0.1\nTASK\nBOGUS x\n")
        self.assertEqual([(code_of(message), str(message)) for _, message in problems],
                         [("G005", "unknown key 'BOGUS'")])

    def test_the_type_checker_is_t(self):
        text = mini(["LET(n, CELL($ws, 2, 1) + 1)", "RETURN(n)"], GOAL_TEST)
        found = [(code, message) for code, message in self.codes(text) if message.startswith("type:")]
        self.assertTrue(found, self.codes(text))
        self.assertTrue(all(code.startswith("T") for code, _ in found), found)

    def test_the_fixture_builder_is_x(self):
        text = mini(["RETURN(1)"], GOAL_TEST).replace('values("Open"x3,"Closed"x3)', 'values("Open"x2,"Closed"x2)')
        spec, _ = parse(text)
        tree = next(k["tree"] for k in spec["keys"] if k["key"] == "FIXTURE")
        with self.assertRaises(ValueError) as caught:
            sheet_rows(parse_fixture(tree))
        self.assertRegex(code_of(caught.exception.args[0]), r"^X\d{3}$")

    def test_a_failed_assertion_is_r(self):
        text = mini(["RETURN(1)"], ["fixture=t_01.xlsx call=P(S) assert=RESULT = 2 AND ERROR_COUNT = 0",
                                    "fixture=t_01.xlsx call=P(S) assert=ERROR_COUNT = 0"])
        failed = run_tests(text)[0].failed
        self.assertEqual(len(failed), 1)
        self.assertRegex(code_of(failed[0]), r"^R\d{3}$")
        self.assertIn("is false", failed[0])

    def test_a_run_that_cannot_go_on_is_r(self):
        text = mini(["RETURN(1)"], ["fixture=missing_01.xlsx call=P(S) assert=RESULT = 1",
                                    "fixture=t_01.xlsx call=P(S) assert=ERROR_COUNT = 0"])
        with self.assertRaises(VclError) as caught:
            run_tests(text)
        self.assertRegex(code_of(caught.exception.args[0]), r"^R\d{3}$")


class CommandLine(unittest.TestCase):
    def test_the_validator_shows_the_code_after_the_message(self):
        bad = ROOT / "tests" / "data" / "bad_for_codes.vcl"
        bad.write_text("VCL 0.4\nTASK\nPROC     x\n", encoding="utf-8")
        try:
            out = subprocess.run([sys.executable, str(ROOT / "compiler" / "validator.py"), str(bad)],
                                 capture_output=True, text=True, cwd=ROOT).stdout
        finally:
            bad.unlink()
        self.assertRegex(out, r"error: missing KIND \[V\d{3}\]")

    def test_vclc_shows_the_code_of_a_failed_check(self):
        bad = ROOT / "tests" / "data" / "bad_for_codes.vcl"
        bad.write_text("VCL 0.4\nTASK\nPROC     x\n", encoding="utf-8")
        try:
            out = subprocess.run([sys.executable, str(ROOT / "compiler" / "vclc.py"), "check", str(bad)],
                                 capture_output=True, text=True, cwd=ROOT).stdout
        finally:
            bad.unlink()
        self.assertRegex(out, r"error: missing KIND \[V\d{3}\]")


if __name__ == "__main__":
    unittest.main()
