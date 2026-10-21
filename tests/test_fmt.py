"""VEP 0.6 Phase 3 (D62): vclc fmt gives one canonical text, and never changes what a spec means."""
import random
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
import validator  # noqa: E402
import vcl_fmt  # noqa: E402
from vcl_goal import add_goals  # noqa: E402
from vcl_fmt import FormatError, format_text, format_value, gap, tokens_of  # noqa: E402

EXAMPLES = ROOT / "vcl" / "examples"
PATHS = sorted((ROOT / "specs").glob("*.vcl")) + [EXAMPLES / n for n in ("example1_rev3.vcl", "example3_files.vcl",
                                                                         "example4_build.vcl")]


def read(path):
    return Path(path).read_text(encoding="utf-8")


def messages(text):
    return [(p.severity, p.message) for p in validator.validate_text(text)]


def mess_up(text, rng):
    """The same spec written untidily: other key padding, other OP indentation, spaces where the canonical
    form has none or several where it has one, blank lines, trailing spaces. Only layout changes."""
    op_pad = rng.choice([1, 2, 3, 12])   # after OP, for every step: the first step sets the base column
    lines = []
    for raw in text.rstrip("\n").split("\n"):
        key = raw.split(" ")[0]
        if raw in ("---", "TASK", "BUILD") or raw.startswith(("VCL ", "# ")):
            lines.append(raw.replace("# ASSUME: ", "# ASSUME:" + " " * rng.choice([1, 3]), 1))
        else:
            after = raw[len(key):]
            spaces = len(after) - len(after.lstrip(" "))
            value = after.lstrip(" ")
            if key == "OP":
                lines.append(key + " " * (op_pad + spaces - (vcl_fmt.KEY_WIDTH - len(key))) + messy_tokens(key, value, rng))
            else:
                lines.append(key + " " * rng.choice([1, 2, 5, 11]) + (value if key == "FIXTURE" else messy_tokens(key, value, rng)))
        if rng.random() < 0.2:
            lines.append(rng.choice(["", "   ", "\t"]))
        if rng.random() < 0.2 and lines[-1].strip():
            lines[-1] += rng.choice(["  ", "\t"])
    out = "\n".join(lines) + "\n"
    return out.replace("\n", "\r\n") if rng.random() < 0.3 else out


def messy_tokens(key, value, rng):
    if key in ("GOAL",):
        return value
    tokens = tokens_of(key, value)
    out = []
    for i, token in enumerate(tokens):
        if i:
            prev, canonical = tokens[i - 1], gap(tokens[i - 1], token, "other")
            if canonical == " ":
                out.append(" " * rng.choice([1, 1, 2, 4]))
            elif prev.type in ("LPAR", "LSQB") or token.type in ("RPAR", "RSQB", "COMMA", "SEMICOLON"):
                out.append(" " * rng.choice([0, 0, 1]))
        out.append(str(token))
    return "".join(out)


class Canonical(unittest.TestCase):
    def test_the_committed_specs_and_examples_are_canonical(self):
        for path in PATHS:
            with self.subTest(spec=path.name):
                self.assertEqual(format_text(read(path)), read(path))

    def test_untidy_specs_come_back_to_the_canonical_text(self):
        for path in PATHS:
            for seed in range(6):
                with self.subTest(spec=path.name, seed=seed):
                    untidy = mess_up(read(path), random.Random(f"{path.name}{seed}"))
                    self.assertNotEqual(untidy, read(path))
                    if messages(untidy) != messages(read(path)):
                        self.fail("the untidy copy no longer validates the same")
                    self.assertEqual(format_text(untidy), read(path))

    def test_formatting_twice_changes_nothing(self):
        for path in PATHS:
            once = format_text(mess_up(read(path), random.Random(path.name)))
            self.assertEqual(format_text(once), once)

    def test_a_spec_with_a_goal_stays_canonical(self):
        text = add_goals(read(EXAMPLES / "example4_build.vcl"),
                         {"": "Mark old open NCRs Overdue and summarise them.", "Run": "Start the review on the real sheets.",
                          "RunHeadless": "Run both steps of the review.", "MarkOverdue": "Mark the old open ones Overdue.",
                          "WriteSummary": "Write the status summary sheet."})
        self.assertEqual(format_text(text), text)
        spaced = text.replace("GOAL     Mark the old open ones Overdue.", "GOAL   Mark   the old open   ones Overdue.  ")
        self.assertEqual(format_text(spaced), text)


class Layout(unittest.TestCase):
    def test_the_layout_rules(self):
        text = read(ROOT / "specs" / "task_09.vcl")
        loose = "\n".join(l.replace(", ", ",   ") if l.startswith(("PRE", "ARGS")) else l for l in text.split("\n"))
        untidy = loose.replace("PRE      ", "PRE\t   ", 1).replace("\n", "\r\n") + "\n\n"
        self.assertEqual(format_text(untidy), text)
        self.assertTrue(format_text(untidy).endswith("\n") and not format_text(untidy).endswith("\n\n"))
        self.assertNotIn("\r", format_text(untidy))

    def test_a_byte_order_mark_goes(self):
        text = read(ROOT / "specs" / "task_09.vcl")
        self.assertEqual(format_text("﻿" + text), text)

    def test_nesting_is_two_spaces_per_level_from_the_first_step(self):
        text = read(ROOT / "specs" / "task_09.vcl")
        shifted = "\n".join(("OP" + " " * 2 + l[9:]) if l.startswith("OP ") else l for l in text.split("\n"))
        self.assertNotEqual(shifted, text)
        self.assertEqual(format_text(shifted), text)

    def test_a_question_reply(self):
        self.assertEqual(format_text("# QUESTION:   Which sheet?  \n"), "# QUESTION: Which sheet?\n")

    def test_spacing_of_one_line(self):
        cases = [("OP", 'IF  SAME(CELL($ws,r,cStatus),"CANCELLED")  AND  NOT  IS_EMPTY( x )',
                  'IF SAME(CELL($ws, r, cStatus), "CANCELLED") AND NOT IS_EMPTY(x)'),
                 ("OP", "LET(n,-1+ 2*3)", "LET(n, -1 + 2 * 3)"),
                 ("OP", "FORMAT(x,bold=yes,fill=red)", "FORMAT(x, bold=yes, fill=red)"),
                 ("PRE", 'COL($ws,"A")>0;COL($ws,"B") > 0', 'COL($ws, "A") > 0; COL($ws, "B") > 0'),
                 ("MODIFIES", '$ws[ "A","B" ]', '$ws["A", "B"]'),
                 ("ARGS", "a As Worksheet,b As Worksheet", "a As Worksheet, b As Worksheet"),
                 ("TEST", "fixture=a_01.xlsx   call=P( S , T )  runs=2 assert=RESULT=1 AND COUNTIF(S![Status],\"x\")>=2",
                  'fixture=a_01.xlsx call=P(S, T) runs=2 assert=RESULT = 1 AND COUNTIF(S![Status], "x") >= 2'),
                 ("FLOW", "a.B->c.D  ->e.F", "a.B -> c.D -> e.F")]
        for key, given, want in cases:
            with self.subTest(given=given):
                self.assertEqual(format_value(key, given), want)


class Safety(unittest.TestCase):
    def test_a_spec_with_errors_is_refused_with_the_validators_message(self):
        with self.assertRaises(FormatError) as caught:
            format_text("VCL 0.4\nTASK\nPROC     x\n")
        self.assertIn("missing KIND", caught.exception.problems)

    def test_quoted_text_and_fixture_rules_never_change(self):
        text = read(ROOT / "specs" / "task_08.vcl")
        out = format_text(text)
        for line in text.split("\n"):
            if line.startswith(("FIXTURE", "TEST")):
                self.assertIn(line, out)

    def test_a_formatter_that_changed_the_meaning_is_caught(self):
        real = vcl_fmt.format_value
        try:
            vcl_fmt.format_value = lambda key, value: real(key, value).replace("> 0", ">= 0")
            with self.assertRaises(FormatError):
                format_text(read(ROOT / "specs" / "task_09.vcl"))
        finally:
            vcl_fmt.format_value = real


class CommandLine(unittest.TestCase):
    def run_fmt(self, *args):
        done = subprocess.run([sys.executable, str(ROOT / "compiler" / "vclc.py"), "fmt", *args], capture_output=True,
                              text=True, cwd=ROOT)
        return done.returncode, done.stdout

    def test_check_passes_for_canonical_files_and_names_the_others(self):
        code, out = self.run_fmt("--check", "specs/task_09.vcl")
        self.assertEqual((code, out.strip()), (0, "1 file in canonical form"))
        with tempfile.TemporaryDirectory() as folder:
            untidy = Path(folder) / "untidy.vcl"
            untidy.write_text(read(ROOT / "specs" / "task_09.vcl").replace("PRE      ", "PRE   "), encoding="utf-8")
            code, out = self.run_fmt("--check", str(untidy))
            self.assertEqual(code, 1)
            self.assertIn("not in canonical form", out)

    def test_write_rewrites_the_file(self):
        with tempfile.TemporaryDirectory() as folder:
            untidy = Path(folder) / "untidy.vcl"
            untidy.write_text(read(ROOT / "specs" / "task_09.vcl").replace("PRE      ", "PRE   "), encoding="utf-8")
            code, out = self.run_fmt("--write", str(untidy))
            self.assertEqual(code, 0)
            self.assertEqual(untidy.read_text(encoding="utf-8"), read(ROOT / "specs" / "task_09.vcl"))
            self.assertEqual(self.run_fmt("--check", str(untidy))[0], 0)

    def test_printing_is_the_default(self):
        code, out = self.run_fmt("specs/task_09.vcl")
        self.assertEqual((code, out), (0, read(ROOT / "specs" / "task_09.vcl")))


if __name__ == "__main__":
    unittest.main()
