"""VEP 0.6 Phase 2 (D61): the VCL 0.6 line and the GOAL line, the English sentence every spec carries."""
import datetime
import re
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
import validator  # noqa: E402
from vcl_codegen import compile_build, compile_text  # noqa: E402
from vcl_goal import add_goals  # noqa: E402
from vcl_interp import run_tests  # noqa: E402
from vcl_mutate import mutate_report  # noqa: E402
from vcl_samples import old_example  # noqa: E402

DAY = datetime.date(2026, 10, 21)
EXAMPLES = ROOT / "vcl" / "examples"
BUILD = old_example("example4_build.vcl")   # the VCL 0.5 BUILD it was before its GOAL lines
TASK = old_example("example1_rev3.vcl")   # RemoveCancelledAndSortByDate, as VCL 0.5
GOALS = {"": "Mark old open NCRs Overdue and rebuild a summary of the statuses.",
         "Run": "Start the NCR review on the workbook's own sheets and show how many NCRs went Overdue.",
         "RunHeadless": "Run both steps of the NCR review on the sheets it is given.",
         "MarkOverdue": "Mark every Open NCR older than the limit as Overdue.",
         "WriteSummary": "Rebuild the summary sheet with one row per NCR status."}
TASK_GOAL = "Delete every cancelled order, then sort the rest by order date, newest first."


def problems(text):
    return [(p.severity, p.message) for p in validator.validate_text(text)]


def errors(text):
    return [m for s, m in problems(text) if s == "error"]


def goals_for(text):
    """A generic GOAL for every part of a spec, for the compatibility tests."""
    if "\nBUILD\n" in text:
        out = {"": "The tool the spec describes, in one run."}
        out.update({proc: f"Do the {proc} step of the tool." for proc in re.findall(r"^PROC +(\w+)", text, re.M)})
        return out
    return "Do what the spec says on the sheets it is given."


def without_goals(text):
    """The reverse of add_goals, for comparing."""
    return "\n".join(l for l in text.split("\n") if not l.startswith("GOAL ")).replace("VCL 0.6", "VCL 0.5")


def as_it_was(newer, old_version):
    """A 0.6 spec as the older version, each GOAL line becoming a comment line, so every other line keeps
    its number: the compiled VBA labels its lines with the spec's line numbers (for Erl)."""
    return re.sub(r"(?m)^GOAL .*$", "# ASSUME: stands where the GOAL was", newer).replace("VCL 0.6", old_version)


def outcome(text):
    return [(r.passed, r.result, r.errors, r.failed, r.log, r.state, r.files, r.messages, r.hidden, r.texts)
            for r in run_tests(text, DAY)]


class AddGoals(unittest.TestCase):
    def test_a_task_gets_its_goal_before_proc_and_the_new_version(self):
        text = add_goals(TASK, TASK_GOAL)
        lines = text.split("\n")
        self.assertEqual(lines[:2], ["VCL 0.6", "TASK"])
        at = lines.index("GOAL     " + TASK_GOAL)
        self.assertTrue(lines[at + 1].startswith("PROC     "))
        self.assertTrue(lines[at - 1].startswith("# ASSUME:"))   # after the assumptions, which stay on top
        self.assertEqual(without_goals(text).replace("VCL 0.5", "VCL 0.5"), TASK.replace("VCL 0.5", "VCL 0.5"))

    def test_a_build_gets_a_goal_for_itself_and_for_each_task(self):
        text = add_goals(BUILD, GOALS)
        self.assertEqual(text.count("\nGOAL     "), 5)
        self.assertEqual(text.count("VCL 0.6\n"), 5)
        self.assertNotIn("VCL 0.5", text)
        self.assertTrue(text.index("GOAL     " + GOALS[""]) < text.index("NAME     "))
        self.assertEqual(without_goals(text), BUILD)
        self.assertEqual(errors(text), [])

    def test_a_missing_or_extra_goal_is_refused(self):
        with self.assertRaises(ValueError):
            add_goals(BUILD, {k: v for k, v in GOALS.items() if k != "MarkOverdue"})
        with self.assertRaises(ValueError):
            add_goals(BUILD, dict(GOALS, Other="A step the BUILD doesn't have."))


class GoalRules(unittest.TestCase):
    def setUp(self):
        self.task = add_goals(TASK, TASK_GOAL)

    def test_a_goal_makes_a_valid_0_6_spec(self):
        self.assertEqual(problems(self.task), [])
        self.assertEqual(problems(add_goals(BUILD, GOALS)), [])

    def test_goal_is_required_from_0_6(self):
        found = errors(self.task.replace("GOAL     " + TASK_GOAL + "\n", ""))
        self.assertEqual(found, ["missing GOAL: write one plain-English sentence saying what the spec does"])

    def test_goal_is_an_error_before_0_6(self):
        older = self.task.replace("VCL 0.6", "VCL 0.5", 1)
        self.assertEqual(errors(older), ["GOAL is VCL 0.6: start the spec with 'VCL 0.6' to use it"])

    def test_goal_is_checked(self):
        line = "GOAL     " + TASK_GOAL
        cases = [(line + "\n" + line, "GOAL appears more than once"),
                 ("GOAL     Too short", "GOAL must be 10 to 200 characters, got 9"),
                 ("GOAL     " + "x" * 201, "GOAL must be 10 to 200 characters, got 201"),
                 ("GOAL     Delete the cancelled orders, café style", "GOAL must be plain printable ASCII text and not start with #"),
                 ("GOAL     # a comment is not a sentence", "GOAL must be plain printable ASCII text and not start with #")]
        for new, message in cases:
            with self.subTest(message=message):
                self.assertEqual(errors(self.task.replace(line, new)), [message])

    def test_goal_comes_first(self):
        goal, proc = "GOAL     " + TASK_GOAL + "\n", "PROC     RemoveCancelledAndSortByDate\n"
        swapped = self.task.replace(goal + proc, proc + goal)
        self.assertEqual(errors(swapped), ["GOAL is out of order: it must come before PROC"])

    def test_a_goal_that_only_repeats_the_name_is_a_warning(self):
        text = self.task.replace(TASK_GOAL, "Remove cancelled and sort by date.")
        self.assertEqual(problems(text), [("warning",
                                           "GOAL only repeats the PROC name: say what the spec does and for which sheets")])

    def test_a_build_needs_a_goal_and_so_does_each_of_its_tasks(self):
        text = add_goals(BUILD, GOALS)
        self.assertEqual(errors(text.replace("GOAL     " + GOALS[""] + "\n", "")),
                         ["missing GOAL: write one plain-English sentence saying what the spec does"])
        no_task_goal = text.replace("GOAL     " + GOALS["MarkOverdue"] + "\n", "")
        self.assertEqual(errors(no_task_goal), ["missing GOAL: write one plain-English sentence saying what the spec does"])


class Compatibility(unittest.TestCase):
    """A spec that was valid is still valid with a GOAL and the 0.6 line, and gives the same results."""

    def texts(self):
        specs = [(p.name, p.read_text(encoding="utf-8")) for p in sorted((ROOT / "specs").glob("*.vcl"))]
        return specs + [(n, old_example(n)) for n in ("example1_rev3.vcl", "example3_files.vcl", "example4_build.vcl")]

    def test_every_spec_and_example_gives_the_same_results_with_a_goal(self):
        for name, text in self.texts():
            old_version = text.split("\n", 1)[0]
            with self.subTest(spec=name):
                newer = add_goals(text, goals_for(text))
                self.assertEqual(problems(text), problems(newer))
                self.assertEqual(outcome(text), outcome(newer))
                same_lines = as_it_was(newer, old_version)   # the older spec, with every line where it is in newer
                self.assertEqual(problems(same_lines), problems(newer))
                self.assertEqual(outcome(same_lines), outcome(newer))
                if "\nBUILD\n" in text:
                    self.assertEqual(compile_build(same_lines, name), compile_build(newer, name))
                else:
                    self.assertEqual(compile_text(same_lines, name), compile_text(newer, name))

    def test_goal_lines_are_not_mutated(self):
        newer = add_goals(BUILD, GOALS)
        killed, total, survivors = mutate_report(newer, snapshot=True, today=DAY)
        self.assertEqual((killed, total, survivors), (34, 34, []))   # Example 4's score is unchanged


if __name__ == "__main__":
    unittest.main()
