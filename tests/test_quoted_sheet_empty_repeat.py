"""VCL 0.7 additions found by corpus batch 9 (D138).

1. A TEST call may pass a sheet whose name has spaces or symbols, in quotes as an assertion names it: 'Enq Summary'
   (a look-alike KGC template keeps the real sheet names). Only for a Worksheet argument.
2. A DOCUMENT's REPEAT may have no rows, from=[]: the repeated row, column or section is removed. A REPEAT section
   from=[] needs no field bindings; a row or a column still names its fields, so KRONOS knows which row it removes.
"""
import datetime as dt
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
sys.path.insert(0, str(ROOT / "vcl"))

import kronos_book  # noqa: E402
import kronos_engine  # noqa: E402
from parser import parse_spec, split_top  # noqa: E402
from validator import validate_text  # noqa: E402
from vba_check import check_text  # noqa: E402
from vcl_fmt import format_text  # noqa: E402
from vcl_interp import VclError, run_tests  # noqa: E402
from vcl_parse import _parse_value  # noqa: E402
from vcl_pilot import spec_procedure  # noqa: E402

TODAY = dt.date(2026, 11, 20)
QUOTED = (ROOT / "tests" / "data" / "quoted_sheet.vcl").read_text(encoding="utf-8")
PROPOSAL = (ROOT / "tests" / "data" / "doc_proposal.vcl").read_text(encoding="utf-8")


def errors(text):
    return [p.message for p in validate_text(text) if p.severity == "error"]


def without_equipment(proposal, tq_rows="from=[]"):
    """The practice proposal with no equipment sections and no technical queries."""
    out = []
    for ln in proposal.splitlines():
        if ln.startswith("REPEAT   section Equipment"):
            ln = "REPEAT   section Equipment from=[]"
        elif ln.startswith("REPEAT   row table=4"):   # the nozzle table is inside the removed section
            continue
        elif ln.startswith("REPEAT   row table=6"):
            ln = f"REPEAT   row table=6 {tq_rows} TQ_No=1, Requirement=2, Query=3"
        elif ln.startswith("TEST"):
            continue
        out.append(ln)
    out.append('TEST     fixture=notes_01.xlsx call=FillProposal("P-1", "20 Nov 2026", FALSE) assert=SECTION_COUNT(Equipment) '
               '= 0 AND TEXT_HAS("Two gussets") = 0 AND TEXT_HAS("TQ-01") = 0 AND TABLE_ROWS(2) = 4 AND ERROR_COUNT = 0')
    return "\n".join(out) + "\n"


class QuotedSheet(unittest.TestCase):
    def test_the_grammar_reads_a_quoted_sheet_in_a_call(self):
        problems = []
        tree = _parse_value("test_value", "fixture=e.xlsx call=P('Enq Summary', 'Bob''s', 3) assert=RESULT = 1", 1,
                            problems)
        self.assertEqual(problems, [])
        self.assertEqual([str(t) for t in tree.children if getattr(t, "type", "") == "QSHEET"],
                         ["'Enq Summary'", "'Bob''s'"])

    def test_the_line_parser_unquotes_it_and_remembers_where(self):
        self.assertEqual(split_top("'a, b', \"c, d\", 3", ","), ["'a, b'", ' "c, d"', " 3"])
        spec = parse_spec(QUOTED)
        self.assertEqual(spec.tests[0].args, ["Enq Summary"])
        self.assertEqual(spec.tests[0].quoted, (0,))

    def test_the_recipe_is_proven_everywhere(self):
        self.assertEqual(errors(QUOTED), [])
        self.assertEqual(format_text(QUOTED), QUOTED)
        results = run_tests(QUOTED, TODAY)
        self.assertTrue(len(results) == 3 and all(r.passed for r in results), [r.failed for r in results])
        sim = check_text(QUOTED, "quoted_sheet", TODAY)
        self.assertEqual((sim.tests, sim.differences), (3, []))

    def test_kronos_offers_the_sheet_by_its_name(self):
        self.assertEqual(kronos_book.arguments(QUOTED)[0]["default"], "Enq Summary")
        self.assertEqual(kronos_engine.sheet_hints(QUOTED), {"ws": "Enq Summary"})

    def test_a_template_recipe_has_no_practice_run_yet(self):
        with self.assertRaises(ValueError) as caught:
            spec_procedure("quoted_sheet.vcl", QUOTED)
        self.assertIn("look-alike KGC template", str(caught.exception))

    def test_what_is_refused(self):
        cases = [
            (QUOTED.replace("'BOM-RM'", "'No Such Sheet'"), "has no sheet No Such Sheet"),
            (QUOTED.replace("ARGS     ws As Worksheet", "ARGS     ws As String")
                   .replace("COL($ws, ", "COL(SHEET(THIS_BOOK, $ws), "), "a sheet name in quotes is only for a Worksheet"),
            (QUOTED.replace("('Enq Summary')", '("Enq Summary")', 1), "has no sheet"),
        ]
        for text, want in cases:
            with self.subTest(want=want):
                found = errors(text)
                self.assertTrue(any(want in m for m in found), found)
        old = QUOTED.replace("VCL 0.7", "VCL 0.6", 1)
        self.assertTrue(any("a sheet name in quotes in a TEST call" in m for m in errors(old)), errors(old))

    def test_the_interpreter_stops_on_a_missing_quoted_sheet(self):
        text = QUOTED.replace("'BOM-RM'", "'No Such Sheet'")
        with self.assertRaises(VclError):
            run_tests(text, TODAY)


class EmptyRepeat(unittest.TestCase):
    def test_the_grammar_reads_from_empty(self):
        problems = []
        tree = _parse_value("repeat_value", "section Equipment from=[]", 1, problems)
        self.assertEqual(problems, [])
        self.assertEqual(tree.children[2].data, "empty_list")

    def test_no_rows_removes_the_section_and_the_rows(self):
        text = without_equipment(PROPOSAL)
        self.assertEqual(errors(text), [])
        self.assertEqual(format_text(text), text)
        results = run_tests(text, TODAY)
        self.assertTrue(results and all(r.passed for r in results), [r.failed for r in results])

    def test_what_is_refused(self):
        text = without_equipment(PROPOSAL)
        cases = [
            (text.replace("REPEAT   section Equipment from=[]", "REPEAT   section Equipment from=[] Item=2"),
             "removes the section: it has no fields to bind"),
            (without_equipment(PROPOSAL, tq_rows="from=[]").replace(" TQ_No=1, Requirement=2, Query=3", ""),
             "REPEAT row names its fields"),
            (text.replace("REPEAT   section Equipment from=[]", "REPEAT   section Equipment from=[] match=1=Tag"),
             "from=[] has no rows"),
        ]
        for bad, want in cases:
            with self.subTest(want=want):
                found = errors(bad)
                self.assertTrue(any(want in m for m in found), found)


if __name__ == "__main__":
    unittest.main()
