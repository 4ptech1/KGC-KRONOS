"""Unit tests for compiler/vcl_fixtures.py, the one fixture generator."""
import datetime
import re
import sys
import unittest
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
from vba_lint import lint  # noqa: E402
from vcl_fixtures import build_workbooks, fixtures_module, parse_fixture, vba_for_file, vba_literal  # noqa: E402
from vcl_interp import run_tests  # noqa: E402
from vcl_parse import parse  # noqa: E402


def trees_of(text):
    spec, problems = parse(text)
    assert not problems, problems
    return [k["tree"] for k in spec["keys"] if k["key"] == "FIXTURE"]


def spec_trees(name):
    return trees_of((ROOT / "specs" / name).read_text(encoding="utf-8"))


def synthetic(*fixture_lines):
    """FIXTURE lines inside the smallest spec that parses."""
    return trees_of("\n".join(["VCL 0.1", "TASK", "PROC     P", "KIND     Sub", "MODULE   modX",
                               "OP       LOG(1)", "ERR      standard"]
                              + ["FIXTURE  " + line for line in fixture_lines]))


def column(sheet, header):
    c = next(c for c in range(1, sheet.max_col() + 1) if sheet.get(1, c) == header)
    rows = max(r for r, _ in sheet.cells) if sheet.cells else 1
    return [sheet.get(r, c) for r in range(2, rows + 1)]


class SameSpecSameData(unittest.TestCase):
    def test_every_real_fixture_builds_the_same_data_twice(self):
        for p in sorted((ROOT / "specs").glob("*.vcl")):
            with self.subTest(spec=p.name):
                first = build_workbooks(spec_trees(p.name))
                again = build_workbooks(spec_trees(p.name))
                self.assertEqual({f: {s: sh.cells for s, sh in b.items()} for f, b in first.items()},
                                 {f: {s: sh.cells for s, sh in b.items()} for f, b in again.items()})

    def test_columns_are_not_correlated(self):
        rule = 'values("a"x10,"b"x10,"c"x10)'
        books = build_workbooks(synthetic(f"seed_01.xlsx sheet=One cols=A:X,B:Y rows=30 data=A:{rule};B:{rule}",
                                          f"seed_01.xlsx sheet=Two cols=A:X rows=30 data=A:{rule}"))
        one_x, one_y = column(books["seed_01.xlsx"]["One"], "X"), column(books["seed_01.xlsx"]["One"], "Y")
        self.assertEqual(Counter(one_x), Counter(one_y))
        self.assertNotEqual(one_x, one_y)                                     # another column, another order
        self.assertNotEqual(one_x, column(books["seed_01.xlsx"]["Two"], "X"))  # another sheet, another order

    def test_seed_rule_is_stable(self):
        # The dataset and the expected results depend on this exact output: never change it.
        # The same values come out of the pilot's harness/fixtures.py (PR #4).
        bom = build_workbooks(spec_trees("task_01.vcl"))["bom_grn_01.xlsx"]["BOM"]
        self.assertEqual(column(bom, "Required Qty")[:5], [3, 9, 32, 23, 3])
        self.assertEqual(column(bom, "Description")[:2], ["NCHWIRSNKCMU", "BXDERLKKKNUG"])


class CountsMatchTheRules(unittest.TestCase):
    def test_values_counts(self):
        nde = build_workbooks(spec_trees("task_03.vcl"))["nde_repair_01.xlsx"]["NDE"]
        self.assertEqual(Counter(column(nde, "Welder ID")), Counter({"W-01": 20, "W-02": 20, "W-03": 15, None: 5}))
        self.assertEqual(Counter(column(nde, "Result")), Counter({"Accepted": 48, "Repair": 8, None: 4}))

    def test_blank_counts(self):
        book = build_workbooks(spec_trees("task_01.vcl"))["bom_grn_01.xlsx"]
        required = column(book["BOM"], "Required Qty")
        self.assertEqual(required.count(None), 3)
        self.assertTrue(all(1 <= v <= 50 for v in required if v is not None))
        self.assertEqual(column(book["GRN"], "Received Qty").count(None), 4)
        trace = build_workbooks(spec_trees("task_02.vcl"))["traceability_01.xlsx"]["Traceability"]
        heat = column(trace, "Plate/Heat No")
        self.assertEqual(heat.count(None), 4)
        self.assertTrue(all(len(v) == 8 and v.isalpha() for v in heat if v is not None))

    def test_seq_dates_and_columns_without_rules(self):
        book = build_workbooks(spec_trees("task_01.vcl"))["bom_grn_01.xlsx"]
        self.assertEqual(column(book["BOM"], "Part No"), list(range(1001, 1041)))
        weld = build_workbooks(spec_trees("task_05.vcl"))["weld_log_01.xlsx"]["WeldLog"]
        self.assertTrue(all(datetime.date(2026, 1, 1) <= d <= datetime.date(2026, 9, 30)
                            for d in column(weld, "Weld Date")))
        inventory = build_workbooks(spec_trees("task_08.vcl"))["consumables_01.xlsx"]["Inventory"]
        self.assertEqual([inventory.get(r, 4) for r in range(2, 10)], [None] * 8)   # Balance has no rule

    def test_values_counts_must_match_rows(self):
        with self.assertRaises(ValueError):
            build_workbooks(synthetic('bad_01.xlsx sheet=S cols=A:X rows=5 data=A:values("a"x4)'))

    def test_doubled_quotes_in_values(self):
        book = build_workbooks(synthetic('q_01.xlsx sheet=S cols=A:X rows=2 data=A:values("say ""hi"""x2)'))
        self.assertEqual(column(book["q_01.xlsx"]["S"], "X"), ['say "hi"', 'say "hi"'])


class InterpreterUsesTheSameData(unittest.TestCase):
    def test_task_01_matches_the_pilots_independent_model(self):
        # batches/batch_000_pilot.expected.csv on PR #4 came from a separate Python model of
        # task_01 on the pilot generator's data. The interpreter must reach the same numbers.
        results = run_tests((ROOT / "specs" / "task_01.vcl").read_text(encoding="utf-8"))
        self.assertEqual([r.result for r in results], [26, 26, 26, -1])
        cells = dict(((r, c), v) for (r, c), v in results[0].state["Shortages"][0])
        remarks = Counter(v for (r, c), v in cells.items() if c == 5 and r > 1)
        self.assertEqual(remarks, Counter({"Short": 16, "Not received": 7, "Required qty missing": 3}))
        self.assertEqual(sum(1 for (r, c) in cells if c == 1), 27)    # COUNTA(Shortages![Part No])
        book = build_workbooks(spec_trees("task_01.vcl"))["bom_grn_01.xlsx"]
        self.assertEqual(sum(v for v in column(book["BOM"], "Required Qty") if v), 986)
        self.assertEqual(sum(v for v in column(book["GRN"], "Received Qty") if v), 661)


class GeneratedVba(unittest.TestCase):
    def setUp(self):
        fixtures = [parse_fixture(t) for t in spec_trees("task_01.vcl")]
        self.vba = vba_for_file("bom_grn_01.xlsx", fixtures)

    def writes(self, proc):
        body = self.vba.split(f"Private Sub {proc}(")[1].split("End Sub")[0]
        return re.findall(r"FixtureWriteRow ws, (\d+), Array\((.*)\)", body)

    def test_headers_are_in_row_1(self):
        rows = self.writes("ForgeFixture_bom_grn_01_BOM")
        self.assertEqual(rows[0], ("1", '"Part No", "Description", "Required Qty"'))
        self.assertEqual([int(n) for n, _ in rows], list(range(1, 42)))

    def test_rows_0_gives_headers_only(self):
        self.assertEqual(self.writes("ForgeFixture_bom_grn_01_Shortages"),
                         [("1", '"Part No", "Required", "Received", "Shortage", "Remark"')])

    def test_blank_cells_are_empty(self):
        values = [args for _, args in self.writes("ForgeFixture_bom_grn_01_BOM")[1:]]
        self.assertEqual(sum(args.endswith(", Empty") for args in values), 3)

    def test_literals(self):
        self.assertEqual(vba_literal(None), "Empty")
        self.assertEqual(vba_literal(datetime.date(2026, 3, 9)), "DateSerial(2026, 3, 9)")
        self.assertEqual(vba_literal('say "hi"'), '"say ""hi"""')
        self.assertEqual(vba_literal("Part No"), '"Part No"')
        for text in ("25", "W-01", "1-2", "TRUE", "=A1", "-x"):  # Excel would convert these
            self.assertEqual(vba_literal(text), f'"\'{text}"')
        self.assertEqual(vba_literal(12.5), "12.5")
        with self.assertRaises(ValueError):
            vba_literal("café")

    def test_big_fixtures_are_split_into_small_procedures(self):
        fixtures = [parse_fixture(t) for t in synthetic("big_01.xlsx sheet=Data cols=A:ID rows=600 data=A:seq(1,1)")]
        vba = vba_for_file("big_01.xlsx", fixtures)
        self.assertIn("Private Sub ForgeFixture_big_01_Data_3(ByVal ws As Worksheet)", vba)
        written = [int(n) for n in re.findall(r"FixtureWriteRow ws, (\d+),", vba)]
        self.assertEqual(sorted(written), list(range(1, 602)))

    def test_fixtures_module_for_all_specs_lints_clean(self):
        specs = [(p.name, spec_trees(p.name)) for p in sorted((ROOT / "specs").glob("*.vcl"))]
        module = fixtures_module(specs)
        self.assertEqual(lint(module), [])
        for name in ("bom_grn_01", "traceability_01", "nde_repair_01", "nde_summary_01", "weld_log_01",
                     "welder_qual_01", "itp_01", "consumables_01", "ncr_01", "calibration_01"):
            self.assertIn(f"Public Sub ForgeFixture_{name}(ByVal wb As Workbook)", module)

    def test_one_fixture_file_name_means_one_set_of_sheets(self):
        a = synthetic("same_01.xlsx sheet=S cols=A:X rows=2 data=A:int(1,9)")
        b = synthetic("same_01.xlsx sheet=S cols=A:X rows=3 data=A:int(1,9)")
        with self.assertRaises(ValueError):
            fixtures_module([("a.vcl", a), ("b.vcl", b)])
        self.assertIn("ForgeFixture_same_01", fixtures_module([("a.vcl", a), ("c.vcl", a)]))


if __name__ == "__main__":
    unittest.main()
