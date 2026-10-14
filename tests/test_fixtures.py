"""Unit tests for harness/fixtures.py. Run: python -m unittest discover tests"""
import datetime
import os
import re
import sys
import unittest
from collections import Counter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path[:0] = [os.path.join(ROOT, "compiler"), os.path.join(ROOT, "harness")]

from fixtures import build_file, vba_for_file, vba_literal  # noqa: E402
from parser import parse_spec  # noqa: E402


def read(path):
    with open(os.path.join(ROOT, path), encoding="utf-8") as f:
        return f.read()


def fixtures_of(path):
    return parse_spec(read(path)).fixtures


def column(sheet, header):
    return [row[sheet.headers.index(header)] for row in sheet.rows]


def synthetic(*fixture_lines):
    """Parsed FIXTURE lines without a full spec around them."""
    text = "VCL 0.1\nTASK\n" + "\n".join("FIXTURE  " + line for line in fixture_lines) + "\n"
    spec = parse_spec(text)
    assert not [p for p in spec.problems if "FIXTURE" in p.message], spec.problems
    return spec.fixtures


class SameSpecSameData(unittest.TestCase):
    def test_every_real_fixture_builds_the_same_data_twice(self):
        for i in range(1, 11):
            fixtures = fixtures_of(f"specs/task_{i:02d}.vcl")
            for file in {fx.file for fx in fixtures}:
                with self.subTest(spec=i, file=file):
                    first = build_file(fixtures, file)
                    again = build_file(fixtures_of(f"specs/task_{i:02d}.vcl"), file)
                    self.assertEqual({k: v.rows for k, v in first.items()}, {k: v.rows for k, v in again.items()})
                    self.assertEqual(vba_for_file(file, first), vba_for_file(file, again))

    def test_columns_are_not_correlated(self):
        rule = 'values("a"x10,"b"x10,"c"x10)'
        fixtures = synthetic(f"seed_01.xlsx sheet=One cols=A:X,B:Y rows=30 data=A:{rule};B:{rule}",
                             f"seed_01.xlsx sheet=Two cols=A:X rows=30 data=A:{rule}")
        sheets = build_file(fixtures, "seed_01.xlsx")
        one_x, one_y = column(sheets["One"], "X"), column(sheets["One"], "Y")
        self.assertEqual(Counter(one_x), Counter(one_y))
        self.assertNotEqual(one_x, one_y)                      # another column, another order
        self.assertNotEqual(one_x, column(sheets["Two"], "X"))  # another sheet, another order

    def test_seed_rule_is_stable(self):
        # The dataset and the expected results depend on this exact output: never change it.
        sheets = build_file(fixtures_of("specs/task_01.vcl"), "bom_grn_01.xlsx")
        bom = sheets["BOM"]
        self.assertEqual(column(bom, "Required Qty")[:5], [3, 9, 32, 23, 3])
        self.assertEqual(column(bom, "Description")[:2], ["NCHWIRSNKCMU", "BXDERLKKKNUG"])


class CountsMatchTheRules(unittest.TestCase):
    def test_values_counts(self):
        nde = build_file(fixtures_of("specs/task_03.vcl"), "nde_repair_01.xlsx")["NDE"]
        self.assertEqual(Counter(column(nde, "Welder ID")), Counter({"W-01": 20, "W-02": 20, "W-03": 15, None: 5}))
        self.assertEqual(Counter(column(nde, "Result")), Counter({"Accepted": 48, "Repair": 8, None: 4}))

    def test_blank_counts(self):
        sheets = build_file(fixtures_of("specs/task_01.vcl"), "bom_grn_01.xlsx")
        required = column(sheets["BOM"], "Required Qty")
        self.assertEqual(required.count(None), 3)
        self.assertTrue(all(1 <= v <= 50 for v in required if v is not None))
        self.assertEqual(column(sheets["GRN"], "Received Qty").count(None), 4)
        heat = column(build_file(fixtures_of("specs/task_02.vcl"), "traceability_01.xlsx")["Traceability"],
                      "Plate/Heat No")
        self.assertEqual(heat.count(None), 4)
        self.assertTrue(all(len(v) == 8 and v.isalpha() for v in heat if v is not None))

    def test_seq_dates_and_columns_without_rules(self):
        sheets = build_file(fixtures_of("specs/task_01.vcl"), "bom_grn_01.xlsx")
        self.assertEqual(column(sheets["BOM"], "Part No"), list(range(1001, 1041)))
        weld = build_file(fixtures_of("specs/task_05.vcl"), "weld_log_01.xlsx")["WeldLog"]
        dates = column(weld, "Weld Date")
        self.assertTrue(all(datetime.date(2026, 1, 1) <= d <= datetime.date(2026, 9, 30) for d in dates))
        inventory = build_file(fixtures_of("specs/task_08.vcl"), "consumables_01.xlsx")["Inventory"]
        self.assertEqual(column(inventory, "Balance"), [None] * 8)

    def test_values_counts_must_match_rows(self):
        fixtures = synthetic('bad_01.xlsx sheet=S cols=A:X rows=5 data=A:values("a"x4)')
        with self.assertRaises(ValueError):
            build_file(fixtures, "bad_01.xlsx")


class GeneratedVba(unittest.TestCase):
    def setUp(self):
        self.vba = vba_for_file("bom_grn_01.xlsx", build_file(fixtures_of("specs/task_01.vcl"), "bom_grn_01.xlsx"))

    def writes(self, proc):
        body = self.vba.split(f"Private Sub {proc}(")[1].split("End Sub")[0]
        return re.findall(r"ForgeWriteRow ws, (\d+), Array\((.*)\)", body)

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
        fixtures = synthetic("big_01.xlsx sheet=Data cols=A:ID rows=600 data=A:seq(1,1)")
        vba = vba_for_file("big_01.xlsx", build_file(fixtures, "big_01.xlsx"))
        self.assertIn("Private Sub ForgeFixture_big_01_Data_3(ByVal ws As Worksheet)", vba)
        written = [int(n) for n in re.findall(r"ForgeWriteRow ws, (\d+),", vba)]
        self.assertEqual(sorted(written), list(range(1, 602)))


if __name__ == "__main__":
    unittest.main()
