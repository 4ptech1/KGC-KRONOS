"""The KRONOS demo (D88, docs/demo/): an invented NCR register and three recipes.

Every committed demo file must be what docs/demo/make_demo.py makes today: the workbook, the three changed copies
(made by compiler/kronos_book.py, the code KRONOS Offline's page runs) and facts.json, from which the Word guide
quotes its numbers. The recipes must be PROVEN and in canonical form.
"""
import datetime as dt
import io
import json
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEMO = ROOT / "docs" / "demo"
sys.path.insert(0, str(ROOT / "compiler"))
sys.path.insert(0, str(DEMO))

try:
    import openpyxl
except ImportError:   # CI installs requirements.txt
    openpyxl = None

from kronos_engine import prove  # noqa: E402
from vcl_fmt import format_text  # noqa: E402

RECIPES = ["recipe_1_highlight.vcl", "recipe_2_tidy.vcl", "recipe_3_summary.vcl"]


def state(data):
    """Everything a person sees in a workbook that the demo depends on, per sheet, in order."""
    wb = openpyxl.load_workbook(io.BytesIO(data))
    out = []
    for ws in wb.worksheets:
        cells = []
        for row in ws.iter_rows():
            for c in row:
                if c.value is None and c.fill.fill_type is None:
                    continue
                value = c.value.date() if isinstance(c.value, dt.datetime) and c.value.time() == dt.time() else c.value
                cells.append((c.coordinate, value, c.number_format, c.fill.fill_type and str(c.fill.fgColor.rgb),
                               bool(c.font.bold)))
        widths = sorted((k, d.width) for k, d in ws.column_dimensions.items() if d.width)
        out.append((ws.title, cells, widths, ws.freeze_panes))
    return out


class Recipes(unittest.TestCase):
    def test_each_recipe_is_proven_and_canonical(self):
        for name in RECIPES:
            text = (DEMO / name).read_text(encoding="utf-8")
            with self.subTest(recipe=name):
                self.assertEqual(format_text(text), text, "run: python compiler/vclc.py fmt --write")
                proof = prove(text, dt.date(2026, 11, 16))
                self.assertTrue(proof.ok, [str(f.message) for f in proof.errors])


@unittest.skipIf(openpyxl is None, "openpyxl is not installed")
class Files(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import make_demo
        cls.make = make_demo
        cls.data = make_demo.build()
        cls.results = make_demo.run((DEMO / make_demo.WORKBOOK).read_bytes())

    def test_the_workbook_is_what_the_script_makes(self):
        self.assertEqual(state((DEMO / self.make.WORKBOOK).read_bytes()), state(self.data))

    def test_each_changed_copy_is_what_kronos_gives_today(self):
        for (name, _, after), (_, result, copy) in zip(self.make.RECIPES, self.results):
            with self.subTest(recipe=name):
                self.assertTrue(result["ok"], result["message"])
                self.assertEqual(result["warnings"], [], "the demo workbook needs nothing ticked")
                self.assertEqual(state((DEMO / after).read_bytes()), state(copy))

    def test_facts_json_is_what_the_guide_should_say(self):
        committed = json.loads((DEMO / "facts.json").read_text(encoding="utf-8"))
        self.assertEqual(committed, json.loads(json.dumps(self.make.facts(self.data, self.results))))
        runs = committed["runs"]
        self.assertEqual(committed["rows"], 40)
        self.assertEqual(runs[1]["result"], "2")                      # the two NCRs typed twice
        self.assertEqual(runs[2]["result"], "40")
        self.assertEqual(committed["chained_summary"]["result"], "38")  # tidy first: each NCR once
        self.assertEqual(int(runs[0]["result"]), len(runs[0]["highlighted"]))

    def test_the_original_is_untouched_by_every_run(self):
        before = (DEMO / self.make.WORKBOOK).read_bytes()
        self.make.run(before)
        self.assertEqual((DEMO / self.make.WORKBOOK).read_bytes(), before)

    def test_only_invented_values(self):
        """Every text in the register comes from make_demo's own invented lists."""
        allowed = set(self.make.HEADERS) | set(self.make.JOBS) | set(self.make.ITEMS) | set(self.make.DEFECTS) \
            | set(self.make.DEPARTMENTS) | set(self.make.STATUSES) | {"Major", "Minor"}
        ws = openpyxl.load_workbook(io.BytesIO(self.data))[self.make.SHEET]
        for row in ws.iter_rows(values_only=True):
            for v in row:
                if isinstance(v, str) and not v.startswith("NCR-2026-"):
                    self.assertIn(v, allowed)


class Guide(unittest.TestCase):
    def test_the_guide_and_its_pictures_are_there(self):
        self.assertTrue((DEMO / "KRONOS_Demo_NCR_Register.docx").read_bytes().startswith(b"PK"))
        for name in ["workbook.png"] + [f"recipe{n}-{part}.png" for n in (1, 2, 3) for part in ("proven", "apply", "copy")]:
            self.assertTrue((DEMO / "shots" / name).read_bytes().startswith(b"\x89PNG"), name)


if __name__ == "__main__":
    unittest.main()
