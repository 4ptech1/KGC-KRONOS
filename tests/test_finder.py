"""KRONOS's recipe finder (compiler/kronos_finder.py, data/library.json; D129).

The library is our own proven recipes only, never the held-out items; search finds a request's own recipe when it is
worded another way; adapt renames a recipe's columns to a workbook's only where it is sure, and the result is proven
again and makes the same changes on a workbook spelt that way.
"""
import datetime as dt
import io
import json
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))

import kronos_finder as kf  # noqa: E402
from vcl_corpus import items  # noqa: E402

TODAY = dt.date(2026, 11, 21)


class Library(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.lib = kf.load_library()

    def test_the_committed_library_is_what_the_sources_give(self):
        self.assertEqual(kf.LIBRARY.read_text(encoding="utf-8"), kf.library_text(),
                         "run python compiler/kronos_finder.py --build")

    def test_it_holds_every_training_pair_and_the_demo_recipes(self):
        ids = {e["id"] for e in self.lib}
        self.assertEqual({f.name for f in items("verified")} | {"demo-1", "demo-2", "demo-3"}, ids)

    def test_no_held_out_item_is_in_it(self):
        text = kf.LIBRARY.read_text(encoding="utf-8")
        for folder in items("eval"):
            self.assertNotIn(f'"id": "{folder.name}"', text)
            self.assertNotIn((folder / "task.txt").read_text(encoding="utf-8").strip()[:120], text)

    def test_each_entry_says_what_it_reads_and_changes(self):
        c001 = kf.entry("c001", self.lib)
        self.assertEqual(c001["reads"], [{"sheet": "BOM", "headers": ["Part No", "Description", "Required Qty"]},
                                         {"sheet": "GRN", "headers": ["Part No", "Received Qty"]}])
        self.assertTrue(c001["changes"] and c001["goal"])


class Search(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.lib = kf.load_library()

    def test_a_reworded_request_finds_its_own_recipe(self):
        """Each paraphrase, left out of the index, finds its own pair in the top 3 (measured: 405 of 405; top 1: 404)."""
        found = total = 0
        for e in self.lib:
            for p in e["paraphrases"]:
                total += 1
                found += e["id"] in [r["id"] for r in kf.search(p, entries=self.lib, leave_out=p)[:3]]
        self.assertGreater(total, 400)
        self.assertGreaterEqual(found / total, 0.98)

    def test_shop_words_mean_the_same(self):
        self.assertEqual(kf.words("non-conformance"), kf.words("NCRs"))
        self.assertEqual(kf.words("radiography"), kf.words("NDE"))
        self.assertEqual(kf.words("quantities"), kf.words("qty"))

    def test_a_request_about_nothing_we_do_finds_nothing(self):
        self.assertEqual(kf.search("make me a cup of tea", entries=self.lib), [])
        self.assertEqual(kf.search("", entries=self.lib), [])

    def test_the_demo_request_finds_the_demo_recipe(self):
        best = kf.search("highlight the open major NCRs", entries=self.lib)[0]
        self.assertEqual(best["id"], "demo-1")
        self.assertGreater(best["strength"], kf.WEAK)
        for key in ("goal", "matched", "fit", "changes", "assumes", "reads", "score"):
            self.assertIn(key, best)

    def test_a_workbook_with_the_recipes_columns_ranks_it_higher(self):
        request = "list parts that are short"
        plain = kf.search(request, entries=self.lib, k=20)
        target = next(r["id"] for r in plain[3:] if kf.entry(r["id"], self.lib)["reads"])
        heads = {f"S{i}": r["headers"] for i, r in enumerate(kf.entry(target, self.lib)["reads"])}
        ranked = [r["id"] for r in kf.search(request, headings=heads, entries=self.lib, k=20)]
        self.assertLess(ranked.index(target), [r["id"] for r in plain].index(target))
        fit = next(r for r in kf.search(request, headings=heads, entries=self.lib, k=20) if r["id"] == target)["fit"]
        self.assertEqual(fit["found"], fit["total"])


class Adapt(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.c001 = kf.entry("c001")["spec"]

    def test_sure_matches_are_renamed_and_the_recipe_is_proven_again(self):
        heads = {"Material List": ["Part No.", "Desc", "Required Quantity"], "Receipts": ["Part No.", "Received Quantity"]}
        out = kf.adapt(self.c001, heads, today=TODAY)
        self.assertTrue(out["ok"], out["message"])
        self.assertEqual([(r["old"], r["new"], r["why"]) for r in out["renames"]], [
            ("Part No", "Part No.", "the same name written differently"),
            ("Description", "Desc", "a known abbreviation"),
            ("Required Qty", "Required Quantity", "a known abbreviation"),
            ("Received Qty", "Received Quantity", "a known abbreviation")])
        self.assertNotIn('"Required Qty"', out["text"])
        self.assertIn('"Required Quantity"', out["text"])
        self.assertIn("cols=A:Part No.,", out["text"])

    def test_only_the_ticked_renames_are_made(self):
        heads = {"Material List": ["Part No.", "Desc", "Required Quantity"], "Receipts": ["Part No.", "Received Quantity"]}
        out = kf.adapt(self.c001, heads, accepted=["Description"], today=TODAY)
        self.assertEqual([r["old"] for r in out["renames"]], ["Description"])
        self.assertIn('"Required Qty"', out["text"])

    def test_a_column_spelt_two_ways_is_left_alone_and_listed(self):
        heads = {"Material List": ["Part No.", "Desc", "Required Quantity"], "Receipts": ["Part Number", "Received Quantity"]}
        plan = kf.proposal(self.c001, heads)
        self.assertNotIn("Part No", [r["old"] for r in plan["renames"]])
        self.assertTrue(any(m["old"] == "Part No" and "differently" in m["why"] for m in plan["missing"]))

    def test_an_unsure_or_unwritable_heading_is_not_used(self):
        cases = {"no match": ["Something Else", "Other"], "two candidates": ["Part No.", "Part-No", "Desc"],
                 "a comma": ["Part No", "Description, long", "Required Qty"]}
        for what, row in cases.items():
            with self.subTest(what):
                plan = kf.proposal(self.c001, {"BOM": row, "GRN": ["Part No", "Received Qty"]})
                self.assertTrue(plan["missing"])
                for r in plan["renames"]:
                    self.assertNotIn(",", r["new"])
        plan = kf.proposal(self.c001, {"BOM": ["Part No", "Part-No", "Description", "Required Qty"],
                                       "GRN": ["Part No", "Received Qty"]})
        self.assertEqual(plan["renames"], [], "an exact match wins; no guessing")

    def test_nothing_to_rename_gives_the_recipe_back_as_it_is(self):
        out = kf.adapt(self.c001, {"BOM": ["Part No", "Description", "Required Qty"], "GRN": ["Part No", "Received Qty"]},
                       today=TODAY)
        self.assertTrue(out["ok"])
        self.assertEqual(out["text"], self.c001)
        self.assertEqual(out["renames"], [])

    def test_a_rename_that_breaks_the_tests_gives_back_the_original(self):
        from unittest import mock

        class NotProven:
            ok = False
        with mock.patch("kronos_engine.prove", return_value=NotProven()):
            out = kf.adapt(self.c001, {"BOM": ["Part No", "Desc", "Required Qty"], "GRN": ["Part No", "Received Qty"]},
                           today=TODAY)
        self.assertFalse(out["ok"])
        self.assertEqual(out["text"], self.c001)
        self.assertIn("kept the original", out["message"])

    def test_renames_never_feed_each_other(self):
        text = 'OP       LET(a, COL($ws, "A"))\nOP       LET(b, COL($ws, "B"))\n'
        self.assertEqual(kf.rename(text, [("A", "B"), ("B", "C")]),
                         'OP       LET(a, COL($ws, "B"))\nOP       LET(b, COL($ws, "C"))\n')

    def test_the_adapted_demo_recipe_changes_a_workbook_spelt_differently_the_same_way(self):
        """End to end: the demo register with two columns spelt another way; the adapted demo recipe makes the same
        changes on it that the original recipe makes on the original register."""
        import openpyxl
        import kronos_book as kb
        original = (ROOT / "docs" / "demo" / "NCR_Register_Demo.xlsx").read_bytes()
        wb = openpyxl.load_workbook(io.BytesIO(original))
        ws = wb["NCR Register"]
        spelt = {"NCR No": "NCR Number", "Status": "status"}
        for cell in ws[1]:
            if cell.value in spelt:
                cell.value = spelt[cell.value]
        buf = io.BytesIO()
        wb.save(buf)
        other = buf.getvalue()
        recipe = kf.entry("demo-1")["spec"]
        heads = {s["name"]: s["headings"] for s in kb.describe(other, "other.xlsx")["sheets"]}
        out = kf.adapt(recipe, heads, today=TODAY)
        self.assertTrue(out["ok"], out["message"])
        self.assertTrue({(r["old"], r["new"]) for r in out["renames"]} <= set(spelt.items()))
        args = {a["name"]: a["default"] for a in kb.arguments(recipe)}
        sheet_args = [a["name"] for a in kb.arguments(recipe) if a["type"] == "Worksheet"]
        for name in sheet_args:
            args[name] = "NCR Register"
        before, _ = kb.apply(recipe, original, "NCR_Register_Demo.xlsx", args, TODAY)
        after, copy = kb.apply(out["text"], other, "other.xlsx", args, TODAY)
        self.assertTrue(before["ok"] and after["ok"], after["message"])
        fills = [sum(sh["fills"] for st in r["steps"] for sh in st["sheets"]) for r in (before, after)]
        self.assertEqual((before["result"], fills[0]), (after["result"], fills[1]))   # 15 NCRs, the same fills
        self.assertGreater(int(after["result"]), 0)
        self.assertGreater(fills[1], 0)
        self.assertIsNotNone(copy)


class Bridge(unittest.TestCase):
    def test_the_page_calls(self):
        import kronos_web
        found = json.loads(kronos_web.find("highlight the open major NCRs", json.dumps({"NCR Register": ["NCR No"]})))
        self.assertEqual(found["results"][0]["id"], "demo-1")
        self.assertIn("weak", found)
        plan = json.loads(kronos_web.propose("c001", json.dumps({"BOM": ["Part No", "Desc", "Required Qty"]})))
        self.assertEqual([r["old"] for r in plan["renames"]], ["Description"])
        done = json.loads(kronos_web.adapt("c001", json.dumps({"BOM": ["Part No", "Desc", "Required Qty"]}),
                                           json.dumps(["Description"]), TODAY.isoformat()))
        self.assertTrue(done["ok"])
        self.assertIn("B:Desc,", done["text"])   # c001 names Description only in its practice data
        self.assertEqual(done["id"], "c001")


if __name__ == "__main__":
    unittest.main()
