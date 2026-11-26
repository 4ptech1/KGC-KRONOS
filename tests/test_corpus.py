"""VEP 0.6 Phase 5 (D64): the corpus of request and spec pairs, `vclc corpus`, `corpus-add` and `export`."""
import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
import vcl_corpus  # noqa: E402
from vcl_samples import example, without_goals  # noqa: E402

REAL = ROOT / "data"


class TheCorpus(unittest.TestCase):
    def test_every_pair_in_data_is_sound(self):
        self.assertEqual(vcl_corpus.check_all(), [])

    def test_the_first_twelve_pairs_are_the_committed_specs(self):
        pairs = vcl_corpus.load("verified")
        self.assertEqual([p["id"] for p in pairs[:12]], [f"c{n:03d}" for n in range(1, 13)])
        self.assertEqual([p["source"] for p in pairs[:12]], [f"committed spec specs/task_{n:02d}.vcl" for n in range(1, 13)])
        for n, pair in enumerate(pairs[:12], 1):
            original = (ROOT / "specs" / f"task_{n:02d}.vcl").read_text(encoding="utf-8")
            body = "\n".join(l for l in pair["spec"].split("\n") if not l.startswith("GOAL "))
            self.assertEqual(body.split("\n", 1)[1], original.split("\n", 1)[1])   # the spec itself is unchanged

    def test_every_training_pair_has_three_paraphrases_and_no_held_out_pair_has_one(self):   # D104
        for folder in vcl_corpus.items("verified"):
            self.assertEqual(sorted(p.name for p in folder.glob("paraphrase_*.txt")),
                             ["paraphrase_1.txt", "paraphrase_2.txt", "paraphrase_3.txt"], folder.name)
        for folder in vcl_corpus.items("eval"):
            self.assertEqual(list(folder.glob("paraphrase_*.txt")), [], folder.name)

    def test_the_report_counts_what_there_is(self):
        text = vcl_corpus.report()
        self.assertIn(f"{len(vcl_corpus.items('verified'))} pairs in data/verified, "
                      f"{len(vcl_corpus.items('eval'))} in data/eval", text)
        self.assertIn("flag", text)
        self.assertRegex(text, r"Mutants caught: \d+ of \d+")


class Rules(unittest.TestCase):
    """Each rule, on a copy of a real pair in a temporary data folder."""

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.data = Path(self.tmp.name)
        self.saved = vcl_corpus.DATA
        vcl_corpus.DATA = self.data
        (self.data / "verified").mkdir()
        shutil.copytree(REAL / "verified" / "c005", self.data / "verified" / "c005")
        self.folder = self.data / "verified" / "c005"

    def tearDown(self):
        vcl_corpus.DATA = self.saved
        self.tmp.cleanup()

    def meta(self, **changes):
        meta = json.loads((self.folder / "meta.json").read_text(encoding="utf-8"))
        meta.update(changes)
        (self.folder / "meta.json").write_text(vcl_corpus.dump(meta), encoding="utf-8")

    def spec(self, old, new):
        path = self.folder / "spec.vcl"
        text = path.read_text(encoding="utf-8")
        assert old in text, old
        path.write_text(text.replace(old, new, 1), encoding="utf-8")

    def problems(self, **kw):
        return vcl_corpus.check_item(self.folder, **kw)

    def test_the_copy_is_sound(self):
        self.assertEqual(self.problems(), [])

    def test_missing_and_extra_files(self):
        (self.folder / "task.txt").unlink()
        self.assertEqual(self.problems(), ["c005: missing task.txt"])
        (self.folder / "task.txt").write_text("Copy the rows that are not accepted to an exceptions sheet.\n")
        (self.folder / "notes.md").write_text("x")
        self.assertEqual(self.problems(), ["c005: unexpected file notes.md"])
        (self.folder / "notes.md").unlink()
        (self.folder / "paraphrase_1.txt").write_text("Put the weld log rows that aren't Accepted on a Weld Log Exceptions sheet.\n")
        self.assertEqual(self.problems(), [])

    def test_a_paraphrase_keeps_every_fact_and_adds_no_number(self):   # D104
        self.assertEqual(vcl_corpus.facts_of((self.folder / "task.txt").read_text(encoding="utf-8"),
                                             (self.folder / "spec.vcl").read_text(encoding="utf-8")),
                         (["Accepted", "Exceptions"], []))
        para = self.folder / "paraphrase_2.txt"
        para.write_text("Copy the weld log rows whose fit-up or visual isn't accepted to an exceptions sheet.\n")
        self.assertEqual(self.problems(), ["c005: paraphrase_2.txt loses 'Accepted', which the request names",
                                           "c005: paraphrase_2.txt loses 'Exceptions', which the request names"])
        para.write_text("Copy the 3 kinds of weld log rows that aren't Accepted to a Weld Log Exceptions sheet.\n")
        self.assertEqual(self.problems(), ["c005: paraphrase_2.txt adds the number 3, which the request doesn't have"])
        para.write_text((self.folder / "task.txt").read_text(encoding="utf-8").upper().lower())
        self.assertEqual(self.problems()[0], "c005: paraphrase_2.txt is the same as the request or another paraphrase")
        para.write_text("Copy rows that aren\u2019t Accepted to the Weld Log Exceptions sheet, please.\n", encoding="utf-8")
        self.assertEqual(self.problems(), ["c005: paraphrase_2.txt must be 20 to 1500 characters of plain printable ASCII"])

    def test_a_number_of_the_request_must_stay_and_step_markers_dont_count(self):
        task = self.folder / "task.txt"
        task.write_text("(1) Copy every weld log row that isn't Accepted to a Weld Log Exceptions sheet if older than "
                        "30 days. (2) Tell me how many I copied.\n")
        for p in self.folder.glob("paraphrase_*.txt"):
            p.unlink()
        para = self.folder / "paraphrase_1.txt"
        para.write_text("(1) Weld log rows not Accepted and older than 30 days go to a Weld Log Exceptions sheet. (2) Count.\n")
        self.assertEqual(self.problems(), [])
        para.write_text("Weld log rows not Accepted and older than a month go to a Weld Log Exceptions sheet; count them.\n")
        self.assertEqual(self.problems(), ["c005: paraphrase_1.txt loses the number 30"])

    def test_a_held_out_pair_is_never_paraphrased(self):
        (self.data / "eval").mkdir()
        held = self.data / "eval" / "c005"
        self.folder.rename(held)
        self.assertEqual(vcl_corpus.check_item(held), ["c005: a held-out pair is never paraphrased (D59)"])

    def test_the_request_is_plain_text(self):
        (self.folder / "task.txt").write_text("Too short\n")
        self.assertIn("task.txt must be 20 to 1500 characters of plain printable ASCII", self.problems()[0])
        (self.folder / "task.txt").write_text("Copy the rows that aren't ‘accepted’ to a new sheet please.\n")
        self.assertIn("plain printable ASCII", self.problems()[0])

    def test_the_spec_must_be_canonical_valid_and_0_6(self):
        self.spec("PROC     ListWeldLogExceptions", "PROC       ListWeldLogExceptions")
        self.assertEqual(self.problems(), ["c005: spec.vcl is not canonical (run vclc fmt --write)"])
        (self.folder / "spec.vcl").write_text(
            without_goals((REAL / "verified" / "c005" / "spec.vcl").read_text(encoding="utf-8")), encoding="utf-8")
        self.assertTrue(any("must be VCL 0.6" in p for p in self.problems()), self.problems())
        self.spec("PROC     ListWeldLogExceptions", "PROC     listWeldLogExceptions")
        self.assertTrue(any("PROC must be PascalCase" in p for p in self.problems()))

    def test_a_warning_must_be_accepted_in_meta(self):
        shutil.rmtree(self.folder)
        shutil.copytree(REAL / "verified" / "c004", self.data / "verified" / "c004")
        self.folder = self.data / "verified" / "c004"
        self.assertEqual(self.problems(), [])
        self.meta(accepted_warnings=[])
        self.assertTrue(any("not in accepted_warnings" in p and "V053" in p for p in self.problems()), self.problems())

    def test_a_failing_test_is_reported(self):
        self.spec("assert=RESULT = ", "assert=RESULT = 1000 AND RESULT = ")
        self.assertTrue(any("TEST on line" in p and "fails" in p for p in self.problems()), self.problems())

    def test_meta_must_match_the_spec(self):
        self.meta(id="c999")
        self.assertIn("c005: meta.json id is 'c999'", self.problems())
        self.meta(id="c005", kind="BUILD", family="build")
        self.assertTrue(any("kind is 'BUILD'" in p for p in self.problems()))
        self.meta(kind="TASK", family="build")
        self.assertIn("c005: family build and kind BUILD go together", self.problems())
        self.meta(family="nonsense")
        self.assertTrue(any("family must be one of" in p for p in self.problems()))
        self.meta(family="rebuild", registers=["nonsense"])
        self.assertTrue(any("registers must be some of" in p for p in self.problems()))
        self.meta(registers=["weld_log"], ops=["LET"])
        self.assertTrue(any("ops or op_lines are stale" in p for p in self.problems()))

    def test_meta_json_has_one_form(self):
        meta = json.loads((self.folder / "meta.json").read_text(encoding="utf-8"))
        (self.folder / "meta.json").write_text(json.dumps(meta), encoding="utf-8")
        self.assertTrue(any("canonical form" in p for p in self.problems()))
        meta["extra"] = 1
        (self.folder / "meta.json").write_text(vcl_corpus.dump(meta), encoding="utf-8")
        self.assertTrue(any("keys must be" in p for p in self.problems()))

    def test_the_mutation_floor(self):
        total = json.loads((self.folder / "meta.json").read_text(encoding="utf-8"))["mutation"]["total"]
        self.meta(mutation={"killed": total - 5, "total": total, "equivalent": 0, "note": ""})
        self.assertTrue(any("under the 90% floor" in p for p in self.problems()), self.problems())
        self.meta(mutation={"killed": total - 5, "total": total, "equivalent": 5, "note": ""})
        self.assertTrue(any("needs a note" in p for p in self.problems()), self.problems())
        self.meta(mutation={"killed": total - 5, "total": total, "equivalent": 5, "note": "they cannot change a result"})
        self.assertEqual(self.problems(), [])
        self.meta(mutation={"killed": total + 1, "total": total, "equivalent": 0, "note": ""})
        self.assertTrue(any("outside 0 to" in p for p in self.problems()))

    def test_mutation_is_run_again_on_request(self):
        self.assertEqual(self.problems(mutation=True), [])
        self.meta(mutation={"killed": 1, "total": 2, "equivalent": 0, "note": ""})
        self.assertTrue(any("mutation is 25 of 25 now" in p for p in self.problems(mutation=True)), self.problems(mutation=True))

    def test_an_id_cannot_be_in_both_sets(self):
        (self.data / "eval").mkdir()
        shutil.copytree(self.folder, self.data / "eval" / "c005", ignore=shutil.ignore_patterns("paraphrase_*"))
        self.assertEqual(vcl_corpus.check_all(), ["c005: in both data/verified and data/eval"])


class AddingAndExporting(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.saved = vcl_corpus.DATA
        vcl_corpus.DATA = Path(self.tmp.name)

    def tearDown(self):
        vcl_corpus.DATA = self.saved
        self.tmp.cleanup()

    REQUEST = "Delete every cancelled order, then sort what is left by order date, newest first."

    def test_a_recipe_only_kronos_offline_runs_is_stored_without_vba(self):   # D137: Word, and two workbooks
        for item, name, family, registers in (("c901", "doc_estimate.vcl", "documents", ["estimate_sheet", "estimation"]),
                                              ("c902", "wps_from_wps.vcl", "documents", ["wps"])):
            with self.subTest(recipe=name):
                spec = (ROOT / "tests" / "data" / name).read_text(encoding="utf-8")
                found = vcl_corpus.add_item(item, "Fill the practice document from the practice workbook as the recipe says.",
                                            spec, registers, family, source="a test", equivalent=0, note="",
                                            accepted_warnings=["V344"] if item == "c902" else ["V389"])
                floor = [p for p in found if "floor" in p]
                self.assertEqual([p for p in found if p not in floor and "survives" not in p], [])
                self.assertFalse(any("does not build" in p for p in found))
        meta = json.loads((vcl_corpus.DATA / "verified" / "c901" / "meta.json").read_text(encoding="utf-8"))
        self.assertEqual((meta["kind"], meta["ops"]), ("DOCUMENT", ["FIELD", "REPEAT"]))

    def test_a_proven_spec_is_stored_in_canonical_form(self):
        untidy = example("example1_rev3.vcl").replace("PRE      ", "PRE  ")
        problems = vcl_corpus.add_item("c101", self.REQUEST, untidy, ["dispatch"], "clean", source="the prompt's Example 1")
        self.assertEqual(problems, [])
        folder = vcl_corpus.DATA / "verified" / "c101"
        self.assertEqual((folder / "spec.vcl").read_text(encoding="utf-8"), example("example1_rev3.vcl"))
        meta = json.loads((folder / "meta.json").read_text(encoding="utf-8"))
        self.assertEqual((meta["kind"], meta["family"], meta["source"]), ("TASK", "clean", "the prompt's Example 1"))
        self.assertGreater(meta["mutation"]["total"], 0)
        self.assertEqual(vcl_corpus.check_all(), [])

    def test_a_pair_that_fails_a_gate_is_not_stored(self):
        problems = vcl_corpus.add_item("c102", self.REQUEST, example("example1_rev3.vcl"), ["nonsense"], "clean")
        self.assertTrue(any("registers must be some of" in p for p in problems))
        self.assertFalse((vcl_corpus.DATA / "verified" / "c102").exists())
        problems = vcl_corpus.add_item("c103", self.REQUEST, "VCL 0.6\nTASK\nPROC     x\n", ["bom"], "clean")
        self.assertTrue(problems and "can't be formatted" in problems[0])
        self.assertFalse((vcl_corpus.DATA / "verified" / "c103").exists())

    def test_an_id_is_used_once(self):
        self.assertEqual(vcl_corpus.add_item("c101", self.REQUEST, example("example1_rev3.vcl"), ["dispatch"], "clean"), [])
        self.assertEqual(vcl_corpus.add_item("c101", self.REQUEST, example("example1_rev3.vcl"), ["dispatch"], "clean"),
                         ["c101: already there"])

    def test_the_export_has_the_vba_compiled_fresh(self):
        vcl_corpus.add_item("c101", self.REQUEST, example("example1_rev3.vcl"), ["dispatch"], "clean")
        vcl_corpus.add_item("c102", "Mark old open NCRs Overdue and keep a summary of the NCRs by status.",
                            example("example4_build.vcl"), ["ncr"], "build")
        out = Path(self.tmp.name) / "out" / "dataset.jsonl"
        self.assertEqual(vcl_corpus.export(out), 2)
        rows = [json.loads(line) for line in out.read_text(encoding="utf-8").splitlines()]
        self.assertEqual([r["id"] for r in rows], ["c101", "c102"])
        task, build = rows
        self.assertEqual(task["goal"], "Delete every cancelled order from the Orders sheet, then sort the rest by order date, newest first.")
        self.assertIn("Function RemoveCancelledAndSortByDate(", task["vba"])
        self.assertIn("Function RunHeadless(", build["vba"])
        self.assertNotIn("Option Explicit", build["vba"])   # the procedures only: ForgeLib is not in the dataset
        for row in rows:
            self.assertEqual(set(row), {"id", "kind", "registers", "family", "request", "goal", "spec", "vba", "paraphrases"})


if __name__ == "__main__":
    unittest.main()
