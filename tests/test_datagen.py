"""Stage A Phase 7 (D103): the data generator, `vclc datagen`.

Every example must be what it says: a repair's broken spec is refused by the toolchain (or its TESTs are too weak)
and its fixed spec is the proven pair; a contrast spec passes its own TESTs yet ends differently. Held-out pairs
are never used.
"""
import datetime as dt
import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
import vcl_datagen  # noqa: E402
import vclc  # noqa: E402
from kronos_engine import prove  # noqa: E402
from vcl_corpus import FLOOR, items  # noqa: E402
from vcl_interp import run_tests  # noqa: E402
from vcl_mutate import mutate_report  # noqa: E402

TODAY = dt.date(2026, 11, 19)


def rows(path):
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines()]


class OnePair(unittest.TestCase):
    """c020 has mutants of every kind and one failed try of a helper (weak TESTs)."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.out = Path(cls.tmp.name)
        cls.summary = vcl_datagen.generate(cls.out, TODAY, only={"c020"})
        cls.repairs = rows(cls.out / "repairs.jsonl")
        cls.contrast = rows(cls.out / "contrast.jsonl")
        cls.spec = (ROOT / "data" / "verified" / "c020" / "spec.vcl").read_text(encoding="utf-8")

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_every_repair_is_refused_and_fixed_by_the_proven_pair(self):
        self.assertTrue(self.repairs)
        for row in self.repairs:
            with self.subTest(change=row["change"]):
                self.assertEqual(row["fixed"], self.spec)
                self.assertNotEqual(row["broken"], row["fixed"])
                self.assertTrue(row["feedback"])
                proof = prove(row["broken"], TODAY)
                if row["stage"] == "mutation":
                    caught, total, _ = mutate_report(row["broken"], snapshot=True, today=TODAY)
                    self.assertTrue(proof.ok)
                    self.assertLess(caught / total, FLOOR)
                else:
                    self.assertFalse(proof.ok)
                if row["stage"] == "check":
                    self.assertTrue(any("[V" in line or "[T" in line or "[G" in line for line in row["feedback"]),
                                    row["feedback"])

    def test_every_contrast_passes_its_own_tests_and_ends_differently(self):
        self.assertTrue(self.contrast)
        golden = vcl_datagen.snapshot(run_tests(self.spec, TODAY))
        for row in self.contrast:
            with self.subTest(change=row["change"]):
                self.assertTrue(prove(row["mutant"], TODAY).ok)
                self.assertNotEqual(vcl_datagen.snapshot(run_tests(row["mutant"], TODAY)), golden)
                self.assertEqual(row["proven"], self.spec)
                self.assertTrue(row["differs"].startswith("TEST on line "), row["differs"])

    def test_the_helpers_weak_try_is_a_repair(self):
        tries = [r for r in self.repairs if r["source"] == "attempt"]
        self.assertEqual([(r["change"], r["stage"]) for r in tries], [("try1.vcl", "mutation")])
        self.assertIn("under the 90% floor", tries[0]["feedback"][0])

    def test_the_summary_counts_the_files(self):
        self.assertEqual(self.summary["repairs"], len(self.repairs))
        self.assertEqual(self.summary["contrast"], len(self.contrast))
        self.assertEqual(self.summary["today"], "2026-11-19")
        self.assertEqual(json.loads((self.out / "summary.json").read_text(encoding="utf-8")), self.summary)


class Rules(unittest.TestCase):
    def test_held_out_pairs_are_never_used(self):
        held = {p.name for p in items("eval")}
        self.assertTrue(held)
        with tempfile.TemporaryDirectory() as tmp:
            summary = vcl_datagen.generate(tmp, TODAY, only=held)
        self.assertEqual(summary["pairs"], {})

    def test_two_runs_give_the_same_files(self):
        with tempfile.TemporaryDirectory() as a, tempfile.TemporaryDirectory() as b:
            vcl_datagen.generate(a, TODAY, only={"c013"})
            vcl_datagen.generate(b, TODAY, only={"c013"})
            for name in ("repairs.jsonl", "contrast.jsonl"):
                self.assertEqual((Path(a) / name).read_bytes(), (Path(b) / name).read_bytes())

    def test_vclc_runs_it(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(vclc.main(["vclc", "datagen", "-o", tmp, "--today", "2026-11-19", "--only", "c013"]), 0)
            self.assertTrue((Path(tmp) / "repairs.jsonl").stat().st_size > 0)


if __name__ == "__main__":
    unittest.main()
