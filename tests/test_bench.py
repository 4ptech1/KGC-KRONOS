"""The benchmark, `vclc bench` (VEP 0.6 Stage A Phase 8, D135): the only scoreboard.

The reference recipes reach H on every held-out item, a writer sees the request and the contract and nothing of how the
reference works, and a candidate is stopped at the first level it misses, with the reason.
"""
import datetime as dt
import re
import sys
import tempfile
import unittest
from unittest import mock
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
import vcl_bench as B  # noqa: E402
import vclc  # noqa: E402
from vcl_corpus import items  # noqa: E402
from vcl_interp import run_tests  # noqa: E402
from vcl_mutate import mutants  # noqa: E402
from kronos_engine import prove  # noqa: E402

TODAY = dt.date(2026, 11, 25)
EVAL = ROOT / "data" / "eval"


def ref(item):
    return (EVAL / item / "spec.vcl").read_text(encoding="utf-8")


class Items(unittest.TestCase):
    def test_the_reference_recipes_reach_h_on_every_item(self):
        with tempfile.TemporaryDirectory() as tmp:
            for folder in items("eval"):
                (Path(tmp) / f"{folder.name}.vcl").write_text(ref(folder.name), encoding="utf-8")
            rows = B.run(tmp, TODAY)
        self.assertEqual(len(rows), 15)
        self.assertEqual({k: r["reached"] for k, r in rows.items() if r["reached"] != "T+M+H"}, {})
        self.assertEqual((B.summary(rows)["headline"], B.summary(rows)["strict"]), (1.0, 15))

    def test_a_missing_candidate_reaches_nothing(self):
        with tempfile.TemporaryDirectory() as tmp:
            rows = B.run(tmp, TODAY)
        self.assertTrue(all(r["reached"] == "" and r["why"] == "no candidate" for r in rows.values()))
        self.assertEqual(B.summary(rows)["H"], 0)


class Contract(unittest.TestCase):
    def test_a_task_gives_its_proc_kind_args_and_ret(self):
        kind, lines = B.contract(ref("c022"))
        self.assertEqual(kind, "TASK")
        self.assertEqual([ln.split()[0] for ln in lines], ["PROC", "KIND", "ARGS", "RET"])
        self.assertIn("FlagUnqualifiedWelds", lines[0])

    def test_a_build_gives_its_name_and_runheadless(self):
        kind, lines = B.contract(ref("c089"))
        self.assertEqual(kind, "BUILD")
        self.assertEqual(len(lines), 2)
        self.assertTrue(lines[0].startswith("NAME") and "modMain.RunHeadless(" in lines[1])

    def test_the_prompt_shows_the_request_and_the_contract_only(self):
        for folder in items("eval"):
            text = B.prompt(folder)
            with self.subTest(item=folder.name):
                self.assertTrue(text.startswith((folder / "task.txt").read_text(encoding="utf-8").strip()))
                self.assertNotRegex(text, r"(?m)^(FIXTURE|TEST|OP|PRE|MODIFIES|GOAL)\s")
                for line in B.contract(ref(folder.name))[1]:
                    self.assertIn(line, text)

    def test_vclc_writes_the_prompts(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(vclc.main(["vclc", "bench", "--prompts", tmp]), 0)
            self.assertEqual(len(list(Path(tmp).glob("c*.txt"))), 15)


class Levels(unittest.TestCase):
    def test_no_recipe_reaches_nothing(self):
        out = B.score("Sure, I can help with that.", ref("c022"), TODAY)
        self.assertEqual((out["reached"], out["why"]), ("", "no recipe in it"))

    def test_a_recipe_that_does_not_check_stops_before_v(self):
        out = B.score(ref("c022").replace("PROC     ", "PROCX    ", 1), ref("c022"), TODAY)
        self.assertEqual(out["reached"], "")
        self.assertTrue(out["why"].startswith("it does not check"))

    def test_another_name_than_the_contract_is_never_right(self):
        renamed = ref("c022").replace("FlagUnqualifiedWelds", "FlagWelds")
        out = B.score(renamed, ref("c022"), TODAY)
        self.assertEqual(out["reached"], "T+M")
        self.assertIn("FlagUnqualifiedWelds", out["why"])

    def test_weak_tests_of_its_own_do_not_help_a_wrong_recipe(self):
        """A mutant the reference TESTs catch, whose catching TESTs are weakened to check ERROR_COUNT only (the validator
        refuses a recipe with fewer TESTs): its own TESTs pass, the hidden TESTs don't."""
        reference = ref("c030")
        lines = reference.splitlines()
        tests = [ln for ln in lines if ln.startswith("TEST ")]
        for i, _, body in mutants(reference):
            if body is None:
                continue
            new = list(lines)
            new[i] = "OP" + body
            text = "\n".join(new) + "\n"
            try:
                results = run_tests(text, TODAY)
            except Exception:
                continue
            failing = {t for t, r in zip(tests, results) if not r.passed}
            if failing:   # the TESTs that catch it keep their calls but check only ERROR_COUNT
                weak = [ln.split("assert=")[0] + "assert=ERROR_COUNT = 0" if ln in failing else ln for ln in new]
                candidate = "\n".join(weak) + "\n"
                if prove(candidate, TODAY, offline=True).ok:   # its own TESTs pass and it checks
                    break
        else:
            self.fail("no mutant to try")
        out = B.score(candidate, reference, TODAY)
        self.assertTrue(out["T"])
        self.assertFalse(out["H"])
        self.assertIn(out["reached"], ("T", "T+M"))
        self.assertIn("hidden TESTs fail", out["why"])

    def test_a_right_macro_below_the_floor_is_right_but_not_strict(self):
        """M and H are scored apart: with the floor set above 100%, the reference macro misses M and still reaches H."""
        with mock.patch.object(B, "FLOOR", 1.01):
            out = B.score(ref("c030"), ref("c030"), TODAY)
        self.assertEqual((out["T"], out["M"], out["H"]), (True, False, True))
        self.assertEqual(out["reached"], "T+H")
        self.assertEqual(B.summary({"c030": out})["strict"], 0)

    def test_the_swap_puts_the_reference_tests_where_its_own_stood(self):
        candidate = ref("c022")
        swapped = B.swap_tests(candidate.replace("runs=2", "runs=3"), ref("c022"))
        self.assertEqual(swapped, candidate)
        swapped = B.swap_tests(ref("c089"), ref("c089"))
        self.assertEqual(swapped, ref("c089"))


if __name__ == "__main__":
    unittest.main()
