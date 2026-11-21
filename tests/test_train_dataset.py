"""The training dataset export (train/make_dataset.py, D124): what goes in, and what must never.

The repairs come from small made-up rows, so these tests run in seconds; the real `vclc datagen` output is checked by
the export's own guards when Boss or a session makes the dataset.
"""
import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "train"))
sys.path.insert(0, str(ROOT / "compiler"))

import make_dataset as md  # noqa: E402
import prompts  # noqa: E402
from vcl_corpus import load  # noqa: E402


def repair_row(pair, change, source="mutant", broken="BROKEN\n"):
    return {"type": "repair", "source": source, "id": pair["id"], "change": change, "stage": "check",
            "request": pair["request"], "broken": broken, "feedback": ["spec.vcl:2: error: made up [V001]"],
            "fixed": pair["spec"]}


class Prompts(unittest.TestCase):
    def test_the_system_prompt_is_short_and_says_recipe_only(self):
        self.assertLess(len(prompts.SYSTEM), 400)
        self.assertIn("recipe only", prompts.SYSTEM)

    def test_a_repair_prompt_holds_the_request_the_recipe_and_the_messages(self):
        m = prompts.repair_prompt("Do x.", "OLD\n", ["a", "b"])
        self.assertEqual([x["role"] for x in m], ["system", "user"])
        for part in ("Do x.", "OLD", "a\nb", "whole corrected recipe"):
            self.assertIn(part, m[1]["content"])


class Dataset(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.pairs = load("verified")
        cls.held = load("eval")
        rows = []
        for p in cls.pairs:
            rows += [repair_row(p, f"line {i}", broken=f"BROKEN {p['id']} {i}\n") for i in range(1, 16)]
            rows.append(repair_row(p, "try1.vcl", "attempt", f"TRY {p['id']}\n"))
        cls.tmp = tempfile.TemporaryDirectory()
        cls.path = Path(cls.tmp.name) / "repairs.jsonl"
        cls.path.write_text("".join(json.dumps(r) + "\n" for r in rows), encoding="utf-8")
        cls.train, cls.val, cls.manifest = md.build(cls.path, per_pair=10)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_it_has_both_kinds_and_a_validation_part(self):
        self.assertTrue(self.train and self.val)
        self.assertEqual({e["type"] for e in self.train + self.val}, {"write", "repair"})
        self.assertEqual(len(self.pairs), len(self.manifest["pairs"]["train"]) + len(self.manifest["pairs"]["val"]))

    def test_every_pair_gives_a_write_example_for_its_request_and_each_paraphrase(self):
        by_id = {e["id"]: e for e in self.train + self.val}
        for p in self.pairs:
            for k, text in enumerate([p["request"]] + p["paraphrases"]):
                e = by_id[f"{p['id']}/write{k}"]
                self.assertEqual([m["role"] for m in e["messages"]], ["system", "user", "assistant"])
                self.assertEqual(e["messages"][1]["content"], text.strip())
                self.assertEqual(e["messages"][2]["content"], p["spec"])

    def test_the_held_out_pairs_are_nowhere(self):
        self.assertGreater(len(self.held), 0)
        everything = json.dumps(self.train + self.val, ensure_ascii=False)
        for item in self.held:
            self.assertNotIn(item["spec"].strip()[:200].replace("\n", "\\n"), everything)
            self.assertFalse([e for e in self.train + self.val if e["id"].startswith(item["id"] + "/")])
        self.assertEqual(self.manifest["held_out_checked"], len(self.held))

    def test_the_guard_stops_a_leak_of_a_request_or_a_recipe(self):
        item = self.held[0]
        bad = [{"id": "x/write0", "type": "write", "source": "", "messages": prompts.write_prompt(item["request"]) +
                [{"role": "assistant", "content": "ok"}]}]
        with self.assertRaises(md.LeakError):
            md.check_no_leaks(bad, [item])
        bad = [{"id": "x/write0", "type": "write", "source": "", "messages": prompts.write_prompt("q") +
                [{"role": "assistant", "content": item["spec"]}]}]
        with self.assertRaises(md.LeakError):
            md.check_no_leaks(bad, [item])

    def test_the_split_is_by_pair(self):
        sides = {}
        for side, rows in (("train", self.train), ("val", self.val)):
            for e in rows:
                sides.setdefault(e["id"].split("/")[0], set()).add(side)
        self.assertTrue(all(len(s) == 1 for s in sides.values()))
        self.assertEqual({p for p, s in sides.items() if s == {"val"}}, set(self.manifest["pairs"]["val"]))

    def test_mutant_repairs_are_capped_and_the_real_tries_are_all_kept(self):
        for p in self.pairs:
            rows = [e for e in self.train + self.val if e["id"].startswith(p["id"] + "/repair/")]
            self.assertEqual(sum(1 for e in rows if e["source"] == "mutant"), 10)
            self.assertEqual(sum(1 for e in rows if e["source"] == "attempt"), 1)

    def test_a_repair_example_answers_with_the_proven_recipe(self):
        e = next(e for e in self.train if e["type"] == "repair")
        pair = next(p for p in self.pairs if p["id"] == e["id"].split("/")[0])
        self.assertEqual(e["messages"][2]["content"], pair["spec"])
        self.assertIn("The toolchain's messages", e["messages"][1]["content"])
        self.assertNotIn("```", e["messages"][2]["content"])

    def test_the_same_input_gives_the_same_files(self):
        with tempfile.TemporaryDirectory() as a, tempfile.TemporaryDirectory() as b:
            for out in (a, b):
                t, v, m = md.build(self.path, per_pair=10)
                md.write_all(out, t, v, m)
            for name in ("train.jsonl", "val.jsonl", "manifest.json"):
                self.assertEqual((Path(a) / name).read_bytes(), (Path(b) / name).read_bytes())

    def test_a_repair_for_an_unknown_pair_stops_the_export(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "repairs.jsonl"
            p.write_text(json.dumps(repair_row({"id": "c999", "request": "r", "spec": "s"}, "x")) + "\n")
            with self.assertRaises(md.LeakError):
                md.build(p)


if __name__ == "__main__":
    unittest.main()
