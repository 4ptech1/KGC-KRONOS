"""KRONOS Offline's wizard (D89): compiler/kronos_wizard.py drives KRONOS.exe's own engine with copy and paste.

The big check: for the same replies, the wizard gives the AI exactly the messages KRONOS.exe gives Codex, and saves
exactly the .bas KRONOS.exe saves.
"""
import datetime as dt
import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))

import kronos_engine as engine  # noqa: E402
import kronos_web  # noqa: E402
import kronos_wizard as kw  # noqa: E402
from kronos_ai import FakeAI, render  # noqa: E402

TODAY = dt.date(2026, 11, 16)
EXAMPLE = (ROOT / "vcl" / "examples" / "example1_rev3.vcl").read_text(encoding="utf-8")
BROKEN = EXAMPLE.replace("GOAL ", "GOLA ", 1)
REQUEST = "Delete every cancelled order from the Orders sheet, then sort the rest by order date, newest first."
QUESTION = "```vcl\n# QUESTION: Which sheet holds the orders?\n```"
PROMPT = engine.PROMPT.read_text(encoding="utf-8")


def reply(spec):
    return f"Here is the spec.\n\n```vcl\n{spec}```\n"


class Wizard(unittest.TestCase):
    def setUp(self):
        self.data = Path(tempfile.mkdtemp(prefix="kronos wizard "))
        self.addCleanup(shutil.rmtree, self.data, True)
        self.wizard = kw.Wizard(self.data / "browser", TODAY)

    def test_the_same_messages_as_kronos_exe(self):
        replies = [QUESTION, reply(BROKEN), reply(EXAMPLE)]
        fake = FakeAI(replies)
        exe = engine.Run(REQUEST, fake, self.data / "exe", TODAY)
        self.assertEqual(exe.start(), "question")
        self.assertEqual(exe.answer("The sheet is Orders."), "proven")

        w = self.wizard
        self.assertEqual(w.ask("r1", REQUEST)["status"], "waiting")
        self.assertEqual(w.message("r1"), render(PROMPT, fake.seen[0]))
        self.assertEqual(w.message("r1", instructions=False), REQUEST + "\n")
        view = w.paste("r1", replies[0])
        self.assertEqual((view["status"], view["question"]), ("question", "Which sheet holds the orders?"))
        self.assertEqual(w.answer("r1", "The sheet is Orders.")["status"], "waiting")
        self.assertEqual(w.message("r1"), render(PROMPT, fake.seen[1]))
        view = w.paste("r1", replies[1])
        self.assertEqual((view["status"], view["newest"], view["tries"]), ("waiting", "repair", 1))
        self.assertEqual(w.message("r1"), render(PROMPT, fake.seen[2]))
        self.assertTrue(w.message("r1", instructions=False).startswith("The VBA Forge toolchain checked your spec"))
        view = w.paste("r1", replies[2])
        self.assertEqual(view["status"], "proven")
        self.assertTrue(view["best"]["proven"])
        self.assertEqual(view["best"]["summary"], engine.summary(exe.proven().proof))
        self.assertEqual(view["starter"], [{"name": "ws", "type": "Worksheet", "default": "Orders"}])
        self.assertEqual([a["proven"] for a in view["attempts"]], [False, True])

        saved = w.save("r1", {"ws": "Orders"})
        path = exe.save_module({"ws": "Orders"})
        self.assertEqual(saved["bas"].encode("ascii"), path.read_bytes())   # byte for byte KRONOS.exe's module
        self.assertEqual(saved["bas_name"], path.name)
        self.assertEqual(saved["practice"].encode("ascii"), exe.practice_file.read_bytes())
        self.assertEqual(saved["macro"], exe.macro)
        self.assertEqual(saved["steps"], engine.import_steps(path.name, exe.macro))

    def test_three_tries_then_the_closest_attempt(self):
        w = self.wizard
        w.ask("r1", REQUEST)
        for n in range(1, 4):
            view = w.paste("r1", reply(BROKEN))
        self.assertEqual(view["status"], "not_proven")
        self.assertIn("no spec passed after 3 tries", view["note"])
        self.assertFalse(view["best"]["proven"])
        self.assertTrue(view["best"]["problems"])
        with self.assertRaises(kw.WizardError):
            w.paste("r1", reply(EXAMPLE))   # the turn is over: KRONOS isn't waiting any more

    def test_a_recipe_from_a_file_and_its_problems_sent_to_the_ai(self):
        w = self.wizard
        view = w.recipe("f1", EXAMPLE, "recipes.docx")
        self.assertEqual(view["status"], "proven")
        self.assertIn("recipes.docx", view["request"])
        view = w.recipe("f2", BROKEN, "recipes.docx")
        self.assertEqual(view["status"], "not_proven")
        view = w.send_problems("f2")
        self.assertEqual((view["status"], view["newest"]), ("waiting", "fix"))
        text = w.message("f2", instructions=False)
        self.assertTrue(text.startswith("I edited your spec by hand."))
        self.assertIn("GOLA", text)
        self.assertEqual(w.paste("f2", reply(EXAMPLE))["status"], "proven")

    def test_a_workbook_argument_needs_no_name_the_starter_passes_this_workbook(self):
        summary = (ROOT / "docs" / "demo" / "recipe_3_summary.vcl").read_text(encoding="utf-8")
        view = self.wizard.recipe("f1", summary, "recipe_3_summary.vcl")
        self.assertEqual(view["starter"], [{"name": "ws", "type": "Worksheet", "default": "NCR"}])
        saved = self.wizard.save("f1", {"ws": "NCR Register"})
        self.assertIn('SummariseNcrs(ThisWorkbook, ThisWorkbook.Worksheets("NCR Register"))', saved["bas"])
        self.assertEqual(saved["macro"], "RunSummariseNcrs")

    def test_the_workbook_goes_below_the_request(self):
        self.assertEqual(kw.request_text(REQUEST, None), REQUEST)
        text = kw.request_text(REQUEST, "The workbook orders.xlsx has 1 sheet:\n")
        self.assertEqual(text, f"{REQUEST}\n\n{kw.BOOK_HEADING}\nThe workbook orders.xlsx has 1 sheet:")
        self.wizard.ask("r1", REQUEST, "The workbook orders.xlsx has 1 sheet:")
        self.assertIn(kw.BOOK_HEADING, self.wizard.message("r1", instructions=False))

    def test_restore_rebuilds_every_run(self):
        actions = [["r1", "ask", REQUEST, None], ["r1", "paste", reply(BROKEN)], ["f1", "recipe", EXAMPLE, "a.vcl"],
                   ["r1", "paste", reply(EXAMPLE)]]
        for action in actions:
            getattr(self.wizard, {"ask": "ask", "paste": "paste", "recipe": "recipe"}[action[1]])(action[0], *action[2:])
        before = {id: self.wizard.view(id) for id in self.wizard.runs}
        again = kw.Wizard(self.data / "again", TODAY).restore(actions)
        self.assertEqual(again, before)

    def test_what_the_person_must_put_right(self):
        w = self.wizard
        for call, words in ((lambda: w.ask("r1", "too short"), "Write a little more"),
                            (lambda: w.paste("nope", "x"), "start it again"),
                            (lambda: w.recipe("f1", "  "), "empty")):
            with self.assertRaises(kw.WizardError) as caught:
                call()
            self.assertIn(words, str(caught.exception))
        w.ask("r1", REQUEST)
        for call, words in ((lambda: w.answer("r1", "x"), "hasn't asked"), (lambda: w.paste("r1", " "), "Paste"),
                            (lambda: w.save("r1"), "Only a proven"), (lambda: w.send_problems("r1"), "no recipe")):
            with self.assertRaises(kw.WizardError) as caught:
                call()
            self.assertIn(words, str(caught.exception))
        w.paste("r1", reply(EXAMPLE))
        with self.assertRaises(kw.WizardError) as caught:
            w.save("r1", {})
        self.assertIn("Fill in ws", str(caught.exception))


class Bridge(unittest.TestCase):
    def test_the_page_gets_canonical_json_and_words(self):
        folder = Path(tempfile.mkdtemp(prefix="kronos bridge "))
        self.addCleanup(shutil.rmtree, folder, True)
        kronos_web.wizard_reset(str(folder), "2026-11-16")
        view = json.loads(kronos_web.wizard_call("ask", "r1", json.dumps([REQUEST, None])))
        self.assertEqual((view["ok"], view["view"]["status"]), (True, "waiting"))
        text = json.loads(kronos_web.wizard_call("message", "r1", json.dumps([True])))["text"]
        self.assertTrue(text.startswith(PROMPT.rstrip("\n")))
        bad = json.loads(kronos_web.wizard_call("answer", "r1", json.dumps(["x"])))
        self.assertEqual((bad["ok"], bad["message"]), (False, "The AI hasn't asked a question."))
        view = json.loads(kronos_web.wizard_call("paste", "r1", json.dumps([reply(EXAMPLE)])))["view"]
        self.assertEqual(view["status"], "proven")
        saved = json.loads(kronos_web.wizard_call("save", "r1", json.dumps([{"ws": "Orders"}])))
        self.assertTrue(saved["ok"] and saved["saved"]["bas"].startswith("Attribute VB_Name"))
        words = json.loads(kronos_web.words())
        self.assertIn("Calibration", words["example"])
        self.assertEqual(len(words["prompt_sha256"]), 64)


if __name__ == "__main__":
    unittest.main()
