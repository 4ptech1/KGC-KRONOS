"""Item 5 Phase 2a (D74, D75): the KRONOS window's controller, the four-step wizard, with a fake AI and no
display."""
import datetime
import sys
import tempfile
import threading
import time
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "app"))
sys.path.insert(0, str(ROOT / "compiler"))
from kronos_ai import AiError, CodexAI, FakeAI  # noqa: E402
from kronos_controller import Controller  # noqa: E402

DAY = datetime.date(2026, 10, 21)
EXAMPLES = ROOT / "vcl" / "examples"
GOOD = (EXAMPLES / "example1_rev3.vcl").read_text(encoding="utf-8")
BUILD = (EXAMPLES / "example4_build.vcl").read_text(encoding="utf-8")
BROKEN_TEST = GOOD.replace('COUNTIF(Orders![Status], "Cancelled") = 0 AND ERROR_COUNT = 0',
                           'COUNTIF(Orders![Status], "Cancelled") = 5 AND ERROR_COUNT = 0')
BROKEN_CHECK = GOOD.replace("OP       LET(removed, 0)\n", "OP       FROB(removed, 0)\n")
QUESTION = "```vcl\n# QUESTION: Which sheet holds the orders?\n```"
REQUEST = "In the Orders sheet, delete every cancelled order, then sort the rest by date, newest first."


def block(spec):
    return f"```vcl\n{spec}```\n"


class Wizard(unittest.TestCase):
    def controller(self, *replies, ready=True, threaded=False):
        self.ai = FakeAI(replies)
        c = Controller(Path(tempfile.mkdtemp()) / "KGC KRONOS", ai=self.ai, today=DAY, threaded=threaded)
        if ready:
            c.confirm_checklist(True, True)
        return c

    def ask(self, c, text=REQUEST):
        c.ask(text)
        c.poll()
        return c

    def test_the_checklist_comes_first(self):
        c = self.controller(block(GOOD), ready=False)
        self.assertFalse(c.can_ask(REQUEST))
        with self.assertRaises(ValueError):
            c.ask(REQUEST)
        self.assertEqual([(item, ok) for item, ok, detail in c.checklist()],
                         [("codex_found", True), ("codex_flags", True), ("data_controls", False),
                          ("data_folder", True), ("practice", False)])
        c.confirm_checklist(True, False)
        self.assertFalse(c.can_ask(REQUEST))
        c.confirm_checklist(True, True)
        self.assertTrue(c.can_ask(REQUEST))
        self.assertFalse(c.can_ask("too short"))
        again = Controller(c.data, ai=self.ai)   # kept in settings.json
        self.assertTrue(again.checklist_done)

    def test_a_proven_recipe_first_time(self):
        c = self.ask(self.controller(block(GOOD)))
        self.assertEqual((c.step, c.banner, c.working, c.calls), ("review", "proven", False, 1))
        self.assertEqual(c.step_bar(), [("ask", "done"), ("question", "skipped"), ("review", "current"),
                                        ("save", "todo")])
        self.assertEqual(c.lines, ["Try 1: it passed every check and 3 of 3 practice tests"])
        review = c.review()
        self.assertEqual(review["summary"]["changes"],
                         ['anything on the sheet called "ws" in the recipe (Orders in the practice tests)'])
        self.assertEqual(len(review["summary"]["assumes"]), 3)
        self.assertEqual([(t["line"], t["passed"]) for t in review["tests"]], [(29, True), (30, True), (31, True)])

    def test_a_question_and_its_answer(self):
        c = self.ask(self.controller(QUESTION, block(GOOD)))
        self.assertEqual((c.step, c.question), ("question", "Which sheet holds the orders?"))
        self.assertEqual(c.step_bar()[1], ("question", "current"))
        with self.assertRaises(ValueError):
            c.change("not now")
        c.answer("The sheet is called Orders.")
        c.poll()
        self.assertEqual((c.step, c.banner), ("review", "proven"))
        self.assertEqual(c.step_bar()[1], ("question", "done"))

    def test_a_repair_and_the_working_panel(self):
        c = self.controller(block(BROKEN_TEST), block(GOOD), threaded=False)
        c.ask(REQUEST)
        events = c.poll()
        kinds = [e.get("kind") for e in events if e["type"] == "progress"]
        self.assertEqual(kinds[:6], ["asking", "checking", "testing", "tried", "asking", "checking"])
        self.assertEqual(c.lines, ["Try 1: 2 of 3 practice tests passed, so KRONOS sent the problems back to "
                                   "the practice AI", "Try 2: it passed every check and 3 of 3 practice tests"])
        self.assertEqual((c.banner, c.calls), ("proven", 2))

    def test_not_proven_after_three_tries(self):
        c = self.ask(self.controller(block(BROKEN_CHECK), block(BROKEN_TEST), block(BROKEN_CHECK)))
        self.assertEqual((c.step, c.banner), ("review", "not_proven"))
        self.assertEqual(c.lines[-1], "Try 3: the recipe broke 1 rule of the checker")   # not sent back
        self.assertIn("no spec passed after 3 tries", c.review()["note"])
        c.set_assumptions(True)
        self.assertFalse(c.can_save())
        with self.assertRaises(ValueError):
            c.go_save()

    def test_try_again_after_not_proven(self):   # the same request, as a new run
        c = self.ask(self.controller(block(BROKEN_TEST), block(BROKEN_TEST), block(BROKEN_TEST), block(GOOD)))
        first = c.run
        c.again()
        c.poll()
        self.assertEqual((c.step, c.banner, c.request), ("review", "proven", REQUEST))
        self.assertIsNot(c.run, first)
        self.assertEqual(len(c.history()), 2)

    def test_no_reply_and_try_again(self):
        c = self.ask(self.controller(AiError("Codex gave no reply within 300 seconds"), block(GOOD)))
        self.assertEqual((c.step, c.notice), ("ask", "the AI gave no reply: Codex gave no reply within 300 seconds"))
        c.retry()
        c.poll()
        self.assertEqual((c.step, c.banner, c.notice), ("review", "proven", ""))

    def test_cancel_stops_the_ai(self):   # threaded, as in the window
        started = threading.Event()

        def slow(conversation):
            started.set()
            self.ai.cancel_event.wait(10)
            raise AiError("cancelled")

        c = self.controller(slow, block(GOOD), threaded=True)
        c.ask(REQUEST)
        self.assertTrue(started.wait(10))
        self.assertTrue(c.working)
        c.cancel()
        self.wait_for(c)
        self.assertEqual((c.step, c.notice), ("ask", "Cancelled."))
        c.retry()
        self.wait_for(c)
        self.assertEqual((c.step, c.banner), ("review", "proven"))

    def wait_for(self, c):
        for _ in range(500):   # at most 10 seconds
            c.poll()
            if not c.working:
                return
            time.sleep(0.02)
        self.fail("the worker never finished")

    def test_a_change_after_review(self):
        c = self.ask(self.controller(block(GOOD), "Done: I kept the cancelled rows.\n" + block(GOOD)))
        c.set_assumptions(True)
        c.change("Please also keep the cancelled rows on a sheet called Cancelled.")
        self.assertFalse(c.assumptions_ok)   # a new recipe needs a new look at its assumptions
        c.poll()
        self.assertEqual((c.step, c.banner, c.calls), ("review", "proven", 2))
        self.assertEqual([m["role"] for m in self.ai.seen[1]], ["user", "assistant", "user"])

    def test_save_with_the_real_sheet_names(self):
        c = self.ask(self.controller(block(GOOD)))
        self.assertFalse(c.can_save())
        c.set_assumptions(True)
        c.go_save()
        self.assertEqual(c.sheet_fields(), [("ws", "Worksheet", "Orders")])
        self.assertEqual(c.check_value("Worksheet", ""), "Type the real sheet name.")
        self.assertTrue(c.check_value("Worksheet", "Orders/2026").startswith("It must be a sheet name Excel allows"))
        self.assertTrue(c.check_value("Long", "12x").startswith("It must be a whole number"))
        self.assertEqual(c.check_value("Worksheet", "Order Book"), "")
        with self.assertRaisesRegex(ValueError, "ws: It must be a sheet name"):
            c.save({"ws": "Orders/2026"})
        self.assertEqual(c.default_save_path(), c.data / "macros" / "modKronosRemoveCancelledAndSort.bas")
        path = c.save({"ws": "Order Book"})
        self.assertEqual(path, c.default_save_path())
        self.assertIn('ThisWorkbook.Worksheets("Order Book")', path.read_text(encoding="ascii"))
        dest = Path(tempfile.mkdtemp()) / "chosen.bas"
        self.assertEqual(c.save({"ws": "Orders"}, dest), dest)
        self.assertTrue(dest.exists())
        self.assertEqual(c.import_steps()[2][:43], "3. Click File > Import File..., choose chos")
        c.back()
        self.assertEqual(c.step, "review")

    def test_a_build_names_its_own_sheets(self):
        c = self.ask(self.controller(block(BUILD)))
        self.assertEqual(c.review()["summary"]["changes"],
                         ['the Status column of the sheet called "ncr" in the recipe (NCR in the practice tests)',
                          'anything on the sheet called "out" in the recipe (Summary in the practice tests)'])
        c.set_assumptions(True)
        c.go_save()
        self.assertIsNone(c.sheet_fields())
        c.save({})
        self.assertIn("choose Run and click Run", c.import_steps()[3])

    def test_the_recipe_is_hidden_unless_turned_on(self):
        c = self.ask(self.controller(block(GOOD), block(GOOD)))
        with self.assertRaisesRegex(ValueError, "Show the recipe"):
            c.recheck(BROKEN_TEST)
        c.set_setting("show_recipe", True)
        self.assertEqual(c.recipe(), GOOD)
        c.recheck(BROKEN_TEST)   # the person's edit, proven by hand
        c.poll()
        self.assertEqual((c.step, c.banner), ("review", "not_proven"))
        self.assertTrue(any(line == 29 and text.endswith("[R051]") for line, text in c.problems()))
        c.fix(BROKEN_TEST)   # Ask Codex to fix: the edit goes back with its problems
        c.poll()
        self.assertEqual((c.banner, c.calls), ("proven", 2))
        self.assertIn("I edited your spec by hand", self.ai.seen[1][-1]["text"])
        self.assertEqual(c.tidy(GOOD), GOOD)
        with self.assertRaises(ValueError):
            c.tidy(GOOD.replace("OP       LET(removed, 0)\n", "OP       FROB(removed, 0)\n"))

    def test_history_and_use_again(self):
        c = self.ask(self.controller(block(GOOD)))
        (row,) = c.history()
        self.assertEqual((row["request"], row["status"]), (REQUEST, "proven"))
        self.assertIn("Status: PROVEN", c.history_report(row["name"]))
        c.new_request()
        self.assertEqual((c.step, c.run), ("ask", None))
        c.use_again(row["name"])   # no AI call: the fake has no reply left
        c.poll()
        self.assertEqual((c.step, c.banner, c.calls, len(c.history())), ("review", "proven", 0, 2))
        with self.assertRaises(ValueError):
            c.history_report("../settings.json")

    def test_back_and_new_request(self):
        c = self.ask(self.controller(QUESTION))
        c.back()
        self.assertEqual(c.step, "ask")
        self.assertEqual(c.request, REQUEST)   # kept, so the person can change it
        c.new_request()
        self.assertEqual((c.step, c.request, c.asked), ("ask", "", False))

    def test_a_failure_inside_kronos_is_reported(self):
        c = self.controller(block(GOOD))
        c._start(lambda: 1 / 0, "ask")
        c.poll()
        self.assertEqual((c.working, c.step), (False, "ask"))
        self.assertEqual(c.notice, "Something went wrong inside KRONOS: ZeroDivisionError: division by zero")

    def test_the_checklist_finds_no_codex(self):
        ai = CodexAI(Path(tempfile.mkdtemp()), "P", which=lambda name: None)
        c = Controller(Path(tempfile.mkdtemp()), ai=ai, today=DAY, threaded=False)
        found = dict((item, (ok, detail)) for item, ok, detail in c.checklist())
        self.assertFalse(found["codex_found"][0])
        self.assertIn("codex login", found["codex_found"][1])
        self.assertEqual(c.ai_label, "Codex")


if __name__ == "__main__":
    unittest.main()
