"""Item 5 Phase 1 (D68 to D72): the KGC KRONOS engine, its AI adapters, the starter macro and `vclc ask`,
all with a fake AI (nothing here needs Codex, a network or Excel)."""
import contextlib
import datetime
import io
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
import kronos_engine as K  # noqa: E402
from kronos_ai import CODEX_COMMAND, CREATE_NO_WINDOW, AiError, CodexAI, FakeAI, render  # noqa: E402
from vba_lint import lint  # noqa: E402
from vcl_codegen import CompileError, assemble, compile_text, starter_args, starter_sub  # noqa: E402

DAY = datetime.date(2026, 10, 21)
EXAMPLES = ROOT / "vcl" / "examples"
GOOD = (EXAMPLES / "example1_rev3.vcl").read_text(encoding="utf-8")    # RemoveCancelledAndSortByDate(ws)
FILES = (EXAMPLES / "example3_files.vcl").read_text(encoding="utf-8")   # FileJobReports(tracker, dump, jobs)
BUILD = (EXAMPLES / "example4_build.vcl").read_text(encoding="utf-8")   # NcrReview, modMain.Run
FAILING_TEST = 'COUNTIF(Orders![Status], "Cancelled") = 0 AND ERROR_COUNT = 0'
BROKEN_TEST = GOOD.replace(FAILING_TEST, 'COUNTIF(Orders![Status], "Cancelled") = 5 AND ERROR_COUNT = 0')
BROKEN_CHECK = GOOD.replace("OP       LET(removed, 0)\n", "OP       FROB(removed, 0)\n")
QUESTION = "```vcl\n# QUESTION: Which sheet holds the orders?\n```"
REQUEST = "In the Orders sheet, delete every cancelled order, then sort the rest by date, newest first."


def block(spec, before=""):
    return f"{before}\n```vcl\n{spec}```\n" if before else f"```vcl\n{spec}```\n"


class ReadReply(unittest.TestCase):
    def test_one_vcl_block_with_sentences_before_it(self):   # rule 13 allows two sentences before a reissue
        reply = K.read_reply(block(GOOD, "I fixed the date test. Here is the spec."))
        self.assertEqual((reply.kind, reply.text), ("spec", GOOD))
        self.assertEqual(reply.outside, "I fixed the date test. Here is the spec.")

    def test_the_tag_ignores_case_and_windows_line_endings(self):
        reply = K.read_reply(("```VCL\n" + GOOD + "```\n").replace("\n", "\r\n"))
        self.assertEqual((reply.kind, reply.text), ("spec", GOOD))

    def test_a_reply_without_a_block_that_starts_with_vcl(self):
        self.assertEqual(K.read_reply(GOOD), K.Reply("spec", GOOD, ""))

    def test_a_build_keeps_its_separator_lines(self):
        reply = K.read_reply(block(BUILD))
        self.assertEqual((reply.kind, reply.text), ("spec", BUILD))
        self.assertEqual(reply.text.count("\n---\n"), 4)

    def test_other_blocks_are_text_outside(self):
        reply = K.read_reply("```text\nnotes\n```\n" + block(GOOD))
        self.assertEqual((reply.kind, reply.text, reply.outside), ("spec", GOOD, "```text\nnotes\n```"))

    def test_a_question(self):
        self.assertEqual(K.read_reply(QUESTION), K.Reply("question", "Which sheet holds the orders?", ""))
        self.assertEqual(K.read_reply("# QUESTION: Which folder?").text, "Which folder?")

    def test_replies_that_cant_be_used(self):
        for text, code in [("Sure! I can help with that.", "K003"), (block(GOOD) + block(GOOD), "K001"),
                           ("```vcl\n\n```", "K002")]:
            reply = K.read_reply(text)
            self.assertEqual((reply.kind, reply.problem.code), ("none", code), text)


class Prove(unittest.TestCase):
    def test_a_good_task(self):
        proof = K.prove(GOOD, DAY)
        self.assertTrue(proof.ok)
        self.assertEqual((proof.kind, proof.name, len(proof.tests)), ("TASK", "RemoveCancelledAndSortByDate", 3))
        self.assertEqual(proof.canonical, GOOD)
        self.assertEqual(lint(proof.module), [])
        self.assertTrue(proof.module.startswith("' KGC KRONOS module: RemoveCancelledAndSortByDate\n"))

    def test_a_check_error_stops_before_the_tests(self):
        proof = K.prove(BROKEN_CHECK, DAY)
        self.assertFalse(proof.ok)
        self.assertEqual(proof.tests, [])
        self.assertTrue(proof.errors and all(f.stage == "check" and f.code for f in proof.errors))
        self.assertTrue(K.feedback(proof)[0].startswith("spec.vcl:19: error: "))

    def test_a_failing_test_is_fed_back_with_its_code(self):
        proof = K.prove(BROKEN_TEST, DAY)
        self.assertEqual((proof.ok, len(proof.failing), proof.module), (False, 1, ""))
        lines = K.feedback(proof)
        self.assertIn("FAIL test   spec.vcl  (2/3 tests)", lines)
        self.assertTrue(any(line.strip().startswith("ASSERT: ") and line.endswith("[R051]") for line in lines))

    def test_a_question_is_not_proven(self):
        proof = K.prove("# QUESTION: Which sheet?\n", DAY)
        self.assertEqual([f.code for f in proof.errors], ["K004"])

    def test_a_build(self):
        proof = K.prove(BUILD, DAY)
        self.assertTrue(proof.ok)
        self.assertEqual((proof.kind, proof.name, len(proof.procedures)), ("BUILD", "NcrReview", 4))

    def test_test_strength(self):
        self.assertEqual(K.prove(GOOD, DAY).strength, None)
        caught, total = K.prove(GOOD, DAY, mutation=True).strength
        self.assertTrue(0 < caught <= total)


class Runs(unittest.TestCase):
    def setUp(self):
        self.data = Path(tempfile.mkdtemp())

    def run_with(self, *replies):
        ai = FakeAI(replies)
        return K.Run(REQUEST, ai, self.data, DAY), ai

    def test_a_spec_first_time(self):
        run, ai = self.run_with(block(GOOD, "Here it is."))
        self.assertEqual(run.start(), "proven")
        self.assertEqual(ai.seen, [[{"role": "user", "text": REQUEST}]])
        self.assertRegex(run.name, r"^\d{4}-\d\d-\d\d_\d{6}_in-the-orders-sheet$")
        self.assertEqual(sorted(p.name for p in run.folder.iterdir()),
                         ["attempt_1.vcl", "conversation.jsonl", "report.json", "report.txt", "request.txt"])
        report = json.loads((run.folder / "report.json").read_text(encoding="utf-8"))
        self.assertEqual((report["status"], report["best"], report["attempts"][0]["proven"]), ("proven", 1, True))
        lines = (run.folder / "conversation.jsonl").read_text(encoding="utf-8").splitlines()
        self.assertEqual([json.loads(line)["kind"] for line in lines], ["request", "reply"])

    def test_a_question_and_the_answer(self):
        run, ai = self.run_with(QUESTION, block(GOOD))
        self.assertEqual(run.start(), "question")
        self.assertEqual(run.question, "Which sheet holds the orders?")
        self.assertIn("The AI asks: Which sheet holds the orders?", run.report_lines())
        self.assertEqual(run.answer("The sheet is called Orders."), "proven")
        self.assertEqual([m["role"] for m in ai.seen[1]], ["user", "assistant", "user"])
        self.assertEqual(ai.seen[1][-1]["text"], "The sheet is called Orders.")
        self.assertEqual(run.ai_tries, 1)   # a question is not a try

    def test_a_repair(self):
        run, ai = self.run_with(block(BROKEN_TEST), block(GOOD, "The test counted wrong; fixed."))
        self.assertEqual(run.start(), "proven")
        repair = ai.seen[1][-1]["text"]
        self.assertTrue(repair.startswith("The VBA Forge toolchain checked your spec and it did not pass."))
        self.assertIn("[R051]", repair)
        self.assertEqual((run.best().number, len(run.attempts)), (2, 2))

    def test_a_reply_without_a_spec_is_a_try(self):
        run, ai = self.run_with("I would loop over the rows.", block(GOOD))
        self.assertEqual(run.start(), "proven")
        self.assertIn("[K003]", ai.seen[1][-1]["text"])
        self.assertFalse((run.folder / "attempt_1.vcl").exists())

    def test_give_up_after_three_tries(self):
        run, ai = self.run_with(block(BROKEN_CHECK), block(BROKEN_TEST), block(BROKEN_CHECK), block(GOOD))
        self.assertEqual(run.start(), "not_proven")
        self.assertEqual(len(ai.seen), 3)   # the first spec and 2 more tries (D68)
        self.assertEqual(run.best().number, 2)   # one failing TEST beats a spec that doesn't check
        self.assertEqual(run.note, "no spec passed after 3 tries: attempt 2 came closest; edit it and Re-check")
        self.assertEqual(run.recheck(GOOD), "proven")   # the person edits the spec and presses Re-check
        self.assertEqual((run.best().number, run.best().source), (4, "person"))

    def test_the_ai_gives_no_reply_and_a_retry(self):
        run, ai = self.run_with(AiError("Codex gave no reply within 300 seconds"), block(GOOD))
        self.assertEqual(run.start(), "ai_error")
        self.assertEqual(run.note, "the AI gave no reply: Codex gave no reply within 300 seconds")
        self.assertEqual(run.retry(), "proven")
        self.assertEqual(ai.seen[1], [{"role": "user", "text": REQUEST}])

    def test_too_many_questions(self):
        run, ai = self.run_with(*[QUESTION] * 4)
        status = run.start()
        for n in range(3):
            self.assertEqual(status, "question")
            status = run.answer(f"answer {n}")
        self.assertEqual(status, "not_proven")
        self.assertIn("more than 3 questions", run.note)

    def test_wrong_order(self):
        run, ai = self.run_with(block(GOOD))
        with self.assertRaises(ValueError):
            run.answer("no question was asked")
        with self.assertRaises(ValueError):
            run.save_module({"ws": "Orders"})
        run.start()
        with self.assertRaises(ValueError):
            run.start()

    def test_the_data_folder_is_never_in_the_repo(self):   # D69
        with self.assertRaises(ValueError):
            K.Run(REQUEST, FakeAI([]), ROOT / "logs", DAY)

    def test_two_runs_in_the_same_second(self):
        now = datetime.datetime(2026, 10, 31, 9, 15, 0)
        first = K.Run(REQUEST, FakeAI([]), self.data, DAY, now=now)
        second = K.Run(REQUEST, FakeAI([]), self.data, DAY, now=now)
        self.assertEqual(second.name, first.name + "_2")


class Module(unittest.TestCase):
    def setUp(self):
        self.data = Path(tempfile.mkdtemp())

    def proven(self, spec):
        run = K.Run(REQUEST, FakeAI([block(spec)]), self.data, DAY)
        self.assertEqual(run.start(), "proven")
        return run

    def test_a_task_with_its_starter_macro(self):
        run = self.proven(GOOD)
        self.assertEqual(run.starter_args(), [("ws", "Worksheet")])
        path = run.save_module({"ws": "Orders"})
        self.assertEqual(path.name, "modKronosRemoveCancelledAndSort.bas")   # 31 characters at most
        raw = path.read_bytes()
        raw.decode("ascii")
        self.assertNotIn(b"\n", raw.replace(b"\r\n", b""))   # Windows line endings only
        first, module = raw.decode("ascii").split("\r\n", 1)
        self.assertEqual(first, 'Attribute VB_Name = "modKronosRemoveCancelledAndSort"')
        module = module.replace("\r\n", "\n")
        self.assertEqual(lint(module), [])
        self.assertIn("Public Sub RunRemoveCancelledAndSortByDate()", module)
        self.assertIn('RemoveCancelledAndSortByDate(ThisWorkbook.Worksheets("Orders"))', module)
        report = (run.folder / "report.txt").read_text(encoding="utf-8")
        self.assertIn("4. Close the VBA editor. Press Alt+F8, choose RunRemoveCancelledAndSortByDate and click Run.",
                      report)

    def test_a_missing_value(self):
        run = self.proven(GOOD)
        with self.assertRaisesRegex(ValueError, r"needs a value for \$ws \(Worksheet\)"):
            run.save_module({})
        path = run.save_module({}, starter=False)
        self.assertNotIn("Public Sub RunRemove", path.read_text(encoding="ascii"))
        self.assertIn("No starter macro", (run.folder / "report.txt").read_text(encoding="utf-8"))

    def test_a_build_starts_with_its_own_run(self):
        run = self.proven(BUILD)
        self.assertIsNone(run.starter_args())
        path = run.save_module()
        self.assertEqual(path.name, "modKronosNcrReview.bas")
        self.assertEqual(run.macro, "Run")
        self.assertEqual(lint(path.read_text(encoding="ascii").split("\n", 1)[1].replace("\r\n", "\n")), [])


class Starter(unittest.TestCase):   # the compiler's starter macro (D68, D72)
    def module(self, spec, values, body=GOOD):
        return assemble([compile_text(body), starter_sub(spec, values)], "test")

    def test_sheets_and_text_values(self):
        values = {"tracker": "Job Tracker", "dump": 'C:\\Dump "new"', "jobs": "C:\\Jobs"}
        module = assemble([compile_text(FILES), starter_sub(FILES, values)], "test")
        self.assertEqual(lint(module), [])
        self.assertIn('FileJobReports(ThisWorkbook.Worksheets("Job Tracker"), "C:\\Dump ""new""", "C:\\Jobs")', module)

    def test_a_sub_with_arguments(self):
        sub = GOOD.replace("KIND     Function", "KIND     Sub").replace("RET      Long\n", "")
        text = starter_sub(sub, {"ws": "Orders"})
        self.assertIn('Call RemoveCancelledAndSortByDate(ThisWorkbook.Worksheets("Orders"))', text)
        self.assertEqual(lint(assemble([compile_text(GOOD), text], "test")), [])

    def test_no_starter_for_a_sub_without_arguments(self):
        sub = GOOD.replace("KIND     Function", "KIND     Sub").replace("RET      Long\n", "") \
            .replace("ARGS     ws As Worksheet\n", "")
        self.assertIsNone(starter_args(sub))
        self.assertIsNone(starter_sub(sub, {}))

    def test_values_by_type(self):
        spec = GOOD.replace("ARGS     ws As Worksheet",
                            "ARGS     ws As Worksheet, n As Long, x As Double, ok As Boolean")
        text = starter_sub(spec, {"ws": "Orders", "n": "-5", "x": "1.5", "ok": "yes"})
        self.assertIn('ThisWorkbook.Worksheets("Orders"), -5, 1.5, True)', text)

    def test_values_that_are_refused(self):
        spec = GOOD.replace("ARGS     ws As Worksheet", "ARGS     ws As Worksheet, n As Long")
        for values, words in [({"ws": "Orders/2026", "n": "1"}, "a sheet name Excel allows"),
                              ({"ws": "Bestellungen \u00fcber", "n": "1"}, "plain ASCII text"),
                              ({"ws": "Orders", "n": "1\u0662"}, "a whole number"),
                              ({"ws": "Orders", "n": "99999999999"}, "a whole number"),
                              ({"ws": "Orders"}, r"needs a value for \$n")]:
            with self.assertRaisesRegex(CompileError, words, msg=values):
                starter_sub(spec, values)

    def test_types_the_starter_cant_pass(self):
        with self.assertRaisesRegex(CompileError, r"can't pass \$ws \(Range\)"):
            starter_sub(GOOD.replace("ws As Worksheet", "ws As Range"), {"ws": "A1"})
        with self.assertRaisesRegex(CompileError, "can't show a Range result"):
            starter_sub(GOOD.replace("RET      Long", "RET      Range"), {"ws": "Orders"})
        with self.assertRaisesRegex(CompileError, "BUILD starts with its modMain.Run"):
            starter_sub(BUILD, {})


class Recorder:
    """Stands in for subprocess.run: keeps each call and writes the reply where Codex would."""

    def __init__(self, reply="```vcl\n# QUESTION: Which sheet?\n```", returncode=0, stderr=b"", stdout=b"",
                 raises=None):
        self.reply, self.returncode, self.stderr, self.stdout, self.raises = reply, returncode, stderr, stdout, raises
        self.calls = []

    def __call__(self, args, **kwargs):
        self.calls.append((args, kwargs))
        if self.raises:
            raise self.raises
        if "--output-last-message" in args and self.reply is not None:
            Path(args[args.index("--output-last-message") + 1]).write_text(self.reply, encoding="utf-8")
        return subprocess.CompletedProcess(args, self.returncode, self.stdout, self.stderr)


class Codex(unittest.TestCase):
    CONVERSATION = [{"role": "user", "text": REQUEST}]

    def codex(self, runner, windows=False, which=lambda name: "/opt/codex/bin/codex", **kwargs):
        self.work = Path(tempfile.mkdtemp()) / "codex"
        return CodexAI(self.work, "THE PROMPT", runner=runner, which=which, windows=windows, **kwargs)

    def test_a_call(self):
        runner = Recorder()
        self.assertEqual(self.codex(runner).ask(self.CONVERSATION), "```vcl\n# QUESTION: Which sheet?\n```")
        (args, kwargs), = runner.calls
        self.assertEqual(args, ["/opt/codex/bin/codex", "exec", "--sandbox", "read-only", "--skip-git-repo-check",
                                "--output-last-message", str(self.work / "last_message.txt"), "-"])
        self.assertEqual((kwargs["cwd"], kwargs["timeout"], kwargs["capture_output"]), (str(self.work), 300, True))
        self.assertNotIn("creationflags", kwargs)
        sent = kwargs["input"].decode("utf-8")   # the whole prompt goes on stdin, never on the command line (D72)
        self.assertTrue(sent.startswith("THE PROMPT\n\nCONVERSATION\n"))
        self.assertTrue(sent.endswith(f"===== USER =====\n{REQUEST}\n"))
        self.assertIn("Run no commands", (self.work / "AGENTS.md").read_text(encoding="utf-8"))

    def test_windows_opens_no_console(self):
        runner = Recorder()
        self.codex(runner, windows=True).ask(self.CONVERSATION)
        self.assertEqual(runner.calls[0][1]["creationflags"], CREATE_NO_WINDOW)

    def test_failures(self):
        cases = [(Recorder(), lambda name: None, "was not found: install it, run `codex login`"),
                 (Recorder(raises=subprocess.TimeoutExpired("codex", 300)), None, "no reply within 300 seconds"),
                 (Recorder(raises=PermissionError("denied")), None, "could not start"),
                 (Recorder(returncode=1, stderr=b"progress\nerror: not logged in\n"), None,
                  "exit code 1: progress | error: not logged in"),
                 (Recorder(reply=None), None, "finished without a reply"),
                 (Recorder(reply="  \n"), None, "finished without a reply")]
        for runner, which, words in cases:
            codex = self.codex(runner, **({"which": which} if which else {}))
            (self.work).mkdir(parents=True, exist_ok=True)
            (self.work / "last_message.txt").write_text("an old reply", encoding="utf-8")   # never read again
            with self.assertRaisesRegex(AiError, words.replace("|", r"\|")):
                codex.ask(self.CONVERSATION)

    def test_a_command_line_without_out(self):
        with self.assertRaisesRegex(AiError, "has no {out}"):
            self.codex(Recorder(), command=["codex", "exec", "-"]).ask(self.CONVERSATION)

    def test_check_the_flags(self):
        full = " ".join(CODEX_COMMAND).encode()
        self.assertEqual(self.codex(Recorder(stdout=full)).check(), [])
        found = self.codex(Recorder(stdout=b"--sandbox --output-last-message")).check()
        self.assertEqual(found, ["`codex exec --help` doesn't list --skip-git-repo-check: update the Codex command "
                                 "line in settings.json"])
        self.assertIn("was not found", self.codex(Recorder(), which=lambda name: None).check()[0])

    def test_the_conversation(self):
        text = render("P", [{"role": "user", "text": "a"}, {"role": "assistant", "text": "b"},
                            {"role": "user", "text": "c"}])
        self.assertEqual(text.split("\n")[-9:], ["===== USER =====", "a", "", "===== YOUR EARLIER REPLY =====", "b",
                                                 "", "===== USER =====", "c", ""])


class Settings(unittest.TestCase):
    def test_defaults_and_overrides(self):
        data = Path(tempfile.mkdtemp())
        self.assertEqual(K.load_settings(data), {"ai": "codex", "codex_command": CODEX_COMMAND, "codex_timeout": 300})
        (data / "settings.json").write_text('{"codex_timeout": 600}', encoding="utf-8")
        self.assertEqual(K.load_settings(data)["codex_timeout"], 600)
        self.assertIsInstance(K.make_ai(K.load_settings(data), data), CodexAI)
        for text, words in [('{"model": "x"}', "unknown settings model"), ('{"ai": "gpt"}', "codex or fake"),
                            ('{"codex_command": "codex exec"}', "a list of strings")]:
            (data / "settings.json").write_text(text, encoding="utf-8")
            with self.assertRaisesRegex(ValueError, words):
                K.load_settings(data)
        with self.assertRaisesRegex(ValueError, "needs its replies"):
            K.make_ai({"ai": "fake"}, data)


class Ask(unittest.TestCase):   # python compiler/vclc.py ask
    def ask(self, *args, replies=()):
        import vclc
        data = Path(tempfile.mkdtemp())
        path = data / "replies.json"
        path.write_text(json.dumps(list(replies)), encoding="utf-8")
        out = io.StringIO()
        with contextlib.redirect_stdout(out), mock.patch.object(K, "ask_person", return_value=None):
            code = vclc.main(["vclc.py", "ask", REQUEST, "--ai", "fake", "--replies", str(path), "--data",
                              str(data / "KGC KRONOS"), "--today", "2026-10-21", *args])
        runs = list((data / "KGC KRONOS" / "runs").iterdir()) if (data / "KGC KRONOS" / "runs").exists() else []
        return code, out.getvalue(), runs

    def test_a_proven_module(self):
        code, out, (run,) = self.ask("--arg", "ws=Orders", replies=[block(GOOD)])
        self.assertEqual(code, 0)
        self.assertIn("Status: PROVEN: the spec passed every check and every TEST", out)
        self.assertTrue((run / "modKronosRemoveCancelledAndSort.bas").exists())
        self.assertIn(f"Run folder: {run}", out)

    def test_a_question_answered_on_the_command_line(self):
        code, out, _ = self.ask("--answer", "Orders", "--arg", "$ws=Orders", replies=[QUESTION, block(GOOD)])
        self.assertEqual(code, 0)
        self.assertIn("The AI asks: Which sheet holds the orders?", out)

    def test_an_unanswered_question(self):
        code, out, _ = self.ask(replies=[QUESTION])
        self.assertEqual(code, 1)
        self.assertIn('No answer given: run it again with --answer "..."', out)

    def test_no_starter_without_the_sheet_names(self):
        code, out, _ = self.ask(replies=[block(GOOD)])
        self.assertEqual(code, 0)
        self.assertIn("No starter macro", out)

    def test_not_proven(self):
        code, out, _ = self.ask(replies=[block(BROKEN_TEST)] * 3)
        self.assertEqual(code, 1)
        self.assertIn("Status: NOT PROVEN (no spec passed after 3 tries", out)

    def test_bad_options(self):
        self.assertEqual(self.ask("--arg", "ws")[0], 1)
        import vclc
        with contextlib.redirect_stdout(io.StringIO()) as out:
            self.assertEqual(vclc.main(["vclc.py", "ask", REQUEST, "--data", str(ROOT / "logs")]), 1)
        self.assertIn("inside the repo", out.getvalue())


class Packaging(unittest.TestCase):
    def test_the_engine_needs_nothing_from_tests(self):   # KRONOS.exe won't have tests/ (docs/kronos-app.md)
        code = ("import sys; sys.path.insert(0, 'compiler'); import kronos_engine; "
                "print(sorted(m for m in sys.modules\n"
                "             if m.startswith(('test', 'conformance', 'vcl_selftest', 'vclc'))))")
        done = subprocess.run([sys.executable, "-c", code], cwd=ROOT, capture_output=True, text=True, check=True)
        self.assertEqual(done.stdout.strip(), "[]")


if __name__ == "__main__":
    unittest.main()
