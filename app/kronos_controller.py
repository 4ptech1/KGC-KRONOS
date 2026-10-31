"""The KGC KRONOS window's controller (item 5 Phase 2a, docs/kronos-app.md section 2, D74, D75).

The window (app/kronos_app.py) holds no logic: it shows what this controller says and calls its
methods. The controller keeps the wizard's state, 1 Ask, 2 Questions, 3 Review and 4 Save, and drives
the engine (compiler/kronos_engine.py). Engine work, asking the AI and proving a recipe, runs on a
worker thread. Its steps and its end come back as events on a queue, and the window reads them with
poll() every 100 ms, because Tk may be touched only from its own thread. With threaded=False the work
runs at once, for the tests. Nothing here needs a display.

    step      "ask", "question", "review" or "save" (the working panel shows over it while working)
    banner    on Review: "proven" or "not_proven"
    notice    a line for the current step: why the AI gave no reply, or that the work was cancelled
"""
import queue
import re
import sys
import threading
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
import kronos_engine as engine  # noqa: E402
from kronos_ai import AiError, CodexAI  # noqa: E402
from vcl_codegen import CompileError, starter_value  # noqa: E402
from vcl_fmt import FormatError, format_text  # noqa: E402

MIN_REQUEST = 20   # characters: shorter than this is not a request Codex can work from
STEPS = ("ask", "question", "review", "save")
STAGES = {"checking": "Checking the recipe", "testing": "Testing it on practice data",
          "building": "Building the macro", "measuring": "Measuring how strong the tests are"}


class Controller:
    def __init__(self, data, ai=None, today=None, threaded=True):
        self.data = engine.data_folder(data)
        self.settings = engine.load_settings(self.data)
        self.ai = ai if ai is not None else engine.make_ai(self.settings, self.data)
        self.today, self.threaded = today, threaded
        self.events = queue.Queue()
        self.run = None
        self.step, self.working, self.origin = "ask", False, "ask"
        self.notice = ""
        self.asked = False            # Codex asked a question in this run (the step bar shows Questions)
        self.assumptions_ok = False   # the person ticked "I checked these assumptions"
        self.saved = None             # the path of the module the person saved
        self.lines = []               # the working panel: one line per try, in plain words
        self.stage = ""               # the working panel: the step now running
        self.started_at = 0.0

    # ---------- the first-run checklist ----------

    @property
    def checklist_done(self):
        return self.settings["data_controls_confirmed"] and self.settings["practice_confirmed"]

    def checklist(self):
        """[(item, ok, detail)]: what the first-run dialog shows."""
        items = []
        if isinstance(self.ai, CodexAI):
            try:
                found, problems = self.ai.program(), []
                problems = self.ai.check()
                items += [("codex_found", True, found), ("codex_flags", not problems, "; ".join(problems))]
            except AiError as exc:
                items += [("codex_found", False, str(exc)), ("codex_flags", False, "")]
        else:
            items += [("codex_found", True, "the practice AI"), ("codex_flags", True, "")]
        return items + [("data_controls", self.settings["data_controls_confirmed"], ""),
                        ("data_folder", True, str(self.data)),
                        ("practice", self.settings["practice_confirmed"], "")]

    def confirm_checklist(self, data_controls, practice):
        self.settings["data_controls_confirmed"] = bool(data_controls)
        self.settings["practice_confirmed"] = bool(practice)
        engine.save_settings(self.data, self.settings)

    def set_setting(self, key, value):
        changed = dict(self.settings, **{key: value})
        engine.check_settings(changed)
        self.settings = changed
        engine.save_settings(self.data, self.settings)

    # ---------- the steps ----------

    def can_ask(self, text):
        return bool(self.checklist_done and not self.working and len(text.strip()) >= MIN_REQUEST)

    def ask(self, text):
        """Step 1: a new request, so a new run."""
        if not self.can_ask(text):
            raise ValueError(f"a request needs the first-run checklist and at least {MIN_REQUEST} characters")
        self.run = engine.Run(text, self.ai, self.data, self.today, self.settings["measure_strength"],
                              progress=self._progress)
        self.asked, self.assumptions_ok, self.saved = False, False, None
        self._start(self.run.start, "ask")

    def answer(self, text):
        """Step 2: the answer to Codex's question."""
        self._need("question")
        self._start(lambda: self.run.answer(text), "question")

    def change(self, text):
        """Review's "Ask for a change": a new turn with its own tries."""
        self._need("review")
        self.assumptions_ok = False
        self._start(lambda: self.run.follow_up(text), "review")

    def retry(self):
        """Try again after no reply or Cancel, from the step the work started on."""
        if self.run is None or self.run.status not in ("ai_error", "cancelled") or self.working:
            raise ValueError("there is nothing to try again")
        self._start(self.run.retry, self.step)

    def cancel(self):
        if self.working and self.run is not None:
            self.run.cancel()

    def _need(self, step):
        if self.working or self.step != step or self.run is None:
            raise ValueError(f"not on the {step} step")

    def _start(self, work, origin):
        if self.run is not None:
            self.run.ready()   # here, on the window's thread, so a Cancel pressed from now on is never lost
        self.working, self.origin, self.notice = True, origin, ""
        self.lines, self.stage, self.started_at = [], "", time.monotonic()
        if not self.threaded:
            self._work(work)
            return
        threading.Thread(target=self._work, args=(work,), daemon=True).start()

    def _work(self, work):
        try:
            work()
            self.events.put({"type": "done"})
        except Exception as exc:   # the window must hear about any failure, or it would wait for ever
            self.events.put({"type": "failed", "message": f"{type(exc).__name__}: {exc}"})

    def _progress(self, kind, info):
        self.events.put({"type": "progress", "kind": kind, **info})

    def poll(self):
        """Read the worker's events (from the window's thread); returns them, after updating the state."""
        out = []
        while True:
            try:
                event = self.events.get_nowait()
            except queue.Empty:
                return out
            out.append(event)
            if event["type"] == "progress":
                self._heard(event)
            elif event["type"] == "done":
                self.working = False
                self._settle()
            else:
                self.working, self.step = False, self.origin
                self.notice = f"Something went wrong inside KRONOS: {event['message']}"

    def _heard(self, event):
        kind, name = event["kind"], self.ai_label
        if kind == "asking":
            self.stage = f"{name} is writing the recipe (try {event['try']} of {event['of']})"
        elif kind == "tried":
            then = "" if event["proven"] else f", so KRONOS sent the problems back to {name}"
            self.lines.append(f"Try {event['try']}: {event['result']}{then}")
        else:
            self.stage = STAGES.get(kind, kind)

    def _settle(self):
        status = self.run.status
        if status == "question":
            self.step, self.asked = "question", True
        elif status in ("proven", "not_proven"):
            self.step = "review"
            if self.lines and status == "not_proven" and self.lines[-1].endswith(f"back to {self.ai_label}"):
                self.lines[-1] = self.lines[-1].split(", so KRONOS")[0]   # the last try was not sent back
        else:   # ai_error or cancelled: back where the work started, with Try again
            self.step = self.origin
            self.notice = "Cancelled." if status == "cancelled" else self.run.note

    @property
    def ai_label(self):
        return "Codex" if getattr(self.ai, "name", "") == "codex" else "The practice AI"

    @property
    def calls(self):
        return self.run.calls if self.run else 0

    def elapsed(self):
        return int(time.monotonic() - self.started_at) if self.working else 0

    def back(self):
        if self.working:
            return
        self.step = {"question": "ask", "review": "ask", "save": "review"}.get(self.step, self.step)

    def new_request(self):
        if self.working:
            raise ValueError("wait for the work to finish, or cancel it")
        self.run, self.step, self.notice, self.asked = None, "ask", "", False
        self.assumptions_ok, self.saved, self.lines = False, None, []

    @property
    def request(self):
        return self.run.request if self.run else ""

    @property
    def question(self):
        return self.run.question if self.run else ""

    def step_bar(self):
        """[(step, state)]: state is done, current, todo or skipped (Questions when Codex asked none)."""
        here = STEPS.index(self.step)
        out = []
        for i, step in enumerate(STEPS):
            if i == here:
                state = "current"
            elif step == "question" and not self.asked and here > i:
                state = "skipped"
            else:
                state = "done" if i < here else "todo"
            out.append((step, state))
        return out

    def conversation(self):
        return self.run.conversation() if self.run else []

    # ---------- step 3, Review ----------

    @property
    def banner(self):
        return self.run.status if self.step == "review" and self.run else ""

    def shown(self):
        """The attempt Review and Save show: the latest turn's best."""
        return self.run.best() if self.run else None

    def review(self):
        """What Review shows: the summary in words, the tests, and the result of each try."""
        attempt = self.shown()
        proof = attempt.proof if attempt else None
        words = engine.summary(proof) if proof and proof.kind else {"does": "", "steps": [], "assumes": [],
                                                                    "changes": [], "warnings": []}
        return {"banner": self.banner, "note": self.run.note if self.run else "", "summary": words,
                "tests": test_rows(proof) if proof else [], "tries": list(self.lines),
                "strength": proof.strength if proof else None}

    def set_assumptions(self, ok):
        self.assumptions_ok = bool(ok)

    def can_save(self):
        return self.step == "review" and not self.working and self.banner == "proven" and self.assumptions_ok

    def go_save(self):
        if not self.can_save():
            raise ValueError("only a proven recipe, with its assumptions ticked, goes on to Save")
        self.step = "save"

    # ---------- step 4, Save ----------

    def sheet_fields(self):
        """[(argument, type, hint)] for "Your real sheet names"; None for a BUILD, which names its own."""
        attempt = self.shown()
        needs = self.run.starter_args()
        if needs is None:
            return None
        hints = engine.sheet_hints(attempt.proof.spec)
        return [(name, typ, hints.get(name, "")) for name, typ in needs]

    @staticmethod
    def check_value(typ, value):
        """"" when the value is fine, else what is wrong, in words."""
        if not value.strip():
            return "Type the real sheet name." if typ == "Worksheet" else "Type a value."
        try:
            starter_value("x", typ, value)
        except CompileError as exc:
            return re.sub(r"^the value for \$x \(\w+\) ", "It ", str(exc))
        return ""

    def default_save_path(self):
        attempt = self.shown()
        folder = Path(self.settings["macro_folder"]) if self.settings["macro_folder"] else self.data / "macros"
        return folder / f"{engine.module_name(attempt.proof.name)}.bas"

    def save(self, values, dest=None):
        """Write the module (with its starter macro) to dest, or to the default place; returns the path."""
        if self.step != "save" or self.working:
            raise ValueError("not on the Save step")
        fields = self.sheet_fields() or []
        wrong = [f"{name}: {problem}" for name, typ, hint in fields
                 for problem in [self.check_value(typ, values.get(name, ""))] if problem]
        if wrong:
            raise ValueError("; ".join(wrong))
        self.saved = self.run.save_module(values, dest=Path(dest) if dest else self.default_save_path())
        return self.saved

    def import_steps(self):
        attempt = self.shown()
        name = self.saved.name if self.saved else f"{engine.module_name(attempt.proof.name)}.bas"
        macro = self.run.macro or ("Run" if attempt.proof.kind == "BUILD" else "Run" + attempt.proof.name)
        return engine.import_steps(name, macro)

    # ---------- the recipe (only with Show the recipe on) ----------

    def recipe(self):
        attempt = self.shown()
        return attempt.proof.spec if attempt else ""

    def problems(self):
        """[(spec line, message)] of the recipe shown, for the Recipe window's list."""
        attempt = self.shown()
        if attempt is None:
            return []
        return [(f.line if f.stage in ("check", "test") else 0, f"{f.message}{f' [{f.code}]' if f.code else ''}")
                for f in attempt.proof.findings] + \
               [(t.line, f"{m}{f' [{c}]' if c else ''}") for t in attempt.proof.failing for m, c in t.failed]

    def _need_recipe(self):
        if not self.settings["show_recipe"]:
            raise ValueError("turn on Show the recipe in Settings first")
        if self.working or self.run is None or self.step not in ("review", "save"):
            raise ValueError("there is no recipe to work on")

    def recheck(self, text):
        self._need_recipe()
        self.assumptions_ok = False
        self._start(lambda: self.run.recheck(text), "review")

    def fix(self, text):
        self._need_recipe()
        self.assumptions_ok = False
        self._start(lambda: self.run.fix(text), "review")

    @staticmethod
    def tidy(text):
        """The recipe in the canonical layout (`vclc fmt`), or ValueError with its first problem."""
        try:
            return format_text(text)
        except FormatError as exc:
            raise ValueError(str(exc.args[0][0]) if exc.args and exc.args[0] else "the recipe has errors") from None

    # ---------- History ----------

    def history(self, limit=200):
        """[{name, request, status, when}] of past runs, newest first."""
        runs = self.data / "runs"
        rows = []
        for folder in sorted(runs.iterdir(), reverse=True)[:limit] if runs.exists() else []:
            report = engine.read_report(folder)
            if report is not None:
                when = re.match(r"(\d{4}-\d\d-\d\d)_(\d\d)(\d\d)", folder.name)
                rows.append({"name": folder.name, "request": report.get("request", ""),
                             "status": report.get("status", ""),
                             "when": f"{when[1]} {when[2]}:{when[3]}" if when else ""})
        return rows

    def _run_folder(self, name):
        folder = self.data / "runs" / name
        if Path(name).name != name or not folder.is_dir():
            raise ValueError(f"no run named {name!r}")
        return folder

    def history_report(self, name):
        return (self._run_folder(name) / "report.txt").read_text(encoding="utf-8")

    def use_again(self, name):
        """A past run's shown attempt, proven again in a new run (no AI call), then Review."""
        if self.working:
            raise ValueError("wait for the work to finish, or cancel it")
        folder = self._run_folder(name)
        report = engine.read_report(folder) or {}
        spec = folder / f"attempt_{report.get('best')}.vcl"
        if not report.get("best") or not spec.exists():
            raise ValueError("that run has no recipe to use again")
        self.run = engine.Run(report["request"], self.ai, self.data, self.today, self.settings["measure_strength"],
                              progress=self._progress)
        self.asked, self.assumptions_ok, self.saved = False, False, None
        text = spec.read_text(encoding="utf-8")
        self._start(lambda: self.run.recheck(text), "ask")


def test_rows(proof):
    """[{line, passed, checks, why}]: each TEST, what it checks (its assert= part) and why it failed."""
    lines = proof.spec.splitlines()
    rows = []
    for t in proof.tests:
        text = lines[t.line - 1] if 0 < t.line <= len(lines) else ""
        checks = text.split("assert=", 1)[1].strip() if "assert=" in text else text
        rows.append({"line": t.line, "passed": t.passed, "checks": checks,
                     "why": [m for m, c in t.failed]})
    return rows
