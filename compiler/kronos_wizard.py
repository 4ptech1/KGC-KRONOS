"""KRONOS Offline's wizard (D89): KRONOS.exe's own engine, with the AI step done by copy and paste.

A browser can't start Codex, so the person copies KRONOS's message into an AI chat and pastes the AI's reply back.
Everything else is kronos_engine.Run exactly as KRONOS.exe uses it: the conversation, read_reply, prove, the repair
messages, MAX_TRIES and MAX_QUESTIONS, the best attempt, the saved .bas and its practice run. Only the AI is new:
PasteAI answers with the reply the person pasted, and with none it says "waiting", so the run stops where KRONOS.exe
would wait for Codex.

    wizard = Wizard(data_folder, today)
    wizard.ask(id, request)                 # a new run; it waits for the AI's first reply
    wizard.message(id, instructions=True)   # what to copy: the prompt and the whole conversation, as Codex gets it
    wizard.paste(id, reply)                 # the AI's reply: a question, or a recipe proven (or repaired up to 3 tries)
    wizard.answer(id, text)                 # the person's answer to the AI's question
    wizard.recipe(id, text, file)           # a recipe read from a file, proven with no AI
    wizard.send_problems(id)                # a recipe from a file that isn't proven goes to the AI with its problems
    wizard.save(id, values)                 # the .bas, the practice run, the import steps
    wizard.view(id)                         # what the page shows

Each call returns plain data. The page keeps the actions it made, so that after Cancel (which stops the browser's
Python) restore() rebuilds every run from them.
"""
import datetime as dt
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import kronos_engine as engine  # noqa: E402
from kronos_ai import AiError, render  # noqa: E402
from kronos_recipes import title_of  # noqa: E402
from vcl_codegen import CompileError, starter_value  # noqa: E402

WAITING = "waiting for the AI's reply"
MIN_REQUEST = 20   # characters, as KRONOS.exe asks (app/kronos_controller.py)
BOOK_HEADING = "The workbook it will run on (KRONOS read it; the data is mine):"


class WizardError(Exception):
    """Something the person can put right, in plain words."""


class PasteAI:
    """The AI of a copy-and-paste chat: it answers with the reply the person pasted, once."""
    name = "an AI chat (copy and paste)"

    def __init__(self):
        self.pending = None

    def ask(self, conversation):
        if self.pending is None:
            raise AiError(WAITING)
        text, self.pending = self.pending, None
        return text


def request_text(request, book=None):
    """The request as the AI gets it: with the chosen workbook's description below it when there is one (D89)."""
    request = request.strip()
    if book and book.strip():
        return f"{request}\n\n{BOOK_HEADING}\n{book.strip()}"
    return request


class Wizard:
    def __init__(self, data, today=None):
        self.data = data
        self.today = dt.date.fromisoformat(today) if isinstance(today, str) else today
        self.runs = {}

    # ---------- starting ----------

    def _run(self, id):
        if id not in self.runs:
            raise WizardError("KRONOS lost this request (the page was reloaded); start it again.")
        return self.runs[id]

    def ask(self, id, request, book=None, step=None):
        if len(request.strip()) < MIN_REQUEST:
            raise WizardError("Write a little more: say which sheet, which columns, and what should change.")
        run = engine.Run(request_text(request, book), PasteAI(), self.data, self.today, progress=_progress(step))
        self.runs[id] = run
        run.start()
        return self.view(id)

    def recipe(self, id, text, file="", step=None):
        if not text.strip():
            raise WizardError("The recipe is empty.")
        where = f" from {file}" if file else ""
        run = engine.Run(f"A recipe{where}: {title_of(text)}", PasteAI(), self.data, self.today, progress=_progress(step))
        self.runs[id] = run
        run.recheck(text)
        return self.view(id)

    # ---------- the conversation ----------

    def message(self, id, instructions=True):
        """What the person copies to the AI. With instructions: the Orchestrator prompt and the whole conversation,
        exactly what KRONOS.exe sends Codex. Without (the claude.ai Orchestrator, which holds the instructions and the
        conversation): only the newest message."""
        run = self._run(id)
        conversation = run.conversation()
        if not conversation or conversation[-1]["role"] != "user":
            raise WizardError("There is nothing to send to the AI right now.")
        if instructions:
            return render(engine.PROMPT.read_text(encoding="utf-8"), conversation)
        return conversation[-1]["text"].rstrip("\n") + "\n"

    def paste(self, id, reply, step=None):
        run = self._run(id)
        if not waiting(run):
            raise WizardError("KRONOS isn't waiting for an AI reply on this request.")
        if not reply.strip():
            raise WizardError("Paste the AI's whole reply first.")
        run.progress = _progress(step)
        run.ai.pending = reply
        run.retry()
        return self.view(id)

    def answer(self, id, text, step=None):
        run = self._run(id)
        if run.status != "question":
            raise WizardError("The AI hasn't asked a question.")
        if not text.strip():
            raise WizardError("Write your answer first.")
        run.progress = _progress(step)
        run.answer(text)
        return self.view(id)

    def send_problems(self, id, step=None):
        run = self._run(id)
        run.progress = _progress(step)
        try:
            run.send_problems()
        except ValueError as ex:
            raise WizardError(str(ex)) from None
        return self.view(id)

    # ---------- saving ----------

    def save(self, id, values=None):
        """The proven recipe's module (with a starter macro filled with the real sheet names), its practice run and
        the import steps, as KRONOS.exe's Save writes them (Run.save_module)."""
        run = self._run(id)
        if run.proven() is None:
            raise WizardError("Only a proven recipe can be saved.")
        needs = run.starter_args() or []
        clean = {}
        for name, typ in needs:
            text = str((values or {}).get(name, "")).strip()
            if not text:
                raise WizardError(f"Fill in {name} first.")
            try:
                starter_value(name, typ, text)
            except CompileError as ex:
                raise WizardError(str(ex)[0].upper() + str(ex)[1:] + ".") from None
            clean[name] = text
        try:
            path = run.save_module(clean)
        except ValueError as ex:
            raise WizardError(str(ex)) from None
        practice = run.practice_file
        out = {"bas_name": path.name, "bas": path.read_bytes().decode("ascii"), "macro": run.macro,
               "steps": engine.import_steps(path.name, run.macro),
               "practice_name": practice.name if practice else "",
               "practice": practice.read_bytes().decode("ascii") if practice else "",
               "practice_note": run.practice_note,
               "practice_steps": engine.practice_steps(path.name, practice.name) if practice else []}
        return out

    # ---------- what the page shows ----------

    def view(self, id):
        run = self._run(id)
        best = run.best()
        proof = best.proof if best else None
        status = "waiting" if waiting(run) else run.status
        newest = next((m for m in reversed(run.messages) if m["role"] == "user"), None)
        out = {"id": id, "status": status, "note": "" if status == "waiting" else run.note, "question": run.question,
               "request": run.request, "newest": newest["kind"] if newest else "",
               "tries": run.turn_tries, "max_tries": engine.MAX_TRIES,
               "attempts": [{"number": a.number, "source": a.source, "words": engine.plain_result(a),
                             "proven": a.proof is not None and a.proof.ok} for a in run.attempts],
               "best": None, "starter": None}
        if proof is not None:
            out["best"] = {"number": best.number, "proven": proof.ok, "kind": proof.kind, "name": proof.name,
                           "title": title_of(proof.spec), "spec": proof.spec, "summary": engine.summary(proof),
                           "tests": engine.test_rows(proof), "problems": engine.feedback(proof),
                           "findings": [{"stage": f.stage, "line": f.line, "severity": f.severity,
                                         "message": str(f.message), "code": f.code} for f in proof.findings]}
            if proof.ok:
                needs = run.starter_args()
                hints = engine.sheet_hints(proof.spec)
                out["starter"] = None if needs is None else [{"name": n, "type": t, "default": hints.get(n, "")}
                                                             for n, t in needs]
        return out

    # ---------- after Cancel ----------

    def restore(self, actions, step=None):
        """Rebuild every run from the page's actions, in order: [id, "ask", request, book] | [id, "recipe", text, file]
        | [id, "paste", reply] | [id, "answer", text] | [id, "problems"]. Returns the views of the runs."""
        self.runs = {}
        for action in actions:
            id, kind, rest = action[0], action[1], action[2:]
            if kind == "ask":
                self.ask(id, rest[0], rest[1] if len(rest) > 1 else None, step)
            elif kind == "recipe":
                self.recipe(id, rest[0], rest[1] if len(rest) > 1 else "", step)
            elif kind == "paste":
                self.paste(id, rest[0], step)
            elif kind == "answer":
                self.answer(id, rest[0], step)
            elif kind == "problems":
                self.send_problems(id, step)
        return {id: self.view(id) for id in self.runs}


def waiting(run):
    """True when the run stopped because it needs the AI's reply (PasteAI had none to give)."""
    return run.status == "ai_error" and WAITING in run.note


def _progress(step):
    step = step or (lambda name: None)
    return lambda kind, info: step(kind) if kind in ("checking", "testing", "building") else None
