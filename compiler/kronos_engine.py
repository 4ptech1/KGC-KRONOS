"""The KGC KRONOS engine (item 5 Phase 1, docs/kronos-app.md, D68 to D72): from a request in plain
English to a proven VBA module, without a window.

    python compiler/vclc.py ask "<request>" [--ai codex|fake] [--replies FILE] [--data DIR]
        [--answer TEXT]... [--arg NAME=VALUE]... [--today YYYY-MM-DD] [--mutation] [--check-ai]

An AI (kronos_ai.py) writes the VCL. Everything after that is this repo's offline toolchain, called
in process (not through vclc, whose imports need tests/): the validator, the TESTs in the
reference interpreter, the compiler and the lint. A spec with errors or a failing TEST goes back
to the AI with the coded messages, as vclc prints them, for up to 2 more tries; a question goes to
the person, whose answer continues the conversation. A proven spec becomes one .bas module with a
starter macro (vcl_codegen.starter_sub), saved with Windows line endings.

Every run gets a folder in the data folder, never in the repo (D69):
  <data>/runs/<date>_<time>_<short name>/  request.txt, conversation.jsonl, attempt_N.vcl,
                                           report.txt, report.json and the .bas
--ai fake reads its replies from a JSON list of strings (--replies FILE), for tests and dry runs.
Exit code 0 when the spec is proven and the module saved, 1 otherwise.
"""
import argparse
import datetime
import json
import re
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))
from errcodes import code_of, coded, tag  # noqa: E402
from kronos_ai import CODEX_COMMAND, CODEX_TIMEOUT, AiError, CodexAI, FakeAI  # noqa: E402
from validator import MODIFIES_RE, validate_text  # noqa: E402
from vba_lint import lint  # noqa: E402
from vcl_codegen import (CompileError, assemble, compile_build_procs, compile_text, starter_args,  # noqa: E402
                         starter_sub, task_signature)
from vcl_fmt import FormatError, format_text  # noqa: E402
from vcl_interp import VclError, run_tests  # noqa: E402
from vcl_mutate import mutate_report  # noqa: E402
from vcl_parse import parse  # noqa: E402

PROMPT = ROOT / "vcl" / "orchestrator_prompt.txt"
MAX_TRIES = 3        # in each turn, the AI's first spec and 2 more tries (D68, D74)
MAX_QUESTIONS = 3    # then the person puts the details in the request instead
SPEC_NAME = "spec.vcl"   # the name the repair messages give the spec, as vclc prints a file
QUESTION = "# QUESTION:"
FENCE = re.compile(r"```[ \t]*([A-Za-z0-9_+-]*)[ \t]*")


def default_data_folder():
    return Path.home() / "Documents" / "KGC KRONOS"


def data_folder(path):
    """The data folder, made if missing. Never inside the repo, which is public (D69)."""
    folder = Path(path).expanduser().resolve()
    if folder == ROOT or ROOT in folder.parents:
        raise ValueError(f"the data folder {folder} is inside the repo: KRONOS keeps its runs outside it (D69)")
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def normalize(text):
    text = text.lstrip("﻿").replace("\r\n", "\n").replace("\r", "\n").strip("\n")
    return text + "\n" if text else ""


# ---------- reading a reply ----------

@dataclass
class Reply:
    kind: str            # "spec", "question" or "none"
    text: str = ""       # the spec, or the question
    outside: str = ""    # text outside the vcl block (rule 13 allows two sentences), kept for the log
    problem: str = ""    # why a "none" reply can't be used (coded)


def is_question(text):
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    return bool(lines) and lines[0].startswith(QUESTION)


def question_text(text):
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    first = lines[0][len(QUESTION):].strip()
    return " ".join([first] + [ln.lstrip("#").strip() for ln in lines[1:]]).strip()


def read_reply(text):
    """The one vcl code block of an AI reply; otherwise the whole reply if it starts with VCL or
    # QUESTION: (rule 1 asks for the block, rule 13 allows two sentences before it)."""
    lines = normalize(text).split("\n")
    blocks, outside, i = [], [], 0
    while i < len(lines):
        m = FENCE.fullmatch(lines[i].strip())
        if not m or lines[i].strip() == "```":
            outside.append(lines[i])
            i += 1
            continue
        end = i + 1
        while end < len(lines) and lines[end].strip() != "```":
            end += 1
        if m.group(1).lower() == "vcl":
            blocks.append("\n".join(lines[i + 1:end]))
        else:
            outside += lines[i:end + 1]
        i = end + 1
    if len(blocks) > 1:
        return Reply("none", problem=coded(f"the reply has {len(blocks)} vcl code blocks: reply with exactly one "
                                           "(rule 1)", "K001"))
    if blocks:
        body, rest = normalize(blocks[0]), "\n".join(outside).strip()
        if not body:
            return Reply("none", outside=rest, problem=coded("the vcl code block is empty: put the whole spec in it, "
                                                              "or only # QUESTION: <one short question>", "K002"))
    else:
        body, rest = normalize(text), ""
        if not (body.startswith("VCL") or body.startswith(QUESTION)):
            return Reply("none", problem=coded("the reply has no vcl code block: reply with exactly one code block "
                                               "tagged vcl (rule 1)", "K003"))
    if is_question(body):
        return Reply("question", question_text(body), rest)
    return Reply("spec", body, rest)


# ---------- proving a spec ----------

@dataclass
class Finding:
    stage: str       # "check", "test", "build" or "lint"
    line: int        # the spec line (for "lint", the module line; 0 when there is none)
    severity: str    # "error" or "warning"
    message: str
    code: str = ""


@dataclass
class TestLine:
    line: int
    passed: bool
    result: object
    errors: int
    failed: list = field(default_factory=list)   # (message, code) of each failed assertion


@dataclass
class Proof:
    spec: str
    kind: str = ""                # "TASK" or "BUILD"
    name: str = ""                # the TASK's PROC or the BUILD's NAME
    findings: list = field(default_factory=list)
    tests: list = field(default_factory=list)
    procedures: list = field(default_factory=list)   # the compiled VBA procedures, without ForgeLib
    module: str = ""              # ForgeLib and the procedures, lint-clean, when the spec is proven
    canonical: object = None      # the spec as `vclc fmt` writes it (None if it can't)
    strength: object = None       # (caught, total) deliberate bugs, with mutation=True
    seconds: dict = field(default_factory=dict)

    @property
    def errors(self):
        return [f for f in self.findings if f.severity == "error"]

    @property
    def warnings(self):
        return [f for f in self.findings if f.severity == "warning"]

    @property
    def failing(self):
        return [t for t in self.tests if not t.passed]

    @property
    def ok(self):
        return bool(self.kind and self.tests and self.module and not self.errors and not self.failing)

    def distance(self):
        """How far from proven: the stage that failed (check 3, test 2, build or lint 1), then how many
        problems. A spec whose only problem is a failing TEST is closer than one that doesn't check."""
        stages = {f.stage for f in self.errors}
        stage = 3 if "check" in stages else 2 if "test" in stages or self.failing else 1 if not self.module else 0
        return stage, len(self.errors) + len(self.failing)


def module_title(name):
    return f"KGC KRONOS module: {name}"


def prove(text, today=None, mutation=False, step=None):
    """Check, test, build and lint one spec, in that order, stopping at the first stage that fails.
    step(name) is told each stage as it starts: checking, testing, building, measuring."""
    step = step or (lambda name: None)
    proof = Proof(normalize(text))
    if is_question(proof.spec):
        message = coded("this is a question, not a spec: KRONOS shows it to the person", "K004")
        proof.findings.append(Finding("check", 1, "error", message, code_of(message)))
        return proof
    step("checking")
    clock = time.perf_counter()
    for p in validate_text(proof.spec):
        proof.findings.append(Finding("check", p.line, p.severity, p.message, p.code))
    proof.seconds["check"] = round(time.perf_counter() - clock, 3)
    if proof.errors:
        return proof
    spec, _ = parse(proof.spec)
    proof.kind = spec["kind"]
    try:
        proof.canonical = format_text(proof.spec)
    except FormatError:
        proof.canonical = None
    step("testing")
    clock = time.perf_counter()
    try:
        results = run_tests(proof.spec, today)
    except VclError as exc:
        message = exc.args[0] if exc.args else str(exc)
        proof.findings.append(Finding("test", 0, "error", message, code_of(message)))
        return proof
    proof.tests = [TestLine(r.line, r.passed, r.result, r.errors, [(m, code_of(m)) for m in r.failed])
                   for r in results]
    proof.seconds["test"] = round(time.perf_counter() - clock, 3)
    if proof.failing or not proof.tests:
        return proof
    step("building")
    clock = time.perf_counter()
    try:
        if proof.kind == "BUILD":
            proof.name, proof.procedures = compile_build_procs(proof.spec, SPEC_NAME)
        else:
            proof.procedures = [compile_text(proof.spec, SPEC_NAME)]
            proof.name = next(str(k["tree"].children[0]) for k in spec["keys"] if k["key"] == "PROC")
    except CompileError as exc:
        proof.findings.append(Finding("build", 0, "error", str(exc)))
        return proof
    module = assemble(proof.procedures, module_title(proof.name))
    found = lint(module)
    proof.findings += [Finding("lint", line, "error", message) for line, message in found]
    proof.seconds["build"] = round(time.perf_counter() - clock, 3)
    if found:
        return proof
    proof.module = module
    if mutation:
        step("measuring")
        clock = time.perf_counter()
        caught, total, _ = mutate_report(proof.spec, snapshot=True, today=today)
        proof.strength = (caught, total)
        proof.seconds["mutation"] = round(time.perf_counter() - clock, 3)
    return proof


def feedback(proof):
    """The proof's problems as vclc prints them: the repair feedback the corpus helpers used."""
    out = [f"{SPEC_NAME}:{f.line}: {f.severity}: {f.message}{tag(f.message)}"
           for f in proof.findings if f.stage == "check"]
    for f in proof.findings:
        if f.stage == "test":
            out.append(f"FAIL test   {SPEC_NAME}: {f.message}{tag(f.message)}")
    if proof.failing:
        out.append(f"FAIL test   {SPEC_NAME}  ({len(proof.tests) - len(proof.failing)}/{len(proof.tests)} tests)")
        for t in proof.failing:
            out.append(f"       line {t.line}: RESULT={t.result} ERROR_COUNT={t.errors}")
            out += [f"         {message}{f' [{code}]' if code else ''}" for message, code in t.failed]
    out += [f"FAIL build  {SPEC_NAME}: {f.message}" for f in proof.findings if f.stage == "build"]
    lint_found = [f for f in proof.findings if f.stage == "lint"]
    if lint_found:
        out.append("FAIL lint   the compiled module")
        out += [f"       line {f.line}: {f.message}" for f in lint_found]
    return out


def repair_message(reply, proof):
    """What goes back to the AI after a try that is not proven."""
    if proof is None:
        return (f"Your reply could not be used: {reply.problem}{tag(reply.problem)}. Reply with the whole spec in "
                f"one vcl code block, or with only {QUESTION} <one short question> in it.")
    return ("The VBA Forge toolchain checked your spec and it did not pass. Its messages follow, as vclc prints "
            "them. Fix every problem and reply with the whole corrected spec in one vcl code block.\n\n"
            + "\n".join(feedback(proof)))


def fix_message(proof):
    """What goes to the AI with a spec the person edited and that is not proven (Ask Codex to fix, D74)."""
    return ("I edited your spec by hand. Here it is with the toolchain's messages, as vclc prints them. Fix "
            "every problem, keep my changes where you can, and reply with the whole corrected spec in one vcl "
            "code block.\n\n```vcl\n" + proof.spec + "```\n\n" + "\n".join(feedback(proof)))


# ---------- the spec in plain words (the window's Review and Save steps, D74) ----------

def change_words(value):
    """One MODIFIES value in words."""
    m = MODIFIES_RE.fullmatch(value.strip())
    if not m:
        return value.strip()
    arg = f'"{m["arg"]}"'
    if m["cols"]:
        cols = [c.strip().strip('"') for c in m["cols"].split(",") if c.strip()]
        names = cols[0] if len(cols) == 1 else ", ".join(cols[:-1]) + " and " + cols[-1]
        return f"the {names} column{'' if len(cols) == 1 else 's'} of the sheet you give as {arg}"
    if m["fills"]:
        return f"the fill colours of the sheet you give as {arg}"
    if m["files"]:
        return f"files in the folder you give as {arg}: it only adds copies, never deletes or moves"
    if m["sheets"]:
        return f"the sheets of the workbook you give as {arg}: it may add or copy sheets"
    return f"anything on the sheet you give as {arg}"


def summary(proof):
    """What a spec does, assumes and changes, in words: {does, steps, assumes, changes, warnings}.
    For a BUILD, what changes is what its modMain.RunHeadless declares (D49)."""
    parts = proof.spec.split("\n---\n")
    lines = proof.spec.splitlines()
    goals = [ln[4:].strip() for ln in lines if re.match(r"GOAL\s", ln)]
    assumes = [ln.split(":", 1)[1].strip() for ln in lines if ln.startswith("# ASSUME:")]
    owner = next((part for part in parts if re.search(r"^PROC\s+RunHeadless\s*$", part, re.M)), parts[0])
    changes = [change_words(ln.split(None, 1)[1]) for ln in owner.splitlines()
               if re.match(r"MODIFIES\s", ln) and len(ln.split(None, 1)) == 2]
    return {"does": goals[0] if goals else "", "steps": goals[1:] if proof.kind == "BUILD" else [],
            "assumes": assumes, "changes": changes, "warnings": [f.message for f in proof.warnings]}


def sheet_hints(spec_text):
    """{argument: the practice sheet name its first TEST passes}, for the Save step's placeholders."""
    try:
        proc, kind, args, ret = task_signature(spec_text)
    except CompileError:
        return {}
    for ln in spec_text.splitlines():
        m = re.match(rf"TEST\s.*\bcall=\s*{re.escape(proc)}\(([^)]*)\)", ln)
        if m:
            given = [v.strip() for v in m.group(1).split(",")]
            return {name: value for (name, typ), value in zip(args, given) if typ == "Worksheet" and value}
    return {}


def plain_result(attempt):
    """One try in plain words, for the working panel."""
    proof = attempt.proof
    if proof is None:
        return "the reply had no recipe in it"
    if proof.ok:
        return f"it passed every check and {len(proof.tests)} of {len(proof.tests)} practice tests"
    if any(f.stage == "check" for f in proof.errors):
        count = sum(f.stage == "check" for f in proof.errors)
        return f"the recipe broke {count} rule{'' if count == 1 else 's'} of the checker"
    if proof.failing or any(f.stage == "test" for f in proof.errors):
        return f"{len(proof.tests) - len(proof.failing)} of {len(proof.tests)} practice tests passed"
    return "the macro could not be built"


# ---------- the module ----------

def module_name(name):
    """The VBA module's name: modKronos plus the spec's name, at most 31 characters as Excel allows."""
    return ("modKronos" + name)[:31]


def bas_text(name, module):
    """The .bas file: an Attribute VB_Name line for File > Import, and Windows line endings."""
    return (f'Attribute VB_Name = "{name}"\n' + module).replace("\r\n", "\n").replace("\n", "\r\n")


def import_steps(file_name, macro):
    return ["1. Make a copy of the workbook and open the copy in Excel 2016.",
            "2. Press Alt+F11 to open the VBA editor.",
            f"3. Click File > Import File..., choose {file_name} and click Open. A module named "
            f"{file_name[:-4]} appears under Modules.",
            "4. Close the VBA editor. Press Alt+F8, choose " + macro + " and click Run.",
            "5. To keep the macro, save the workbook as an Excel Macro-Enabled Workbook (.xlsm).",
            "One KRONOS module per workbook: each module carries its own ForgeLib, and two would clash. "
            "If Excel shows a yellow bar saying macros are disabled, click Enable Content."]


# ---------- one run ----------

@dataclass
class Attempt:
    number: int
    source: str                   # "ai" or "person" (an edited spec, Re-check)
    reply: object = None          # the Reply, for an AI attempt
    proof: object = None          # the Proof, when there was a spec to prove


def read_report(folder):
    """A run folder's report.json as a dict, or None when it is missing or unreadable (History)."""
    try:
        report = json.loads((Path(folder) / "report.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    return report if isinstance(report, dict) else None


def short_name(request):
    words = re.findall(r"[a-z0-9]+", request.lower())[:4]
    return "-".join(words)[:30].strip("-") or "request"


class Run:
    """One request, start to end: the conversation, every attempt, and the run folder.

    The conversation goes in turns (D74): the request, an answer to a question, a change the person
    asks for, an edited spec sent back to the AI, or an edited spec re-checked by hand. Each turn gets
    MAX_TRIES tries by the AI, and what the run shows is the best attempt of the latest turn.
    progress(kind, info) hears each step: asking (with the try), checking, testing, building,
    measuring, and tried (with the try's result in words). cancel() may be called from another thread.
    """

    def __init__(self, request, ai, data, today=None, mutation=False, now=None, progress=None):
        self.request = request.strip()
        if not self.request:
            raise ValueError("the request is empty")
        self.ai, self.today, self.mutation = ai, today, mutation
        self.progress = progress or (lambda kind, info: None)
        self.started = now or datetime.datetime.now()
        self.folder = self._new_folder(data_folder(data) / "runs")
        self.messages, self.attempts = [], []
        self.status, self.question, self.note = "new", "", ""
        self.questions = self.calls = self.turn_start = 0
        self.cancelled = False
        self.module_file, self.macro = None, ""
        self._write("request.txt", self.request + "\n")

    def _new_folder(self, runs):
        base = f"{self.started:%Y-%m-%d_%H%M%S}_{short_name(self.request)}"
        for n in range(1, 100):
            folder = runs / (base if n == 1 else f"{base}_{n}")
            try:
                folder.mkdir(parents=True)
                return folder
            except FileExistsError:
                continue
        raise ValueError(f"too many runs named {base}")

    def _write(self, name, text):
        (self.folder / name).write_text(text, encoding="utf-8", newline="\n")

    @property
    def name(self):
        return self.folder.name

    def conversation(self):
        return [{"role": m["role"], "text": m["text"]} for m in self.messages]

    def _say(self, role, kind, text):
        message = {"role": role, "kind": kind, "text": text,
                   "time": datetime.datetime.now().isoformat(timespec="seconds")}
        self.messages.append(message)
        with open(self.folder / "conversation.jsonl", "a", encoding="utf-8", newline="\n") as f:
            f.write(json.dumps(message, ensure_ascii=False) + "\n")

    def _attempt(self, source, reply, proof):
        attempt = Attempt(len(self.attempts) + 1, source, reply, proof)
        self.attempts.append(attempt)
        if proof is not None:
            self._write(f"attempt_{attempt.number}.vcl", proof.spec)
        return attempt

    def _stop(self, status, note=""):
        self.status, self.note = status, note
        self.save_report()
        return self.status

    @property
    def ai_tries(self):
        return sum(a.source == "ai" for a in self.attempts)

    @property
    def turn_tries(self):
        return sum(a.source == "ai" for a in self.attempts[self.turn_start:])

    def _new_turn(self):
        self.turn_start = len(self.attempts)

    def cancel(self):
        """Stop the AI call in progress (from another thread); the run ends this turn as cancelled. It stays
        cancelled until ready() (the window calls it, on its own thread, before the next piece of work)."""
        self.cancelled = True
        getattr(self.ai, "cancel", lambda: None)()

    def ready(self):
        self.cancelled = False
        getattr(self.ai, "reset", lambda: None)()

    def start(self):
        if self.status != "new":
            raise ValueError("this run has already started")
        self._say("user", "request", self.request)
        self._new_turn()
        return self._converse()

    def answer(self, text):
        """The person's answer to the AI's question."""
        if self.status != "question":
            raise ValueError("the AI has asked no question")
        if not text.strip():
            raise ValueError("the answer is empty")
        self.question = ""
        self._say("user", "answer", text.strip())
        self._new_turn()
        return self._converse()

    def follow_up(self, text):
        """A change the person asks for once the run has a result (rule 13: the AI reissues the spec)."""
        if self.status not in ("proven", "not_proven"):
            raise ValueError("a change can be asked for only once there is a result")
        if not text.strip():
            raise ValueError("the change is empty")
        self._say("user", "change", text.strip())
        self._new_turn()
        return self._converse()

    def retry(self):
        """Ask the AI again after it gave no reply or was cancelled; the turn keeps its tries."""
        if self.status not in ("ai_error", "cancelled"):
            raise ValueError("only a run whose AI gave no reply, or was cancelled, is retried")
        return self._converse()

    def _converse(self):
        while self.turn_tries < MAX_TRIES:   # each pass adds an attempt or returns
            if self.cancelled:
                return self._stop("cancelled", "cancelled")
            self.progress("asking", {"try": self.turn_tries + 1, "of": MAX_TRIES})
            self.calls += 1
            try:
                text = self.ai.ask(self.conversation())
            except AiError as exc:
                if self.cancelled:
                    return self._stop("cancelled", "cancelled")
                return self._stop("ai_error", f"the AI gave no reply: {exc}")
            if self.cancelled:
                return self._stop("cancelled", "cancelled")
            self._say("assistant", "reply", text)
            reply = read_reply(text)
            if reply.kind == "question":
                self.questions += 1
                if self.questions > MAX_QUESTIONS:
                    return self._stop("not_proven", f"the AI asked more than {MAX_QUESTIONS} questions: put the "
                                                    "details in the request and start a new run")
                self.question = reply.text
                return self._stop("question")
            step = lambda name: self.progress(name, {})   # noqa: E731
            proof = prove(reply.text, self.today, self.mutation, step) if reply.kind == "spec" else None
            attempt = self._attempt("ai", reply, proof)
            self.progress("tried", {"try": self.turn_tries, "result": plain_result(attempt),
                                    "proven": proof is not None and proof.ok})
            if proof is not None and proof.ok:
                return self._stop("proven")
            if self.turn_tries < MAX_TRIES:
                self._say("user", "repair", repair_message(reply, proof))
        best = self.best()
        return self._stop("not_proven", f"no spec passed after {MAX_TRIES} tries"
                          + (f": attempt {best.number} came closest; edit it and Re-check" if best else ""))

    def recheck(self, text):
        """The person's edited spec, proven like the AI's (no AI call): a turn of its own."""
        if not text.strip():
            raise ValueError("the spec is empty")
        self._new_turn()
        proof = prove(text, self.today, self.mutation, lambda name: self.progress(name, {}))
        self._attempt("person", None, proof)
        self.question = ""   # an edited spec answers any question still open
        return self._stop("proven" if proof.ok else "not_proven", "" if proof.ok else "the edited spec is not proven")

    def fix(self, text):
        """The person's edited spec: proven first, and if it is not, sent to the AI with its problems."""
        if self.recheck(text) == "proven":
            return self.status
        self._say("user", "fix", fix_message(self.attempts[-1].proof))
        return self._converse()

    def best(self):
        """The latest turn's latest proven attempt, else the one closest to proven (Proof.distance; the later
        one on a tie)."""
        proved = [a for a in self.attempts[self.turn_start:] if a.proof is not None]
        if not proved:
            return None
        return min(reversed(proved), key=lambda a: (False, -a.number) if a.proof.ok else (True, a.proof.distance()))

    def proven(self):
        best = self.best()
        return best if best is not None and best.proof.ok else None

    def starter_args(self):
        """The arguments the starter macro needs values for ([] for none, None when there is no starter)."""
        attempt = self.proven()
        if attempt is None:
            raise ValueError("no proven spec yet")
        return None if attempt.proof.kind == "BUILD" else starter_args(attempt.proof.spec)

    def save_module(self, values=None, starter=True, dest=None):
        """Write the proven spec's .bas to the run folder, and to dest when given (the window's Save As);
        with a starter macro for a TASK unless starter=False. Returns the path written for the person
        (dest, or the run folder's copy). ValueError when a value is missing or wrong."""
        attempt = self.proven()
        if attempt is None:
            raise ValueError("no proven spec yet")
        proof, values = attempt.proof, dict(values or {})
        procedures, macro = list(proof.procedures), "Run"
        needs = self.starter_args()
        if needs is None and proof.kind == "TASK":
            macro = proof.name
        elif needs is not None:
            missing = [f"${n} ({t})" for n, t in needs if n not in values]
            if starter and missing:
                raise ValueError("the starter macro needs a value for " + ", ".join(missing))
            if starter:
                try:
                    procedures.append(starter_sub(proof.spec, values))
                except CompileError as exc:
                    raise ValueError(str(exc)) from None
                macro = "Run" + proof.name
            else:
                macro = ""
        module = assemble(procedures, module_title(proof.name))
        found = lint(module)
        if found:
            raise ValueError("the module failed lint: " + "; ".join(f"line {n}: {m}" for n, m in found[:5]))
        name = module_name(proof.name)
        path = self.folder / f"{name}.bas"
        path.write_bytes(bas_text(name, module).encode("ascii"))
        if dest is not None:
            Path(dest).parent.mkdir(parents=True, exist_ok=True)
            Path(dest).write_bytes(bas_text(name, module).encode("ascii"))
        self.module_file, self.macro = path.name, macro
        self.save_report()
        return Path(dest) if dest is not None else path

    # ---------- the report ----------

    STATUS_WORDS = {"new": "NOT STARTED", "question": "WAITING FOR AN ANSWER: the AI asked a question",
                    "proven": "PROVEN: the spec passed every check and every TEST",
                    "not_proven": "NOT PROVEN", "ai_error": "NO REPLY from the AI", "cancelled": "CANCELLED"}

    def report_lines(self):
        out = [f"KGC KRONOS run {self.name}", f"Request: {self.request}", f"AI: {self.ai.name}",
               f"Status: {self.STATUS_WORDS[self.status]}" + (f" ({self.note})" if self.note else ""),
               f"Tries: {self.ai_tries} by the AI, {len(self.attempts) - self.ai_tries} by hand; "
               f"{self.questions} question{'' if self.questions == 1 else 's'}"]
        if self.status == "question":
            out.append(f"The AI asks: {self.question}")
        best = self.best()
        if best is not None and best.proof.ok:
            words = summary(best.proof)
            out += ["", f"What it does: {words['does']}"]
            out += ["What it assumes:"] + [f"  - {a}" for a in words["assumes"]] if words["assumes"] else []
            out += ["What it changes:"] + [f"  - {c}" for c in words["changes"]] if words["changes"] else []
        for a in self.attempts:
            out += [""] + self.attempt_lines(a, a is best)
        if self.module_file:
            out += ["", f"Module: {self.module_file}"]
            if self.macro:
                out += import_steps(self.module_file, self.macro)
            else:
                out.append("No starter macro: call the procedure from your own macro, or give the real sheet names.")
        return out

    def attempt_lines(self, attempt, best):
        head = f"Attempt {attempt.number} ({'AI' if attempt.source == 'ai' else 'edited by hand'})"
        proof = attempt.proof
        if proof is None:
            return [f"{head}: not a spec: {attempt.reply.problem}{tag(attempt.reply.problem)}"]
        out = [f"{head}: {'PROVEN' if proof.ok else 'NOT PROVEN'}{' (the one shown)' if best else ''}",
               f"  Checks: {len(proof.errors)} error(s), {len(proof.warnings)} warning(s)"]
        if proof.tests:
            out.append(f"  TESTs: {len(proof.tests) - len(proof.failing)} of {len(proof.tests)} passed")
        if proof.module:
            out.append(f"  VBA: {len(proof.module.splitlines())} lines, lint-clean")
        if proof.strength:
            out.append(f"  Test strength: {proof.strength[0]} of {proof.strength[1]} deliberate bugs caught")
        if proof.canonical is not None and proof.canonical != proof.spec:
            out.append("  Layout: `vclc fmt` would tidy it (shown, not forced)")
        out += ["  " + line for line in feedback(proof)]
        return out

    def report(self):
        def attempt_json(a):
            p = a.proof
            item = {"number": a.number, "source": a.source, "file": f"attempt_{a.number}.vcl" if p else None,
                    "reply": a.reply.kind if a.reply else None, "problem": a.reply.problem if a.reply else ""}
            if p is not None:
                item.update({"proven": p.ok, "kind": p.kind, "name": p.name, "seconds": p.seconds,
                             "canonical": p.canonical == p.spec if p.canonical is not None else None,
                             "strength": list(p.strength) if p.strength else None,
                             "findings": [vars(f) for f in p.findings],
                             "tests": [{"line": t.line, "passed": t.passed, "result": t.result, "errors": t.errors,
                                        "failed": [{"message": m, "code": c} for m, c in t.failed]} for t in p.tests]})
            return item
        best = self.best()
        return {"run": self.name, "request": self.request, "ai": self.ai.name, "status": self.status,
                "note": self.note, "question": self.question, "questions": self.questions,
                "today": self.today.isoformat() if self.today else None,
                "best": best.number if best else None, "module": self.module_file,
                "attempts": [attempt_json(a) for a in self.attempts]}

    def save_report(self):
        self._write("report.txt", "\n".join(self.report_lines()) + "\n")
        self._write("report.json", json.dumps(self.report(), indent=1, default=str) + "\n")


# ---------- settings and the AI ----------

SETTINGS = {"ai": "codex", "codex_command": CODEX_COMMAND, "codex_timeout": CODEX_TIMEOUT,
            # the window (D74): the recipe and test strength are off by default; the first-run checklist's
            # two ticks; where Save suggests putting the module ("" is <data folder>/macros)
            "show_recipe": False, "measure_strength": False, "data_controls_confirmed": False,
            "practice_confirmed": False, "macro_folder": ""}
FLAGS = ("show_recipe", "measure_strength", "data_controls_confirmed", "practice_confirmed")


def check_settings(settings, where="settings"):
    if settings["ai"] not in ("codex", "fake"):
        raise ValueError(f"the ai setting must be codex or fake, got {settings['ai']!r}")
    if not (isinstance(settings["codex_command"], list) and settings["codex_command"]
            and all(isinstance(a, str) for a in settings["codex_command"])):
        raise ValueError("codex_command must be a list of strings, starting with the program")
    if isinstance(settings["codex_timeout"], bool) or not isinstance(settings["codex_timeout"], int) \
            or not 10 <= settings["codex_timeout"] <= 3600:
        raise ValueError("codex_timeout must be a whole number of seconds from 10 to 3600")
    bad = [k for k in FLAGS if not isinstance(settings[k], bool)]
    if bad:
        raise ValueError(f"{where}: {', '.join(bad)} must be true or false")
    if not isinstance(settings["macro_folder"], str):
        raise ValueError("macro_folder must be a folder name, or empty")


def load_settings(data):
    """settings.json in the data folder over the defaults; a key it doesn't know is an error."""
    settings = {k: (list(v) if isinstance(v, list) else v) for k, v in SETTINGS.items()}
    path = Path(data) / "settings.json"
    if path.exists():
        loaded = json.loads(path.read_text(encoding="utf-8"))
        unknown = sorted(set(loaded) - set(SETTINGS))
        if unknown:
            raise ValueError(f"{path}: unknown settings {', '.join(unknown)}")
        settings.update(loaded)
    check_settings(settings, str(path))
    return settings


def save_settings(data, settings):
    """Write settings.json in the data folder: the known keys only, checked first."""
    check_settings(settings)
    known = {k: settings[k] for k in SETTINGS}
    (Path(data) / "settings.json").write_text(json.dumps(known, indent=1, sort_keys=True) + "\n",
                                              encoding="utf-8", newline="\n")


def make_ai(settings, data, replies=None):
    if settings["ai"] == "fake":
        if replies is None:
            raise ValueError("the fake AI needs its replies (--replies FILE)")
        return FakeAI(replies)
    return CodexAI(Path(data) / "codex", PROMPT.read_text(encoding="utf-8"), settings["codex_command"],
                   settings["codex_timeout"])


# ---------- vclc ask ----------

def say(text=""):
    """Print without failing on a character the console can't show (Windows, D72)."""
    encoding = getattr(sys.stdout, "encoding", None) or "utf-8"
    print(str(text).encode(encoding, "replace").decode(encoding))


def ask_person(prompt):
    if not sys.stdin.isatty():
        return None
    try:
        return input(prompt).strip() or None
    except EOFError:
        return None


def parse_args(argv):
    ap = argparse.ArgumentParser(prog="vclc ask", description="A request in plain English to a proven VBA module.")
    ap.add_argument("request", nargs="?", default="")
    ap.add_argument("--ai", choices=("codex", "fake"))
    ap.add_argument("--replies", help="the fake AI's replies: a JSON list of strings")
    ap.add_argument("--data", help="the data folder (default: Documents/KGC KRONOS)")
    ap.add_argument("--answer", action="append", default=[], help="an answer to the AI's next question")
    ap.add_argument("--arg", action="append", default=[], metavar="NAME=VALUE",
                    help="the real sheet name (or value) behind an argument, for the starter macro")
    ap.add_argument("--today", type=datetime.date.fromisoformat, help="pins TODAY for the TESTs")
    ap.add_argument("--mutation", action="store_true", help="also measure test strength (slower)")
    ap.add_argument("--check-ai", action="store_true", help="check the AI's command line and stop")
    return ap.parse_args(argv)


def main(argv):
    try:
        args = parse_args(argv)
    except SystemExit as exc:
        return 0 if exc.code == 0 else 1
    try:
        data = data_folder(args.data or default_data_folder())
        settings = load_settings(data)
        if args.ai:
            settings["ai"] = args.ai
        replies = json.loads(Path(args.replies).read_text(encoding="utf-8")) if args.replies else None
        ai = make_ai(settings, data, replies)
    except (OSError, ValueError) as exc:
        say(f"vclc ask: {exc}")
        return 1
    if args.check_ai:
        problems = ai.check() if hasattr(ai, "check") else []
        for problem in problems:
            say(problem)
        say(f"the {ai.name} command line looks right" if not problems else f"{len(problems)} problem(s)")
        return 1 if problems else 0
    values = {}
    for item in args.arg:
        name, sep, value = item.partition("=")
        if not sep:
            say(f"vclc ask: --arg needs NAME=VALUE, got {item!r}")
            return 1
        values[name.strip().lstrip("$")] = value
    try:
        run = Run(args.request, ai, data, args.today, args.mutation)
    except ValueError as exc:
        say(f"vclc ask: {exc}")
        return 1
    say(f"Asking {ai.name} for the VCL spec ...")
    run.start()
    answers = list(args.answer)
    for _ in range(MAX_QUESTIONS):   # each answer is one more turn; the run stops asking after MAX_QUESTIONS
        if run.status != "question":
            break
        say(f"The AI asks: {run.question}")
        answer = answers.pop(0) if answers else ask_person("Your answer: ")
        if answer is None:
            say('No answer given: run it again with --answer "..."')
            break
        run.answer(answer)
    if run.status == "proven":
        needs = run.starter_args() or []
        for name, typ in needs:
            if name not in values:
                hint = "the real sheet name" if typ == "Worksheet" else f"the value ({typ})"
                typed = ask_person(f"{hint} for ${name}: ")
                if typed is not None:
                    values[name] = typed
        missing = [n for n, _ in needs if n not in values]
        try:
            run.save_module(values, starter=not missing)
        except ValueError as exc:
            run.note = f"the module was not saved: {exc}"
            run.save_report()
    for line in run.report_lines():
        say(line)
    say(f"\nRun folder: {run.folder}")
    return 0 if run.status == "proven" and run.module_file else 1
