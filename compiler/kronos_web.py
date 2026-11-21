"""The bridge between KRONOS Offline's page and the engine (D86, docs/kronos-offline.md).

KRONOS Offline runs this repo's toolchain inside the browser, through Pyodide (Python compiled to WebAssembly).
The page's worker (web/worker.js) writes the very same compiler/, vcl/ and vba/ files into Pyodide's file system
and calls check(); the tests call it in normal Python. Both get the same canonical JSON text, so "the browser gives
the same results" is one string comparison. Every answer comes from kronos_engine: this module only reads the
pasted reply, proves it and turns the proof into plain data; it holds no engine logic of its own.

    import kronos_web
    kronos_web.check(text)                       # the JSON text below
    kronos_web.check(text, today="2026-11-14")   # pins TODAY for the TESTs (a date or an ISO date)

The JSON: reply ("spec", "question" or "none"), question, proven, words (the result in plain words), kind, name,
findings, tests, problems (as vclc prints them: what the AI gets back), module (the VBA module's name) and bas (the
.bas text with Windows line endings, "" unless proven).

The upload (D87, compiler/kronos_book.py) has three more calls, each answering canonical JSON text:

    kronos_web.describe_book(data, name)                    # the workbook's sheets, headings, sample rows and AI text
    kronos_web.arguments(text)                              # the recipe's ARGS with defaults, and whether it uses files
    kronos_web.apply_changes(text, data, name, args_json, today=None, step=None)
                                                            # (JSON text, the changed copy's bytes or b"")
Each answers {"ok": false, "message": ...} for a file or recipe it can't use, in words for the person.

Recipes from a file (D89, compiler/kronos_recipes.py): a .vcl, .txt, .md or Word .docx file the person opens.

    kronos_web.read_recipe_file(data, name)                 # {"ok", "name", "recipes": [{"title", "text"}], "message"}

The wizard (D89, compiler/kronos_wizard.py): KRONOS.exe's engine with copy and paste for the AI step.

    kronos_web.wizard_reset(data_folder, today=None)        # a fresh wizard (the page's start, and after Cancel)
    kronos_web.wizard_call(action, id, args_json, step=None)
        # action: ask, recipe, message, paste, answer, problems, save, view, restore -> {"ok", ...}
    kronos_web.words()                                      # EXAMPLE, TIPS, NEVER and the prompt's SHA-256
    kronos_web.apply_changes_many(texts_json, data, name, args_list_json, today=None, step=None)
        # several recipes, one after another, on one copy of the workbook (kronos_book.apply_many)
"""
import datetime
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from errcodes import tag  # noqa: E402
from kronos_engine import Attempt, bas_text, feedback, module_name, plain_result, prove, read_reply  # noqa: E402


def result(text, today=None, step=None):
    """The pasted reply (an AI's whole reply, or a recipe on its own), proven by the engine, as plain data.
    step(name) hears each stage as it starts: checking, testing, building."""
    if isinstance(today, str):
        today = datetime.date.fromisoformat(today)
    reply = read_reply(text)
    out = {"reply": reply.kind, "question": "", "proven": False, "kind": "", "name": "", "findings": [], "tests": [],
           "problems": [], "module": "", "bas": "", "offline_only": False}
    if reply.kind == "question":
        out.update(question=reply.text, words="the AI asked a question instead of writing a recipe")
        return out
    if reply.kind == "none":
        out.update(problems=[f"{reply.problem}{tag(reply.problem)}"], words=plain_result(Attempt(1, "person")))
        return out
    proof = prove(reply.text, today, step=step, offline=True)   # KRONOS Offline proves its own recipes (D118)
    out.update(proven=proof.ok, words=plain_result(Attempt(1, "person", reply, proof)), kind=proof.kind,
               name=proof.name, problems=feedback(proof),
               findings=[{"stage": f.stage, "line": f.line, "severity": f.severity, "message": str(f.message),
                          "code": f.code} for f in proof.findings],
               tests=[{"line": t.line, "passed": t.passed, "result": t.result, "errors": t.errors,
                       "failed": [{"message": str(m), "code": c} for m, c in t.failed]} for t in proof.tests])
    out["offline_only"] = proof.offline_only   # D118: proven, but no VBA (a DOCUMENT or a read= recipe)
    if proof.ok and proof.module:
        out.update(module=module_name(proof.name), bas=bas_text(module_name(proof.name), proof.module))
    return out


def canonical(data):
    """One JSON text for one result: sorted keys, no spaces, every character kept as it is."""
    return json.dumps(data, sort_keys=True, ensure_ascii=False, separators=(",", ":"), default=str)


def check(text, today=None, step=None):
    """result() as canonical JSON text: what the page's worker hands to the page."""
    return canonical(result(text, today, step))


# ---------------------------------------------------------------- the upload (D87)

def _book():
    import kronos_book   # here, so that checking a recipe never needs the workbook reader
    return kronos_book


def _bytes(data):
    """A workbook's bytes, from bytes or from the Uint8Array the page's worker hands over."""
    return bytes(data.to_py() if hasattr(data, "to_py") else data)


def describe_book(data, name):
    """The uploaded workbook, described (canonical JSON text)."""
    book = _book()
    try:
        return canonical({"ok": True, **book.describe(_bytes(data), str(name))})
    except book.BookError as ex:
        return canonical({"ok": False, "message": str(ex)})


def arguments(text):
    """What a recipe's procedure takes, and whether it works with files (canonical JSON text)."""
    book = _book()
    try:
        return canonical({"ok": True, "args": book.arguments(text), "files": book.uses_files(text)})
    except book.BookError as ex:
        return canonical({"ok": False, "message": str(ex)})


def apply_changes(text, data, name, args_json, today=None, step=None):
    """Make a recipe's changes on a copy of the uploaded workbook: (canonical JSON text, the copy's bytes or b"").
    step(name) hears each stage as it starts: checking, testing, building, reading, running, writing."""
    book = _book()
    args = json.loads(args_json) if isinstance(args_json, str) else dict(args_json or {})
    try:
        result, out = book.apply(text, _bytes(data), str(name), args, today, step)
    except book.BookError as ex:
        return canonical({"ok": False, "message": str(ex)}), b""
    return canonical(result), out or b""


# ---------- recipes from a file (D89) ----------

def read_recipe_file(data, name):
    """The recipes in a file the person opened (canonical JSON text)."""
    import kronos_recipes
    return canonical(kronos_recipes.read_recipes(_bytes(data), str(name)))


# ---------- the wizard (D89) ----------

_WIZARD = {}


def wizard_reset(data_folder="/kronos-data", today=None):
    """A fresh wizard: the page's start, and its first call after Cancel stopped the browser's Python."""
    import kronos_wizard
    _WIZARD["wizard"] = kronos_wizard.Wizard(data_folder, today)
    return canonical({"ok": True})


def wizard_call(action, id, args_json, step=None):
    """One wizard action, as canonical JSON: {"ok": true, "view" | "text" | "saved" | "views": ...} or
    {"ok": false, "message": ...} for something the person can put right."""
    import kronos_wizard
    if "wizard" not in _WIZARD:
        wizard_reset()
    w = _WIZARD["wizard"]
    args = json.loads(args_json) if isinstance(args_json, str) else list(args_json or [])
    try:
        if action == "ask":
            return canonical({"ok": True, "view": w.ask(id, args[0], args[1] if len(args) > 1 else None, step)})
        if action == "recipe":
            return canonical({"ok": True, "view": w.recipe(id, args[0], args[1] if len(args) > 1 else "", step)})
        if action == "message":
            return canonical({"ok": True, "text": w.message(id, bool(args[0]) if args else True)})
        if action == "paste":
            return canonical({"ok": True, "view": w.paste(id, args[0], step)})
        if action == "answer":
            return canonical({"ok": True, "view": w.answer(id, args[0], step)})
        if action == "problems":
            return canonical({"ok": True, "view": w.send_problems(id, step)})
        if action == "save":
            return canonical({"ok": True, "saved": w.save(id, args[0] if args else {})})
        if action == "view":
            return canonical({"ok": True, "view": w.view(id)})
        if action == "restore":
            return canonical({"ok": True, "views": w.restore(args[0] if args else [], step)})
    except kronos_wizard.WizardError as ex:
        return canonical({"ok": False, "message": str(ex)})
    raise ValueError(f"no wizard action {action!r}")


def words():
    """KRONOS's own words for the page (the same as KRONOS.exe's), and the Orchestrator prompt's SHA-256, by which
    the page knows whether the prompt changed since the person last put it in their Orchestrator project."""
    import hashlib
    import kronos_words
    from kronos_engine import PROMPT
    prompt = PROMPT.read_bytes()
    return canonical({"example": kronos_words.EXAMPLE, "tips": kronos_words.TIPS, "never": kronos_words.NEVER,
                      "prompt": prompt.decode("utf-8"), "prompt_sha256": hashlib.sha256(prompt).hexdigest()})


def apply_changes_many(texts_json, data, name, args_list_json, today=None, step=None):
    """Several recipes, one after another, on one copy of the uploaded workbook: (canonical JSON, the copy's bytes or b"")."""
    book = _book()
    texts = json.loads(texts_json) if isinstance(texts_json, str) else list(texts_json)
    args_list = json.loads(args_list_json) if isinstance(args_list_json, str) else list(args_list_json or [])
    try:
        result, out = book.apply_many(texts, _bytes(data), str(name), args_list, today, step)
    except book.BookError as ex:
        return canonical({"ok": False, "message": str(ex)}), b""
    return canonical(result), out or b""
