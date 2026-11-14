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
           "problems": [], "module": "", "bas": ""}
    if reply.kind == "question":
        out.update(question=reply.text, words="the AI asked a question instead of writing a recipe")
        return out
    if reply.kind == "none":
        out.update(problems=[f"{reply.problem}{tag(reply.problem)}"], words=plain_result(Attempt(1, "person")))
        return out
    proof = prove(reply.text, today, step=step)
    out.update(proven=proof.ok, words=plain_result(Attempt(1, "person", reply, proof)), kind=proof.kind,
               name=proof.name, problems=feedback(proof),
               findings=[{"stage": f.stage, "line": f.line, "severity": f.severity, "message": str(f.message),
                          "code": f.code} for f in proof.findings],
               tests=[{"line": t.line, "passed": t.passed, "result": t.result, "errors": t.errors,
                       "failed": [{"message": str(m), "code": c} for m, c in t.failed]} for t in proof.tests])
    if proof.ok:
        out.update(module=module_name(proof.name), bas=bas_text(module_name(proof.name), proof.module))
    return out


def canonical(data):
    """One JSON text for one result: sorted keys, no spaces, every character kept as it is."""
    return json.dumps(data, sort_keys=True, ensure_ascii=False, separators=(",", ":"), default=str)


def check(text, today=None, step=None):
    """result() as canonical JSON text: what the page's worker hands to the page."""
    return canonical(result(text, today, step))
