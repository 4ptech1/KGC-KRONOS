"""The benchmark, the only scoreboard (VEP 0.6 Stage A Phase 8, `docs/vep-0.6-stage-a.md` section 7; D135).

    python compiler/vclc.py bench --prompts DIR                 write what a writer sees for each held-out item
    python compiler/vclc.py bench CANDIDATES [--today D] [--json FILE]   score the candidate recipes

The items are the held-out pairs in data/eval: never trained on, paraphrased or repaired (D59). A writer (Claude as the
Orchestrator today, Codex, our model from March 2027) sees an item's request and its **contract**: the names and types
the hidden TESTs call (a TASK's PROC, KIND, ARGS and RET; a BUILD's NAME and its modMain.RunHeadless SIG). Nothing else
of the reference recipe.

A candidate is a file CANDIDATES/<id>.vcl (a recipe, or a whole AI reply with one vcl block). It is scored in four
levels; V needs a recipe, T needs V, and M and H each need T, independently (a right macro with weak TESTs is still
right, and strong TESTs don't make a wrong macro right):
    V  valid: it checks with no errors (the validator, the grammar, the type checker)
    T  tested: its own TESTs pass and it builds (as KRONOS proves it; a recipe only KRONOS Offline runs needs no VBA)
    M  strong: its own TESTs catch at least 90% of the deliberate bugs (the snapshot mutation score, D59's "strong
       tests required"; a candidate can't explain equivalent mutants, so the plain share counts)
    H  right: its procedure passes the hidden reference TESTs, with its own FIXTURE and TEST lines replaced by the
       reference's, so weak TESTs of its own can't help it
The headline is the share of items at H, the macro right. Beside it, "strict": the share at both H and M, a recipe the
corpus would store as it is (D64's floor). Missing candidates count as nothing reached.
"""
import datetime as dt
import json
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from kronos_engine import normalize, prove, read_reply  # noqa: E402
from vcl_corpus import DATA, FLOOR, items  # noqa: E402
from vcl_interp import VclError, run_tests  # noqa: E402
from vcl_mutate import mutate_report  # noqa: E402

LEVELS = ("V", "T", "M", "H")
KEYS = ("FIXTURE", "TEST")


def _key(line):
    m = re.match(r"([A-Z_]+)\s", line)
    return m.group(1) if m else ""


def contract(spec_text):
    """The lines of the reference recipe the hidden TESTs depend on: names and types, nothing of how it works."""
    head = spec_text.split("\n---\n")[0]
    if re.search(r"^BUILD\s*$", head, re.M):
        lines = [ln for ln in head.splitlines() if _key(ln) == "NAME"
                 or (_key(ln) == "SIG" and "modMain.RunHeadless(" in ln)]
        kind = "BUILD"
    else:
        lines = [ln for ln in head.splitlines() if _key(ln) in ("PROC", "KIND", "ARGS", "RET")]
        kind = "TASK"
    return kind, lines


def prompt(folder):
    """What a writer sees for one item: the request and the contract."""
    request = (folder / "task.txt").read_text(encoding="utf-8").strip()
    kind, lines = contract((folder / "spec.vcl").read_text(encoding="utf-8"))
    what = ("Write a VCL 0.7 BUILD. Its NAME and its modMain.RunHeadless signature must be exactly:" if kind == "BUILD"
            else "Write a VCL 0.7 TASK. Its PROC, KIND, ARGS and RET must be exactly:")
    return f"{request}\n\n{what}\n" + "\n".join(lines) + "\n"


def swap_tests(candidate, reference):
    """The candidate with its own FIXTURE and TEST lines replaced by the reference's, at the place its own stood."""
    ref = [ln for ln in reference.split("\n---\n")[0].splitlines() if _key(ln) in KEYS]
    parts = candidate.split("\n---\n")
    for n, part in enumerate(parts):
        lines = part.splitlines()
        where = [i for i, ln in enumerate(lines) if _key(ln) in KEYS]
        if where:
            kept = [ln for ln in lines if _key(ln) not in KEYS]
            at = where[0]
            parts[n] = "\n".join(kept[:at] + ref + kept[at:])
            return "\n---\n".join(parts) + ("\n" if not parts[-1].endswith("\n") else "")
    head = parts[0].rstrip("\n").splitlines()
    parts[0] = "\n".join(head + ref)
    return "\n---\n".join(parts) + "\n"


def _contract_differs(candidate, reference):
    want_kind, want = contract(reference)
    got_kind, got = contract(candidate)
    if want_kind != got_kind:
        return f"it is a {got_kind}, the contract asks for a {want_kind}"
    squash = lambda ls: [" ".join(ln.split()) for ln in ls]
    missing = [ln for ln in squash(want) if ln not in squash(got)]
    return f"the contract line {missing[0]!r} is not in it" if missing else ""


def score(candidate, reference, today=None):
    """{level: True/False for V, T, M, H, reached: the last level met ('' for none), why: what stopped it}."""
    out = {lv: False for lv in LEVELS}
    out.update(reached="", why="")
    reply = read_reply(candidate)
    if reply.kind != "spec":
        out["why"] = "no recipe in it"
        return out
    text = normalize(reply.text)
    proof = prove(text, today, offline=True)
    if any(f.stage == "check" for f in proof.errors):
        out["why"] = "it does not check: " + next(f.message for f in proof.errors if f.stage == "check")
        return out
    out["V"], out["reached"] = True, "V"
    if not proof.ok:
        out["why"] = "its own TESTs fail or it does not build"
        return out
    out["T"], out["reached"] = True, "T"
    caught, total, _ = mutate_report(text, snapshot=True, today=today)
    out["mutants"] = [caught, total]
    why = []
    if total and caught / total < FLOOR:
        why.append(f"its TESTs catch {caught} of {total} deliberate bugs, under {FLOOR:.0%}")
    else:
        out["M"] = True
    differs = _contract_differs(text, reference)
    if differs:
        why.append(differs)
    else:
        try:
            results = run_tests(swap_tests(text, reference), today)
            failed = [r.line for r in results if not r.passed]
            if failed or not results:
                why.append(f"{len(failed)} of {len(results)} hidden TESTs fail")
            else:
                out["H"] = True
        except (VclError, KeyError, TypeError, ValueError, IndexError) as ex:
            why.append(f"the hidden TESTs can't run on it: {ex}")
    out["reached"] = "+".join(["T"] + (["M"] if out["M"] else []) + (["H"] if out["H"] else []))
    out["why"] = "; ".join(why)
    return out


def run(candidates, today=None, kind="eval"):
    """Score every held-out item: {item id: score}, a missing candidate scoring nothing."""
    rows = {}
    for folder in items(kind):
        reference = (folder / "spec.vcl").read_text(encoding="utf-8")
        path = Path(candidates) / f"{folder.name}.vcl"
        if not path.exists():
            rows[folder.name] = {**{lv: False for lv in LEVELS}, "reached": "", "why": "no candidate"}
            continue
        rows[folder.name] = score(path.read_text(encoding="utf-8"), reference, today)
    return rows


def summary(rows):
    n = len(rows)
    counts = {lv: sum(r[lv] for r in rows.values()) for lv in LEVELS}
    strict = sum(r["H"] and r["M"] for r in rows.values())
    return {"items": n, **counts, "strict": strict, "headline": round(counts["H"] / n, 3) if n else 0.0}


def report(rows):
    lines = [f"{'item':6} {'reached':8} why"]
    for item, r in rows.items():
        lines.append(f"{item:6} {r['reached'] or '-':8} {r['why']}")
    s = summary(rows)
    lines.append("")
    lines.append(f"{s['items']} items: V {s['V']}, T {s['T']}, M {s['M']}, H {s['H']} (the headline: "
                 f"{s['H']} of {s['items']} right; strict, right and strong: {s['strict']})")
    return "\n".join(lines)


def main(argv):
    if "--prompts" in argv:
        out = Path(argv[argv.index("--prompts") + 1])
        out.mkdir(parents=True, exist_ok=True)
        for folder in items("eval"):
            (out / f"{folder.name}.txt").write_text(prompt(folder), encoding="utf-8")
        print(f"wrote {len(items('eval'))} prompts into {out}")
        return 0
    args = [a for a in argv if not a.startswith("--")]
    today = dt.date.fromisoformat(argv[argv.index("--today") + 1]) if "--today" in argv else None
    if "--today" in argv:
        args.remove(argv[argv.index("--today") + 1])
    if "--json" in argv:
        args.remove(argv[argv.index("--json") + 1])
    if len(args) != 1:
        print("usage: vclc bench --prompts DIR | vclc bench CANDIDATES [--today YYYY-MM-DD] [--json FILE]")
        return 1
    rows = run(args[0], today)
    print(report(rows))
    if "--json" in argv:
        Path(argv[argv.index("--json") + 1]).write_text(
            json.dumps({"today": (today or dt.date.today()).isoformat(), "summary": summary(rows), "items": rows},
                       indent=1, sort_keys=True) + "\n", encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
