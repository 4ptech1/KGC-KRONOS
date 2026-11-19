"""The data generator (VEP 0.6 Stage A Phase 7, D103): training examples made from the proven pairs by the toolchain.

    python compiler/vclc.py datagen [-o DIR] [--today YYYY-MM-DD] [--only c013,c020]

Writes, into DIR (default dist/datagen, git-ignored, made fresh like `vclc export`, never stored):
    repairs.jsonl    a spec that is not proven, the toolchain's messages, and the proven spec that fixes it
    contrast.jsonl   a spec one step away from the proven one that passes its own TESTs yet behaves differently:
                     "read carefully" data, never trained as a fix
    summary.json     how many of each, per pair, the date the TESTs ran with and the toolchain commit

Two sources of repairs:
- **mutants:** every mutant of a proven pair (`vcl_mutate.mutants`, the same small bugs the mutation score uses)
  that the toolchain refuses: it doesn't check, or a TEST fails. The messages are what KRONOS sends back to the AI
  (`kronos_engine.feedback`), so the model learns the repair loop it will run in.
- **the helpers' failed tries** in `data/repair/attempts/<id>/tryN.vcl`, for a pair stored in data/verified: proven
  again today, so the messages never go stale. A try that passes every check but whose TESTs catch too few mutants
  gets the corpus gate's message and the mutants that survive. A try that passes every gate is not a repair (what was
  wrong with it, a TASK for a build request, needs a reader) and is left out, and so is a try identical to the fix.

Only data/verified is used. The held-out pairs in data/eval are never mutated, repaired or paraphrased (D59).
A mutant the interpreter can't run at all (a crash, not a VclError) is counted and left out.
"""
import datetime as dt
import json
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from kronos_engine import feedback, normalize, prove  # noqa: E402
from vcl_corpus import DATA, FLOOR, items  # noqa: E402
from vcl_interp import VclError, run_tests  # noqa: E402
from vcl_mutate import mutate_report, mutants  # noqa: E402

__all__ = ["generate", "pair_examples", "attempt_examples", "snapshot"]


def snapshot(results):
    """What a run of the TESTs leaves, for telling two specs apart (the mutation score's snapshot)."""
    return [(r.result, r.state, r.files, r.messages, r.hidden, r.texts) for r in results]


def first_difference(want, got, lines):
    """In plain words: the first TEST whose run differs, and in what."""
    names = ("RESULT", "the sheets", "the files", "the LOG and MSG texts", "the hidden rows", "the file texts")
    for line, w, g in zip(lines, want, got):
        for name, a, b in zip(names, w, g):
            if a != b:
                if name == "the sheets":
                    sheets = sorted(s for s in set(a) | set(b) if a.get(s) != b.get(s))
                    return f"TEST on line {line}: sheet {', '.join(sheets)} ends differently"
                return f"TEST on line {line}: {name} differ"
    return "the runs differ"


def mutant_text(lines, i, body):
    out = list(lines)
    if body is None:
        del out[i]
    else:
        out[i] = "OP" + body
    return "\n".join(out) + "\n"


def pair_examples(pair_id, request, spec, today):
    """The repairs and contrast examples one proven pair gives, and how many mutants were left out."""
    repairs, contrast, crashed = [], [], 0
    golden = run_tests(spec, today)
    want = snapshot(golden)
    lines = spec.splitlines()
    seen = set()
    for i, what, body in mutants(spec):
        text = mutant_text(lines, i, body)
        if text in seen or text == spec:
            continue
        seen.add(text)
        where = f"line {i + 1}: {what}"
        try:
            proof = prove(text, today)
        except Exception:   # the interpreter can't run it at all: not a lesson
            crashed += 1
            continue
        if proof.errors or proof.failing:
            repairs.append({"type": "repair", "source": "mutant", "id": pair_id, "change": where,
                            "stage": "check" if any(f.stage == "check" for f in proof.errors) else "test",
                            "request": request, "broken": text, "feedback": feedback(proof), "fixed": spec})
            continue
        if not proof.ok:   # it checks and tests but doesn't build: never seen so far, so not guessed at
            crashed += 1
            continue
        got = snapshot(run_tests(text, today))
        if got != want:
            contrast.append({"type": "contrast", "id": pair_id, "change": where, "request": request,
                             "proven": spec, "mutant": text,
                             "differs": first_difference(want, got, [r.line for r in golden])})
    return repairs, contrast, crashed


def attempt_examples(pair_id, request, spec, today):
    """The repairs from the helpers' failed tries of one stored pair, and the tries left out (with why)."""
    folder = DATA / "repair" / "attempts" / pair_id
    out, skipped = [], []
    for path in sorted(folder.glob("try*.vcl"), key=lambda p: int(p.stem[3:])) if folder.is_dir() else []:
        text = normalize(path.read_text(encoding="utf-8"))
        if text == normalize(spec):
            skipped.append((path.name, "identical to the stored spec"))
            continue
        proof = prove(text, today)
        lines = feedback(proof)
        if proof.ok:
            caught, total, survivors = mutate_report(text, snapshot=True, today=today)
            if total and caught / total >= FLOOR:
                skipped.append((path.name, "passes every gate"))
                continue
            lines = [f"weak tests: the TESTs catch {caught} of {total} deliberate bugs, under the {FLOOR:.0%} floor; "
                     "add TEST cases or FIXTURE rows that catch these:"]
            lines += [f"       survives: line {line}: {what}" for line, what in survivors]
        out.append({"type": "repair", "source": "attempt", "id": pair_id, "change": path.name,
                    "stage": "mutation" if proof.ok else "check" if any(f.stage == "check" for f in proof.errors)
                    else "test" if proof.errors or proof.failing else "build",
                    "request": request, "broken": text, "feedback": lines, "fixed": spec})
    return out, skipped


def commit():
    try:
        return subprocess.run(["git", "rev-parse", "--short", "HEAD"], cwd=HERE.parent, capture_output=True,
                              text=True, check=True).stdout.strip()
    except (OSError, subprocess.CalledProcessError):
        return "unknown"


def generate(dest, today=None, only=None, progress=None):
    """Write repairs.jsonl, contrast.jsonl and summary.json into dest; returns the summary."""
    today = today or dt.date.today()
    dest = Path(dest)
    dest.mkdir(parents=True, exist_ok=True)
    summary = {"today": today.isoformat(), "commit": commit(), "pairs": {}, "repairs": 0, "contrast": 0,
               "left_out": 0, "attempts_left_out": []}
    with (dest / "repairs.jsonl").open("w", encoding="utf-8") as rep, \
            (dest / "contrast.jsonl").open("w", encoding="utf-8") as con:
        for folder in items("verified"):
            if only and folder.name not in only:
                continue
            if progress:
                progress(folder.name)
            request = (folder / "task.txt").read_text(encoding="utf-8").strip()
            spec = (folder / "spec.vcl").read_text(encoding="utf-8")
            repairs, contrast, crashed = pair_examples(folder.name, request, spec, today)
            tries, skipped = attempt_examples(folder.name, request, spec, today)
            for row in repairs + tries:
                rep.write(json.dumps(row, sort_keys=True) + "\n")
            for row in contrast:
                con.write(json.dumps(row, sort_keys=True) + "\n")
            summary["pairs"][folder.name] = {"mutant_repairs": len(repairs), "attempt_repairs": len(tries),
                                             "contrast": len(contrast), "left_out": crashed}
            summary["repairs"] += len(repairs) + len(tries)
            summary["contrast"] += len(contrast)
            summary["left_out"] += crashed
            summary["attempts_left_out"] += [f"{folder.name}/{name}: {why}" for name, why in skipped]
    (dest / "summary.json").write_text(json.dumps(summary, indent=1, sort_keys=True) + "\n", encoding="utf-8")
    return summary


def main(argv):
    dest = HERE.parent / "dist" / "datagen"
    today, only = None, None
    if "-o" in argv:
        dest = Path(argv[argv.index("-o") + 1])
    if "--today" in argv:
        today = dt.date.fromisoformat(argv[argv.index("--today") + 1])
    if "--only" in argv:
        only = set(argv[argv.index("--only") + 1].split(","))
    s = generate(dest, today, only, progress=lambda name: print(f"  {name}", file=sys.stderr, flush=True))
    print(f"{s['repairs']} repairs and {s['contrast']} contrast examples from {len(s['pairs'])} pairs "
          f"({s['left_out']} mutants left out, {len(s['attempts_left_out'])} tries left out), in {dest}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
