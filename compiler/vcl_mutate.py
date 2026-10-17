"""Mutation scoring: how many small, plausible bugs do a spec's TESTs catch?

mutate_report(text) -> (killed, total, survivors). Each mutant changes one OP line: a flipped
comparison, AND/OR swap, number +1, changed text literal, deleted step, reversed sort, sum
changed to overwrite, a reversed loop, or SKIP_IF and EXIT_LOOP_IF swapped. Mutants that no longer parse or type-check are
discarded. A surviving mutant means the TESTs can't tell that behaviour apart from the real one:
add a test case or fixture rows that tell them apart.
"""
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from vcl_interp import VclError, run_tests  # noqa: E402
from vcl_parse import parse  # noqa: E402
from vcl_types import Checker  # noqa: E402

SWAPS = [(" <= ", " < "), (" < ", " <= "), (" >= ", " > "), (" > ", " >= "), (" = ", " <> "),
         (" <> ", " = "), (" AND ", " OR "), (" OR ", " AND "), ("order=asc", "order=desc"),
         ("order=desc", "order=asc"), ("mode=sum", "mode=overwrite"), ("mode=count", "mode=keep"),
         ("mode=keep", "mode=overwrite"), ("mode=overwrite", "mode=keep")]


def mutants(text):
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if not line.startswith("OP "):
            continue
        body = line[2:]
        for old, new in SWAPS:
            start = 0
            while True:
                j = body.find(old, start)
                if j < 0:
                    break
                yield i, f"'{old.strip()}' -> '{new.strip()}'", body[:j] + new + body[j + len(old):]
                start = j + len(old)
        for m in re.finditer(r"(?<![\w$.])(\d+)(?![\w.])", body):
            n = int(m.group(1))
            yield i, f"{n} -> {n + 1}", body[:m.start()] + str(n + 1) + body[m.end():]
        for m in re.finditer(r'"((?:[^"]|"")+)"', body):
            yield i, f'"{m.group(1)}" -> changed', body[:m.start()] + '"zz' + m.group(1) + '"' + body[m.end():]
        stripped = body.strip()
        loop_jump = stripped.startswith(("SKIP_IF ", "EXIT_LOOP_IF "))
        if (re.match(r"^[A-Z_]+\(", stripped) or loop_jump) and not stripped.startswith(("RETURN(",)):
            yield i, "step deleted", None
        for old, new in (("SKIP_IF ", "EXIT_LOOP_IF "), ("EXIT_LOOP_IF ", "SKIP_IF ")):
            if stripped.startswith(old):
                yield i, f"'{old.strip()}' -> '{new.strip()}'", body.replace(old, new, 1)
        if "dir=down" in body:
            yield i, "loop reversed", body.replace("dir=down", "dir=up")


def mutate_report(text, snapshot=False, today=None):
    """snapshot=True also counts a mutant as caught when any test's final workbook (values, fills
    and number formats) or RESULT differs from the original spec's: the check generated VBA will face.
    today pins TODAY for every run (default: the real date)."""
    lines = text.splitlines()
    killed, total, survivors = 0, 0, []
    golden = [(r.result, r.state) for r in run_tests(text, today)] if snapshot else None
    for i, what, new_body in mutants(text):
        new_lines = list(lines)
        if new_body is None:
            del new_lines[i]
        else:
            new_lines[i] = "OP" + new_body
        mutant = "\n".join(new_lines)
        spec, problems = parse(mutant)
        if problems or not spec or Checker(spec).run():
            continue
        total += 1
        try:
            results = run_tests(mutant, today)
            caught = any(not r.passed for r in results)
            if snapshot and not caught:
                caught = [(r.result, r.state) for r in results] != golden
        except (VclError, KeyError, TypeError, ValueError, IndexError):
            caught = True
        if caught:
            killed += 1
        else:
            survivors.append((i + 1, what))
    return killed, total, survivors


if __name__ == "__main__":
    snap = "--snapshot" in sys.argv
    files = [a for a in sys.argv[1:] if not a.startswith("--")]
    grand_k = grand_t = 0
    for p in files:
        k, t, surv = mutate_report(Path(p).read_text(encoding="utf-8"), snapshot=snap)
        grand_k += k
        grand_t += t
        score = 100 * k / t if t else 100
        print(f"{score:5.1f}%  {p}  ({k}/{t} mutants caught)")
        for line, what in surv[:6]:
            print(f"        survives: line {line}: {what}")
        if len(surv) > 6:
            print(f"        ... and {len(surv) - 6} more")
    if grand_t:
        print(f"\nOverall mutation score: {100 * grand_k / grand_t:.1f}% ({grand_k}/{grand_t})")
