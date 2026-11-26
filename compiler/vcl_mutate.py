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
from vcl_interp import VclError, proc_name, run_tests  # noqa: E402
from vcl_parse import parse  # noqa: E402
from vcl_types import spec_problems  # noqa: E402

SWAPS = [(" <= ", " < "), (" < ", " <= "), (" >= ", " > "), (" > ", " >= "), (" = ", " <> "),
         (" <> ", " = "), (" AND ", " OR "), (" OR ", " AND "), ("order=asc", "order=desc"),
         ("order=desc", "order=asc"), ("mode=sum", "mode=overwrite"), ("mode=count", "mode=keep"),
         ("mode=keep", "mode=overwrite"), ("mode=overwrite", "mode=keep")]


def untested_lines(text):
    """The OP lines of a BUILD's modMain.Run (0-based): no TEST calls Run, and E14 keeps it to
    CALLing RunHeadless, LET, IF and MSG, so a mutant there can never be caught (D53)."""
    spec, problems = parse(text)
    if problems or not spec or spec.get("kind") != "BUILD":
        return set()

    def walk(nodes):
        for n in nodes:
            yield n["line"] - 1
            yield from walk(n.get("children", []))
            yield from walk(n.get("else", []))

    return {i for task in spec["tasks"] if task and proc_name(task) == "modMain.Run" for i in walk(task["ops"])}


def mutants(text):
    lines = text.splitlines()
    skip = untested_lines(text)
    for i, line in enumerate(lines):
        if not line.startswith("OP ") or i in skip:
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


DOC_KEYS = ("FIELD ", "REPEAT ", "KEEP ")


def doc_mutants(text):
    """A DOCUMENT's deliberate bugs (D115): each number one higher and each quoted text changed, on its FIELD,
    REPEAT and KEEP lines, and each FIELD line's value turned into empty text. (i, what, the whole new line).
    No mutant is the line itself (a FIELD that is already "" emptied), and a REPEAT from=[] keeps its <Name>=<col>
    numbers: with no rows, a column can change nothing (D139)."""
    for i, line in enumerate(text.splitlines()):
        if not line.startswith(DOC_KEYS):
            continue
        key, body = line[:9], line[9:]
        found = []
        empty = re.search(r"\bfrom=\[\]", body) if key.startswith("REPEAT") else None
        for m in re.finditer(r"(?<![\w$.\"])(\d+)(?![\w.])", body):
            if empty and m.start() > empty.start():
                continue
            n = int(m.group(1))
            found.append((f"{n} -> {n + 1}", key + body[:m.start()] + str(n + 1) + body[m.end():]))
        for m in re.finditer(r'"((?:[^"]|"")+)"', body):
            found.append((f'"{m.group(1)}" -> changed', key + body[:m.start()] + '"zz' + m.group(1) + '"' + body[m.end():]))
        if key.startswith("FIELD"):
            name = body.split("=", 1)[0].strip()
            found.append((f"{name} -> empty", f'{key}{name} = ""'))
        if key.startswith("KEEP"):
            found.append(("condition negated", key + body.replace(" IF ", " IF NOT ", 1)))
        for what, new_line in found:
            if new_line != line:
                yield i, what, new_line


def mutate_report(text, snapshot=False, today=None):
    """snapshot=True also counts a mutant as caught when any test's final workbook (values, fills
    and number formats), hidden rows (D42), final mock file system (VCL 0.3) with its file texts
    (D45), LOG and MSG texts (D36) or RESULT differs from the original spec's: the check generated
    VBA will face.
    today pins TODAY for every run (default: the real date)."""
    lines = text.splitlines()
    killed, total, survivors = 0, 0, []
    seen = lambda results: [(r.result, r.state, r.files, r.messages, r.hidden, r.texts, r.document)   # noqa: E731
                            for r in results]
    golden = seen(run_tests(text, today)) if snapshot else None
    spec0, _ = parse(text)
    document = bool(spec0) and spec0.get("kind") == "DOCUMENT"
    found = ((i, what, line) for i, what, line in doc_mutants(text)) if document else \
        ((i, what, None if body is None else "OP" + body) for i, what, body in mutants(text))
    for i, what, new_line in found:
        new_lines = list(lines)
        if new_line is None:
            del new_lines[i]
        else:
            new_lines[i] = new_line
        mutant = "\n".join(new_lines)
        spec, problems = parse(mutant)
        if problems or not spec:
            continue
        if document:   # a DOCUMENT mutant the validator refuses is no bug a TEST must catch
            from validator import validate_text
            if any(p.severity == "error" for p in validate_text(mutant)):
                continue
        elif spec_problems(mutant, spec):   # a BUILD: every TASK, with its SIGs
            continue
        total += 1
        try:
            results = run_tests(mutant, today)
            caught = any(not r.passed for r in results)
            if snapshot and not caught:
                caught = seen(results) != golden
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
