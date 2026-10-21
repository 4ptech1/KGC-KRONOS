"""The corpus: plain-English requests with the VCL specs the toolchain proved (VEP 0.6 Phase 5, D64).

    python compiler/vclc.py corpus [--mutation] [--report]   check every pair in data/verified and data/eval
    python compiler/vclc.py corpus-add --id c013 --request FILE --spec FILE --registers a,b --family f
                          [--set eval] [--source TEXT] [--equivalent N --note TEXT]   prove a spec, then store it
    python compiler/vclc.py export [-o FILE]                  the dataset as JSON Lines (default dist/dataset.jsonl)

One pair is a folder data/<set>/<id>/ with
    task.txt   the request, in the words Boss would type
    spec.vcl   the proven spec: VCL 0.6, canonical (vclc fmt), with its GOAL lines
    meta.json  what the pair is (below), and the mutation score it was stored with
    paraphrase_N.txt   the same request worded differently (optional)
The VBA is never stored: it is derived and goes stale, so `export` compiles it fresh.

A pair is stored only when the whole toolchain accepts it: valid with no warnings (but those meta.json
accepts), every TEST passes, the build is lint-clean, the layout is canonical, and the TESTs catch at least
90% of the mutants (the plan's "strong tests required"; mutants that cannot change a result are named in
meta.json). `corpus --mutation` runs the mutation again and compares it with what meta.json says.
"""
import datetime
import json
import re
import sys
from collections import Counter
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import validator  # noqa: E402
from parser import parse_spec  # noqa: E402
from vba_lint import lint  # noqa: E402
from vcl_codegen import CompileError, assemble, compile_build, compile_build_procs, compile_text, module_title  # noqa: E402
from vcl_fixtures import fixtures_module  # noqa: E402
from vcl_fmt import FormatError, format_text  # noqa: E402
from vcl_interp import VclError, run_tests  # noqa: E402
from vcl_mutate import mutate_report  # noqa: E402
from vcl_parse import parse as grammar_parse  # noqa: E402

DATA = HERE.parent / "data"
SETS = ("verified", "eval")
REGISTERS = ("bom", "grn", "weld_log", "nde", "ncr", "itp", "welder_quals", "calibration", "consumables",
             "material_certs", "wps_pqr", "dispatch", "job_tracker")
# the plan for 150 tasks (VEP 0.6): 100 basic, 30 with the item 2 ops, 20 BUILDs
QUOTAS = {"flag": 30, "summarise": 25, "reconcile": 20, "clean": 15, "rebuild": 10,
          "files": 12, "sheets_formulas": 10, "find_filter": 8, "build": 20}
FLOOR = 0.90
FILES = {"task.txt", "spec.vcl", "meta.json"}
META_KEYS = {"id", "kind", "registers", "family", "ops", "op_lines", "source", "mutation", "accepted_warnings"}


def items(kind):
    """The pair folders of a set, in name order."""
    root = DATA / kind
    return sorted(p for p in root.iterdir() if p.is_dir()) if root.is_dir() else []


def dump(meta):
    return json.dumps(meta, indent=1, sort_keys=True) + "\n"


def ops_used(text):
    """The op keywords a spec uses and the number of steps, over a TASK or every TASK of a BUILD."""
    spec = parse_spec(text)
    names, steps = set(), 0
    for part in [spec] + list(spec.tasks):
        for node in part.all_ops():
            steps += 1
            for found in re.finditer(r"(?:^|\bTHEN )([A-Z][A-Z_]+)", node.text):
                names.add(found.group(1))
    names -= {"AND", "OR", "NOT", "THEN"}
    return sorted(names), steps


def build_problems(text, name):
    """Compile and lint a spec and its fixture builders; the problems found."""
    spec, _ = grammar_parse(text)
    try:
        if spec["kind"] == "BUILD":
            _, module = compile_build(text, name)
        else:
            module = assemble([compile_text(text, name)], module_title(1))
        found = list(lint(module))
        trees = [k["tree"] for k in spec["keys"] if k["key"] == "FIXTURE"]
        found += lint(fixtures_module([(name, trees)]))
    except (CompileError, ValueError) as err:
        return [f"does not build: {err}"]
    return [f"VBA lint: {line}" for line in found]


def check_item(folder, mutation=False):
    """The problems with one pair (an empty list when it is sound)."""
    name = folder.name
    files = {p.name for p in folder.iterdir()}
    problems = [f"{name}: missing {f}" for f in sorted(FILES - files)]
    extra = sorted(f for f in files - FILES if not re.fullmatch(r"paraphrase_\d+\.txt", f))
    problems += [f"{name}: unexpected file {f}" for f in extra]
    if problems:
        return problems
    request = (folder / "task.txt").read_text(encoding="utf-8")
    spec_text = (folder / "spec.vcl").read_text(encoding="utf-8")
    meta_text = (folder / "meta.json").read_text(encoding="utf-8")
    if not (20 <= len(request.strip()) <= 1500) or any(ord(c) > 126 or (ord(c) < 32 and c not in "\n\t") for c in request):
        problems.append(f"{name}: task.txt must be 20 to 1500 characters of plain printable ASCII")
    try:
        meta = json.loads(meta_text)
    except ValueError as err:
        return problems + [f"{name}: meta.json is not JSON ({err})"]
    if dump(meta) != meta_text:
        problems.append(f"{name}: meta.json is not in its canonical form (sorted keys, 1-space indent)")
    if set(meta) != META_KEYS:
        return problems + [f"{name}: meta.json keys must be {sorted(META_KEYS)}, got {sorted(meta)}"]
    try:
        canonical = format_text(spec_text)
    except FormatError as err:
        return problems + [f"{name}: spec.vcl can't be formatted: {err.problems[0]}"]
    if canonical != spec_text:
        problems.append(f"{name}: spec.vcl is not canonical (run vclc fmt --write)")
    found = validator.validate_text(spec_text)
    errors = [p for p in found if p.severity == "error"]
    problems += [f"{name}: spec.vcl line {p.line}: {p.message} [{p.code}]" for p in errors]
    accepted = set(meta["accepted_warnings"])
    problems += [f"{name}: spec.vcl line {p.line}: warning {p.message} [{p.code}] (not in accepted_warnings)"
                 for p in found if p.severity == "warning" and p.code not in accepted]
    if errors:
        return problems
    if not spec_text.startswith("VCL 0.6\n"):
        problems.append(f"{name}: spec.vcl must be VCL 0.6")
    try:
        results = run_tests(spec_text)
    except VclError as err:
        return problems + [f"{name}: the TESTs can't run: {err}"]
    problems += [f"{name}: TEST on line {r.line} fails: {r.failed[0] if r.failed else r.result}" for r in results if not r.passed]
    problems += [f"{name}: {line}" for line in build_problems(spec_text, name + ".vcl")]
    spec = parse_spec(spec_text)
    ops, steps = ops_used(spec_text)
    if meta["id"] != name:
        problems.append(f"{name}: meta.json id is {meta['id']!r}")
    if meta["kind"] != spec.kind:
        problems.append(f"{name}: meta.json kind is {meta['kind']!r}, the spec is a {spec.kind}")
    if meta["ops"] != ops or meta["op_lines"] != steps:
        problems.append(f"{name}: meta.json ops or op_lines are stale (the spec uses {len(ops)} ops in {steps} steps)")
    if not meta["registers"] or any(r not in REGISTERS for r in meta["registers"]):
        problems.append(f"{name}: registers must be some of {', '.join(REGISTERS)}")
    if meta["family"] not in QUOTAS:
        problems.append(f"{name}: family must be one of {', '.join(QUOTAS)}")
    elif (meta["family"] == "build") != (spec.kind == "BUILD"):
        problems.append(f"{name}: family build and kind BUILD go together")
    mutants = meta["mutation"]
    if set(mutants) != {"killed", "total", "equivalent", "note"}:
        problems.append(f"{name}: mutation must hold killed, total, equivalent and note")
    else:
        live = mutants["total"] - mutants["equivalent"]
        if not 0 <= mutants["killed"] <= live:
            problems.append(f"{name}: mutation killed {mutants['killed']} is outside 0 to {live}")
        elif live and mutants["killed"] / live < FLOOR:
            problems.append(f"{name}: the TESTs catch {mutants['killed']} of {live} mutants, under the {FLOOR:.0%} floor")
        if mutants["equivalent"] and not mutants["note"]:
            problems.append(f"{name}: mutation equivalent {mutants['equivalent']} needs a note saying why")
        if mutation:
            killed, total, _ = mutate_report(spec_text, snapshot=True)
            if (killed, total) != (mutants["killed"], mutants["total"]):
                problems.append(f"{name}: mutation is {killed} of {total} now, meta.json says "
                                f"{mutants['killed']} of {mutants['total']}")
    return problems


def check_all(mutation=False):
    """Every pair of every set, and that no id is in both."""
    problems, seen = [], {}
    for kind in SETS:
        for folder in items(kind):
            if folder.name in seen:
                problems.append(f"{folder.name}: in both data/{seen[folder.name]} and data/{kind}")
            seen[folder.name] = kind
            problems += check_item(folder, mutation)
    return problems


def load(kind="verified"):
    """The pairs of a set as dicts, for the report and the export."""
    out = []
    for folder in items(kind):
        meta = json.loads((folder / "meta.json").read_text(encoding="utf-8"))
        meta["request"] = (folder / "task.txt").read_text(encoding="utf-8").strip()
        meta["spec"] = (folder / "spec.vcl").read_text(encoding="utf-8")
        meta["paraphrases"] = [p.read_text(encoding="utf-8").strip() for p in sorted(folder.glob("paraphrase_*.txt"))]
        out.append(meta)
    return out


def goal_of(spec_text):
    return next(line.split(None, 1)[1] for line in spec_text.split("\n") if line.startswith("GOAL "))


def vba_of(spec_text, name):
    """The VBA the compiler writes for a spec (a BUILD's procedures, without ForgeLib)."""
    spec = parse_spec(spec_text)
    if spec.kind == "BUILD":
        return "\n".join(compile_build_procs(spec_text, name)[1])
    return compile_text(spec_text, name)


def export(dest):
    """The dataset as JSON Lines: request, GOAL, spec, freshly compiled VBA, paraphrases."""
    pairs = load("verified")
    dest = Path(dest)
    dest.parent.mkdir(parents=True, exist_ok=True)
    with dest.open("w", encoding="utf-8") as out:
        for pair in pairs:
            out.write(json.dumps({"id": pair["id"], "kind": pair["kind"], "registers": pair["registers"],
                                  "family": pair["family"], "request": pair["request"], "goal": goal_of(pair["spec"]),
                                  "spec": pair["spec"], "vba": vba_of(pair["spec"], pair["id"] + ".vcl"),
                                  "paraphrases": pair["paraphrases"]}, sort_keys=True) + "\n")
    return len(pairs)


def report():
    pairs = load("verified")
    held = load("eval")
    lines = [f"{len(pairs)} pairs in data/verified, {len(held)} in data/eval (held out)", ""]
    kinds = Counter(p["kind"] for p in pairs)
    lines.append("Families (have of planned):")
    have = Counter(p["family"] for p in pairs + held)
    for family, quota in QUOTAS.items():
        lines.append(f"  {family:16} {have[family]:3} of {quota}")
    lines.append(f"  {'all':16} {sum(have.values()):3} of {sum(QUOTAS.values())}   ({kinds['TASK']} TASKs, {kinds['BUILD']} BUILDs)")
    lines.append("")
    regs = Counter(r for p in pairs + held for r in p["registers"])
    lines.append("Registers: " + ", ".join(f"{r} {regs[r]}" for r in REGISTERS))
    ops = Counter(o for p in pairs + held for o in p["ops"])
    lines.append("Ops used: " + ", ".join(f"{o} {n}" for o, n in sorted(ops.items())))
    live = sum(p["mutation"]["total"] - p["mutation"]["equivalent"] for p in pairs + held)
    killed = sum(p["mutation"]["killed"] for p in pairs + held)
    if live:
        lines.append(f"Mutants caught: {killed} of {live} that can be caught ({100 * killed / live:.1f}%)")
    return "\n".join(lines)


def add_item(item_id, request, spec_text, registers, family, kind_set="verified", source="", equivalent=0, note="",
             accepted_warnings=()):
    """Prove a spec and store it as data/<set>/<id>/. Returns the problems; nothing is stored when there are any."""
    folder = DATA / kind_set / item_id
    if folder.exists():
        return [f"{item_id}: already there"]
    try:
        spec_text = format_text(spec_text)
    except FormatError as err:
        return [f"{item_id}: the spec can't be formatted: {err.problems[0]}"]
    spec = parse_spec(spec_text)
    ops, steps = ops_used(spec_text)
    killed, total, survivors = mutate_report(spec_text, snapshot=True)
    meta = {"id": item_id, "kind": spec.kind, "registers": list(registers), "family": family, "ops": ops,
            "op_lines": steps, "source": source, "accepted_warnings": sorted(accepted_warnings),
            "mutation": {"killed": killed, "total": total, "equivalent": equivalent, "note": note}}
    folder.mkdir(parents=True)
    (folder / "task.txt").write_text(request.strip() + "\n", encoding="utf-8")
    (folder / "spec.vcl").write_text(spec_text, encoding="utf-8")
    (folder / "meta.json").write_text(dump(meta), encoding="utf-8")
    problems = check_item(folder)
    if problems:
        for p in folder.iterdir():
            p.unlink()
        folder.rmdir()
        if survivors and any("floor" in line for line in problems):
            problems += [f"{item_id}: survives: line {line}: {what}" for line, what in survivors[:6]]
    return problems


def option(args, name, default=None):
    return args[args.index(name) + 1] if name in args and args.index(name) + 1 < len(args) else default


def main(argv):
    command, args = argv[0], argv[1:]
    if command == "corpus":
        problems = check_all(mutation="--mutation" in args)
        for line in problems:
            print(line)
        if "--report" in args or not problems:
            print(report())
        print(f"corpus: {len(problems)} problem{'s' if len(problems) != 1 else ''}" if problems else "corpus: every pair is sound")
        return 1 if problems else 0
    if command == "corpus-add":
        try:
            request = Path(option(args, "--request")).read_text(encoding="utf-8")
            spec_text = Path(option(args, "--spec")).read_text(encoding="utf-8")
            item_id, family = option(args, "--id"), option(args, "--family")
            registers = option(args, "--registers", "").split(",")
        except (TypeError, OSError):
            print("usage: corpus-add --id ID --request FILE --spec FILE --registers a,b --family F [--set eval]")
            return 1
        problems = add_item(item_id, request, spec_text, registers, family, kind_set=option(args, "--set", "verified"),
                            source=option(args, "--source", "orchestrator helper"),
                            equivalent=int(option(args, "--equivalent", 0)), note=option(args, "--note", ""),
                            accepted_warnings=[w for w in option(args, "--accept", "").split(",") if w])
        for line in problems:
            print(line)
        print(f"stored data/{option(args, '--set', 'verified')}/{item_id}" if not problems else f"not stored: {len(problems)} problems")
        return 1 if problems else 0
    if command == "export":
        count = export(option(args, "-o", str(HERE.parent / "dist" / "dataset.jsonl")))
        print(f"exported {count} pairs")
        return 0
    return 1
