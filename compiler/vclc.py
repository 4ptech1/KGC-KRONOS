"""vclc: the VBA Forge toolchain.

  python compiler/vclc.py check  FILES...            the validator: spec.md rules, grammar, type check
  python compiler/vclc.py test   FILES...            run every TEST in the reference interpreter
  python compiler/vclc.py build  FILES... [-o DIR]   compile, assemble, lint, write expected results
  python compiler/vclc.py mutate FILES... [--snapshot]   mutation score of each spec's tests
  python compiler/vclc.py all    FILES... [-o DIR]   check + test + build, stop at the first failure
  python compiler/vclc.py pilot-check RESULTS FILES...  compare an Excel pilot run with the interpreter

Options: -o DIR (default dist) and --today YYYY-MM-DD, which pins TODAY for test, build and
mutate runs (default: the real date).

build writes into DIR:
  forge_module.bas      one paste-ready module: the ForgeLib runtime plus one procedure per spec
  forge_selftest.bas    the ForgeLib self-test: every conformance case, run through ForgeLib in
                        Excel (vcl_selftest.py)
  forge_pilot.bas       the Excel pilot of the specs in vcl_pilot.PILOT_SPECS: runs their TESTs
                        and writes lines to paste back for pilot-check (vcl_pilot.py)
  forge_fixtures.bas    the fixture builders: ForgeFixture_<file>(wb) rebuilds each synthetic
                        test workbook with exactly the data the interpreter tested on, and
                        ForgeSandbox_<file>() builds its VCL 0.3 folders in a new %TEMP% sandbox
  procedures/*.bas      each procedure on its own
  expected/*.json       for every TEST: RESULT, ERROR_COUNT and the final workbook state (values,
                        fills, and the number formats set by FORMAT), and for VCL 0.3 folder
                        FIXTUREs the final mock file system

Exit code 0 when everything passes, 1 otherwise. Nothing here needs Excel.
"""
import datetime
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from vba_lint import lint, public_helpers  # noqa: E402
from vcl_codegen import FORGELIB, CompileError, assemble, compile_text, module_title  # noqa: E402
from vcl_fixtures import fixtures_module  # noqa: E402
from vcl_interp import VclError, run_tests  # noqa: E402
from vcl_mutate import mutate_report  # noqa: E402
from vcl_pilot import PILOT_SPECS, check as pilot_check, pilot_module  # noqa: E402
from vcl_selftest import selftest_module  # noqa: E402
from vcl_parse import parse  # noqa: E402
from validator import validate_file  # noqa: E402


def cmd_check(files):
    bad = 0
    for p in files:
        problems = validate_file(p)
        errors = [x for x in problems if x.severity == "error"]
        bad += bool(errors)
        print(f"{'OK  ' if not errors else 'FAIL'} check  {p}")
        for x in problems:
            print(f"       line {x.line}: {x.severity}: {x.message}")
    return bad == 0


def cmd_test(files, today):
    bad = 0
    for p in files:
        try:
            res = run_tests(Path(p).read_text(encoding="utf-8"), today)
        except VclError as e:
            print(f"FAIL test   {p}: {e}")
            bad += 1
            continue
        failed = [r for r in res if not r.passed]
        bad += bool(failed)
        print(f"{'OK  ' if not failed else 'FAIL'} test   {p}  ({len(res) - len(failed)}/{len(res)} tests)")
        for r in failed:
            print(f"       line {r.line}: RESULT={r.result} ERROR_COUNT={r.errors}")
            for problem in r.failed:   # a failed assertion with what it measured, or an UNCHANGED difference
                print(f"         {problem}")
    return bad == 0


def lint_report(path, module, helpers=()):
    found = lint(module, helpers)
    path.write_text(module, encoding="utf-8")
    print(f"{'OK  ' if not found else 'FAIL'} lint   {path}  ({len(module.splitlines())} lines)")
    for line, msg in found:
        print(f"       line {line}: {msg}")
    return not found


def cmd_build(files, out_dir, today):
    out = Path(out_dir)
    (out / "procedures").mkdir(parents=True, exist_ok=True)
    (out / "expected").mkdir(parents=True, exist_ok=True)
    procs, fixture_sets, ok = [], [], True
    run_day = today or datetime.date.today()
    for p in files:
        path = Path(p)
        text = path.read_text(encoding="utf-8")
        try:
            vba = compile_text(text, path.name)
        except CompileError as e:
            print(f"FAIL build  {p}: {e}")
            ok = False
            continue
        procs.append(vba)
        spec, _ = parse(text)
        fixture_sets.append((path.name, [k["tree"] for k in spec["keys"] if k["key"] == "FIXTURE"]))
        (out / "procedures" / (path.stem + ".bas")).write_text(vba, encoding="utf-8")
        expected = {"spec": path.name, "today": run_day.isoformat(),
                    "tests": [{"test_line": r.line, "passed": r.passed, "result": r.result,
                               "error_count": r.errors,
                               "final_state": {sheet: {"cells": [[rc[0], rc[1], v] for rc, v in cells],
                                                       "fills": [[rc[0], rc[1], f] for rc, f in fills],
                                                       "number_formats": [[rc[0], rc[1], f] for rc, f in formats],
                                                       **({"hidden_rows": r.hidden[sheet]} if sheet in r.hidden else {})}
                                               for sheet, (cells, fills, formats) in r.state.items()},
                               **({"final_files": [list(f) for f in r.files]} if r.files else {}),
                               **({"final_texts": [list(f) for f in r.texts]} if r.texts else {})}
                              for r in run_tests(text, run_day)]}
        (out / "expected" / (path.stem + ".json")).write_text(json.dumps(expected, indent=1, default=str),
                                                               encoding="utf-8")
        print(f"OK   build  {p}")
    if not procs:
        return False
    ok = lint_report(out / "forge_module.bas", assemble(procs, module_title(len(procs)))) and ok
    forgelib_helpers = public_helpers(FORGELIB.read_text(encoding="utf-8"))
    ok = lint_report(out / "forge_selftest.bas", selftest_module(), forgelib_helpers) and ok
    piloted = [(Path(p).name, Path(p).read_text(encoding="utf-8")) for p in files if Path(p).name in PILOT_SPECS]
    if piloted:
        ok = lint_report(out / "forge_pilot.bas", pilot_module(piloted), forgelib_helpers) and ok
    try:
        fixtures = fixtures_module(fixture_sets)
    except ValueError as e:
        print(f"FAIL build  fixtures: {e}")
        return False
    return lint_report(out / "forge_fixtures.bas", fixtures) and ok


def cmd_mutate(files, snapshot, today):
    tk = tt = 0
    for p in files:
        k, t, surv = mutate_report(Path(p).read_text(encoding="utf-8"), snapshot=snapshot, today=today)
        tk, tt = tk + k, tt + t
        print(f"{100 * k / t if t else 100:5.1f}%  mutate {p}  ({k}/{t})")
        for line, what in surv[:4]:
            print(f"        survives: line {line}: {what}")
        if len(surv) > 4:
            print(f"        ... and {len(surv) - 4} more")
    if tt:
        print(f"Overall mutation score: {100 * tk / tt:.1f}%")
    return True


def main(argv):
    if len(argv) < 2 or argv[1] not in ("check", "test", "build", "mutate", "all", "pilot-check"):
        print(__doc__)
        return 1
    cmd = argv[1]
    if cmd == "pilot-check":
        if len(argv) < 4:
            print("pilot-check needs the pasted results file and the piloted spec files")
            return 1
        pasted = Path(argv[2]).read_text(encoding="utf-8").splitlines()
        problems = pilot_check(pasted, [(Path(p).name, Path(p).read_text(encoding="utf-8")) for p in argv[3:]])
        for problem in problems:
            print(problem)
        print("PILOT OK: Excel and the interpreter agree" if not problems else f"PILOT: {len(problems)} differences")
        return 0 if not problems else 1
    out_dir = "dist"
    today = None
    files = []
    args = argv[2:]
    i = 0
    while i < len(args):
        if args[i] in ("-o", "--today"):
            if i + 1 >= len(args):
                print(f"{args[i]} needs a value")
                return 1
            if args[i] == "-o":
                out_dir = args[i + 1]
            else:
                try:
                    today = datetime.date.fromisoformat(args[i + 1])
                except ValueError:
                    print(f"--today needs a date like 2026-10-16, got {args[i + 1]!r}")
                    return 1
            i += 2
            continue
        if not args[i].startswith("--"):
            files.append(args[i])
        i += 1
    if not files:
        print("no spec files given")
        return 1
    if cmd == "check":
        return 0 if cmd_check(files) else 1
    if cmd == "test":
        return 0 if cmd_test(files, today) else 1
    if cmd == "build":
        return 0 if cmd_build(files, out_dir, today) else 1
    if cmd == "mutate":
        return 0 if cmd_mutate(files, "--snapshot" in args, today) else 1
    ok = cmd_check(files) and cmd_test(files, today) and cmd_build(files, out_dir, today)
    print("\nALL STAGES PASSED" if ok else "\nSTOPPED: see the FAIL lines above")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
