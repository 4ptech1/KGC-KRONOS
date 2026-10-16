"""vclc: the VBA Forge toolchain (MVP).

  python compiler/vclc.py check  FILES...            parse + type check
  python compiler/vclc.py test   FILES...            run every TEST in the reference interpreter
  python compiler/vclc.py build  FILES... [-o DIR]   compile, assemble one paste-ready module, lint it,
                                                     and write expected results for later Excel checks
  python compiler/vclc.py mutate FILES... [--snapshot]   mutation score of each spec's tests
  python compiler/vclc.py all    FILES... [-o DIR]   check + test + build, stop at the first failure

Exit code 0 when everything passes, 1 otherwise. Nothing here needs Excel.
"""
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from vba_lint import lint  # noqa: E402
from vcl_codegen import CompileError, assemble, compile_text  # noqa: E402
from vcl_interp import VclError, run_tests  # noqa: E402
from vcl_mutate import mutate_report  # noqa: E402
from vcl_parse import parse  # noqa: E402
from vcl_types import Checker  # noqa: E402


def cmd_check(files):
    bad = 0
    for p in files:
        spec, problems = parse(Path(p).read_text(encoding="utf-8"))
        if not problems and spec and spec["kind"] == "TASK":
            problems = Checker(spec).run()
        bad += bool(problems)
        print(f"{'OK  ' if not problems else 'FAIL'} check  {p}")
        for line, msg in problems:
            print(f"       line {line}: {msg}")
    return bad == 0


def cmd_test(files):
    bad = 0
    for p in files:
        try:
            res = run_tests(Path(p).read_text(encoding="utf-8"))
        except VclError as e:
            print(f"FAIL test   {p}: {e}")
            bad += 1
            continue
        failed = [r for r in res if not r.passed]
        bad += bool(failed)
        print(f"{'OK  ' if not failed else 'FAIL'} test   {p}  ({len(res) - len(failed)}/{len(res)} tests)")
        for r in failed:
            print(f"       line {r.line}: RESULT={r.result} ERROR_COUNT={r.errors}")
    return bad == 0


def cmd_build(files, out_dir):
    out = Path(out_dir)
    (out / "procedures").mkdir(parents=True, exist_ok=True)
    (out / "expected").mkdir(parents=True, exist_ok=True)
    procs, ok = [], True
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
        (out / "procedures" / (path.stem + ".bas")).write_text(vba, encoding="utf-8")
        expected = [{"test_line": r.line, "passed": r.passed, "result": r.result, "error_count": r.errors,
                     "final_state": {sheet: {"cells": [[rc[0], rc[1], v] for rc, v in cells],
                                             "fills": [[rc[0], rc[1], f] for rc, f in fills]}
                                     for sheet, (cells, fills) in r.state.items()}}
                    for r in run_tests(text)]
        (out / "expected" / (path.stem + ".json")).write_text(json.dumps(expected, indent=1, default=str), encoding="utf-8")
        print(f"OK   build  {p}")
    if not procs:
        return False
    module = assemble(procs, f"VBA Forge module: {len(procs)} procedures")
    found = lint(module)
    (out / "forge_module.bas").write_text(module, encoding="utf-8")
    print(f"{'OK  ' if not found else 'FAIL'} lint   {out / 'forge_module.bas'}  ({len(module.splitlines())} lines)")
    for line, msg in found:
        print(f"       line {line}: {msg}")
    return ok and not found


def cmd_mutate(files, snapshot):
    tk = tt = 0
    for p in files:
        k, t, surv = mutate_report(Path(p).read_text(encoding="utf-8"), snapshot=snapshot)
        tk, tt = tk + k, tt + t
        print(f"{100 * k / t if t else 100:5.1f}%  mutate {p}  ({k}/{t})")
        for line, what in surv[:4]:
            print(f"        survives: line {line}: {what}")
    if tt:
        print(f"Overall mutation score: {100 * tk / tt:.1f}%")
    return True


def main(argv):
    if len(argv) < 2 or argv[1] not in ("check", "test", "build", "mutate", "all"):
        print(__doc__)
        return 1
    cmd = argv[1]
    out_dir = "dist"
    files = []
    args = argv[2:]
    i = 0
    while i < len(args):
        if args[i] == "-o":
            out_dir = args[i + 1]
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
        return 0 if cmd_test(files) else 1
    if cmd == "build":
        return 0 if cmd_build(files, out_dir) else 1
    if cmd == "mutate":
        return 0 if cmd_mutate(files, "--snapshot" in args) else 1
    ok = cmd_check(files) and cmd_test(files) and cmd_build(files, out_dir)
    print("\nALL STAGES PASSED" if ok else "\nSTOPPED: see the FAIL lines above")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
