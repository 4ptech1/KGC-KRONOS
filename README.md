# VBA Forge

VBA Forge is an offline, VBA-only code generator for Excel automation in pressure-vessel
fabrication registers (BOM, weld log, NDE, NCR, calibration and similar). Claude writes each task
as a spec in VCL (VBA Command Language), and VBA Forge turns the spec into VBA and proves the
macro works with tests. The verified spec-to-VBA pairs become the dataset for a fine-tuned local
model that takes over generation from March 2027.

## Current phase

**Phase 0, Foundation (Oct 14 to 31, 2026).** Everything runs in the cloud (Claude Code on the
web, GitHub Codespaces, GitHub Actions) until the new PC arrives in March 2027. Every check is
automated; running the generated VBA in Excel is optional and never blocks work.

See `CLAUDE.md` for the working rules, `docs/vcl-mvp-plan.md` for the build sequence,
`DECISIONS.md` for the decisions so far and `docs/VBA_Forge_Project_Roadmap.docx` for the full
plan.

## Check, test and build the specs

From the repo root, with Python 3.11:

```
python -m pip install --require-hashes -r requirements.txt    # once: lark, pinned
python -m unittest discover tests
python compiler/validator.py specs/*.vcl
python compiler/vclc.py all specs/*.vcl -o dist
python compiler/vclc.py mutate specs/*.vcl --snapshot          # how strong each spec's tests are
```

The validator checks the spec.md rules, the formal grammar (`vcl/grammar.lark`) and the types
(every calculation and date comparison guarded). It prints one line per problem, as
`path:line: error|warning: message`, and ends with `N files checked: N errors, N warnings`. It
exits 1 if there is any error or no files were given; warnings alone exit 0.

`vclc all` runs the validator, then every TEST in the reference interpreter, then builds:

- `dist/forge_module.bas`: one paste-ready module, the ForgeLib runtime plus one procedure per
  spec. It is committed, so you can copy it any time; CI fails if it is out of date.
- `dist/forge_fixtures.bas`: one `ForgeFixture_<file>(wb)` builder per fixture file, with exactly
  the data the interpreter tested on.
- `dist/expected/*.json`: for every TEST, the expected RESULT, ERROR_COUNT and final workbook
  state, so any future Excel run can be compared automatically.

GitHub Actions runs all of this on every push and pull request.

## Dataset count

Dataset count: 0 verified pairs

## Progress log

- **2026-10-14:** Repo set up. Added the 10 task specs (VCL 0.1 rev 2), `vcl/spec.md`, the
  Orchestrator prompt, the starter checker and the roadmap; all 10 specs pass the checker. Added
  the project folders, `.gitignore` and the spec-check CI workflow.
- **2026-10-14:** VCL 0.1 revision 3: synthetic fixtures only (`fixture=target` removed),
  COUNTIFS/SUMIFS assertions, per-column fixture seeds, and exact goal TESTs for task_07 and
  task_08.
- **2026-10-14:** `compiler/parser.py` and `compiler/validator.py` replace the starter checker,
  with unit tests in `tests/`. All 10 specs validate with 0 errors (task_04 gets a size
  warning).
- **2026-10-16:** MVP toolchain integrated: grammar, type checker, reference interpreter, VBA
  compiler, ForgeLib runtime, VBA linter, mutation scorer and the `vclc` CLI. One fixture
  generator for everything. VCL 0.1 revision 4 records the M2 decisions (`DECISIONS.md`);
  task_06 rewritten so its date check is provably guarded. The validator also runs the grammar
  and the type checker. All 10 specs check, pass their 39 TESTs and compile to lint-clean VBA.
- **2026-10-17:** VCL 0.1 revision 5: fixtures that catch wrong macros. New FIXTURE parts
  `stale=<n>`, `+boundary(...)`, `+dupes(n)` and `reldates(lo, hi)`, and rule 19 (validator
  warnings). All 10 specs' fixtures and goal TESTs upgraded (no OP line changed); the
  whole-workbook mutation score rose from 84% to 96.4%, and every surviving mutant is
  equivalent.
