# VBA Forge: project brief for Claude Code

Read this first in every session. It is the standing brief; `docs/VBA_Forge_Project_Roadmap.docx`
has the full plan and `docs/vcl-mvp-plan.md` the current build sequence.

## What this is

VBA Forge is an offline, VBA-only code generator for Excel automation at KGC Engineering Projects
(pressure-vessel fabrication paper trail: BOM, weld log, NDE, NCR, calibration and similar
registers). Claude writes task specs in VCL (VBA Command Language); VBA Forge turns each spec into
VBA and proves the macro works with tests. From March 2027 a fine-tuned local model will generate
the VBA. Until then, the goal is a verified dataset of spec-to-VBA pairs.

## Working policy

- **I don't test, debug or verify anything.** You verify everything automatically (unit tests,
  the `vclc` pipeline, CI) and keep going without waiting for me.
- **One PR per item.** When a PR's CI is green, merge it yourself with a merge commit, then start
  the next item.
- If something needs my decision, pick the more conservative option, record it in
  `DECISIONS.md`, and carry on.
- Never weaken a test, a check or a threshold to get to green.
- Running the generated VBA in Excel is optional, happens only when I choose, and never blocks
  work.

## Working with me

- Call me Boss. I'm learning, so explain in plain words and keep it brief.
- Do one thing at a time and keep changes small.
- When I must do something by hand (for example update the Orchestrator Project on claude.ai),
  give exact clicks and what I should see afterwards.
- Ask only when truly necessary. Otherwise state your assumption and continue.
- When I ask for "baby steps": one action per reply, written as "🍼 Next action" plus
  "Expected:", then wait for my result.

## Hard constraints until March 2027

- My work PC is locked down: no local installs, no Excel Trust Center changes. All code runs in
  the cloud (Claude Code on the web, GitHub Codespaces, GitHub Actions).
- Excel 2016 only. No XLOOKUP, FILTER(), LET(), LAMBDA, dynamic arrays or any other
  Microsoft 365 feature, in VBA or in formulas.
- Synthetic data only. Never put real KGC files, names, figures or customer data in the repo,
  tests or prompts. Excel workbooks are git-ignored on purpose.
- Synthetic fixtures only: every TEST uses a named FIXTURE file. `fixture=target` is banned
  until March 2027, and nothing we generate ever touches real workbooks.
- No local-model work (llama.cpp, fine-tuning, GGUF) until the new PC arrives in March 2027.

## Sources of truth

- `vcl/spec.md`: VCL 0.1 revision 5, the language reference.
- `vcl/orchestrator_prompt.txt`: the claude.ai prompt that writes specs. It contains the same
  reference text (a unit test keeps them identical), so the two always change together.
- `vcl/grammar.lark`: the formal syntax. `vcl/semantics.py`: the reference meaning of the value
  functions, with conformance cases in `tests/conformance_cases.py` and the generated tables in
  `docs/vcl-semantics.md`.
- `DECISIONS.md`: every language decision and every decision you made on your own.
- Any VCL change means: update spec.md (with its changelog) and the prompt together, the grammar,
  `compiler/parser.py` and `compiler/validator.py`, the type checker, the interpreter, the
  compiler and ForgeLib as needed, the conformance cases and `docs/vcl-semantics.md`, their
  tests and `DECISIONS.md`; re-run every check; rebuild `dist/forge_module.bas`; and tell me so I
  can update the Orchestrator Project.
- `docs/VBA_Forge_Project_Roadmap.docx`: phases, gates and design decisions.

## Repo layout

```
CLAUDE.md            this brief
DECISIONS.md         language decisions and Claude's own decisions
README.md            overview, progress log, dataset count
requirements.txt     pinned dependencies (lark, with its hash)
vcl/                 spec.md, orchestrator_prompt.txt, grammar.lark, semantics.py, examples/
specs/               one .vcl file per task (task_01.vcl, ...)
compiler/            vclc.py (the toolchain), validator.py + parser.py, vcl_parse.py (grammar),
                     vcl_types.py (types), vcl_interp.py (reference interpreter),
                     vcl_codegen.py (VBA compiler), vcl_fixtures.py (the one fixture
                     generator), vba_lint.py, vcl_mutate.py, gen_semantics_doc.py
vba/                 ForgeLib.bas, the VBA runtime every generated procedure calls
dist/                vclc build output; only forge_module.bas is committed
templates/           one .bas template per op pattern (later)
harness/             Excel batch runner (deferred: PR #4)
batches/             generated batch modules (later)
data/                verified/, repair/, eval/, results/
library/             approved .bas, .cls, .frm files
tests/               unit tests for the Python tools
docs/                roadmap, plans and design documents
logs/
.github/workflows/   CI
(train/ and model/ arrive in March 2027)
```

## Rules for code in this repo

**Python**
- Python 3.11, standard library first. Add a dependency only when needed and pin it, with its
  hash, in `requirements.txt` (today: `lark==1.3.1`).
- Every script runs from the repo root: `python compiler/<script>.py ...`.
- Checkers exit 0 when everything passes and 1 when they find problems, printing one line per
  problem.

**Specs**
- Specs are inputs, not code to fix. Never edit a spec just to make a check pass; report the
  problem to me instead (an edit I ask for, like the task_06 rewrite, is fine).
- Before every commit that touches `specs/`, `vcl/`, `compiler/`, `vba/` or `tests/`, run:
  ```
  python -m unittest discover tests
  python compiler/validator.py specs/*.vcl
  python compiler/vclc.py all specs/*.vcl -o dist
  python compiler/gen_semantics_doc.py --check
  ```
  and commit `dist/forge_module.bas` if it changed. CI runs the same and fails if it is stale.

**VBA (anything that will run in Excel)**
- `Option Explicit`, Excel 2016 compatible, late-bound `Scripting.Dictionary` (no References).
- Standard error handler that writes to ForgeLog and restores ScreenUpdating and Calculation.
- No Select or Activate. Never delete source rows because data is missing or pending.
- No Kill, RmDir, Shell, registry access, HTTP, API Declare, or SaveAs over an existing file.
- Every loop has a clear bound: before March, Ctrl+Break is the only watchdog.
- Every VCL meaning goes through a ForgeLib helper; `vba_lint.py` must report 0 problems.

## How a spec is verified (no Excel needed)

`python compiler/vclc.py all specs/*.vcl -o dist`:
1. **check**: the validator (spec.md rules, the grammar, the type checker)
2. **test**: every TEST runs in the reference interpreter on fixture data from
   `vcl_fixtures.py`, with the meanings in `vcl/semantics.py`
3. **build**: the compiler writes `forge_module.bas` (ForgeLib plus one procedure per spec),
   `forge_fixtures.bas` (one builder per fixture file), `procedures/` and `expected/` (RESULT,
   ERROR_COUNT and the final workbook state of every TEST), then lints the VBA

`python compiler/vclc.py mutate specs/*.vcl --snapshot` scores how many deliberate bugs each
spec's tests catch.

The Excel batch runner (`harness/ForgeTest.bas`, PR #4) is deferred; nothing depends on it.

## Workflow for every change

1. Make one small change.
2. Verify it: the checks above, for what you touched.
3. Read your own diff.
4. Commit with a clear message that says what changed and why.
5. Push and open a pull request. When CI is green, merge it yourself with a merge commit, then
   start the next item.

## Status and next steps

Phase 0, Foundation (Oct 14 to 31, 2026). Gate 0: all specs validate, type-check, pass their
TESTs in the reference interpreter and compile to lint-clean VBA, in CI.

Done: 10 specs, VCL 0.1 revision 5, the Orchestrator prompt, the parser and validator, and the
MVP toolchain (grammar, type checker, interpreter, compiler, ForgeLib, linter, mutation scorer,
`vclc`), with one fixture generator. Fixtures from the mutation evidence (revision 5:
`stale=<n>`, `+boundary`, `+dupes`, `reldates`, rule 19): snapshot mutation score 96.4%, and
every survivor is an equivalent mutant (D14). PR #4 (Excel pilot) is open with the `deferred`
label.

Known: task_04 exceeds the ~60-line TASK guideline; accepted for now, revisit in March.

Next, in order (`docs/vcl-mvp-plan.md`):
1. VCL 0.2 core as VEPs: SAME, CLEAN_TEXT, SKIP_IF, EXIT_LOOP_IF, WRITE_ROW, MODIFIES with
   automatic UNCHANGED, PARSE_DATE, IS_ERROR, IIF, MOD; then migrate the 10 specs.
2. The rest of revision 3 in the compiler and interpreter (DEDUPE with `cols=[...]`, FILTER,
   SET_FORMULA, sheet and file ops), plus list syntax.

Update this section at the end of every session.
