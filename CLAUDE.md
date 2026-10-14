# VBA Forge: project brief for Claude Code

Read this first in every session. It is the standing brief; `docs/VBA_Forge_Project_Roadmap.docx`
has the full plan.

## What this is

VBA Forge is an offline, VBA-only code generator for Excel automation at KGC Engineering Projects
(pressure-vessel fabrication paper trail: BOM, weld log, NDE, NCR, calibration and similar
registers). Claude writes task specs in VCL (VBA Command Language); VBA Forge turns each spec into
VBA and proves the macro works with tests. From March 2027 a fine-tuned local model will generate
the VBA. Until then, the goal is a verified dataset of spec-to-VBA pairs.

## Working with me

- Call me Boss. I'm learning, so explain in plain words and keep it brief.
- Do one thing at a time and keep changes small.
- When I must do something by hand (Excel, GitHub website), give exact clicks and what I should
  see afterwards.
- Ask only when truly necessary. Otherwise state your assumption and continue.
- When I ask for "baby steps": one action per reply, written as "🍼 Next action" plus
  "Expected:", then wait for my result.

## Hard constraints until March 2027

- My work PC is locked down: no local installs, no Excel Trust Center changes. All code runs in
  the cloud (Claude Code on the web, GitHub Codespaces, GitHub Actions).
- Excel testing is manual. I paste one generated batch module (`ForgeTest_batch_NNN.bas`) into
  Excel, run `RunBatch`, and paste the Results sheet back as CSV.
- Excel 2016 only. No XLOOKUP, FILTER(), LET(), LAMBDA, dynamic arrays or any other
  Microsoft 365 feature, in VBA or in formulas.
- Synthetic data only. Never put real KGC files, names, figures or customer data in the repo,
  tests or prompts. Excel workbooks are git-ignored on purpose.
- Synthetic fixtures only: every TEST uses a named FIXTURE file. `fixture=target` is banned
  until March 2027, and the batch runner never touches real workbooks.
- No local-model work (llama.cpp, fine-tuning, GGUF) until the new PC arrives in March 2027.

## Sources of truth

- `vcl/spec.md`: VCL 0.1 revision 3, the language reference.
- `vcl/orchestrator_prompt.txt`: the claude.ai prompt that writes specs. It contains the same
  reference text, so the two must always change together.
- Any VCL change means: update both files and the changelog in `spec.md`, update
  `compiler/parser.py`, `compiler/validator.py` and their tests, re-run the validator on every
  spec, and tell me so I can update the Orchestrator Project.
- `docs/VBA_Forge_Project_Roadmap.docx`: phases, gates and design decisions.

## Repo layout

```
CLAUDE.md            this brief
README.md            overview, progress log, dataset count
vcl/                 spec.md, orchestrator_prompt.txt
specs/               one .vcl file per task (task_01.vcl, ...)
compiler/            parser.py, validator.py now; router.py, assembler.py later
templates/           one .bas template per op pattern
harness/             ForgeTest.bas, batch_builder.py, results_import.py, lint.py
batches/             generated batch modules, ready to paste into Excel
data/                verified/, repair/, eval/, results/
library/             approved .bas, .cls, .frm files
tests/               unit tests for the Python tools
docs/                roadmap and design documents
logs/
.github/workflows/   CI that runs the spec checks
(train/ and model/ arrive in March 2027)
```

## Rules for code in this repo

**Python**
- Python 3.11, standard library first. Add a dependency only when needed (openpyxl is expected
  later) and pin it in `requirements.txt`.
- Every script runs from the repo root: `python compiler/<script>.py ...`.
- Checkers exit 0 when everything passes and 1 when they find problems, printing one line per
  problem.

**Specs**
- Specs are inputs, not code to fix. Never edit a spec just to make a check pass; report the
  problem to me instead.
- Run `python -m unittest discover tests` and `python compiler/validator.py specs/*.vcl`
  before every commit that touches `specs/`, `vcl/`, `compiler/` or `tests/`.

**VBA (anything that will run in Excel)**
- `Option Explicit`, Excel 2016 compatible, late-bound `Scripting.Dictionary` (no References).
- Standard error handler that writes to ForgeLog and restores ScreenUpdating and Calculation.
- No Select or Activate. Never delete source rows because data is missing or pending.
- No Kill, RmDir, Shell, registry access, HTTP, API Declare, or SaveAs over an existing file.
- Every loop has a clear bound: before March, Ctrl+Break is the only watchdog.

## Batch runner contract (harness/ForgeTest.bas)

The runner engine is fixed VBA in `harness/ForgeTest.bas`; `batch_builder.py` appends the
procedures, fixture builders and test calls for each batch. Per `vcl/spec.md`, it must:

- run each TEST line on its own temporary workbook: build the FIXTURE data in VBA, record the
  `PRE_` values, call the procedure with `Application.Run` (`runs=n` times), check the
  assertions, close without saving
- write the test id to the Results sheet before each test, so a hang shows which test it was
- support `COL`, `ROW_RANGE`, `COPY_ROW`, `FILL_COUNT`, `COUNTIFS`, `SUMIFS`, `[Header]`
  ranges, `RESULT`, `ERROR_COUNT` and `PRE_` measures
- treat a failed PRE as: one ForgeLog entry, nothing changed, the Function returns -1
- write Results columns: Test ID | Spec | Pass/Fail | Error No | Line | Description | Seconds

## Workflow for every change

1. Make one small change.
2. Verify it: run the validator and the unit tests for what you touched.
3. Read your own diff.
4. Commit with a clear message that says what changed and why.
5. Push and open or update a pull request; CI must pass. Don't merge to main unless I ask.

## Status and next steps

Phase 0, Foundation (Oct 14 to 31, 2026). Gate 0: all specs validate, and a pilot batch runs in
Excel.

Done: 10 specs, `vcl/spec.md`, the Orchestrator prompt, and the repo skeleton (folders,
`.gitignore`, README, CI). Waiting for review, to be merged in this order:
1. VCL 0.1 revision 3 (synthetic fixtures only, COUNTIFS/SUMIFS, exact goal TESTs for task_07
   and task_08): `vcl-rev3` PR.
2. `compiler/parser.py` and `compiler/validator.py`, with unit tests, replacing the starter
   checker: `parser-validator` PR. All 10 specs validate with 0 errors.

Known: task_04 exceeds the ~60-line TASK guideline; accepted for now, revisit in March.

Next, in order:
1. Pilot `harness/ForgeTest.bas`: `RunBatch` with 3 hand-written tests for `task_01`. I run it
   in Excel and paste the results into `data/results/pilot.csv`.
2. Phase 1: `harness/batch_builder.py` (lint, FIXTURE to VBA, assemble a batch),
   `harness/results_import.py`, `harness/lint.py`, and the first 15 templates.

Update this section at the end of every session.
