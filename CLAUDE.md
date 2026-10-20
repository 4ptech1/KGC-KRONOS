# VBA Forge: project brief for Claude Code

Read this first in every session. It is the standing brief; `docs/VBA_Forge_Project_Roadmap.docx`
has the full plan and `docs/vcl-mvp-plan.md` the current build sequence.

## What this is

VBA Forge is an offline, VBA-only code generator for Excel automation at KGC Engineering Projects
(pressure-vessel fabrication paper trail: BOM, weld log, NDE, NCR, calibration and similar
registers). Claude writes task specs in VCL (VBA Command Language); VBA Forge turns each spec into
VBA and proves the macro works with tests. From March 2027 our own model, trained from zero on our
data (D25), will write VCL specs from plain English; the compiler stays the only thing that writes
VBA (D31, `docs/ai-plan.md`). Until then, the goal is a verified dataset of tasks and their proven
specs, each with its VBA.

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
- No model training or local-model work until the new PC arrives in March 2027.
- **Original model and context data (D25, standing rule).** The VBA Forge model is designed and
  trained from zero by us: no pre-trained weights, no fine-tuning of someone else's model, no
  scraped or downloaded code, datasets or prompts, no retrieval over outside sources. Everything it
  learns from or reads is in-house: made by this repo's tools, by me, or by Claude through our
  pipeline, and verified by our toolchain. General tools (Python, PyTorch) are allowed, and Claude
  stays the design and build tool (Orchestrator, Claude Code). Compiler first, own model second:
  the model writes VCL, never VBA (D31).

## Sources of truth

- `vcl/spec.md`: VCL 0.4, the language reference.
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
                     generator), vba_lint.py, vcl_mutate.py, gen_semantics_doc.py,
                     vcl_formula.py (the SET_FORMULA check)
vba/                 ForgeLib.bas, the VBA runtime every generated procedure calls
dist/                vclc build output; only forge_module.bas, forge_selftest.bas and
                     forge_pilot.bas are committed
templates/           one .bas template per op pattern (later)
harness/             Excel batch runner (deferred: PR #4)
batches/             generated batch modules (later)
data/                verified/, repair/, eval/, results/
library/             approved .bas, .cls, .frm files
tests/               unit tests for the Python tools
docs/                roadmap, plans, design documents, kgc-kronos-context.md (AI context)
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
  and commit `dist/forge_module.bas`, `dist/forge_selftest.bas` and `dist/forge_pilot.bas` if they
  changed. CI runs the same and fails if any of them is stale.

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

Done: 12 specs (11 in VCL 0.2, task_12 in 0.3), VCL 0.4 in spec.md and the Orchestrator prompt (D46), the parser and validator, and the
MVP toolchain (grammar, type checker, interpreter, compiler, ForgeLib, linter, mutation scorer,
`vclc`), with one fixture generator. Fixtures from the mutation evidence (revision 5:
`stale=<n>`, `+boundary`, `+dupes`, `reldates`, rule 19). Every SKIP_IF has blank key rows
(D22). FORMAT fill names work (D23). task_11 HighlightRejectedWelds is
Boss's weld-log spec, unchanged (D24). task_12 RouteJobPacks is the router MVP spec, VCL 0.3 (D32).
Snapshot mutation score 95.4% (418 of 438); 19 of the 20 survivors are
equivalents, and 1 (task_11 `LOOP_COLS from=1`) needs a fixture change, reported to Boss. The repo is KGC-KRONOS (renamed 17 Oct 2026). PR #4 (Excel pilot) is open with the `deferred`
label.

Known: task_04 is bigger than a TASK should be: 48 OP lines (the validator warns above 40) and 107
lines of VBA (rule 9: about 60); accepted for now, revisit in March.

Next, in order (`docs/vcl-mvp-plan.md` calls them items 4, 5 and 6-7):
1. The Job-Pack Router MVP, `docs/vep-0.3-file-ops.md` (approved by Boss): sandbox only, COPY_FILE
   first (no MOVE), VCL 0.3. Phase 0 (blank key rows) and Phase 1 (the language surface: the VCL
   0.3 line, SAFE_NAME, PATH_JOIN, FOLDER_EXISTS, the path types, MAKE_FOLDER, COPY_FILE,
   MODIFIES files and folder FIXTUREs; D26, D27) and Phase 2 (the mock file system: the file op
   statuses, FILE_EXISTS, FILE_COUNT and UNCHANGED for files; E11, F7, D28) are done. Boss said
   "go Phase 3" on 19 Oct 2026, and Phase 3 is done: 3a (the compiler, ForgeLib's file helpers
   and the lint file-safety rules; D29, PR #25) and 3b (the sandbox builder and the sandbox
   length check; D30, PR #26). Boss said "go Phase 4" on 20 Oct 2026, and Phase 4 is done: 4a
   (`specs/task_12.vcl` RouteJobPacks; D32, PR #28) and 4b (spec.md and the Orchestrator prompt
   teach VCL 0.3: hard rule 21, Example 3, the new ROLE line; D33). Boss pastes the new prompt
   into the Orchestrator Project. Boss said "go Phase 5" on 20 Oct 2026, and Phase 5 is done:
   5a (the ForgeLib self-test, `dist/forge_selftest.bas`; D34, PR #30) and 5b (the task_12 pilot,
   `dist/forge_pilot.bas`, `vclc pilot-check` and `docs/excel-pilot.md`; D35). **The router MVP
   is complete.** Running the kit in Excel is optional, whenever Boss chooses; when he pastes
   results back, save them to a file and run `vclc pilot-check`.
2. The rest of revision 3 in the compiler and interpreter. Boss said "go item 2" on 20 Oct 2026.
   Designed in `docs/vep-0.4-rev3-rest.md` (D36): Phase 1, the interpreter catches up with what
   already compiles (no language change), then list syntax with DEDUPE, FIND, the sheet ops,
   FILTER, SET_FORMULA and text files. OPEN_BOOK, SAVE_AS and CLOSE_BOOK wait until March.
   Phase 1 is done (D37): arrays, INSERT_ROW, SELECT_CASE, LOG, MSG, RANGE, SHEET, FILE_NAME
   and FOR_EACH over SHEETS and CELLS now run in the interpreter. Phase 2 is done (D38): the
   `VCL 0.4` line, `cols=[1, 2]` and DEDUPE (S21). spec.md and the prompt move to 0.4 after
   the last language phase, so Boss updates the Orchestrator once. Phase 3 is done (D39): FIND
   and IS_NOTHING (S22). Phase 4a is done (D40): Workbook arguments, `MODIFIES $book sheets`,
   ADD_SHEET and SHEET_EXISTS (S23). Phase 4b is done (D41): COPY_SHEET with `name=`. Phase 5 is
   done (D42): FILTER hides rows by SAME, never by AutoFilter, with VISIBLE_COUNT. Phase 6 is done
   (D43): SET_FORMULA with an Excel 2016 allowlist, tested by its text with FORMULA. Phase 7a is
   done (D44): LIST_FILES and ITEMS (S24). Phase 7b is done (D45): READ_TEXT and WRITE_TEXT with
   statuses (S25), file texts in folder FIXTUREs and FILE_TEXT. The switch is done (E12, D46):
   spec.md and the Orchestrator prompt teach VCL 0.4 (hard rule 22: sheets and formulas), and
   Boss pastes the new prompt into the Orchestrator Project once. **Item 2 is complete.** CALL
   and BUILD specs (the VEP's group D) are their own item, for Boss to schedule.
3. After the MVP: the AI plan, `docs/ai-plan.md` (D31, approved by Boss on 19 Oct 2026): the
   model writes VCL, never VBA. Its Stage A (a `GOAL` line, a formatter, error codes, the corpus,
   the data generator and the benchmark) starts when Boss says "go item 3", as every earlier
   item did (D47); no training before March 2027.

Update this section at the end of every session, and keep `docs/kgc-kronos-context.md` (the
context file Boss gives to other Claude chats) in step with it.
