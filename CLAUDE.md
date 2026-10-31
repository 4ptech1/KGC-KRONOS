# VBA Forge: project brief for Claude Code

Read this first in every session. It is the standing brief; `docs/VBA_Forge_Project_Roadmap.docx`
has the full plan and `docs/vcl-mvp-plan.md` the current build sequence.

## What this is

VBA Forge is a VBA-only code generator for Excel automation at KGC Engineering Projects
(pressure-vessel fabrication paper trail: BOM, weld log, NDE, NCR, calibration and similar
registers). An AI writes task specs in VCL (VBA Command Language): Claude today, and ChatGPT
Codex through the KGC KRONOS desktop app (item 5, D68). VBA Forge turns each spec into VBA and
proves the macro works with tests. Only the request-to-VCL step may go online; the toolchain and
the VBA stay offline (D69). From March 2027 our own model, trained from zero on our data (D25,
D70), will write VCL specs from plain English; the compiler stays the only thing that writes VBA
(D31, `docs/ai-plan.md`). Until then, the goal is the KRONOS app and a verified dataset of tasks
and their proven specs, each with its VBA.

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

- My work PC is locked down: no Excel Trust Center changes, and no local installs except
  KRONOS.exe and the Codex CLI, once IT approves them (D69). All development runs in the cloud
  (Claude Code on the web, GitHub Codespaces, GitHub Actions).
- Excel 2016 only. No XLOOKUP, FILTER(), LET(), LAMBDA, dynamic arrays or any other
  Microsoft 365 feature, in VBA or in formulas.
- The repo is public: synthetic data only. Never put real KGC files, names, figures or customer
  data in the repo, its tests, specs or corpus, or in Claude's prompts. Excel workbooks are
  git-ignored on purpose.
- KRONOS may receive real work requests (the company allows sending them to ChatGPT), never
  files. Its runs stay in its own data folder outside the repo; at first only practice requests
  with invented data. Personal ChatGPT accounts turn off "Improve the model for everyone" in
  Data Controls. Later KRONOS moves to the company's storage drive, where real data may live
  (D69, D70).
- Synthetic fixtures only: every TEST uses a named FIXTURE file. `fixture=target` is banned
  until March 2027, and nothing we generate ever touches real workbooks.
- No model training or local-model work until the new PC arrives in March 2027.
- **Original model and context data (D25, standing rule).** The VBA Forge model is designed and
  trained from zero by us: no pre-trained weights, no fine-tuning of someone else's model, no
  scraped or downloaded code, datasets or prompts, no retrieval over outside sources. Everything it
  learns from or reads is in-house: made by this repo's tools, by me, by Claude through our
  pipeline, or by Codex through KRONOS (D70), and verified by our toolchain. General tools (Python,
  PyTorch) are allowed, and Claude stays the design and build tool (Orchestrator, Claude Code).
  Compiler first, own model second: the model writes VCL, never VBA (D31).

## Sources of truth

- `vcl/spec.md`: VCL 0.6, the language reference.
- `vcl/orchestrator_prompt.txt`: the prompt that writes specs, pasted into the claude.ai
  Orchestrator and sent to Codex by the KRONOS app (D68). It contains the same reference text as
  spec.md (a unit test keeps them identical), so the two always change together.
- `vcl/grammar.lark`: the formal syntax. `vcl/semantics.py`: the reference meaning of the value
  functions, with conformance cases in `tests/conformance_cases.py` and the generated tables in
  `docs/vcl-semantics.md`.
- `DECISIONS.md`: every language decision and every decision you made on your own.
- Any VCL change means: update spec.md (with its changelog) and the prompt together, the grammar,
  `compiler/parser.py` and `compiler/validator.py`, the type checker, the interpreter, the
  compiler and ForgeLib as needed, the conformance cases and `docs/vcl-semantics.md`, their
  tests and `DECISIONS.md`; re-run every check; rebuild `dist/forge_module.bas`; and tell me so I
  can update the Orchestrator Project.
- `docs/VBA_Forge_Project_Roadmap.docx`: phases, gates and design decisions. Its phase list is the
  plan of 14 Oct 2026; the current order of work is "Next" below and `docs/vcl-mvp-plan.md` (D73).
- `docs/kronos-app.md`: the KGC KRONOS desktop app (item 5, D68 to D71).

## Repo layout

```
CLAUDE.md            this brief
DECISIONS.md         language decisions and Claude's own decisions
README.md            overview, progress log, dataset count
requirements.txt     pinned dependencies (lark, with its hash)
vcl/                 spec.md, orchestrator_prompt.txt, grammar.lark, semantics.py, examples/
specs/               one .vcl file per task (task_01.vcl, ...)
compiler/            vclc.py (the toolchain), validator.py + parser.py, vcl_parse.py (grammar),
                     errcodes.py + gen_error_codes.py (stable message codes, D60), vcl_goal.py (GOAL lines),
                     vcl_fmt.py (`vclc fmt`, the canonical text, D62),
                     vcl_types.py (types), vcl_interp.py (reference interpreter),
                     vcl_codegen.py (VBA compiler), vcl_fixtures.py (the one fixture
                     generator), vba_lint.py, vcl_mutate.py, gen_semantics_doc.py, gen_context.py,
                     vcl_formula.py (the SET_FORMULA check),
                     kronos_engine.py + kronos_ai.py (the KGC KRONOS engine and `vclc ask`, D72)
vba/                 ForgeLib.bas, the VBA runtime every generated procedure calls
dist/                vclc build output; only forge_module.bas, forge_selftest.bas and
                     forge_pilot.bas are committed (builds/<NAME>.bas: one module per BUILD)
templates/           one .bas template per op pattern (later)
harness/             Excel batch runner (deferred: PR #4)
batches/             generated batch modules (later)
data/                verified/ (the proven request and spec pairs, D64), repair/, eval/ (held out), results/
library/             approved .bas, .cls, .frm files
tests/               unit tests for the Python tools
docs/                roadmap, plans, design documents, kgc-kronos-context.md (AI context)
logs/
.github/workflows/   CI
app/                 the KGC KRONOS window: kronos_controller.py (the wizard's state, D75), kronos_app.py (2b)
(train/ and model/ arrive in March 2027)
```

## Rules for code in this repo

**Python**
- Python 3.11, standard library first. Add a dependency only when needed and pin it, with its
  hash, in `requirements.txt` (today: `lark==1.3.1`).
- Every script runs from the repo root: `python compiler/<script>.py ...`.
- Checkers exit 0 when everything passes and 1 when they find problems, printing one line per
  problem.
- Every message a check gives is written `coded("text", "V012")`. For a new message write
  `coded("text")` and run `python compiler/gen_error_codes.py`: it gives the code and updates
  `compiler/error_codes.json` and `docs/error-codes.md`. Never reuse or change a code.

**Specs**
- Specs are written in canonical form: `python compiler/vclc.py fmt --check specs/*.vcl` (a unit
  test checks it). To tidy a new one, `fmt --write` it.
- Specs are inputs, not code to fix. Never edit a spec just to make a check pass; report the
  problem to me instead (an edit I ask for, like the task_06 rewrite, is fine).
- Before every commit that touches `specs/`, `vcl/`, `compiler/`, `vba/` or `tests/`, run:
  ```
  python -m unittest discover tests
  python compiler/validator.py specs/*.vcl
  python compiler/vclc.py all specs/*.vcl -o dist
  python compiler/gen_semantics_doc.py --check
  python compiler/gen_context.py --check
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
TESTs in the reference interpreter and compile to lint-clean VBA, in CI. **Gate 0 is met.** The
roadmap .docx's later phases are the plan of 14 Oct 2026; the order of work is the list below
(D73).

Done: 12 specs (11 in VCL 0.2, task_12 in 0.3), VCL 0.6 in spec.md and the Orchestrator prompt (D63), the parser and validator, and the
MVP toolchain (grammar, type checker, interpreter, compiler, ForgeLib, linter, mutation scorer,
`vclc`), with one fixture generator. Fixtures from the mutation evidence (revision 5:
`stale=<n>`, `+boundary`, `+dupes`, `reldates`, rule 19). Every SKIP_IF has blank key rows
(D22). FORMAT fill names work (D23). task_11 HighlightRejectedWelds is
Boss's weld-log spec, unchanged (D24). task_12 RouteJobPacks is the router MVP spec, VCL 0.3 (D32).
Two review passes after item 2 are done (D47 and D48, PRs #45 and #46; D51, a parallel
session's PR #48 folded in by #51): Orchestrator dry runs and ForgeLib reviews; the prompt now
says what the toolchain does, and Boss pastes the final prompt once.
Snapshot mutation score 95.4% (418 of 438); 19 of the 20 survivors are
equivalents, and 1 (task_11 `LOOP_COLS from=1`) needs a fixture change, reported to Boss. The repo is KGC-KRONOS (renamed 17 Oct 2026). PR #4 (Excel pilot) is open with the `deferred`
label.

Known: task_04 is bigger than a TASK should be: 48 OP lines (the validator warns above 40) and 107
lines of VBA (rule 9: about 60); accepted for now, revisit in March.

Next, in order (`docs/vcl-mvp-plan.md` calls them items 4, 5, 5b, 6-7 and 9):
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
   Boss pastes the new prompt into the Orchestrator Project once. **Item 2 is complete.**
   After item 2, a review pass (D47: wording fixed where it differed from the tools, clearer TEST
   failures) and ForgeLib fixes from a VBA review (D48).
3. CALL and BUILD specs, `docs/vep-0.5-build.md` (D49). Boss said "continue building the next
   part according to your recommendations" on 21 Oct 2026. Phases: 1 the language surface and the
   validator, 2 the interpreter, 3 the compiler and mutation, 4 the review gaps (D47, D51), 5 a proven
   BUILD and the switch to VCL 0.5 (Boss updates the Orchestrator once). FORM and xlam wait.
   Phase 1 is done (E13, E14, D50): the `VCL 0.5` line, CALL `into=`, and the BUILD rules in the
   validator and type checker. Phase 2 is done (D52): the interpreter runs a 0.5 BUILD's TESTs
   through modMain.RunHeadless. Phase 3 is done (D53): `vclc build` writes `dist/builds/<NAME>.bas`,
   and mutation scores a BUILD. Phase 4a is done (D55): ROW_OF(<range>) and DEDUPE `into=`.
   Phase 4b is done (D56): a header the spec writes to row 1 counts as checked for rule 14.
   Phase 4c is done (D57): an assertion names a sheet with spaces in quotes, `'Open NCRs'!B1`.
   Phase 5 is done (D58, and D67 for two BUILD rules found after it): a proven BUILD, `vcl/examples/example4_build.vcl` (the prompt's
   Example 4; its 3 TESTs pass and catch 34 of 34 mutants), and the switch: spec.md and the
   Orchestrator prompt teach VCL 0.5 (hard rules 7, 9 and 14 changed, new hard rule 23 for BUILD
   and CALL), and Boss pastes the new prompt into the Orchestrator Project once. **Item 3 is
   complete.** FORM and xlam stay refused until after March 2027 (D49).
4. The AI plan, `docs/ai-plan.md` (D31, approved by Boss on 19 Oct 2026): the model writes VCL,
   never VBA. Boss said "go item 4" on 21 Oct 2026 for its Stage A, designed in
   `docs/vep-0.6-stage-a.md` (D59): Phase 1 error codes, 2 the `GOAL` line (VCL 0.6), 3 the
   formatter `vclc fmt`, 4 the switch of spec.md and the prompt to 0.6 (Boss updates the
   Orchestrator once), 5 corpus tooling and the first pairs, 6 the corpus to 150 tasks, 7 the
   data generator, 8 the benchmark. Phase 0 (the design), Phase 1 (error codes: every message
   has a stable code, `docs/error-codes.md`, D60) and Phase 2 (the toolchain reads `VCL 0.6` and the
   `GOAL` line, E15, D61; `vcl_goal.py` adds GOAL lines to a spec) and Phase 3 (`vclc fmt`: one
   canonical text for every spec, which proves it changed nothing, D62) and Phase 4 (the switch:
   spec.md and the Orchestrator prompt teach VCL 0.6, with hard rule 24 for the GOAL line, D63;
   Boss pastes the new prompt once) and Phase 5 (the corpus tooling: `vclc corpus`, `corpus-add`,
   `export`; the 12 committed specs are the pairs `c001` to `c012` in `data/verified/`; the plan
   for the rest is `docs/corpus-plan.md`; D64) are done; Phase 6, the corpus to 150 tasks in
   batches of about 20, is in progress (batches 1 and 2, c013 to c052, are stored: 52 pairs of 150, four of
   them held out; `data/gaps.md` logs what VCL could not say; D65, D66). Batch 3, leaning on the thin
   registers and families (`vclc corpus --report`), waits until KRONOS.exe is done (Phase 3; D71,
   D74); no training before March 2027.
5. **The KGC KRONOS desktop app** (Boss asked on 30 Oct 2026; `docs/kronos-app.md`, D68 to D71): a
   window where a person types a request; KRONOS asks ChatGPT Codex (the Codex CLI, `codex exec`,
   read-only) for the VCL, then checks, tests and builds it, sends problems back for up to 2 more tries,
   lets the person edit the spec, and saves the .bas with import steps and a starter macro. Tkinter,
   one KRONOS.exe built by GitHub Actions; Boss's PC first, then the team on the company drive. Only the
   request-to-VCL step is online. Phase 0 (the design and the decisions, with D67; PR #58) is done.
   Phase 1 is done (D72): the engine (`compiler/kronos_engine.py`), the fake and Codex adapters
   (`compiler/kronos_ai.py`, the prompt on stdin), reply reading (codes K001 to K004), the repair loop
   (3 tries), the run folder outside the repo, the starter macro (`vcl_codegen.starter_sub`) and
   `python compiler/vclc.py ask "<request>"`. **Next: Phase 2, the window, before everything else
   (Boss, 31 Oct 2026; D74).** It is a four-step wizard (Ask, Questions, Review, Save) with the recipe
   hidden behind a setting, designed in `docs/kronos-app.md` section 2, in PRs 2.0 (the design, docs
   only; done, PR #70), 2a (engine hooks and the wizard's controller; done, D75), 2b (the window and
   `--smoke`; next), 2c (first-run checklist, Settings, History, Help) and 2d (Suggest for the corpus);
   every window PR sends Boss screenshots. Then Phase 3 (KRONOS.exe and the guide), then the pilot (Phase 4) once IT approves
   KRONOS.exe and the Codex CLI; corpus batch 3 resumes after Phase 3, and the team version (Phase 5)
   follows the pilot. The parallel session is gone; this is the only session on the repo.

Update this section at the end of every session, and keep `docs/kgc-kronos-context.md` (the
context file Boss gives to other Claude chats) in step with it: edit its sections 1 to 12 by hand,
then run `python compiler/gen_context.py`, which rewrites its appendices from the repo (D54).
