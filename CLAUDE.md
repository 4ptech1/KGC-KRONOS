# VBA Forge: project brief for Claude Code

Read this first in every session. It is the standing brief; `docs/VBA_Forge_Project_Roadmap.docx`
has the full plan and `docs/vcl-mvp-plan.md` the current build sequence.

## What this is

VBA Forge is a VBA-only code generator for Excel automation at KGC Engineering Projects
(pressure-vessel fabrication paper trail: BOM, weld log, NDE, NCR, calibration and similar
registers). An AI writes task specs in VCL (VBA Command Language): Claude today, and ChatGPT
Codex through the KGC KRONOS desktop app (item 5, D68). VBA Forge turns each spec into VBA and
proves the macro works with tests. Only the request-to-VCL step may go online; the toolchain and
the VBA stay offline (D69). From March 2027 our own model, trained on our data (D25,
D70; since D102 either trained from zero or a small open model tuned on it, whichever scores better), will write VCL specs from plain English, and the VBA for them with its context and memory in VCL (D90,
`docs/ai-vba-plan.md`): its VBA is used only when it passes the recipe's own TESTs, else the compiler's (D31). Until then, the goal is the KRONOS app and a verified dataset of tasks
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
- When Boss wants KRONOS offline, in whatever words, build the offline copy as
  `docs/kronos-offline.md` says and give its download link first (D86).

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
  (D69, D70). KRONOS Offline may also be given a real workbook to make a changed copy of it (D87,
  Boss, 16 Nov 2026): that happens in the browser on his PC, the original is never changed, the file
  never goes into the repo or into this Claude Code session, and what he copies from it into an AI
  chat is his choice. Permitted KGC samples (D101, Boss, 19 Nov 2026) may be read in a Claude Code session to learn
  the shape of the registers; nothing from them goes into the repo, its tests, specs, corpus or docs, and recipes
  modelled on them use invented values. **D105, D106 (Boss, 20 Nov 2026):** the KGC samples themselves are published
  in `samples/kgc/` (11 files, listed with their SHA-256 in its README), the one exception to the synthetic-only rule;
  no other real file joins them without Boss's word, and specs, TESTs, FIXTUREs, the corpus and unit tests still use
  invented look-alikes only (`tests/test_samples.py`).
  From March 2027 the company storage is the new PC's 300 GB, reserved for KRONOS (D91,
  `docs/kronos-storage-plan.md`): real data only under `KRONOS\work` (and what is trained on it), models trained
  on real data never leave that PC, and KRONOS never deletes real records.
- Synthetic fixtures only: every TEST uses a named FIXTURE file. `fixture=target` is banned
  until March 2027, and nothing we generate ever touches real workbooks, except the copy KRONOS
  Offline makes of a workbook Boss uploads (D87); no TEST ever uses one.
- No model training or local-model work until the new PC arrives in March 2027 (writing the `train/` code is allowed, D124;
  no model is downloaded or run).
- **Original model and context data (D25, standing rule; its weights rule lifted by D102).** Since D102 (Boss,
  19 Nov 2026) the model may start from a small open-weight model (Apache 2.0 or MIT licence, at most about 7 billion
  parameters, tuned with LoRA, run offline on the new PC), built beside our own model trained from zero (at most about
  350 million parameters); the held-out benchmark chooses. Still banned: scraped or downloaded code, datasets or
  prompts as training data, and retrieval over outside sources. Everything it
  learns from or reads is in-house: made by this repo's tools, by me, by Claude through our
  pipeline, or by Codex through KRONOS (D70), and verified by our toolchain. General tools (Python,
  PyTorch) are allowed, and Claude stays the design and build tool (Orchestrator, Claude Code) and the only
  one who commits. OpenAI Codex on Boss's own PC may advise, practise VCL and propose code, which comes back
  through `codex_changes/` for Claude to review, verify and merge; its rules are in `AGENTS.md` (D83).
  Compiler first, own model second (D31). Since D90 (Boss, 17 Nov 2026) the model may also write the VBA from a
  proven VCL recipe, its context and memory in VCL; that VBA is used only when it passes the recipe's TESTs in the
  VBA simulator, and the compiler's VBA is the fallback (`docs/ai-vba-plan.md`).

## Sources of truth

- `vcl/spec.md`: VCL 0.7, the language reference (D116).
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
- `docs/ai-vba-plan.md`: our model writes the VBA from a proven recipe, and the VBA simulator proves it (item 7, D90).
- `docs/kgc-documents-plan.md`: KRONOS fills KGC's own Excel and Word documents (WPS, estimation, enquiry register;
  D105), the plan another session implements; Phase 0, publishing the samples in `samples/kgc/`, is done (D106).
- `docs/known-gaps.md`: what the dry run of KRONOS Offline version 6 on the real samples found, to fix later (D125).
- `docs/kronos-storage-plan.md`: the new PC's 300 GB, KRONOS's home from March 2027: the layout, budgets, rules
  and monthly report (item 8, D91; the reference files follow it, D92).

## Repo layout

```
CLAUDE.md            this brief
AGENTS.md            the front page for OpenAI Codex: read CLAUDE.md first, then its own rules (D83)
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
                     kronos_engine.py + kronos_ai.py (the KGC KRONOS engine and `vclc ask`, D72),
                     kronos_web.py (the bridge KRONOS Offline's page calls under Pyodide, D86),
                     kronos_book.py (an uploaded workbook: describe, run a proven recipe on it, write the copy, D87),
                     kronos_recipes.py (recipes read from .vcl/.txt/.md/.docx files, D89), kronos_wizard.py (the browser
                     wizard: KRONOS.exe's Run with copy and paste for the AI, D89), kronos_words.py (EXAMPLE, TIPS, NEVER),
                     vba_sim.py (the VBA simulator: runs generated VBA without Excel, D90, D98), vba_check.py (the
                     compiled VBA of every TEST in the simulator against the interpreter, D99),
                     vcl_datagen.py (`vclc datagen`: repairs and contrast examples from the pairs, D103),
                     kronos_finder.py (the recipe finder: data/library.json, search and adapt, D129)
vba/                 ForgeLib.bas, the VBA runtime every generated procedure calls
dist/                vclc build output; only forge_module.bas, forge_selftest.bas and
                     forge_pilot.bas are committed (builds/<NAME>.bas: one module per BUILD)
templates/           one .bas template per op pattern (later)
harness/             Excel batch runner (deferred: PR #4)
batches/             generated batch modules (later)
data/                verified/ (the proven request and spec pairs, D64), repair/, eval/ (held out), results/,
                     library.json (the finder's library, built from verified/ and docs/demo, D129),
                     templates/ (invented look-alike KGC documents, D108)
library/             approved .bas, .cls, .frm files
tests/               unit tests for the Python tools
docs/                roadmap, plans, design documents, kgc-kronos-context.md (AI context)
logs/
.github/workflows/   CI
app/                 the KGC KRONOS window: kronos_controller.py (the wizard's state, D75), kronos_app.py (the window, D76),
                     kronos_dialogs.py (Settings, Recipe, History, Help, D77)
packaging/           kronos.spec: the PyInstaller spec for KRONOS.exe (D79); the build is
                     .github/workflows/build-exe.yml, its tools are in requirements-build.txt;
                     check-release-exe.yml tests an exe that is already published (D84);
                     KRONOS Offline (D86): build_offline.py builds its zip from offline/ (the launcher and the
                     three servers, licences/), offline_pins.json (the downloads, by SHA-256) and web/; its CI is
                     build-offline.yml, its test tools are in offline_test/
web/                 KRONOS Offline's page: index.html, styles.css, app.js, worker.js (Pyodide), logo.svg (D86)
train/               code only, no model loaded or downloaded (D124): prompts.py, make_dataset.py (the chat dataset from data/verified
                     and `vclc datagen`, into the git-ignored dist/train), grammar_gbnf.py + vcl.gbnf (VCL's grammar for llama.cpp, D127),
                     gbnf_check.py, qlora_train.py and export_model.py (the tuning run and the 4-bit export, never run, D128)
(model/ arrives in March 2027; the prepared datasets, the checkpoints and the weights stay
 on the new PC's 300 GB, in KRONOS\train and KRONOS\models: weights never go into the repo, D91)
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

Known gaps of KRONOS Offline version 6 on the real samples (21 Nov 2026, D125, `docs/known-gaps.md`): a recipe that writes over a
formula gives no warning; the look-alike WPS form doesn't match the real forms' rows; a Word marker inside a table cell is
accepted by the check and refused by the fill; a Word recipe counts tables by position. Recorded, not fixed: Boss will say go.

**Boss's ten steps (18 Nov 2026, D101), the current order:** 1 try KRONOS Offline v5 and 2 the IT request and prompt
paste (Boss: done); 3 and 4 the corpus to 150 and the prompt fixes (done, D93 to D97); 5 the VBA simulator (done, D98
to D100); 6 the data generator (done, D103, D104: `vclc datagen` and 405 paraphrases). **Next: KGC's own documents (D105,
`docs/kgc-documents-plan.md`, Phases 0 to 6)**: KRONOS fills the WPS forms, the estimation workbook, the enquiry register and the
Word estimation sheet and technical proposals; Phase 0 (the samples, D106) and Phase 1 (the survey, `docs/kgc-documents.md`,
D107), Phase 2 (the look-alike templates, `data/templates/`, D108) and Phase 3 (VCL 0.7: the VEP D109, template
FIXTUREs D110, ROUND/ROUNDUP/MIN/MAX D111, AI_TEXT D112, merged cells D113, two workbooks D114, DOCUMENT for Word D115,
and the switch of spec.md and the prompt D116) and Phase 4 (KRONOS Offline version 6: the design D117, the engine D118,
a copy that keeps everything D119, two workbooks D120, Word on the page D121, Review and the guide D122, published as the
release `kronos-offline-v6`, D123) are done. **Next: Phase 5** (corpus batch 9, about 20 document jobs) **and Phase 6**,
step 7, the benchmark. Then 8 KRONOS Offline version 8 (the old version 6 scope: History, the practice-run check, the
corpus suggestion, Help), 9 storage S1, 10 the VCL memory and the day-one kit. For Boss: with version 6 in hand, paste
the updated prompt (VCL 0.7; D94, D97, D116) into the Orchestrator Project.

**The model kit (Boss, 21 Nov 2026; `docs/ai-plan.md` "Stage C in detail", `train/README.md`, D124 to D128):** the plan for the open
model is written (a small coder model in 4-bit, kept to KRONOS's job by LoRA, a grammar lock and a short prompt; an optional Codex
meaning check on a proven recipe), and its code before March is done: the chat dataset (`train/make_dataset.py`: 1,935 examples,
split by pair, held-out items guarded), VCL's grammar for the offline runner (`train/vcl.gbnf`: accepts all 154 recipes), the QLoRA
tuning script and the 4-bit export (written and tested without a model, never run; a dry run unless `--go`; the first run in March
is `--smoke`). Still to do before March: the benchmark (step 7), the Codex meaning check in KRONOS.exe and in KRONOS Offline version 8,
and the in-house glossary. The gaps in `docs/known-gaps.md` (D125) wait for Boss's go.

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
4. The AI plan, `docs/ai-plan.md` (D31, approved by Boss on 19 Oct 2026): the model writes VCL; since D90
   (17 Nov 2026, `docs/ai-vba-plan.md`) it may also write the VBA, proven in a VBA simulator against the recipe's TESTs. Boss said "go item 4" on 21 Oct 2026 for its Stage A, designed in
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
   batches of about 20, **is complete** (batches 1 to 8, c001 to c159: 150 pairs of 150, 15 of them held out;
   `data/gaps.md` logs what VCL could not say; D65, D66, D82, D85, D93, D95, D96). Batch 5 (Boss: "Complete
   3 and 4", 18 Nov 2026) stored 15 of 20; c074 and c098 failed on FIXTURE layout and return in batch 6, c106 (a
   BUILD under the floor) too, c094 needs a formula naming a sheet with a space (a language gap) and c107 was one
   TASK for a build request; replacements follow. Preparing step 4, the first prompt fix, showed that the
   interpreter and the compiled VBA compared numbers typed as text differently: D94 gives `<`, `>`, `<=` and `>=`
   one meaning (S26, ForgeCompare) and says it in spec.md and the prompt. Batch 6 stored 17 of 22 and retired c074,
   c098 and c106 after three batches each (D95). Batch 7 and a top-up finished the 150 on 19 Nov 2026 (D96). The
   wording fixes are done too (D97: the FIXTURE `+boundary` rules, the wrong-sheet TEST, entry macros, and five
   smaller facts, in spec.md and the prompt together; Boss pastes the prompt once, with D94's line in it). Next:
   Phase 7, the data generator (Boss's step 6), is done: `vclc datagen` makes 4,432 repairs and 124 contrast examples
   from the 135 training pairs (D103), and every training pair has three checked paraphrases (D104). Next: Phase 8,
   the benchmark (step 7). No training before March 2027.
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
   only; done, PR #70), 2a (engine hooks and the wizard's controller; done, D75, PR #71), 2b (the window,
   the first-run checklist and `--smoke`, which CI runs under xvfb-run; done, D76), 2c (Settings, History,
   Help, the Recipe window, Test Codex, Codex errors in plain words; done, D77) and 2d (Suggest for the
   corpus, from Review or Save, into the data folder's `corpus_candidates`; done, D78); every
   window PR sends Boss screenshots (`python app/kronos_app.py`, Tk needed: in the cloud container,
   Python 3.12 with `python3-tk` under `xvfb-run`). **The window is complete.** Phase 3 is done (D79):
   `packaging/kronos.spec` and `.github/workflows/build-exe.yml` build KRONOS.exe on Windows and run its
   `--self-test`; `compiler/resources.py` finds the data files inside the exe; `docs/kronos-app-guide.md` is
   Boss's guide with a message for IT; a manual run of the workflow with "release" ticked publishes it
   (releases `kronos-v0.4` and `kronos-v0.5` are published). Boss asked on 2 Nov 2026 whether the MVP was ready; it was built
   but never run against the real Codex or in Excel, so two more PRs went in before the pilot: 3b (D80,
   PR #76) ran KRONOS against the real Codex CLI 0.166.0, signed out, and fixed what it showed (it asks
   `codex login status` first, stops a Codex that can't reach OpenAI instead of waiting 5 minutes, stops
   the whole codex.cmd process tree on Windows, `--ephemeral`; a Windows `codex-check` job in CI); 3c (D81)
   adds the practice run: Save also writes `<name>_practice.bas`, which in a new, empty workbook builds the
   TESTs' practice data and runs the macro, and the window's Check a practice run compares Excel's answers
   with the tests (version 0.5, published, SHA-256 on its release page). The published 0.5 was checked again on 14 Nov 2026 (D84):
   the manual workflow `.github/workflows/check-release-exe.yml` downloads KRONOS.exe from a release, compares its SHA-256 and runs
   its `--self-test` on Windows; it passed. **Next: the pilot (Phase 4)**, once IT
   approves KRONOS.exe and the Codex CLI: Boss reported the IT request done on 18 Nov 2026 (D101), and the pilot starts
   when both are installed on his PC (not yet confirmed); the team version (Phase 5) follows the pilot. Still unproven, and what the pilot is for: Codex's
   real reply to the Orchestrator prompt, and any macro running in real Excel. The parallel session is gone; this is the only session on the repo.
6. **KRONOS Offline** (Boss decided on 14 Nov 2026; `docs/kronos-offline.md`, D86): KRONOS in the browser with nothing
   installed, so it can reach the locked-down PC while KRONOS.exe waits for IT. Boss unzips a folder and double-clicks
   Start KRONOS (Windows); a tiny local server (Node.js, Python or Windows PowerShell) shows a page in BlueForce's look
   that runs the repo's own toolchain through Pyodide 0.24.1 (the last that covers Chrome 109), through
   `compiler/kronos_web.py`; the AI step will be copy and paste with an AI chat; no password page. Phase 1, the proof, is
   done: version 1 checks, tests and builds a pasted recipe offline, its answers equal normal Python's, and CI checks the
   zip, the servers (PowerShell on Windows included) and the page on today's Chromium, Chromium 109 and Firefox 115;
   it is published as the release `kronos-offline-v1` (built from main `2ed9f79` by `build-offline.yml`; its SHA-256
   is on the release page, and the downloaded zip passed every check again). Version 2, the same day, only makes the
   Python server (`serve.py`) give every connection its own thread (version 1's answers nobody while an idle
   connection is open; not seen with a real browser); the page, but for one sentence, and the PowerShell server, which
   the work PC uses, are unchanged; it is published as the release `kronos-offline-v2` (main `988b020`), the copy
   Boss was handed. Three rules came out of Phase 1 for every later phase: the page never loads anything synchronously
   while a download is under way; the tests run the page through a stand-in that answers one request at a time as
   `serve.ps1` does; and a test keeps a server's output in a file, never a pipe (a full 4 KB pipe stopped `serve.py`
   in one Windows run). Not proven until Boss tries it: Windows PowerShell 2, Windows 7 and the work PC itself (CI has
   PowerShell 5.1). **Next: Boss tries version 3 on the work PC; then Phase 2, the wizard** (the four steps with the
   copy-paste AI step and the .bas download), then Phase 3 (History, the practice-run check, the corpus suggestion,
   Help), each its own plan, PR and version. **Boss asked on 16 Nov 2026 for an upload button (D87), and it comes first,
   as version 3 on the same one screen:** he chooses a workbook, KRONOS makes a recipe's changes on a copy with the
   reference interpreter, and he downloads the copy. Real workbooks are allowed, anything from the file may go into
   the text copied to the AI, and what a copy can't keep (charts, macros that moved, formulas after row changes) is
   warned about before the download. PR 1, the engine (`compiler/kronos_book.py`, openpyxl 3.1.5, the interpreter's
   per-cell origins; the TESTs of every recipe in the repo give the same workbook through an uploaded file) and
   PR 2, the page (a "Your workbook" card with the description for the AI, "Make the changes on a copy" after a PROVEN
   recipe, warnings to tick, the download; `docs/kronos-offline.md` Phase 1b; the plain-language guide
   `docs/KRONOS_Offline_Guide.docx` is version 3), are done, and version 3 is published as the release `kronos-offline-v3` (main `10fbdf0`; SHA-256 `32d694d4…`), the copy Boss
   is handed. Not proven until he tries it: a real workbook of his, on the work PC, and how long his biggest sheets take (a
   sheet of 5,000 rows with a quarter deleted takes about 50 s in the browser; the time grows with the square of the rows).
   **The demo (D88, 16 Nov 2026):** `docs/demo/` holds an invented NCR register (`NCR_Register_Demo.xlsx`, 40 rows), three
   PROVEN recipes (highlight open Major NCRs, tidy, summary sheet), what KRONOS gives back, and the Word guide
   `KRONOS_Demo_NCR_Register.docx` for colleagues; the scripts that make them are beside them, and `tests/test_demo.py`
   keeps every file what they make today. Only `docs/demo/NCR_Register_Demo*.xlsx` may be committed.
   **Version 4 (D89, 16 Nov 2026): the wizard, with recipes read from files.** Boss asked that KRONOS read the .vcl, .txt,
   .md and .docx files recipes come in; it came together with the wizard (Phase 2). Part 1 (`compiler/kronos_recipes.py`),
   part 2 (`compiler/kronos_wizard.py`: KRONOS.exe's own `Run` with copy and paste for the AI; `kronos_book.apply_many`)
   and part 3 (the page: 1 Ask, 2 Questions, 3 Review, 4 Save, "Open a recipe file", several recipes on one copy) are
   done, and version 4 is published as the release `kronos-offline-v4` (main `4c119eb`; SHA-256 `46f5f8a7…`), the copy Boss
   is handed. Not proven until he tries it: the copy and paste with his real AI chats, his own recipe files, the work PC.
   Version 5 (the same day) only makes Cancel sure: the practice example's panel shows at once, and Cancel before KRONOS's
   words arrive loads them again. It is published as the release `kronos-offline-v5` (main `62d66d4`; SHA-256 `79fdb69c…`),
   the copy Boss is handed. Boss tried it on the work PC on 18 Nov 2026: it works (D101). **Version 6 (D117 to D123, 21 Nov 2026):
   KGC's own documents.** A copy that keeps everything (only the changed cells are written into the file when no row or sheet
   moves), two workbooks ("Workbook to read from"), Word files filled from a template and a workbook, Review that says what a
   recipe fills and lists the AI's own text, and recipes for VCL 0.7. Published as the release `kronos-offline-v6` (main
   `95bf184`; SHA-256 `d4e1fa22…`), the copy Boss is handed; not proven until he tries it: his own Word files and workbooks.
   The old version 6 scope (Phase 3: History, the practice-run check, the corpus suggestion, Help) is version 8.
   **Version 7 (D129, D130, 22 Nov 2026): the recipe finder,** KRONOS's own AI until March (Boss: "Implement the AI and give
   me KRONOS"): "Find a proven recipe" shows the closest of the library's proven recipes (`data/library.json`, never
   `data/eval`) and renames its columns to the workbook's where sure, each rename ticked, then proves it again.
7. **The AI writes the VBA, with its context in VCL** (Boss, 17 Nov 2026; `docs/ai-vba-plan.md`, D90). Phase 0, the design,
   is done. Phase 1, `compiler/vba_sim.py`, a simulator for the VBA KRONOS writes (Boss: "Complete 5", 19 Nov 2026), is
   in three parts: part 1, the language core, is done (D98: the ForgeLib self-test's 241 cases pass in it, and planted
   bugs are caught); part 2 is done (D99: every compiled TASK without folders, 130 in all, ends where the interpreter
   does on every TEST, and it found and fixed two compiler bugs in task_09: a reserved name, `cDate`, and a date
   where a number of days belongs); part 3 is done (D100: files and BUILDs too, so all 165 specs, examples and
   corpus pairs with TESTs end where the interpreter does, and `vclc sim` runs in CI). **Phase 1 is complete.**
   Next: Phase 2, the VCL memory (request, VCL, VBA pairs and a similar-recipe index); the model itself from
   March 2027.
   **D102 (Boss, 19 Nov 2026): D25's ban on pre-trained weights is lifted.** In March two candidates are built and the
   held-out benchmark chooses: our own model from zero (at most about 350 million parameters) and a small open-weight
   coder model (Apache 2.0 or MIT, at most about 7 billion) tuned with LoRA on our data, offline on the new PC.
8. **The new PC's 300 GB, KRONOS's home from March 2027** (Boss, 17 Nov 2026; `docs/kronos-storage-plan.md`, D91): company
   storage, real data allowed, only Boss for now, reserved and filled as KRONOS grows, with a monthly report for his
   manager. The plan (S0) is done, and every reference file follows it (D92). Before March: S1
   `vclc drive init/report/check`, S2 KRONOS saving into `work` (with KRONOS Offline Phase 3), S3 the simulator and
   memory writing to the drive, S4 the day-one kit; from March, S5 to S8. Needed from Boss in March: the drive
   letter, the graphics card, IT's answers and backup.

Update this section at the end of every session, and keep `docs/kgc-kronos-context.md` (the
context file Boss gives to other Claude chats) in step with it: edit its sections 1 to 12 by hand,
then run `python compiler/gen_context.py`, which rewrites its appendices from the repo (D54).
