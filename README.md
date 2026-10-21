# VBA Forge

VBA Forge is an offline, VBA-only code generator for Excel automation in pressure-vessel
fabrication registers (BOM, weld log, NDE, NCR, calibration and similar). Claude writes each task
as a spec in VCL (VBA Command Language), and VBA Forge turns the spec into VBA and proves the
macro works with tests. Each verified spec, with its task in plain English and its proven VBA,
becomes training data for our own model, designed and trained from zero on in-house data only
(`DECISIONS.md` D25). From March 2027 that model writes VCL from plain English, and the compiler
still writes all the VBA (D31, `docs/ai-plan.md`).

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

- `dist/forge_selftest.bas` (committed): the ForgeLib self-test, every conformance case run
  through ForgeLib in Excel (D34)
- `dist/forge_pilot.bas` (committed): the Excel pilot of task_12; paste its results back for
  `python compiler/vclc.py pilot-check RESULTS specs/task_12.vcl` (D35, `docs/excel-pilot.md`)
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
- **2026-10-17:** VCL 0.2, PR A: the six value functions SAME, CLEAN_TEXT, IS_ERROR,
  PARSE_DATE, IIF and MOD, through the whole toolchain (semantics and 52 new conformance
  cases, validator, type checker, interpreter, compiler, ForgeLib). They work only in specs
  that start with `VCL 0.2`; the 10 specs and the Orchestrator prompt stay `VCL 0.1` until
  the migration PR.
- **2026-10-17:** VCL 0.2, PR B: SKIP_IF, EXIT_LOOP_IF and WRITE_ROW. SET_VALUE, WRITE_ROW and
  COPY_ROW now keep text as text in Excel (`ForgeSetValue`), closing a gap between the
  interpreter and Excel for keys like "1001".
- **2026-10-17:** VCL 0.2, PR C: MODIFIES says which sheets, columns and fills a spec may change.
  The validator checks every write against it, and the interpreter checks every TEST cell by
  cell (UNCHANGED). Nothing changes for the 10 specs until PR D adds MODIFIES lines to them.
- **2026-10-17:** VCL 0.2, PR D1: `compiler/vcl_migrate.py` moved the 10 specs to VCL 0.2 (SAME,
  SKIP_IF, WRITE_ROW, MODIFIES), with proven rewrites only: every TEST gives the same RESULT
  and final workbook, and the TESTs miss no more mutants than before (15). The new SKIP_IF
  lines expose 22 mutants no TEST catches, because no fixture has blank key rows; with them
  the snapshot score reads 90.9%. The Orchestrator prompt switches in PR D2.
- **2026-10-17:** VCL 0.2, PR D2: spec.md and the Orchestrator prompt teach VCL 0.2 (the new
  functions, SKIP_IF, EXIT_LOOP_IF, WRITE_ROW, MODIFIES and hard rule 20); Example 1 migrated.
  The Orchestrator Project on claude.ai needs the new prompt.
- **2026-10-17:** Repo renamed KGC-KRONOS. Phase 0 of the Job-Pack Router plan
  (`docs/vep-0.3-file-ops.md`): blank key rows in the fixtures, so every SKIP_IF is proven by a
  TEST. Snapshot mutation score 96.1% (370 of 385); the 15 left are the D14 equivalents.
- **2026-10-17:** FORMAT fill names (`fill=red` and the rest) now work in the interpreter and
  the compiler (PR #18). task_11 HighlightRejectedWelds added: Boss's weld-log spec from the
  Orchestrator, unchanged. It checks, passes its 5 TESTs and compiles to lint-clean VBA. Snapshot
  mutation score 95.1% (389 of 409); task_11 catches 19 of 24, and its 5 survivors are
  explained in D24.
- **2026-10-18:** Standing rule D25, "Original model and context data": the March 2027 model is
  designed and trained from zero on in-house data only (no pre-trained weights, no outside
  code, datasets, prompts or retrieval). The roadmap .docx, CLAUDE.md and the context file say so.
- **2026-10-18:** Job-Pack Router Phase 1a: the `VCL 0.3` line and SAFE_NAME, PATH_JOIN and
  FOLDER_EXISTS. The type checker accepts a path only when it starts from a String `$argument`
  and every part is SAFE_NAME(...) or a safe literal, so no path can climb out of its root.
  Every VCL 0.2 spec gives the same results under 0.3. Snapshot mutation score unchanged (95.1%).
- **2026-10-18:** Job-Pack Router Phase 1b: MAKE_FOLDER and COPY_FILE (each writes a status
  into a String variable instead of stopping the run), `MODIFIES $root files`, and folder
  FIXTUREs with locked files. The validator checks every rule; the mock file system that runs
  them comes in Phase 2.
- **2026-10-18:** Job-Pack Router Phase 2: the interpreter runs FOLDER_EXISTS, MAKE_FOLDER and
  COPY_FILE on a mock file system in memory, with a status for every case (Copied, Already
  there, Conflict, Locked, Missing, No folder, Bad name, Too long, ...), the FILE_EXISTS and
  FILE_COUNT assertions, and an UNCHANGED check for files. A sample router spec passes all its
  TESTs, which catch 29 of its 31 mutants (the other 2 change nothing). Phase 3 waits for Boss.
- **2026-10-19:** Job-Pack Router Phase 3a: the compiler turns SAFE_NAME, PATH_JOIN,
  FOLDER_EXISTS, MAKE_FOLDER and COPY_FILE into calls of new ForgeLib helpers. They use a
  late-bound FileSystemObject, never overwrite, delete or retry, and return the interpreter's
  status for every case. The linter now bans Kill, Shell, deleting, moving, renaming, the
  registry and SaveAs, and allows file writes only inside those helpers (D29).
- **2026-10-19:** Job-Pack Router Phase 3b: `forge_fixtures.bas` gets a sandbox builder,
  `ForgeSandbox_<file>()`, which makes each fixture's folders and files (exact sizes, locked
  files held open) in a new folder below `%TEMP%` and never touches anything else. A TEST now
  fails when a path is so close to the Windows length limit that Excel's longer sandbox root
  could change its status (D30). Phase 3 is done; Phase 4 (the router spec, task_12) waits for Boss.
- **2026-10-19:** The AI plan, `docs/ai-plan.md` (D31, approved by Boss): our own model will
  write VCL, never VBA, and the compiler stays the only thing that writes VBA, so every model
  output is proven by the toolchain without Excel. Planned VCL changes (a `GOAL` line, a
  formatter, error codes, fewer NOTE steps) come after the MVP.
- **2026-10-20:** Job-Pack Router Phase 4a: `specs/task_12.vcl` RouteJobPacks, the router MVP
  spec (VCL 0.3). For each tracker row it makes the job's folder and copies the job's report into
  it from the dump folder, never overwriting or deleting, and writes the status back. Its TESTs
  show every status on a row of its own, a re-run, and two PRE failures, and catch all 29 of its
  mutants. 12 specs; snapshot mutation score 95.4% (418 of 438). Next: 4b, the switch of the
  reference text and the Orchestrator prompt to VCL 0.3 (D32).
- **2026-10-20:** Job-Pack Router Phase 4b: spec.md and the Orchestrator prompt teach VCL 0.3.
  They have the file ops and their statuses, a new hard rule 21 for files, rule 19's checks for
  blank cells and locked files, a corrected ROLE line (the compiler writes the VBA), and a
  proven Example 3 that copies files (D33). Boss pastes the new prompt into the Orchestrator
  Project once.
- **2026-10-20:** Job-Pack Router Phase 5a: `dist/forge_selftest.bas`, the ForgeLib self-test.
  Pasted next to `forge_module.bas` in Excel, it runs all 163 conformance cases through ForgeLib
  and writes PASS or FAIL for each, plus one line to paste back. It is the first check of ForgeLib
  in real Excel 2016, and it runs only when Boss chooses (D34).
- **2026-10-20:** Job-Pack Router Phase 5b: `dist/forge_pilot.bas` runs task_12's TESTs in Excel,
  each on fresh synthetic sheets and a new sandbox, and writes lines to paste back;
  `vclc pilot-check` compares them with the interpreter line by line. `docs/excel-pilot.md`
  gives the steps. The router MVP is complete; running it in Excel is optional (D35).
- **2026-10-20:** Item 2 of the MVP plan, the rest of revision 3, is designed in
  `docs/vep-0.4-rev3-rest.md` (D36): seven phases, starting with the interpreter catching up
  with what already compiles. Ops that would open or save real workbooks wait until March.
- **2026-10-20:** Item 2, Phase 1: the interpreter now runs the array ops, INSERT_ROW,
  SELECT_CASE, LOG, MSG, RANGE, SHEET, FILE_NAME (S20) and FOR_EACH over SHEETS and CELLS, so
  specs that use them can be proven by TESTs. No language change (D37).
- **2026-10-20:** Item 2, Phase 2: DEDUPE runs everywhere (interpreter, compiler, ForgeLib), with
  key columns as a list, `cols=[1, 2]`, in the new `VCL 0.4` line. It keeps the first row of each
  key and never deletes a row with a blank or error key cell (S21, D38).
- **2026-10-20:** Item 2, Phase 3: FIND and IS_NOTHING run everywhere. FIND gives the first cell,
  row by row, that is SAME as the value (or holds it, with `match=part`), or NOTHING; ForgeFind
  loops itself instead of using Excel's Range.Find (S22, D39). Also fixed: the type checker now
  forgets cell guards after WRITE_ARRAY (#35).
- **2026-10-20:** Item 2, Phase 4a: ADD_SHEET and SHEET_EXISTS run everywhere, on a Workbook
  argument declared with `MODIFIES $book sheets`. A sheet is only ever added: a bad name (S23) or
  one already used stops the run first. Also fixed: the D20 UNCHANGED check now runs for VCL 0.4
  specs too (#37).
- **2026-10-20:** Item 2, Phase 4b: COPY_SHEET runs everywhere. It needs `name=` (checked first,
  never a replacement) and copies into the workbook of its sheet, under `MODIFIES $book sheets`
  (D41).
- **2026-10-20:** Item 2, Phase 5: FILTER runs everywhere. Below the header, a row stays visible when
  its cell is SAME as the value and is hidden otherwise; nothing is deleted and AutoFilter is never
  used. The new VISIBLE_COUNT assertion counts what stays visible. LASTROW no longer depends on
  hidden rows in Excel (D42).
- **2026-10-20:** Item 2, Phase 6: SET_FORMULA runs everywhere. A formula is checked against an Excel
  2016 allowlist (no Microsoft 365 function, nothing outside the workbook) and written in Excel's
  English form; TESTs check its text with FORMULA, since only Excel can calculate it (D43).
- **2026-10-20:** Item 2, Phase 7a: LIST_FILES and ITEMS run everywhere. A spec can list the files
  directly in a sandbox folder that match `*`/`?` patterns, in an order that is the same on every PC
  (S24, D44).
- **2026-10-20:** Item 2, Phase 7b: READ_TEXT and WRITE_TEXT run everywhere, each with a status
  (like COPY_FILE): text is printable ASCII only (S25), and WRITE_TEXT only ever adds a new file.
  Folder FIXTUREs can hold text files, and FILE_TEXT checks what was written (D45). Every phase of
  item 2 is now in the toolchain; spec.md and the prompt move to VCL 0.4 next.
- **2026-10-20:** Item 2, the switch: spec.md and the Orchestrator prompt teach VCL 0.4. They have
  every op item 2 made provable, with its meaning, a new hard rule 22 for sheets and formulas, and
  rules 10, 19, 20 and 21 grown to match; ops no TEST can prove (ACTIVE_SHEET, and OPEN_BOOK,
  SAVE_AS and CLOSE_BOOK until March 2027) are no longer taught. Every spec gives the same
  results under 0.4 (E12, D46). Boss pastes the new prompt into the Orchestrator Project once.
  Item 2 is complete.
- **2026-10-21:** Review pass after item 2. Two Orchestrator dry runs wrote nine specs from the
  prompt alone; where the prompt's words differed from the tools, the words were fixed (no
  meaning changed): `[Header]` ranges include the header cell, keys are not trimmed, assertions
  have no NOT, and the VAR types, fill colours and formula references are spelled out. A failing
  TEST now says which assertion failed and what it measured. Stale docs refreshed, and a new test
  keeps the reference text and the validator in step (D47).
- **2026-10-21:** ForgeLib fixes from a strict review of the item 2 VBA (D48). COPY_SHEET can no
  longer rename a sheet of the wrong workbook. Excel and the interpreter now agree on sheet names
  above U+7FFF and on DEDUPE keys. A formula never lands as text in a Text-formatted cell, and
  text-file paths must be plain ASCII.
- **2026-10-21:** CALL and BUILD specs designed in `docs/vep-0.5-build.md` (D49): a BUILD (several
  procedures making one tool) becomes provable the way every op of item 2 did. CALL keeps a
  Function's result, MODIFIES follows the arguments, BUILD TESTs run in the interpreter, and the
  compiler writes one module per BUILD. Forms and add-ins wait. Five phases, then one switch to
  VCL 0.5.
