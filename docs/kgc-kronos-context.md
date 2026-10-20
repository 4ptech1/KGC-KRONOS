# KGC KRONOS: AI context file

**Snapshot: 18 Oct 2026, `main` at ed5a793.** Hand this file to any Claude chat or claude.ai
Project that needs to understand the project. The repo is always the source of truth: if this
file and the repo disagree, the repo wins, and the repo's `CLAUDE.md` has the current status.

- Repo: https://github.com/4ptech1/KGC-KRONOS (renamed from `vba-forge` on 17 Oct 2026; old
  links redirect). Inside the code and docs the project is still called **VBA Forge**.
- Owner: **Boss** (GitHub `4ptech1`). Always call him Boss.

---

## 1. What the project is

VBA Forge (repo KGC-KRONOS) is an offline, VBA-only code generator for Excel automation at KGC
Engineering Projects. KGC does pressure-vessel fabrication, and its paper trail lives in Excel
registers: BOM, GRN, weld log, NDE (RT/UT/MPT/DPT), NCR, ITP, welder qualifications,
calibration, consumables.

How it works:

1. A request in plain English goes to the **Orchestrator** (a claude.ai Project). It replies
   with a **spec** written in **VCL** (VBA Command Language), a small, strict language.
2. The **toolchain** in this repo (Python, run in the cloud) checks the spec, runs its TESTs on
   synthetic data in a reference interpreter, and compiles it to Excel 2016 VBA that calls a
   shared runtime, **ForgeLib**.
3. Every task in plain English plus its proven spec (and the compiler's VBA) is a **verified
   pair**. From March 2027 those pairs train **our own model, designed and trained from zero**
   (D25). **The model writes VCL, never VBA** (D31, `docs/ai-plan.md`): from plain English, and
   as repairs of failing specs. The compiler stays the only thing that writes VBA, so every model
   output is proven by the toolchain before anything reaches Excel.

The goal until March 2027: a large, verified dataset of tasks and proven specs, and a toolchain
that proves every macro without Excel.

## 2. Who does what

| Who | Where | Job |
| --- | --- | --- |
| Boss | Locked-down work PC, browser, phone | Gives requests and decisions. Does no testing or debugging. Runs Excel only when he chooses. |
| Orchestrator | claude.ai Project, instructions = `vcl/orchestrator_prompt.txt` | Turns a request into exactly one VCL spec in one code block, or one `# QUESTION:` line. Never writes VBA. |
| Claude Code | Cloud session on the repo (claude.ai/code) | Builds and verifies everything: code, tests, CI, PRs. Merges its own PR when CI is green. |
| You (the reader of this file) | Any Claude chat | Advise, plan, review specs, explain. Changes to the repo go through Claude Code, as a short brief. |

Boss's working policy (from `CLAUDE.md`):
- He doesn't test, debug or verify. Everything is verified automatically (unit tests, the
  `vclc` pipeline, CI).
- One PR per item. When CI is green, Claude Code merges it with a merge commit and starts the
  next item, unless Boss has paused the work.
- If something needs his decision, pick the more conservative option, record it in
  `DECISIONS.md`, and carry on. Ask only when truly necessary.
- Never weaken a test, a check or a threshold to get to green.
- Plain words, brief answers; he is learning. For anything he must do by hand, give exact
  clicks and what he should see.
- When he asks for **"baby steps"**: one action per reply, as "🍼 Next action" plus
  "Expected:", then wait for his result.

## 3. Hard constraints (until March 2027)

- **Locked-down work PC:** no installs, no Excel Trust Center changes. All code runs in the
  cloud: Claude Code on the web, GitHub Codespaces, GitHub Actions. Macros can run on his PC
  (he confirmed this).
- **Excel 2016 only.** No XLOOKUP, FILTER(), LET(), LAMBDA, dynamic arrays, TEXTJOIN, IFS or
  any other Microsoft 365 or 2019 feature, in VBA or in formulas.
- **Synthetic data only.** No real KGC files, names, figures or customer data in the repo,
  tests or prompts. Excel workbooks are git-ignored on purpose.
- **Synthetic fixtures only.** Every TEST names a FIXTURE file; `fixture=target` is banned, and
  nothing generated ever touches real workbooks or real folders.
- **No model training or local-model work** until the new PC arrives in March 2027.
- **Original model and context data (D25, standing rule).** The VBA Forge model is designed and
  trained from zero by us: no pre-trained weights (Llama, Qwen, ...), no fine-tuning of someone
  else's model, no scraped or downloaded code, datasets or prompts (GitHub, Stack Overflow, public
  VBA collections), no retrieval over outside sources. In-house means made by this repo's tools, by
  Boss, or by Claude through our pipeline, and verified by our toolchain. General tools (Python,
  PyTorch) are allowed; Claude stays the design and build tool. Design: compiler first, own model
  second. The model writes VCL, never VBA (D31), and is used once it passes enough of the held-out
  benchmark; Claude stays the Orchestrator until then. Never suggest a base model or outside data.
- **Generated VBA rules:**
  - `Option Explicit`, and late-bound `Scripting.Dictionary` (no References).
  - The standard error handler, which logs to ForgeLog and restores ScreenUpdating and
    Calculation.
  - No Select or Activate. No Kill, RmDir, Shell, registry access, HTTP, API Declare, or SaveAs
    over an existing file.
  - Every loop is bounded.
  - Every VCL meaning goes through a ForgeLib helper.
- **Specs are inputs, not code to fix.** Never edit a spec just to make a check pass; report
  the problem to Boss. An edit he asks for is fine.

## 4. Repo map

```
CLAUDE.md                 standing brief for Claude Code (rules, status, next steps)
DECISIONS.md              every language decision (G, S, T, E, F ids) and Claude's own (D1-D25)
README.md                 overview, how to run the checks, progress log, dataset count
requirements.txt          one pinned dependency: lark==1.3.1 (with hash)
vcl/spec.md               THE language reference (VCL 0.3)
vcl/orchestrator_prompt.txt  the Orchestrator's instructions; shares the reference text with spec.md
vcl/grammar.lark          formal LALR grammar
vcl/semantics.py          reference meaning of the value functions (S1-S17)
vcl/examples/             example1_rev3.vcl (= the prompt's Example 1), build_example.vcl (a sketch, invalid)
specs/task_01..12.vcl     the 12 task specs (11 in VCL 0.2; task_12, the router, in 0.3)
compiler/vclc.py          the toolchain CLI: check, test, build, all, mutate
compiler/parser.py        hand-written parser   compiler/validator.py  rules, plus grammar and type stages
compiler/vcl_parse.py     grammar parser        compiler/vcl_types.py  type checker (guards, narrowing)
compiler/vcl_interp.py    reference interpreter, TEST assertions, the UNCHANGED check
compiler/vcl_codegen.py   VCL to VBA compiler   compiler/vcl_fixtures.py  the one fixture generator
compiler/vba_lint.py      static VBA checks     compiler/vcl_mutate.py  mutation scorer
compiler/vcl_migrate.py   VCL 0.1 to 0.2 rewriter, with --prove
compiler/gen_semantics_doc.py  writes docs/vcl-semantics.md
compiler/vcl_formula.py   SET_FORMULA's Excel 2016 allowlist and checks (D43)
vba/ForgeLib.bas          the VBA runtime every generated procedure calls
dist/forge_module.bas     committed build output: ForgeLib plus one procedure per spec (paste-ready)
dist/forge_selftest.bas   committed: the ForgeLib self-test, every conformance case run in Excel (D34)
dist/forge_pilot.bas      committed: the Excel pilot of task_12; results checked by `vclc pilot-check` (D35)
docs/excel-pilot.md       Boss's steps for the optional Excel pilot
dist/forge_fixtures.bas, dist/procedures/, dist/expected/   rebuilt by vclc, git-ignored
tests/                    338 unit tests; conformance_cases.py is shared with the docs
tests/data/specs_v01/     the 10 specs and Example 1 frozen as VCL 0.1 (migration proof)
tests/data/task_01_pilot.vcl  task_01 as the Excel pilot ran it (pinned seed tests)
docs/vcl-mvp-plan.md      build sequence        docs/vep-0.2-core.md  VCL 0.2 design (done)
docs/vep-0.3-file-ops.md  file ops and Job-Pack Router design and roadmap (approved, Phases 0 to 4 done)
docs/vcl-semantics.md     generated conformance tables   docs/vcl-m2-findings.md  M2 research
docs/VBA_Forge_Project_Roadmap.docx  the full roadmap (phases, gates)
harness/, batches/        Excel batch runner, deferred (PR #4)
templates/, data/, library/, logs/  placeholders for later phases
.github/workflows/        CI: unit tests, validator, vclc all, dist and docs up to date
```

## 5. How a spec is verified (no Excel needed)

`python compiler/vclc.py all specs/*.vcl -o dist` runs:

1. **check:** the validator, in three stages: the spec.md rules, the grammar, and the type
   checker. The type checker rejects any calculation or date comparison on a raw cell that
   isn't guarded by `IS_NUMBER` or `IS_DATE` in an enclosing IF.
2. **test:** every TEST runs in the reference interpreter, on fixture data from
   `vcl_fixtures.py`, using the meanings in `vcl/semantics.py`. For specs with MODIFIES, the
   automatic **UNCHANGED** check compares every cell before and after the run.
3. **build:** the compiler writes `forge_module.bas`, the fixture builders and the expected
   results, then `vba_lint.py` lints the VBA (0 problems required).

`python compiler/vclc.py mutate specs/*.vcl --snapshot` plants small deliberate bugs (flipped
comparisons, deleted steps, off-by-one numbers) and counts how many the TESTs catch. This is
the measure of test strength.

Before every commit Claude Code runs: the unit tests, the validator, `vclc all`, and
`gen_semantics_doc.py --check`. CI runs the same, and fails if `dist/forge_module.bas` or
`docs/vcl-semantics.md` is stale.

## 6. VCL 0.3 cheat sheet (summary only; the full text is in Appendix A)

**TASK keys, in this order:** `VCL 0.3`, `TASK`, `# ASSUME:` lines, PROC, KIND (Sub or
Function), MODULE, ARGS, RET, VAR, PRE, MODIFIES, OP, PERF, ERR standard, ALLOW, FIXTURE, TEST.
- Arguments are written `$name`.
- OP lines nest by 2 extra spaces.
- Option names (dir, from, to, var, in, mode, into, key, order, header, what, fill, numfmt,
  bold, width, style, cols, col, crit, readonly, format, save, after, name, match) are reserved
  words.

**Structural ops:**
- loops: `LOOP_ROWS dir=up|down`, `LOOP_COLS`, `FOR_EACH`
- `IF ... THEN <op>` (one line), or a block `IF` with optional `ELSE`
- `SELECT_CASE`
- `EXIT_IF` (leaves the procedure)
- `SKIP_IF` and `EXIT_LOOP_IF` (inside loops only)

**Call ops:**
- LET, READ_ARRAY, WRITE_ARRAY, DICT_ADD (mode=overwrite|keep|sum|count), DICT_GET
- SORT, FILTER, DEDUPE
- SET_VALUE, WRITE_ROW, SET_FORMULA, FORMAT, CLEAR, COPY_ROW, INSERT_ROW, DELETE_ROW
- files (VCL 0.3): MAKE_FOLDER and COPY_FILE, each writing a status into a String VAR
- sheet and other file ops: ADD_SHEET … WRITE_TEXT
- CALL, LOG, MSG, RETURN
- NOTE (the escape hatch)

**Value functions:**
- worksheet: COL, CELL, LASTROW, LASTCOL, ROW_RANGE, USED_RANGE, RANGE, SHEET
- arrays and dictionaries: ARR, ROWS, COLS, COUNT, EXISTS
- text: TRIM, UPPER, LOWER, LEN, LEFT, RIGHT
- checks: IS_EMPTY, IS_NUMBER, IS_DATE, IS_ERROR, IS_NOTHING
- constants and files: TODAY, NOW, FILE_NAME
- new in 0.2: SAME, CLEAN_TEXT, PARSE_DATE (with a required format), IIF, MOD
- new in 0.3: SAFE_NAME, PATH_JOIN, FOLDER_EXISTS

**Meanings that surprise people** (S1 to S17, with tables in `docs/vcl-semantics.md`):
- **Emptiness and trimming:** IS_EMPTY is true for spaces-only text, including non-breaking
  spaces. TRIM also strips non-breaking spaces, tabs and line breaks.
- **Numbers and dates:** IS_NUMBER accepts numeric text like `" 25 "`, but not `"1,250"`.
  IS_DATE is true only for real dates.
- **Text form:** every value becomes text the same way on every PC. Dates become yyyy-mm-dd and
  whole numbers lose ".0".
- **`=` versus SAME:** `=` with text compares exactly, so case and spaces count. SAME compares
  trimmed text, ignoring case.
- **Lookups:** COL matches headers ignoring case and spaces, and dictionary keys ignore case,
  so 1001 and "1001" are one key.
- **AND / OR and IIF:** both sides are always evaluated, so a guard on the left never protects
  the right.
- **Exits and PRE:** EXIT_IF leaves the procedure. A failed PRE logs once, changes nothing, and
  makes a Function return -1.
- **Writes:** SET_VALUE, WRITE_ROW and COPY_ROW keep text as text in Excel (S17).
- **Files (VCL 0.3):** every path is `PATH_JOIN($root, SAFE_NAME(...), ...)`, written in place,
  so no path can climb out of its root. MAKE_FOLDER and COPY_FILE never stop the run: they write
  a status (`Created`, `Exists`, `Copied`, `Already there`, `Conflict`, `Locked`, `Missing`,
  `No folder`, `Blocked by file`, `Bad name`, `Too long`). Nothing moves, overwrites or deletes
  a file.

**MODIFIES** (a key after PRE):
- the three forms:
  - `MODIFIES $x`: the whole sheet
  - `MODIFIES $x["Header", ...]`: values in those columns, row 2 down
  - `MODIFIES $x fills`: fills only
  - `MODIFIES $root files` (VCL 0.3): folders and files may be added below that String argument
- **Before the run**, the validator checks every write against these lines.
- **After every TEST**, the interpreter checks that nothing else changed (UNCHANGED).
- **A spec that writes nothing** needs no MODIFIES lines, and every TEST checks it changed
  nothing.

**The 21 hard rules (short form):**
1. One `vcl` code block, nothing else.
2. Never write VBA.
3. Use only the defined syntax.
4. NOTE rarely.
5. At least two TESTs: one proves the goal, one proves nothing else changed. ERROR_COUNT = 0 at
   least once. Add a wrong-sheet TEST that expects RESULT = -1 when there's another sheet.
6. Assumptions go on `# ASSUME:` lines at the top.
7. Ask a `# QUESTION:` only when a guess could destroy data, or when the request can't be
   specified at all.
8. ALLOW only for risky actions the user asked for.
9. TASK for about 60 VBA lines or less, BUILD for more.
10. Delete rows only in `dir=up` loops, and never because data is missing.
11. Use arrays for more than about 1,000 rows.
12. Naming: procedures PascalCase and verb-first, modules `mod` + PascalCase, variables
    camelCase, fixtures snake_case_01.xlsx.
13. If the message isn't a new request, answer in two sentences.
14. Find columns with COL(), and check every header in PRE.
15. Trim IDs, compare typed text with SAME, and guard calculations with IS_NUMBER or IS_DATE
    in a nested IF.
16. Re-runs are safe: the goal TEST uses runs=2.
17. Prefer a Function that returns a count.
18. Synthetic FIXTUREs only.
19. Fixtures must be able to catch a wrong macro:
    - stale rows on rebuilt sheets
    - `+boundary(...)` values exactly on each comparison
    - repeated dictionary keys
    - `reldates` when the spec uses TODAY
    - blank cells in any column a SKIP_IF tests
    - a `locked=` file when the spec copies files
20. Declare every sheet the macro changes with MODIFIES (and `files` for a root it adds to).
21. Files only through MAKE_FOLDER and COPY_FILE on `PATH_JOIN($root, SAFE_NAME(...))` paths;
    every root checked with FOLDER_EXISTS in PRE; nothing moves, overwrites or deletes a file.

**FIXTURE syntax:**
`FIXTURE <file>.xlsx sheet=<Name> cols=A:Header,... rows=<n> [stale=<n>] data=<col>:<rule>;...`
- Rules: values("a"x3,...), seq, int, num, dates, reldates, text.
- Modifiers, in this order: `+boundary(...)`, `+blank(n)`, `+dupes(n)`.
- The data is deterministic: each column has its own seed, made from file, sheet and column.
- VCL 0.3 folders: `FIXTURE <file>.xlsx folder=<Name> files=a.pdf:120,J-1\b.pdf locked=a.pdf`;
  TESTs pass the folder bare, like a sheet. They live in a sandbox, never on real folders.

**TEST assertions:**
- measures: COUNTIF, COUNTIFS, COUNTA, SUM, SUMIFS, VALUE, ROWS_USED, FILL_COUNT
- files (VCL 0.3): `FILE_EXISTS("Folder\path")`, `FILE_COUNT("Folder")`
- run results: RESULT, ERROR_COUNT
- PRE_<measure>, taken before the run
- ranges by header: `Sheet![Header]` and `Sheet![Header]<row>`

## 7. What the toolchain can run today (important when reviewing specs)

The validator accepts the whole language, but the interpreter and compiler implement only part
of it. A spec outside this subset passes the validator but fails at test or build.

| Fully supported (validate, test, compile) | Compile only (no interpreter yet) | Not built yet |
| --- | --- | --- |
| LET, SET_VALUE, WRITE_ROW, CLEAR, FORMAT, COPY_ROW, DELETE_ROW, DICT_ADD, DICT_GET, SORT, RETURN, MAKE_FOLDER, COPY_FILE (VCL 0.3); LOOP_ROWS, LOOP_COLS, FOR_EACH over KEYS(...), IF/ELSE, EXIT_IF, SKIP_IF, EXIT_LOOP_IF; CELL, COL, LASTROW, LASTCOL, ROW_RANGE, USED_RANGE, TRIM, UPPER, LOWER, LEN, LEFT, RIGHT, IS_EMPTY, IS_NUMBER, IS_DATE, IS_ERROR, SAME, CLEAN_TEXT, PARSE_DATE, IIF, MOD, COUNT, EXISTS, KEYS, SAFE_NAME, PATH_JOIN, FOLDER_EXISTS; MODIFIES (with `files`); folder FIXTUREs; assertions FILE_EXISTS, FILE_COUNT; and since item 2 Phase 1 (D37): INSERT_ROW, READ_ARRAY, WRITE_ARRAY, LOG, MSG, SELECT_CASE, ARR, ROWS, COLS, SHEET, RANGE (a cell or a block), FILE_NAME, THIS_BOOK, ACTIVE_BOOK; FOR_EACH over SHEETS or CELLS; and since Phase 2 (D38): DEDUPE, with `cols=[1, 2]` in a `VCL 0.4` spec; and since Phase 3 (D39): FIND, IS_NOTHING, NOTHING (write to a found cell through FIND written in place); and since Phase 4a (D40): ADD_SHEET and SHEET_EXISTS on a `$book` Workbook argument with `MODIFIES $book sheets` (a TEST passes the workbook as its fixture file, `call=P(ncr_01.xlsx, Ncr)`); and since Phase 4b (D41): COPY_SHEET with `name=`; and since Phase 5 (D42): FILTER (hides rows by SAME, whole-sheet MODIFIES, not with SORT in one spec) and the VISIBLE_COUNT assertion; and since Phase 6 (D43): SET_FORMULA (a quoted formula on the Excel 2016 allowlist; not with ops that move rows) and the FORMULA assertion | CALL; FOR_EACH over ITEMS | OPEN_BOOK, SAVE_AS, CLOSE_BOOK, LIST_FILES, READ_TEXT, WRITE_TEXT, NOTE; BUILD specs |

**Fill colours:** `fill=` takes `none`, a name (black, white, red, green, blue, yellow, magenta,
cyan) or `#RRGGBB`. Each name is the colour of the VBA constant of that name (red = `#FF0000`), so
`fill=red` and `fill=#FF0000` give the same result (D23, fixed in PR #18).

## 8. Status (20 Oct 2026)

**Roadmap Phase 0, Foundation (14 to 31 Oct 2026).** Gate 0: all specs validate, type-check,
pass their TESTs in the interpreter, and compile to lint-clean VBA, in CI. **This is met.**

| Measure | Now |
| --- | --- |
| Specs | 12 (task_01 to task_12): 11 in VCL 0.2, and task_12 (the router) in VCL 0.3 |
| Unit tests | 338 pass |
| Validator | 0 errors; 1 known warning (task_04 has 48 OP lines, over the ~40 guideline; accepted until March) |
| Mutation score (snapshot) | 95.4% (418 of 438); task_12 catches 29 of 29; 19 of the 20 survivors are equivalent mutants (D14, D24) that no test can ever catch; 1 is a small fixture gap in task_11 (D24) |
| VBA lint | 0 problems |
| Verified pairs | 0 (none has been run in Excel yet) |
| Orchestrator Project | the VCL 0.3 prompt is ready (router Phase 4b); once Boss pastes it in, replies start with `VCL 0.3` and can copy files |

The 12 specs:

| Spec | Procedure | Sheets (ARGS) |
| --- | --- | --- |
| task_01 | BuildShortageList | bom, grn, out |
| task_02 | FlagMissingTraceability | ws |
| task_03 | BuildWelderRepairRate | nde, out |
| task_04 | BuildNdeSummary | rt, ut, mpt, dpt, out |
| task_05 | ListWeldLogExceptions | weldLog, out |
| task_06 | FlagExpiredWelderQualifications | weldLog, qual |
| task_07 | BuildCallNoticeList | itp, out |
| task_08 | ReconcileConsumableStock | issues, inv |
| task_09 | BuildOpenNcrSummary | ncr, out |
| task_10 | BuildCalibrationDueList | cal, out |
| task_11 | HighlightRejectedWelds | ws |
| task_12 | RouteJobPacks (VCL 0.3) | tracker; dump and jobs are folder paths |

**History in one line each** (PRs on KGC-KRONOS; #1 to #15 were made under the old name):
- #1: repo skeleton and the 10 specs.
- #2: VCL 0.1 revision 3.
- #3: the parser and validator.
- #8: MVP toolchain integrated, with revision 4.
- #9: revision 5 fixtures, which catch wrong macros.
- #10: VCL 0.2 design.
- #11: value functions.
- #12: SKIP_IF, EXIT_LOOP_IF, WRITE_ROW and ForgeSetValue.
- #13: MODIFIES and UNCHANGED.
- #14: the 10 specs migrated by `vcl_migrate.py`.
- #15: spec.md and the prompt teach 0.2.
- #16: blank key rows (score 90.9% to 96.1%).
- #17: this context file.
- #18: FORMAT fill names (`fill=red`) work in the interpreter and compiler.
- #19: task_11 HighlightRejectedWelds, Boss's weld-log spec from the Orchestrator.
- #20 and #21: this context file brought up to date; D25 (our own model, from zero).
- #22 to #24: router Phases 1 and 2 (the VCL 0.3 language surface and the mock file system).
- #25 and #26: router Phase 3 (the compiler, ForgeLib's file helpers, the sandbox builder).
- #27: the AI plan, D31 (the model writes VCL, never VBA).
- #28: task_12 RouteJobPacks, the router MVP spec.
- #29: spec.md and the prompt teach VCL 0.3.
- #30 and #31: the Excel pilot kit (the ForgeLib self-test; the task_12 pilot and `pilot-check`).
- #32: the design of item 2, the rest of revision 3 (`docs/vep-0.4-rev3-rest.md`, D36).
- #33: item 2 Phase 1, the interpreter catches up with what already compiles (D37).
- #34: item 2 Phase 2, the `VCL 0.4` line, list syntax and DEDUPE (S21, D38).
- #35: the type checker forgets cell guards after WRITE_ARRAY.
- #36: item 2 Phase 3, FIND and IS_NOTHING (S22, D39).
- #37: the D20 UNCHANGED check runs for VCL 0.4 specs too.
- #38: item 2 Phase 4a, ADD_SHEET and SHEET_EXISTS (S23, D40).
- #39: item 2 Phase 4b, COPY_SHEET (D41).
- #40: item 2 Phase 5, FILTER and VISIBLE_COUNT (D42).
- #41: item 2 Phase 6, SET_FORMULA and FORMULA (D43).

**Still open on GitHub:**
- #4, the ForgeTest Excel pilot. Deferred, and nothing depends on it.
- #5, #6 and #7: older proposal documents (Phase 1 plan, Phase 2 dataset plan, VCL 0.2
  proposal). They are parked, and #7 is largely superseded by the VCL 0.2 that was actually
  built.

## 9. The plan

**Done: the Job-Pack Router MVP** (`docs/vep-0.3-file-ops.md`, approved by Boss; complete on 20 Oct 2026).

A single-task MVP that proves ForgeLib in Boss's Excel 2016:
- It reads a tracker of Job IDs and report files.
- It creates job folders and copies each PDF test report from a dump folder into its job
  folder.
- It writes a status back to the tracker.

Boss's decisions:
- **Sandbox only:** it never runs on real folders before March 2027.
- **COPY_FILE first:** MOVE_FILE comes later and will need `ALLOW file_delete`.
- **VCL 0.3** carries the new features.

| Router phase | Content | State |
| --- | --- | --- |
| 0 | Blank key rows so every SKIP_IF is proven | **Done** (PR #16) |
| 1a | Language surface, part 1: the `VCL 0.3` line; SAFE_NAME, PATH_JOIN, FOLDER_EXISTS; the Name and Path types (T4); the meaning of SAFE_NAME and PATH_JOIN with conformance cases (S18, S19, moved here from Phase 2, D26) | **Done** (PR #22) |
| 1b | Language surface, part 2: MAKE_FOLDER and COPY_FILE returning a status into a String VAR (E9); `MODIFIES $jobs files` (E10); folder FIXTUREs (F6) | **Done** (PR #23) |
| 2 | A mock file system in fixtures and the interpreter; FOLDER_EXISTS and the file ops run on it with their statuses (E11); FILE_EXISTS and FILE_COUNT assertions (F7); UNCHANGED for files | **Done** (PR #24) |
| 3a | Compiler and ForgeLib: late-bound FSO helpers that return the interpreter's statuses and never overwrite, delete or move; lint bans Kill, Shell, deleting, moving, the registry and SaveAs, and allows file writes only inside the helpers (D29) | **Done** (PR #25) |
| 3b | A sandbox builder in `%TEMP%` for the folder FIXTUREs (`ForgeSandbox_<file>()`); a TEST fails when Excel's longer sandbox root could change a Too long status (D30) | **Done** (PR #26) |
| 4a | `specs/task_12.vcl` RouteJobPacks: every status on a row of its own, a re-run, two PRE failures; 29 of 29 mutants caught (D32) | **Done** (PR #28) |
| 4b | The shared reference text and the prompt switch to 0.3 (hard rule 21, Example 3), and Boss updates the Orchestrator again (exact clicks) (D33) | **Done** (PR #29); Boss pastes the prompt |
| 5a | The ForgeLib self-test: every conformance case run through ForgeLib in Excel (`dist/forge_selftest.bas`, D34) | **Done** (PR #30) |
| 5b | The task_12 pilot (`dist/forge_pilot.bas`), `vclc pilot-check`, and Boss's guide `docs/excel-pilot.md` (D35) | **Done**; running it in Excel is optional |

The key safety design for files:
- **Every path is built** with `PATH_JOIN` from a `$root` argument and `SAFE_NAME` parts.
  SAFE_NAME refuses `..`, `\ / : * ? " < > |`, reserved names like CON, and trailing dots.
- **Every file op returns a status** into a String VAR (E11). MAKE_FOLDER: `Bad name`,
  `Too long`, `Exists`, `Blocked by file`, `No folder`, `Created`. COPY_FILE: `Bad name`,
  `Too long`, `Missing`, `No folder`, `Already there`, `Conflict`, `Locked`, `Copied` (and in
  Excel `Check copy` or `Failed`). An FSO error never aborts the whole run.
- **Tests run on a mock file system** in memory (F7): FIXTURE folders sit below
  `C:\KRONOS_SANDBOX`; nothing real is read or written.
- **No overwrite, no delete, no retry loops.**

**After the MVP** (`docs/vcl-mvp-plan.md`):
1. The rest of revision 3 in the interpreter and compiler (Boss said "go item 2" on 20 Oct
   2026), designed in `docs/vep-0.4-rev3-rest.md` (D36): first the interpreter catches up with
   what already compiles, then list syntax with DEDUPE, FIND, the sheet ops, FILTER (it hides
   rows by SAME, never by AutoFilter criteria), SET_FORMULA (tested by its text) and text files
   (new files only). OPEN_BOOK, SAVE_AS and CLOSE_BOOK wait until March 2027. Phases 1 (D37)
   and 2 (D38: the `VCL 0.4` line in the tools, `cols=[1, 2]`, DEDUPE), 3 (D39: FIND) and 4a (D40:
   ADD_SHEET, SHEET_EXISTS), 4b (D41: COPY_SHEET), 5 (D42: FILTER) and 6 (D43: SET_FORMULA) are
   done; next is Phase 7, text files in the sandbox. The Orchestrator stays on VCL 0.3 until all of 0.4 is in.
2. A corpus of 150 tasks, Orchestrator-written specs, and an expressibility report.
3. A dataset generator: tasks in plain English with their proven specs, plus repair pairs.
   Mutants double as "read carefully" contrast pairs.
4. **March 2027:** the new PC; train our own model from zero on in-house data (D25). It writes
   VCL, never VBA (D31). Excel conformance runs.

**The AI plan** (`docs/ai-plan.md`, D31, approved by Boss on 19 Oct 2026): the model's jobs are
plain English to VCL, and repairing failing specs. NOTE steps shrink as VCL grows; until then
they go to Claude or Boss. Planned VCL changes, each its own VEP after the MVP: a `GOAL` line
(one English sentence per TASK), a canonical formatter, error codes in checker messages, fewer
NOTE steps, a small language, and a minimum mutation score for model-written TESTs. Stages: A and
B before March (corpus, data generator, benchmark, growing VCL), C in March (train a small model
from zero, grammar-constrained), D (repair; used once it passes enough of the benchmark), E
(Boss's daily use).

**Pending Boss decisions right now:**
1. When, if ever, to run the optional Excel pilot (`docs/excel-pilot.md`).
3. task_11 (HighlightRejectedWelds) has one small test gap (D24): if the macro's last-row scan
   skipped column A, no TEST would notice. Closing it needs a FIXTURE whose last row has a value
   only in column A. That is a spec change, so it waits for Boss.

## 10. Decisions worth knowing (full list in `DECISIONS.md`)

- **D1:** one fixture generator, a SplitMix64 PRNG seeded from SHA-256 of
  "file|sheet|column". The data never depends on the Python version.
- **D3:** only `dist/forge_module.bas` is committed (and, since D34, `dist/forge_selftest.bas`).
  The expected results depend on TODAY.
- **D12, D20:** frozen copies are kept instead of rewriting pinned tests.
  `tests/data/specs_v01` proves the 0.1 to 0.2 migration in CI.
- **D14:** the 15 surviving mutants are equivalent, and still count against the score.
- **D16:** VCL 0.2 was designed before it was built, and lands in small PRs. The prompt
  changes once, at the end.
- **D17:** SAME uses a binary compare of lower-case text, not the PC's locale. PARSE_DATE
  rejects years before 1900.
- **D18, S17:** every cell write goes through `ForgeSetValue`, which adds a hidden `'` so Excel
  never turns "1001" or "#N/A" into something else.
- **D19:** MODIFIES checks the kind of change: values, fills, or formats and rows. A write that
  isn't through a `$sheet` argument is an error.
- **D20:** mutation proofs compare missed mutants, not percentages.
- **D22:** blank key rows for SKIP_IF, plus the matching rule 19 warning.
- **D23:** FORMAT fill names work: each name is the VBA colour constant of that name.
- **D24:** Boss's weld-log spec is task_11, unchanged. The router spec becomes task_12.
- **D25:** original model and context data: the model is trained from zero, on in-house data only.
- **D31:** the model writes VCL, never VBA; the compiler writes all VBA (`docs/ai-plan.md`).

## 11. How to help (rules for the AI reading this)

- **Check the language, not your memory.** Before writing or reviewing VCL, read `vcl/spec.md`,
  or the cheat sheet above. Never invent keys, ops or functions; if something can't be said,
  it is a NOTE or a new VEP.
- **Review a spec the way the toolchain does.**
  - Correct syntax and key order.
  - Every input column found with COL() and checked in PRE.
  - Guards on every calculation, done with the check itself, never with a flag.
  - MODIFIES matches the writes.
  - TESTs: a goal TEST with runs=2, a collateral TEST, ERROR_COUNT = 0, and a wrong-sheet TEST.
  - Fixtures follow rule 19.
  - Only the supported subset from section 7 is used.

  Then say that the real proof is Claude Code running `vclc all` and `vclc mutate` on it.
- **Don't claim anything is verified** unless the toolchain or CI did it. Never fake test
  results.
- **Respect the constraints:** no Microsoft 365 functions; no local installs or Trust Center
  changes for Boss; no real KGC data; no real files or folders touched; nothing outside the
  sandbox.
- **For a repo change,** write a short brief for Claude Code: what, why, the constraints, and
  "done when". Claude Code then makes one PR, verifies it and merges it.
- **Any change to VCL itself** means updating spec.md and the prompt together, plus the grammar,
  parser, validator, types, interpreter, compiler, ForgeLib, conformance cases, docs, tests and
  DECISIONS.md. Boss then updates the Orchestrator Project, and you give him exact clicks.
- **Talk to Boss** in plain words, briefly, as "Boss". Use 🍼 baby steps when he asks for them.

## 12. Glossary

| Term | Meaning |
| --- | --- |
| VCL | VBA Command Language: the spec language (`vcl/spec.md`). Current version 0.3 (20 Oct 2026). |
| Spec | One `.vcl` file: one TASK (a procedure) or one BUILD (several procedures). |
| Orchestrator | The claude.ai Project that writes specs from plain requests. |
| ForgeLib | `vba/ForgeLib.bas`, the VBA runtime. Every VCL meaning is a Forge* helper. |
| Fixture | Synthetic test data that a FIXTURE line describes; built the same way every time. |
| stale / boundary / blank / dupes | Fixture parts that let a TEST catch a wrong macro: leftover rows, exact edge values, blank cells, repeated keys. |
| Interpreter | `compiler/vcl_interp.py`, which runs a spec's TESTs without Excel. |
| UNCHANGED | The automatic check that nothing outside MODIFIES changed. |
| Mutant | A copy of a spec with one deliberate bug. "Caught" means a TEST notices it. |
| Equivalent mutant | A mutant that can't change any result, so no test can ever catch it (D14). |
| VEP | A design document for a language change, written before any code (`docs/vep-*.md`). |
| Verified pair | A task in plain English plus its proven spec (and the compiler's VBA); the training data for March 2027 (D31). |
| Gate 0 | The Phase 0 exit test: everything validates, passes, compiles and lints, in CI. |

## Appendix A: the full VCL 0.3 reference

Copied word for word from `vcl/spec.md`, the same text the Orchestrator prompt contains. A unit
test fails if this copy ever differs from spec.md. The parts of spec.md not copied here are the
exact fixture algorithm and the changelogs.

### Hard rules

```text
1.  Reply with exactly one code block tagged vcl. No text before or after it.
2.  Never write VBA. Express all logic with VCL keys, ops and expressions only.
3.  Use only the keys, ops, functions and syntax defined here. Never invent new ones.
4.  If a step truly cannot be expressed with ops, write it as NOTE "<terse technical text>".
    Use NOTE rarely.
5.  Every spec has at least two TEST lines: one proves the goal was met, one proves nothing
    else changed (for read-only work: proves the source data is unchanged). At least one
    TEST includes ERROR_COUNT = 0. When the test workbook has another sheet, add a
    wrong-sheet TEST that passes it and expects RESULT = -1 AND ERROR_COUNT = 1.
6.  Put every assumption at the top of the spec as a line starting with # ASSUME:
    Prefer a sensible assumption over a question.
7.  Ask only when a wrong guess could destroy or overwrite data, or the request cannot be
    specified at all. Then the code block contains only: # QUESTION: <one short question>
8.  Use ALLOW only for risky actions the user explicitly asked for: file_delete,
    file_overwrite, shell, http. Otherwise design around them.
9.  Choose TASK for one procedure of up to about 60 lines of VBA. Choose BUILD when the
    request needs several procedures, a UserForm, or more than about 60 lines.
10. Delete rows only inside LOOP_ROWS dir=up, and only rows the user explicitly wants
    removed. Never delete rows because data is missing or pending: copy them to an
    exceptions sheet instead.
11. For more than about 1,000 rows: READ_ARRAY, process the array, WRITE_ARRAY, and add
    PERF screen_off calc_manual.
12. Naming: procedures PascalCase and verb-first (DeleteRowsByValue); modules mod +
    PascalCase (modCleanup); variables camelCase; fixtures snake_case_01.xlsx.
13. If the user's message is not a new request (for example a question about your last
    spec), answer in at most two plain sentences, then reissue the spec if it changed.
14. Find input columns by header name with COL(), never by fixed position, and check
    every header you use with a PRE line.
15. Assume data is typed by hand: trim IDs before matching, compare typed text with SAME,
    and guard every calculation and date comparison with IS_NUMBER or IS_DATE in a nested
    IF. Guard with the check itself (IF IS_DATE(x)), never with a flag variable set earlier.
16. Make re-runs safe: rebuild output sheets from scratch, and recalculate totals from
    source data instead of adding to or subtracting from stored values. The goal TEST
    uses runs=2 to prove it.
17. Prefer KIND Function returning a count (rows written, flagged or removed), so TESTs
    can check RESULT.
18. Always test with a named synthetic FIXTURE. Never use fixture=target or real workbooks.
19. Make the FIXTURE able to catch a wrong macro: give each sheet the macro rebuilds (or
    whose fills it owns) stale=<n> rows, put the values that sit exactly on each <, <=, >
    or >= comparison in the first rows with +boundary(...), repeat dictionary keys with
    +dupes(<n>) or values() counts of 2 or more, and use reldates(lo, hi) when the spec
    uses TODAY. Give each column a SKIP_IF IS_EMPTY tests at least one blank cell
    (+blank(<n>) or a "" boundary value), and give a spec that copies files a locked= file
    in a FIXTURE folder. Prefer exact assertions to ranges.
20. Declare every sheet the macro changes with MODIFIES: the whole sheet when it is
    rebuilt, cleared, sorted or has rows added or removed; ["<header>", ...] when only
    those columns' values change; fills when only fills change; files for a String
    argument whose folder the macro adds folders or files to. A spec that changes
    nothing has no MODIFIES lines.
21. Work with files only through MAKE_FOLDER and COPY_FILE, on paths built in place as
    PATH_JOIN($<root>, SAFE_NAME(...), ...) from a String argument. Check every root with
    FOLDER_EXISTS in PRE, and write each file op's status where the user can see it. No op
    moves, renames, overwrites or deletes a file.
```

### Reference

```text
TASK FORMAT  (keys in this order; (opt) = optional)
VCL 0.3
TASK
# ASSUME: <assumption>                        (opt, one line each)
PROC     <PascalCase name>
KIND     Sub | Function
MODULE   <modName>
ARGS     <name> As <Type>, ...                (opt)
RET      <Type>                               (Function only)
VAR      <name> As <Type>                     (opt, one line per local variable)
PRE      <condition>; <condition>             (opt, checked before any change)
MODIFIES $<arg> | $<arg>["<header>", ...] | $<arg> fills | $<arg> files   (opt, one line each)
OP       <step>                               (one or more)
PERF     screen_off calc_manual events_off    (opt, any of these)
ERR      standard
ALLOW    file_delete file_overwrite shell http  (opt)
FIXTURE  <name>.xlsx <fixture spec>           (opt)
TEST     fixture=<name>.xlsx call=<Proc>(<args>) [runs=<n>] assert=<assertion>   (two or more)

Layout: pad keys with spaces so values line up. Nesting = 2 extra spaces after the key,
relative to the enclosing OP line.
Types: Long, Double, String, Boolean, Date, Variant, Range, Worksheet, Workbook,
Dictionary, Collection. Loop variables from LOOP_ROWS, LOOP_COLS and FOR_EACH are declared
automatically; never repeat them in VAR.
For form events: MODULE is the form, PROC is <control>_<Event> (btnRun_Click).

BUILD FORMAT
VCL 0.3
BUILD
# ASSUME: <assumption>                        (opt)
NAME     <PascalCase tool name>
TARGET   xlsm | xlam
ENTRY    modMain.Run
MODULE   <modName>                            (one line per module; always modMain)
FORM     <frmName> CONTROLS <name>:<ControlType> ...   (opt)
SIG      <module>.<Proc>(<name> As <Type>, ...) [As <Type>]   (one per procedure)
FLOW     <caller> -> <callee> -> ...          (call order)
FIXTURE  ...                                  (opt)
TEST     fixture=<name>.xlsx call=modMain.RunHeadless(<args>) assert=...   (two or more)
---
<one complete TASK spec per SIG, each starting with VCL 0.3, separated by lines
containing only --->

BUILD rules: modMain always has RunHeadless (runs the whole tool without the form, so the
harness can test it) and Run (shows the form, or calls RunHeadless if there is no form).
Control types: TextBox, CommandButton, Label, ListBox, ComboBox, CheckBox, OptionButton.

OP SYNTAX
Structural ops begin an OP line:
  LOOP_ROWS dir=up|down from=<expr> to=<expr> var=<name>
  LOOP_COLS dir=left|right from=<expr> to=<expr> var=<name>
  FOR_EACH var=<name> in=SHEETS(<book>)|CELLS(<range>)|KEYS(<dict>)|ITEMS(<collection>)
  IF <condition> THEN <action>       (one-line form)
  IF <condition>                     (block form; steps below are nested)
  ELSE
  SELECT_CASE <expr>                 (block form)
  CASE <value>, <value>
  CASE_ELSE
  EXIT_IF <condition>                (leaves the procedure when the condition is true)
  SKIP_IF <condition>                (inside a loop: goes on with its next pass)
  EXIT_LOOP_IF <condition>           (inside a loop: leaves that loop)

Every other op is a call, NAME(arguments, key=value):
  Assign  LET(<var>, <expr>)
  Data    READ_ARRAY(<range>, into=<var>)
          WRITE_ARRAY(<var>, to=<range>)
          DICT_ADD(<dict>, <key>, <value>, mode=overwrite|keep|sum|count)
          DICT_GET(<dict>, <key>, into=<var>)
          SORT(<range>, key=<col>, order=asc|desc, header=yes|no)
          FILTER(<range>, col=<n>, crit=<expr>)
          DEDUPE(<range>, cols=<n>|<n,n>, header=yes|no)
  Cells   SET_VALUE(<target>, <expr>)
          WRITE_ROW(<ws>, <row>, <expr>, <expr>, ..., col=<n>)
                  (the values go to columns n, n+1, ...; without col= they start at column 1)
          SET_FORMULA(<target>, "<formula>", style=a1|r1c1)
          FORMAT(<target>, numfmt="<format>", bold=yes|no, fill=<name>|#RRGGBB, width=auto|<n>)
          CLEAR(<target>, what=all|contents|formats)
          COPY_ROW(<fromWs>, <fromRow>, <toWs>, <toRow>, cols=<n>)   (values of columns 1 to n)
          INSERT_ROW(<ws>, <row>)
          DELETE_ROW(<ws>, <row>)
  Sheets  ADD_SHEET(<book>, name=<expr>, after=<sheet>, into=<var>)
          COPY_SHEET(<sheet>, after=<sheet>, into=<var>)
          OPEN_BOOK(<path>, readonly=yes|no, into=<var>)
          SAVE_AS(<book>, <path>, format=xlsx|xlsm|csv)
          CLOSE_BOOK(<book>, save=yes|no)
  Files   MAKE_FOLDER(<path>, into=<status>)       (the folder, and any missing ones below its root)
          COPY_FILE(<from>, <to>, into=<status>)   (one file; never overwrites, never deletes)
          LIST_FILES(<folder>, "<pattern>", into=<var>)
          READ_TEXT(<path>, into=<var>)
          WRITE_TEXT(<path>, <expr>, mode=new|append)
  Calls   CALL(<module>.<Proc>, <args>)
  Output  LOG(<expr>)   MSG(<expr>)   RETURN(<expr>)
  Escape  NOTE "<terse technical text>"
FORMAT fill names: none (removes the fill), black, white, red, green, blue, yellow, magenta,
cyan. Write any other colour as #RRGGBB.

EXPRESSIONS
Value functions:
  LASTROW(<ws>, <col>)  LASTCOL(<ws>, <row>)  USED_RANGE(<ws>)
  COL(<ws>, "<header>")   column number of a row-1 header, or 0 if it is missing
  CELL(<ws>, <row>, <col>)  ROW_RANGE(<ws>, <row>, <fromCol>, <toCol>)  RANGE(<ws>, "<A1 ref>")
  SHEET(<book>, "<name>")  THIS_BOOK  ACTIVE_BOOK  ACTIVE_SHEET
  FIND(<range>, <expr>, match=whole|part)   returns a Range or NOTHING
  ROWS(<array>)  COLS(<array>)  ARR(<array>, <row>, <col>)
  COUNT(<dict or collection>)  EXISTS(<dict>, <key>)
  LEN  TRIM  UPPER  LOWER  LEFT(<expr>, <n>)  RIGHT(<expr>, <n>)
  IS_EMPTY  IS_NUMBER  IS_DATE  IS_ERROR  IS_NOTHING  TODAY  NOW  FILE_NAME(<path>)
  SAME(<expr>, <expr>)   TRUE when TRIM(a) and TRIM(b) are the same text, ignoring case
  CLEAN_TEXT(<expr>)   TRIM, and every inner run of whitespace becomes one space
  PARSE_DATE(<expr>, "<format>")   a real date, or empty when it isn't one in that format;
                       formats: "dd/mm/yyyy" "dd-mm-yyyy" "dd.mm.yyyy" "mm/dd/yyyy" "yyyy-mm-dd"
  IIF(<condition>, <expr>, <expr>)  MOD(<expr>, <expr>)
  SAFE_NAME(<expr>)   CLEAN_TEXT of it, or "" when that is not safe as one file or folder name
  PATH_JOIN($<root>, <part>, ...)   $<root>, then each part, joined by backslashes
  FOLDER_EXISTS(<path>)   TRUE when that folder exists
Arguments are written $name. Variables and loop variables are bare names.
Text uses "double quotes"; write a quote inside text as two quotes: "Say ""hi""".
Also numbers, TRUE, FALSE, NOTHING.
Operators: = <> < > <= >= AND OR NOT + - * / &
Precedence, tightest first: - (negation), * /, + -, &, comparisons, NOT, AND, OR.
Comparisons don't chain: write a < b AND b < c, never a < b < c.
Option names are reserved words: dir from to var in mode into key order header what fill
numfmt bold width style cols col crit readonly format save after name match. Write an
option as name=value with no space before "="; never use these words as names.
Columns are numbers (1 = A) everywhere except inside "A1 ref" strings.
An array element can be a SET_VALUE target: SET_VALUE(ARR(data, r, 3), 0).

SEMANTICS  (the same on every PC, whatever its regional settings)
  IS_EMPTY    true for empty cells and for text made only of spaces, non-breaking spaces,
              tabs or line breaks; an error cell (#N/A) is not empty
  TRIM        removes spaces, non-breaking spaces, tabs and line breaks at both ends
  IS_NUMBER   true for numbers and for text that is a plain number after TRIM ("25",
              " -3.5 "); false for blanks, TRUE/FALSE, dates, errors, "1,250", "1e3", "25%"
  IS_DATE     true only for real Excel date values; text that looks like a date is false
  Text        a value becomes text the same way everywhere (TRIM, UPPER, LOWER, LEN, LEFT,
              RIGHT, &, keys): whole numbers without ".0", "." as the decimal point, dates
              as yyyy-mm-dd, TRUE/FALSE, errors as their Excel text (#N/A)
  = and <>    with text on either side they compare that text exactly (case and spaces
              count), so a cell holding 25 equals "25"; compare numbers or dates only after
              IS_NUMBER or IS_DATE
  SAME        compares TRIM(a) and TRIM(b) as text, ignoring case: SAME(" ready ", "READY")
              and SAME(25, "25") are TRUE, SAME(25, "25.0") is FALSE; an error is SAME as its
              text, so check IS_ERROR first when that matters
  CLEAN_TEXT  also removes characters below code 32; case is kept
  IS_ERROR    true for Excel error values (#N/A, #DIV/0!, ...); the text "#N/A" is not one
  PARSE_DATE  a real date passes through; text is TRIMmed and must match the format exactly
              (day and month 1 or 2 digits, year 4 digits from 1900, a date that exists);
              anything else gives empty, so guard the result with IS_DATE. VCL never guesses
              whether 03/04/2026 is March or April
  IIF         both values are always evaluated, so the condition never guards them; values
              only, never a Range, Worksheet, Workbook or Dictionary
  MOD         a - b * INT(a / b), as Excel: MOD(-7, 3) = 2, MOD(7.5, 2) = 1.5; MOD(a, 0) is a
              run-time error; numbers only, guarded like any calculation
  SKIP_IF     past a SKIP_IF or EXIT_LOOP_IF line its condition is FALSE, so
              SKIP_IF NOT IS_NUMBER(x) guards x for the rest of that pass
  Writes      SET_VALUE, WRITE_ROW and COPY_ROW store text as exactly that text, even text
              Excel would read as a number, date, TRUE/FALSE, error or formula ("1001", "1-2")
  LEFT RIGHT  a count below 0 gives ""; a count beyond the length gives the whole text
  COL         matches the row-1 header ignoring case, extra spaces and line breaks; the first
              match wins; 0 if not found
  LASTROW     the last row holding any value (spaces too), like End(xlUp); 1 if the column
              is empty. LASTCOL works the same along a row
  Dictionary  keys are compared as text ignoring case, so 1001 and "1001" are one key;
              Dictionary and Collection VARs start empty
  PRE         if a condition is false: one ForgeLog entry, nothing changes, and a Function
              returns -1
  EXIT_IF     leaves the procedure; a Function then returns 0, "" or FALSE (use IF ... THEN
              RETURN(<value>) to return something else)
  AND / OR    both sides are always evaluated, so a guard on the left does not protect the
              right: guard risky comparisons with a nested IF on the check itself
  SAFE_NAME   "" when the CLEAN_TEXT result is empty, over 100 characters, holds any of
              \ / : * ? " < > |, is "." or "..", ends with a dot, or is a device name (CON,
              PRN, AUX, NUL, COM1-COM9, LPT1-LPT9, any case, any extension: con.txt)
  PATH_JOIN   trailing backslashes on $<root> are dropped; "" when $<root> is empty or a
              part is empty or not exactly its own SAFE_NAME, and a file op given "" reports
              Bad name
  Paths       a file op's path is PATH_JOIN($<root>, ...) of a String argument, never a bare
              root, and every part after the root is SAFE_NAME(...) or a quoted literal that
              is its own safe name, written in place (never a variable or a cell), so no path
              can climb out of its root. A SAFE_NAME result or a path is text elsewhere
  FOLDER_EXISTS  TRUE only for a folder, FALSE for a file or ""; in Excel the path must be
              written in full (C:\Jobs, not Jobs or C:\Jobs\)
  MAKE_FOLDER writes one status into its String VAR, checked in this order: Bad name ("" or
              with * or ?), Too long (over 247 characters), Exists, Blocked by file (a file
              has the folder's name, or the name of a folder on the way; then nothing is
              made), No folder (the root is missing), Created (with every missing folder
              between it and the root, never the root itself)
  COPY_FILE   writes, in this order: Bad name, Too long (over 258 characters), Missing (no
              file at <from>), No folder (make the folder first), Already there (same size:
              a re-run is safe), Conflict (another size, or a folder: never overwritten),
              Locked (open in another program), Copied (<from> stays). In Excel also Check
              copy (the copy came out a different size; it is kept) and Failed (logged to
              ForgeLog). A file op never stops the run

MODIFIES  (one line per sheet the macro changes, or per kind of change)
  MODIFIES $<arg>                     anything on that sheet: values, fills, formats, rows
  MODIFIES $<arg>["<header>", ...]    values in those columns, row 2 down
  MODIFIES $<arg> fills               fills anywhere on that sheet
  MODIFIES $<arg> files               folders and files may be added below that String
                                      argument; nothing else on disk may change
  An argument has one whole-sheet line, or at most one [headers] line and one fills line.
  Every write op must go through a $sheet argument whose lines allow that change: values
  need [headers] or the whole sheet, fill-only FORMATs need fills, and anything else
  (other FORMATs, CLEAR what=all or formats, SORT, INSERT_ROW, DELETE_ROW, DEDUPE, FILTER)
  needs the whole sheet. Write through CELL, ROW_RANGE, USED_RANGE or RANGE of a $sheet.
  MAKE_FOLDER's folder and COPY_FILE's <to> must be below a root with a files line;
  COPY_FILE's <from> may be below any String argument.
  Every TEST then checks that nothing else changed: each other cell keeps its value, fill
  and number format, and the header row never changes; on disk nothing is removed or
  changed, and new folders and files appear only below a files root. A spec that changes
  nothing has no MODIFIES lines and gets the same check.

TEST ASSERTIONS  (checked on the test workbook after the run)
  COUNTIF(<Sheet>!<range>, <criteria>)  COUNTA(<Sheet>!<range>)  SUM(<Sheet>!<range>)
  COUNTIFS(<Sheet>!<range>, <criteria>, <Sheet>!<range>, <criteria>, ...)
                 rows where every range meets its criteria
  SUMIFS(<Sheet>!<sumRange>, <Sheet>!<range>, <criteria>, ...)
                 sum of sumRange over the rows where every range meets its criteria
  VALUE(<Sheet>!<cell>)  ROWS_USED(<Sheet>)  SHEET_EXISTS("<name>")
  FILE_EXISTS("<Folder>\<path>")   that file exists below a FIXTURE folder of the TEST
  FILE_COUNT("<Folder>[\<path>]")  files in that folder and in every folder below it
  FILL_COUNT(<Sheet>!<range>, #RRGGBB)   cells in the range with that fill colour
  RESULT         the Function's return value (from the last run)
  ERROR_COUNT    errors written to ForgeLog during the run
  PRE_<measure>  the same measurement taken before the first run, e.g. PRE_COUNTA(Data!A:A)
                 or PRE_COUNTIFS(Data![Status], "Open", Data![Owner], "QA")
COUNTIF, COUNTIFS and SUMIFS work as in Excel 2016: criteria such as "Ready", "<>" (not
blank), "<>B-109" or ">=5"; text matching ignores case. COUNTIFS and SUMIFS ranges must be
the same size, so use [Header] ranges from the same sheet.
Ranges by header: <Sheet>![Header] is the whole column under that row-1 header;
<Sheet>![Header]<n> is that column's cell in row n. Prefer them to column letters.
runs=<n> calls the procedure n times on the same test workbook before checking
(default 1).
Combine assertions with AND. In TEST call arguments, a bare sheet name (Orders) means that
worksheet in the test workbook, and a bare folder name (Dump) means that FIXTURE folder.
fixture=<name>.xlsx names a FIXTURE file from this spec; the harness builds it fresh for
every TEST. fixture=target and real workbooks are not allowed.

FIXTURE SPEC
  FIXTURE <name>.xlsx sheet=<name> cols=<A:Header,B:Header,...> rows=<n> [stale=<n>] data=<col>:<rule>;<col>:<rule>
  Rules: values("a"x<n>,"b"x<n>)  seq(<start>,<step>)  int(<lo>,<hi>)  num(<lo>,<hi>)
         dates(<yyyy-mm-dd>,<yyyy-mm-dd>)  reldates(<lo>,<hi>)  text(<length>)
  reldates(lo, hi) gives dates from TODAY + lo to TODAY + hi days, worked out when the
  file is built, so TESTs that use TODAY give the same results on every day.
  Modifiers, after the rule, in this order, each at most once:
    +boundary(v1,v2,...)  the first data rows hold exactly these values ("" is a blank
                          cell; numbers of days for reldates). For values(...) the values
                          move up, so the counts stay the same.
    +blank(<n>)           leave n random cells below the boundary rows blank.
    +dupes(<n>)           give n random rows the value of another row (repeated keys).
  stale=<n> marks n rows as leftovers of an earlier run, with a red fill: on a rows=0
  sheet, n rows of the text "stale"; on a sheet with data, n random data rows, whose
  values stay.
  values(...) counts add up to rows. rows=0 creates the sheet with its header row only.
  Repeat FIXTURE lines with the same file name to add more sheets.
  FIXTURE <name>.xlsx folder=<Name> [files=<path>[:<bytes>],...] [locked=<path>,...]
  adds a synthetic folder to the same test file: each file below it (\ between folders,
  each part a safe name of letters, digits, _ . and -), with that size or one from its own
  seed. locked= files act as if another program has them open. A folder exists when a
  file is below it. <Name> is one word and not a sheet's name. Nothing real is read or
  written: the folders sit in a sandbox (C:\KRONOS_SANDBOX in the interpreter, a new
  folder below %TEMP% in Excel, whose root may be up to 120 characters), so a TEST fails
  when a file op path could be Too long in Excel but not in the interpreter: keep paths
  far over or well under the limits.
  Sheet names have no spaces, so TEST calls can use them bare. They may be short forms of
  the real sheet names (WelderPerf for "Welder Performance"); ASSUME lines describe the
  real workbook.
  Each column has its own fixed seed, made from file name + sheet + column. The builder
  uses it for everything random in that column: the shuffled order of values(...), the
  int, num, dates, reldates and text values, and which cells +blank(<n>) leaves blank;
  +dupes(<n>) and stale=<n> have seeds of their own. Columns are not correlated,
  and the same spec always builds the same file.
```
