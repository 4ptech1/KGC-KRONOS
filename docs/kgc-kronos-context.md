# KGC KRONOS: AI context file

**Snapshot: 21 Oct 2026, after the review passes of item 2 (D47, D48, D51), VEP 0.5 Phases 1 to 4 (D50, D52, D53, D55 to D57) and this file's generated appendices (D54).** Hand this file to any Claude chat or claude.ai
Project that needs to understand the project. The repo is always the source of truth: if this
file and the repo disagree, the repo wins, and the repo's `CLAUDE.md` has the current status.

- Repo: https://github.com/4ptech1/KGC-KRONOS (renamed from `vba-forge` on 17 Oct 2026; old
  links redirect). Inside the code and docs the project is still called **VBA Forge**.
- Owner: **Boss** (GitHub `4ptech1`). Always call him Boss.

**How this file is built.** Sections 1 to 12 are written by hand and kept current at the end of
every session. The appendices are copied from the repo by `python compiler/gen_context.py`, so
they are never stale (a unit test fails when they differ from the repo):
- A: the full language reference (as in `vcl/spec.md` and the Orchestrator prompt)
- B: the rest of the Orchestrator prompt (role, examples, final check)
- C: every spec in `specs/`, each with the VBA the compiler writes for it
- D: `DECISIONS.md` in full
- E: the meaning of every value function, with its conformance cases
- F: `vba/ForgeLib.bas`, the VBA runtime
- G: `CLAUDE.md`, the plans, and every language design (VEP)

The whole file stays under 3 MB (Boss's cap); today it is about 400 KB, so a chat can read all
of it. When a question needs detail, look it up in the appendices instead of guessing.

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
DECISIONS.md              every language decision (G, S, T, E, F ids) and Claude's own (D1-D46)
README.md                 overview, how to run the checks, progress log, dataset count
requirements.txt          one pinned dependency: lark==1.3.1 (with hash)
vcl/spec.md               THE language reference (VCL 0.4)
vcl/orchestrator_prompt.txt  the Orchestrator's instructions; shares the reference text with spec.md
vcl/grammar.lark          formal LALR grammar
vcl/semantics.py          reference meaning of the value functions (S1-S25)
vcl/examples/             example1_rev3.vcl and example3_files.vcl (= the prompt's Examples 1 and 3), build_example.vcl (a sketch, invalid)
specs/task_01..12.vcl     the 12 task specs (11 in VCL 0.2; task_12, the router, in 0.3)
compiler/vclc.py          the toolchain CLI: check, test, build, all, mutate
compiler/parser.py        hand-written parser   compiler/validator.py  rules, plus grammar and type stages
compiler/vcl_parse.py     grammar parser        compiler/vcl_types.py  type checker (guards, narrowing)
compiler/vcl_interp.py    reference interpreter, TEST assertions, the UNCHANGED check
compiler/vcl_codegen.py   VCL to VBA compiler   compiler/vcl_fixtures.py  the one fixture generator
compiler/vba_lint.py      static VBA checks     compiler/vcl_mutate.py  mutation scorer
compiler/vcl_migrate.py   VCL 0.1 to 0.2 rewriter, with --prove
compiler/gen_semantics_doc.py  writes docs/vcl-semantics.md
compiler/gen_context.py   writes this file's appendices from the repo (D54)
compiler/vcl_formula.py   SET_FORMULA's Excel 2016 allowlist and checks (D43)
vba/ForgeLib.bas          the VBA runtime every generated procedure calls
dist/forge_module.bas     committed build output: ForgeLib plus one procedure per spec (paste-ready)
dist/forge_selftest.bas   committed: the ForgeLib self-test, every conformance case run in Excel (D34)
dist/forge_pilot.bas      committed: the Excel pilot of task_12; results checked by `vclc pilot-check` (D35)
docs/excel-pilot.md       Boss's steps for the optional Excel pilot
dist/forge_fixtures.bas, dist/procedures/, dist/expected/   rebuilt by vclc, git-ignored
tests/                    403 unit tests; conformance_cases.py is shared with the docs
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

## 6. VCL 0.4 cheat sheet (summary only; the full text is in Appendix A)

**TASK keys, in this order:** `VCL 0.4`, `TASK`, `# ASSUME:` lines, PROC, KIND (Sub or
Function), MODULE, ARGS, RET, VAR, PRE, MODIFIES, OP, PERF, ERR standard, ALLOW, FIXTURE, TEST.
- Arguments are written `$name`.
- OP lines nest by 2 extra spaces.
- Option names (dir, from, to, var, in, mode, into, key, order, header, what, fill, numfmt,
  bold, width, style, cols, col, crit, readonly, format, save, after, name, match) are reserved
  words.

**Structural ops:**
- loops: `LOOP_ROWS dir=up|down`, `LOOP_COLS`, `FOR_EACH` over SHEETS, CELLS, KEYS or ITEMS
- `IF ... THEN <op>` (one line), or a block `IF` with optional `ELSE`
- `SELECT_CASE`
- `EXIT_IF` (leaves the procedure)
- `SKIP_IF` and `EXIT_LOOP_IF` (inside loops only)

**Call ops:**
- LET, READ_ARRAY, WRITE_ARRAY, DICT_ADD (mode=overwrite|keep|sum|count), DICT_GET
- SORT, FILTER (hides rows, never deletes), DEDUPE (`cols=[1, 2]`)
- SET_VALUE, WRITE_ROW, SET_FORMULA, FORMAT, CLEAR, COPY_ROW, INSERT_ROW, DELETE_ROW
- sheets (VCL 0.4): ADD_SHEET and COPY_SHEET, each with `name=`
- files: MAKE_FOLDER and COPY_FILE (VCL 0.3), READ_TEXT and WRITE_TEXT (VCL 0.4), each writing
  a status into a String VAR; LIST_FILES (VCL 0.4) into a Collection
- CALL, LOG, MSG, RETURN
- NOTE (the escape hatch)
- no longer taught: OPEN_BOOK, SAVE_AS and CLOSE_BOOK (refused until March 2027), ACTIVE_BOOK
  and ACTIVE_SHEET (D46)

**Value functions:**
- worksheet: COL, CELL, LASTROW, LASTCOL, ROW_RANGE, USED_RANGE, RANGE, SHEET, THIS_BOOK, FIND
- arrays, dictionaries and lists: ARR, ROWS, COLS, COUNT, EXISTS, KEYS, ITEMS
- text: TRIM, UPPER, LOWER, LEN, LEFT, RIGHT
- checks: IS_EMPTY, IS_NUMBER, IS_DATE, IS_ERROR, IS_NOTHING
- constants and files: TODAY, NOW, FILE_NAME
- new in 0.2: SAME, CLEAN_TEXT, PARSE_DATE (with a required format), IIF, MOD
- new in 0.3: SAFE_NAME, PATH_JOIN, FOLDER_EXISTS
- new in 0.4: SHEET_EXISTS

**Meanings that surprise people** (S1 to S25, with tables in `docs/vcl-semantics.md`):
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
- **Writes:** SET_VALUE, WRITE_ROW, COPY_ROW and WRITE_ARRAY keep text as text in Excel (S17).
- **Rows (VCL 0.4):** DEDUPE keeps the first row of each key and never deletes a row with a blank
  or error key cell (S21). FILTER hides each row whose cell isn't SAME as the value, without
  AutoFilter. FIND gives the first SAME cell, row by row, or NOTHING (S22).
- **Sheets and formulas (VCL 0.4):** a sheet is only ever added: a bad name or one already used
  (S23) stops the run before anything changes. SET_FORMULA takes Excel 2016 functions only, and
  a TEST checks a formula's text, never its value.
- **Files (VCL 0.3 and 0.4):** every path is `PATH_JOIN($root, SAFE_NAME(...), ...)`, written in
  place, so no path can climb out of its root. MAKE_FOLDER, COPY_FILE, READ_TEXT and WRITE_TEXT
  never stop the run: they write a status (`Created`, `Exists`, `Copied`, `Read`, `Written`,
  `Already there`, `Conflict`, `Locked`, `Missing`, `No folder`, `Bad name`, `Too long`, ...).
  Nothing moves, overwrites or deletes a file; WRITE_TEXT only adds a new one. Text files hold
  printable ASCII, tabs and line breaks, up to 32,767 characters (S25). LIST_FILES matches `*`
  and `?` against the whole name, ignoring case, in lowercase name order (S24).

**MODIFIES** (a key after PRE):
- the forms:
  - `MODIFIES $x`: the whole sheet (needed by SORT, FILTER, DEDUPE, INSERT_ROW and DELETE_ROW)
  - `MODIFIES $x["Header", ...]`: values in those columns, row 2 down
  - `MODIFIES $x fills`: fills only
  - `MODIFIES $root files` (VCL 0.3): folders and files may be added below that String argument
  - `MODIFIES $book sheets` (VCL 0.4): sheets may be added to that Workbook argument
- **Before the run**, the validator checks every write against these lines.
- **After every TEST**, the interpreter checks that nothing else changed (UNCHANGED): values,
  fills, number formats, hidden rows, the sheet list and the files.
- **A spec that writes nothing** needs no MODIFIES lines, and every TEST checks it changed
  nothing.

**The 22 hard rules (short form):**
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
10. Delete rows only in `dir=up` loops or with DEDUPE, and never because data is missing.
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
    - a `locked=` file when the spec copies or reads files
20. Declare every sheet the macro changes with MODIFIES (`files` for a root it adds to, `sheets`
    for a workbook it adds sheets to).
21. Files only through MAKE_FOLDER, COPY_FILE, LIST_FILES, READ_TEXT and WRITE_TEXT on
    `PATH_JOIN($root, SAFE_NAME(...))` paths; every root checked with FOLDER_EXISTS in PRE;
    nothing moves, renames, overwrites or deletes a file.
22. Sheets added only with ADD_SHEET or COPY_SHEET (with `name=`, skipped when SHEET_EXISTS);
    formulas only with SET_FORMULA and Excel 2016 functions, tested with FORMULA; no op removes
    or replaces a sheet.

**FIXTURE syntax:**
`FIXTURE <file>.xlsx sheet=<Name> cols=A:Header,... rows=<n> [stale=<n>] data=<col>:<rule>;...`
- Rules: values("a"x3,...), seq, int, num, dates, reldates, text.
- Modifiers, in this order: `+boundary(...)`, `+blank(n)`, `+dupes(n)`.
- The data is deterministic: each column has its own seed, made from file, sheet and column.
- Folders (VCL 0.3): `FIXTURE <file>.xlsx folder=<Name> files=a.pdf:120,J-1\b.pdf locked=a.pdf`;
  TESTs pass the folder bare, like a sheet. They live in a sandbox, never on real folders.
  VCL 0.4 adds text files: `files=notes\a.txt:"one line of text"`.

**TEST assertions:**
- measures: COUNTIF, COUNTIFS, COUNTA, SUM, SUMIFS, VALUE, ROWS_USED, FILL_COUNT
- files (VCL 0.3): `FILE_EXISTS("Folder\path")`, `FILE_COUNT("Folder")`; (VCL 0.4)
  `FILE_TEXT("Folder\path")`
- rows and formulas (VCL 0.4): `VISIBLE_COUNT(Sheet!range)`, `FORMULA(Sheet!cell)`
- run results: RESULT, ERROR_COUNT
- PRE_<measure>, taken before the run
- ranges by header: `Sheet![Header]` (the whole column from row 1, header cell included) and
  `Sheet![Header]<row>` (one cell)
- assertions have no NOT, TRUE or FALSE: "not there" is `SHEET_EXISTS("X") = 0`; an assertion
  names a sheet as one word, so it can't read a sheet whose name has spaces
- a Workbook argument is passed as the TEST's own fixture file: `call=P(ncr_01.xlsx, Ncr)`

## 7. What the toolchain can run today (important when reviewing specs)

The validator accepts the whole language, but the interpreter and compiler implement only part
of it. A spec outside this subset passes the validator but fails at test or build.

| Fully supported (validate, test, compile) | Compile only (no interpreter yet) | Not built yet |
| --- | --- | --- |
| LET, SET_VALUE, WRITE_ROW, CLEAR, FORMAT, COPY_ROW, DELETE_ROW, DICT_ADD, DICT_GET, SORT, RETURN, MAKE_FOLDER, COPY_FILE (VCL 0.3); LOOP_ROWS, LOOP_COLS, FOR_EACH over KEYS(...), IF/ELSE, EXIT_IF, SKIP_IF, EXIT_LOOP_IF; CELL, COL, LASTROW, LASTCOL, ROW_RANGE, USED_RANGE, TRIM, UPPER, LOWER, LEN, LEFT, RIGHT, IS_EMPTY, IS_NUMBER, IS_DATE, IS_ERROR, SAME, CLEAN_TEXT, PARSE_DATE, IIF, MOD, COUNT, EXISTS, KEYS, SAFE_NAME, PATH_JOIN, FOLDER_EXISTS; MODIFIES (with `files`); folder FIXTUREs; assertions FILE_EXISTS, FILE_COUNT; and since item 2 Phase 1 (D37): INSERT_ROW, READ_ARRAY, WRITE_ARRAY, LOG, MSG, SELECT_CASE, ARR, ROWS, COLS, SHEET, RANGE (a cell or a block), FILE_NAME, THIS_BOOK, ACTIVE_BOOK; FOR_EACH over SHEETS or CELLS; and since Phase 2 (D38): DEDUPE, with `cols=[1, 2]` in a `VCL 0.4` spec; and since Phase 3 (D39): FIND, IS_NOTHING, NOTHING (write to a found cell through FIND written in place); and since Phase 4a (D40): ADD_SHEET and SHEET_EXISTS on a `$book` Workbook argument with `MODIFIES $book sheets` (a TEST passes the workbook as its fixture file, `call=P(ncr_01.xlsx, Ncr)`); and since Phase 4b (D41): COPY_SHEET with `name=`; and since Phase 5 (D42): FILTER (hides rows by SAME, whole-sheet MODIFIES, not with SORT in one spec) and the VISIBLE_COUNT assertion; and since Phase 6 (D43): SET_FORMULA (a quoted formula on the Excel 2016 allowlist; not with ops that move rows) and the FORMULA assertion; and since Phase 7a (D44): LIST_FILES into a Collection, FOR_EACH over ITEMS, COUNT of a list; and since Phase 7b (D45): READ_TEXT and WRITE_TEXT with a status into `into=`, text files in folder FIXTUREs (`files=a.txt:"text"`) and the FILE_TEXT assertion | CALL; ACTIVE_SHEET (no longer taught: what is active differs from run to run, D46) | OPEN_BOOK, SAVE_AS, CLOSE_BOOK (refused until March 2027 with a clear message, D36, D46), NOTE; BUILD specs |

**Fill colours:** `fill=` takes `none`, a name (black, white, red, green, blue, yellow, magenta,
cyan) or `#RRGGBB`. Each name is the colour of the VBA constant of that name (red = `#FF0000`), so
`fill=red` and `fill=#FF0000` give the same result (D23, fixed in PR #18).

## 8. Status (20 Oct 2026)

**Roadmap Phase 0, Foundation (14 to 31 Oct 2026).** Gate 0: all specs validate, type-check,
pass their TESTs in the interpreter, and compile to lint-clean VBA, in CI. **This is met.**

| Measure | Now |
| --- | --- |
| Specs | 12 (task_01 to task_12): 11 in VCL 0.2, and task_12 (the router) in VCL 0.3 |
| Unit tests | 403 pass |
| Validator | 0 errors; 1 known warning (task_04 has 48 OP lines, over the validator's 40, and 107 lines of VBA, over rule 9's ~60; accepted until March) |
| Mutation score (snapshot) | 95.4% (418 of 438); task_12 catches 29 of 29; 19 of the 20 survivors are equivalent mutants (D14, D24) that no test can ever catch; 1 is a small fixture gap in task_11 (D24) |
| VBA lint | 0 problems |
| Verified pairs | 0 (none has been run in Excel yet) |
| Orchestrator Project | the final VCL 0.4 prompt is ready (D46, with the D47, D48 and D51 review fixes); once Boss pastes it in, replies start with `VCL 0.4`, use every op item 2 added (sheets, formulas, FILTER, DEDUPE, FIND, text files), and are one TASK (no BUILD or CALL until they can be tested) |

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
- #42: item 2 Phase 7a, LIST_FILES and ITEMS (S24, D44).
- #43: item 2 Phase 7b, READ_TEXT and WRITE_TEXT (S25, D45).
- #44: spec.md and the prompt teach VCL 0.4, with hard rule 22 (E12, D46). Item 2 is complete.
- #45: review pass after item 2: wording fixed where it differed from the tools, clearer TEST failures (D47).
- #46: ForgeLib fixes from a VBA review: COPY_SHEET safety, sheet names, DEDUPE keys, formulas on Text cells, ASCII text-file paths (D48).
- #47: the design of CALL and BUILD specs (`docs/vep-0.5-build.md`, D49).
- #49: VEP 0.5 Phase 1, the `VCL 0.5` line, CALL `into=` and the BUILD rules (E13, E14, D50).
- #48 and #51: a second review pass, from a parallel session, folded in by #51 (D51): ForgeSort
  gives every sort option, LIST_FILES leaves out hidden and system files, more lint bans, and
  more rules say what the toolchain checks.
- #52: VEP 0.5 Phase 2, the interpreter runs BUILD TESTs (D52).
- #53: VEP 0.5 Phase 3, one module per BUILD from `vclc build`, and mutation over a BUILD (D53).
- #54: this file grows to about 400 KB, with appendices copied from the repo by `gen_context.py` (D54).
- #55: VEP 0.5 Phase 4a, ROW_OF and DEDUPE `into=` (D55).
- #56: VEP 0.5 Phase 4b, a header the spec writes counts as checked for rule 14 (D56).
- #57: VEP 0.5 Phase 4c, sheet names with spaces in assertions, in quotes (D57).

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

**After the MVP** (the numbers are CLAUDE.md's and Boss's; `docs/vcl-mvp-plan.md` calls the
router item 4, this item 2 its item 5, and item 3 its items 6 and 7):
2. The rest of revision 3 in the interpreter and compiler (Boss said "go item 2" on 20 Oct
   2026), designed in `docs/vep-0.4-rev3-rest.md` (D36): first the interpreter catches up with
   what already compiles, then list syntax with DEDUPE, FIND, the sheet ops, FILTER (it hides
   rows by SAME, never by AutoFilter criteria), SET_FORMULA (tested by its text) and text files
   (new files only). OPEN_BOOK, SAVE_AS and CLOSE_BOOK wait until March 2027. Phases 1 (D37)
   and 2 (D38: the `VCL 0.4` line in the tools, `cols=[1, 2]`, DEDUPE), 3 (D39: FIND) and 4a (D40:
   ADD_SHEET, SHEET_EXISTS), 4b (D41: COPY_SHEET), 5 (D42: FILTER), 6 (D43: SET_FORMULA) and 7a
   (D44: LIST_FILES) and 7b (D45: READ_TEXT, WRITE_TEXT) are done, and so is the switch (E12,
   D46): spec.md and the Orchestrator prompt teach VCL 0.4, and Boss updates the Orchestrator
   once. **Item 2 is complete.**
2. CALL and BUILD specs (`docs/vep-0.5-build.md`, D49; Boss said "continue building the next
   part according to your recommendations" on 21 Oct 2026): BUILD TESTs run in the interpreter,
   CALL keeps a Function's result with `into=`, MODIFIES follows the arguments, and the compiler
   writes one module per BUILD. FORM and xlam wait. Then one switch to VCL 0.5. Phase 1 is done
   (E13, E14, D50): the validator checks 0.5 BUILDs. Phase 2 is done (D52): the interpreter runs
   their TESTs. Phase 3 is done (D53): one module per BUILD, and mutation. Phase 4a is done
   (D55): ROW_OF gives the row FIND found, and DEDUPE `into=` counts the rows it removed (both
   taught at the switch to 0.5). Phase 4b is done (D56): a header the spec writes to row 1
   counts as checked for rule 14. Phase 4c is done (D57): an assertion names a sheet with
   spaces in quotes. Phase 5 (a proven BUILD and the switch to 0.5) is next.
3. A corpus of 150 tasks, Orchestrator-written specs, and an expressibility report.
4. A dataset generator: tasks in plain English with their proven specs, plus repair pairs.
   Mutants double as "read carefully" contrast pairs.
5. **March 2027:** the new PC; train our own model from zero on in-house data (D25). It writes
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
1. Paste the VCL 0.4 prompt (`vcl/orchestrator_prompt.txt`) into the Orchestrator Project's
   instructions, once.
2. When, if ever, to run the optional Excel pilot (`docs/excel-pilot.md`).
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
- **D36, D46:** item 2 made the rest of revision 3 provable by TESTs, and VCL 0.4 is taught in
  one step. OPEN_BOOK, SAVE_AS and CLOSE_BOOK stay in the language but are refused until
  March 2027, because they open and save real workbooks.
- **D42, D43:** FILTER hides rows itself (never AutoFilter, whose criteria read through the PC's
  settings). SET_FORMULA writes Excel 2016 functions only, and a TEST checks a formula's text,
  never its value (only Excel can calculate it).
- **S24, S25:** LIST_FILES gives names in one fixed order. READ_TEXT and WRITE_TEXT handle
  printable ASCII only, and WRITE_TEXT never overwrites a file.
- **D47:** `Sheet![Header]` is the header's whole column from row 1, the header cell included
  (as `A:A` is): 5 values under a header give COUNTA = 6. Committed specs count on it.
- **D49:** CALL and BUILD specs: one module per BUILD, CALL `into=`, MODIFIES that follows the
  arguments; FORM and xlam wait (no UserForm can be pasted, and THIS_BOOK would be the add-in).
- **D48:** a review of ForgeLib fixed five places where Excel would differ from the interpreter
  (one a safety fix in COPY_SHEET). Non-ASCII dictionary keys may still group differently in
  Excel; ASCII keys always agree.
- **D51:** a second review pass, folded in from a parallel session: rule 5 asks for a PRE_
  measure; about 40 OP lines per TASK, and no BUILD or CALL until they can be tested; a formula
  names only sheets every TEST workbook has. ForgeSort gives every sort option; LIST_FILES leaves
  out hidden and system files; the linter bans FileCopy, SetAttr, Workbooks.Open and more. Known
  gaps: Excel's locale-aware text sort, hidden rows in a sort, vbTextCompare keys on a Turkish PC.

## 11. How to help (rules for the AI reading this)

- **Check the language, not your memory.** Before writing or reviewing VCL, read Appendix A (the
  same text as `vcl/spec.md`), or the cheat sheet above; the specs in Appendix C show it in use.
  Never invent keys, ops or functions; if something can't be said,
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
| VCL | VBA Command Language: the spec language (`vcl/spec.md`). Current version 0.4 (20 Oct 2026). |
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

<!-- Everything below is generated by python compiler/gen_context.py from the repo: do not edit by hand. -->

## Appendix A: the full VCL 0.4 reference

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
    else changed by comparing a measure with its PRE_ value, for example
    COUNTA(Orders![Status]) = PRE_COUNTA(Orders![Status]) (for read-only work: proves the
    source data is unchanged). At least one TEST includes ERROR_COUNT = 0. When the test
    workbook has another sheet, add a wrong-sheet TEST that passes it and expects
    RESULT = -1 AND ERROR_COUNT = 1.
6.  Put every assumption at the top of the spec as a line starting with # ASSUME:
    Prefer a sensible assumption over a question.
7.  Ask only when a wrong guess could destroy or overwrite data, when the request cannot be
    specified at all, or when one TASK can't do all of it (rule 9). Then the code block
    contains only: # QUESTION: <one short question>
8.  Use ALLOW only for risky actions the user explicitly asked for: file_delete,
    file_overwrite, shell, http. No VCL 0.4 op deletes or overwrites a file, runs a program
    or uses the web, so design around them.
9.  Write one TASK: one procedure of up to about 40 OP lines (about 60 lines of VBA). BUILD
    and CALL can be checked but not yet tested, so don't use them for now. When one TASK
    can't do the whole request (it needs several procedures, a UserForm, more OP lines, or
    ops one spec may not combine), ask which part to do first (rule 7).
10. Delete rows only inside LOOP_ROWS dir=up or with DEDUPE, and only rows the user
    explicitly wants removed. Never delete rows because data is missing or pending: copy
    them to an exceptions sheet instead (DEDUPE never deletes a row with a blank key).
11. For more than about 1,000 rows: READ_ARRAY, process the array, WRITE_ARRAY, and add
    PERF screen_off calc_manual.
12. Naming: procedures PascalCase and verb-first (DeleteRowsByValue); modules mod +
    PascalCase (modCleanup); variables camelCase; fixtures snake_case_01.xlsx.
13. If the user's message is not a new request (for example a question about your last
    spec), answer in at most two plain sentences, then reissue the spec if it changed.
14. Find input columns by header name with COL(), never by fixed position, and check
    every header you use with a PRE line.
15. Assume data is typed by hand: trim IDs before matching, compare typed text with SAME,
    and guard every calculation and date comparison on a value from the workbook (a cell,
    an array element, a dictionary value) with IS_NUMBER or IS_DATE in a nested IF; your
    own counters need no guard. Guard with the check itself (IF IS_DATE(x)), never with a
    flag variable set earlier.
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
    uses TODAY. Give each column that a SKIP_IF IS_EMPTY tests at least one blank cell
    (+blank(<n>) or a "" boundary value), and give a spec that copies or reads files a
    locked= file in a FIXTURE folder. Prefer exact assertions to ranges.
20. Declare every sheet the macro changes with MODIFIES: the whole sheet when it is
    rebuilt, cleared, sorted, filtered or has rows added or removed; ["<header>", ...]
    when only those columns' values change; fills when only fills change; files for a
    String argument whose folder the macro adds folders or files to; sheets for a
    Workbook argument the macro adds sheets to. A spec that changes nothing has no
    MODIFIES lines.
21. Work with files only through MAKE_FOLDER, COPY_FILE, LIST_FILES, READ_TEXT and
    WRITE_TEXT, on paths built in place as PATH_JOIN($<root>, SAFE_NAME(...), ...) from a
    String argument (LIST_FILES may list the root itself). Check every root with
    FOLDER_EXISTS in PRE, and write each file op's status where the user can see it. No op
    moves, renames, overwrites or deletes a file.
22. Add sheets only with ADD_SHEET or COPY_SHEET, each with name=, on a Workbook argument
    declared with MODIFIES $<book> sheets, and skip the add when SHEET_EXISTS says the
    sheet is there, so a re-run is safe. Write formulas only with SET_FORMULA, using
    Excel 2016 functions; a TEST checks their text with FORMULA, never their value. A spec
    that SET_FORMULAs may not INSERT_ROW, DELETE_ROW, DEDUPE or SORT. A formula's text can't
    use COL(), so check in PRE that each column it names holds its header, such as
    COL($ws, "Method") = 3, and it may name another sheet only when every TEST workbook has
    a sheet of that name. No op removes or replaces a sheet.
```

### Reference

```text
TASK FORMAT  (keys in this order; (opt) = optional)
VCL 0.4
TASK
# ASSUME: <assumption>                        (opt, one line each)
PROC     <PascalCase name>
KIND     Sub | Function
MODULE   <modName>
ARGS     <name> As <Type>, ...                (opt)
RET      <Type>                               (Function only)
VAR      <name> As <Type>                     (opt, one line per local variable)
PRE      <condition>; <condition>             (opt, checked before any change)
MODIFIES $<arg> | $<arg>["<header>", ...] | $<arg> fills | $<arg> files | $<arg> sheets   (opt, one line each)
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
automatically; never repeat them in VAR. READ_ARRAY fills a Variant VAR, LIST_FILES a
Collection, ADD_SHEET and COPY_SHEET a Worksheet; status VARs and READ_TEXT's text VAR are
String. VARs start at 0, "" or FALSE; Dictionary and Collection VARs start empty.
For form events: MODULE is the form, PROC is <control>_<Event> (btnRun_Click).

BUILD FORMAT  (not yet tested: don't use it, rule 9)
VCL 0.4
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
<one complete TASK spec per SIG, each starting with VCL 0.4, separated by lines
containing only --->

BUILD rules: modMain always has RunHeadless (runs the whole tool without the form, so the
harness can test it) and Run (shows the form, or calls RunHeadless if there is no form).
Control types: TextBox, CommandButton, Label, ListBox, ComboBox, CheckBox, OptionButton.

OP SYNTAX
Structural ops begin an OP line:
  LOOP_ROWS dir=up|down from=<expr> to=<expr> var=<name>
  LOOP_COLS dir=left|right from=<expr> to=<expr> var=<name>
  FOR_EACH var=<name> in=SHEETS(<book>)|CELLS(<range>)|KEYS(<dict>)|ITEMS(<list>)
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
          DEDUPE(<range>, cols=<n>|[<n>, <n>, ...], header=yes|no)
  Cells   SET_VALUE(<target>, <expr>)
          WRITE_ROW(<ws>, <row>, <expr>, <expr>, ..., col=<n>)
                  (the values go to columns n, n+1, ...; without col= they start at column 1)
          SET_FORMULA(<target>, "<formula>", style=a1|r1c1)
          FORMAT(<target>, numfmt="<format>", bold=yes|no, fill=<name>|#RRGGBB, width=auto|<n>)
          CLEAR(<target>, what=all|contents|formats)
          COPY_ROW(<fromWs>, <fromRow>, <toWs>, <toRow>, cols=<n>)   (values of columns 1 to n)
          INSERT_ROW(<ws>, <row>)
          DELETE_ROW(<ws>, <row>)
  Sheets  ADD_SHEET(<book>, name=<expr>, after=<sheet>, into=<var>)    (a new empty sheet)
          COPY_SHEET(<sheet>, name=<expr>, after=<sheet>, into=<var>)  (a copy, in its workbook)
          OPEN_BOOK, SAVE_AS and CLOSE_BOOK open and save real workbooks: refused until March 2027
  Files   MAKE_FOLDER(<path>, into=<status>)       (the folder, and any missing ones below its root)
          COPY_FILE(<from>, <to>, into=<status>)   (one file; never overwrites, never deletes)
          LIST_FILES(<folder>, "<pattern>", into=<list>)   (file names in the folder, in name order)
          READ_TEXT(<path>, <text var>, into=<status>)     (a text file's text into <text var>)
          WRITE_TEXT(<path>, <expr>, into=<status>)        (a new text file; never overwrites)
  Calls   CALL(<module>.<Proc>, <args>)   (not yet tested: don't use it, rule 9)
  Output  LOG(<expr>)   MSG(<expr>)   RETURN(<expr>)
  Escape  NOTE "<terse technical text>"
FORMAT fill names: none (removes the fill), black, white, red, green, blue, yellow, magenta,
cyan. Write any other colour as #RRGGBB. FILL_COUNT takes #RRGGBB: black #000000, white
#FFFFFF, red #FF0000, green #00FF00, blue #0000FF, yellow #FFFF00, magenta #FF00FF, cyan
#00FFFF.

EXPRESSIONS
Value functions:
  LASTROW(<ws>, <col>)  LASTCOL(<ws>, <row>)  USED_RANGE(<ws>)
  COL(<ws>, "<header>")   column number of a row-1 header, or 0 if it is missing
  CELL(<ws>, <row>, <col>)  ROW_RANGE(<ws>, <row>, <fromCol>, <toCol>)  RANGE(<ws>, "<A1 ref>")
  SHEET(<book>, "<name>")  SHEET_EXISTS(<book>, <name>)  THIS_BOOK
  FIND(<range>, <expr>, match=whole|part)   the first matching cell, or NOTHING
  ROWS(<array>)  COLS(<array>)  ARR(<array>, <row>, <col>)
  COUNT(<dict or list>)  EXISTS(<dict>, <key>)
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
Also numbers, TRUE, FALSE, NOTHING. A list [<n>, <n>, ...] is only for DEDUPE cols=.
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
  CLEAN_TEXT  also removes the characters below code 32 other than tabs and line breaks;
              case is kept
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
  Dictionary  keys are compared as text ignoring case, not trimmed (TRIM them first), so
              1001 and "1001" are one key;
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
              root (LIST_FILES may list the root itself), and every part after the root is
              SAFE_NAME(<any text>) or a quoted literal that is its own safe name, written in
              place (a bare variable or cell is refused), so no path can climb out of its
              root. A SAFE_NAME result or a path is text elsewhere
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
  LIST_FILES  the names (not paths) of the files directly in the folder whose whole name
              matches the pattern (* any run of characters, ? one character, case ignored),
              sorted by lowercase name, into a Collection VAR (replacing what it held); a
              missing or bad folder gives an empty list. ITEMS walks them in that order. The
              pattern is quoted text, never an expression. Hidden and system files (such as
              the ~$ file Excel keeps for an open workbook) are left out
  Text        READ_TEXT and WRITE_TEXT handle printable ASCII, tabs and line breaks, at most
              32,767 characters; anything else is Not text, because it reads differently on
              each PC. Their paths are printable ASCII too, else Bad name
  READ_TEXT   writes, in this order: Bad name, Too long, Missing, Too big, Locked, Not text,
              Read; with Read, <text var> holds the file's text without one final line
              break, else "". In Excel also Failed
  WRITE_TEXT  writes, in this order: Bad name, Too long, Not text, No folder, Conflict (a
              folder or a different file is there: never overwritten), Locked, Already there
              (the same text: a re-run is safe), Written (exactly the text, no line break
              added). In Excel also Check write and Failed
  RANGE       a cell (B2) or a block (B2:D9); whole columns, whole rows and names stop a TEST
  THIS_BOOK   the TEST's workbook; pass sheets as arguments rather than using the active one
  READ_ARRAY  always a 2-D array from 1, even for one cell; WRITE_ARRAY writes it from the
              top-left cell of to= and stores text as exactly that text
  SELECT_CASE a CASE matches when = would be TRUE; the first match runs, else CASE_ELSE
  FOR_EACH    over CELLS gives each cell's value, row by row; over SHEETS each sheet
  INSERT_ROW  a new empty row, visible, with no fill and no number format
  LOG MSG     LOG writes a line to ForgeLog and is not an error; MSG shows a message
  FILE_NAME   the text after the last backslash
  DEDUPE      keeps the first row of each key and deletes each later row with that key
              (whole sheet rows); cols= are sheet column numbers, as for SORT; key cells
              compare as their text, case ignored, not trimmed (" W-1 " and "W-1" differ); a
              row with an empty or error key cell stays
  FIND        looks row by row, left to right: match=whole (the default) means SAME,
              match=part means the cell's text holds TRIM(<expr>), case ignored (an empty
              text finds NOTHING). Test it with IS_NOTHING before writing to it. FIND gives
              that cell only, and VCL has no row number of a cell: to change another column
              of the matching row, loop with LOOP_ROWS and SAME
  FILTER      below the range's first row, a row stays visible when its cell in sheet column
              col= is SAME as crit=, and is hidden otherwise; it never deletes and never uses
              AutoFilter. A spec that FILTERs may not SORT
  SET_FORMULA writes the formula's English text (=SUM(B2:B9), or R1C1 such as =RC[-2]*RC[-1]
              with style=r1c1). Only Excel 2016 functions: no XLOOKUP, FILTER, LET, TEXTJOIN,
              IFS, INDIRECT, TEXT, VALUE, no @, no A1#, nothing outside the workbook. style=a1
              writes one cell; style=r1c1 writes the same text to every cell of the range.
              Later steps and TESTs can't read a formula's value; a spec that SET_FORMULAs
              may not INSERT_ROW, DELETE_ROW, DEDUPE or SORT. The text is fixed, so its
              references are fixed columns or offsets: check them in PRE, such as
              COL($ws, "Qty") = 3 for =RC3*RC4. A cell formatted as Text (@) stops the run,
              because Excel would keep the formula as text
  Sheet names ADD_SHEET and COPY_SHEET take 1 to 31 characters, no spaces at the ends, none
              of \ / ? * [ ] :, no apostrophe first or last, not History; a bad name or one
              already used (case ignored) stops the run before anything is added. Without
              after= the new sheet goes after the last one

MODIFIES  (one line per sheet the macro changes, or per kind of change)
  MODIFIES $<arg>                     anything on that sheet: values, fills, formats, rows
  MODIFIES $<arg>["<header>", ...]    values in those columns, row 2 down
  MODIFIES $<arg> fills               fills anywhere on that sheet
  MODIFIES $<arg> files               folders and files may be added below that String
                                      argument; nothing else on disk may change
  MODIFIES $<arg> sheets              sheets may be added to that Workbook argument; no sheet
                                      is removed, and each sheet that was there changes only
                                      as its own MODIFIES lines allow
  An argument has one whole-sheet line, or at most one [headers] line and one fills line.
  Every write op must go through a $sheet argument whose lines allow that change: values
  need [headers] or the whole sheet, fill-only FORMATs need fills, and anything else
  (other FORMATs, CLEAR what=all or formats, SORT, INSERT_ROW, DELETE_ROW, DEDUPE, FILTER)
  needs the whole sheet. Write through CELL, ROW_RANGE, USED_RANGE, RANGE or FIND of a
  $sheet, or through the into= VAR of ADD_SHEET or COPY_SHEET, which writes to the new
  sheet and which nothing else may set.
  MAKE_FOLDER's folder, COPY_FILE's <to> and WRITE_TEXT's file must be below a root with a
  files line; COPY_FILE's <from>, READ_TEXT's file and LIST_FILES' folder may be below any
  String argument.
  Every TEST then checks that nothing else changed: each other cell keeps its value, fill
  and number format, each row stays visible or hidden, and the header row never changes
  (unless the whole sheet is declared);
  no sheet is removed, and new sheets appear only with a sheets line; on disk nothing is
  removed or changed, and new folders and files appear only below a files root. A spec
  that changes nothing has no MODIFIES lines and gets the same check.

TEST ASSERTIONS  (checked on the test workbook after the run)
  COUNTIF(<Sheet>!<range>, <criteria>)  COUNTA(<Sheet>!<range>)  SUM(<Sheet>!<range>)
  COUNTIFS(<Sheet>!<range>, <criteria>, <Sheet>!<range>, <criteria>, ...)
                 rows where every range meets its criteria
  SUMIFS(<Sheet>!<sumRange>, <Sheet>!<range>, <criteria>, ...)
                 sum of sumRange over the rows where every range meets its criteria
  VALUE(<Sheet>!<cell>)  SHEET_EXISTS("<name>")
  ROWS_USED(<Sheet>)               the last used row (row 1, the headers, counts)
  FILE_EXISTS("<Folder>\<path>")   that file exists below a FIXTURE folder of the TEST
  FILE_COUNT("<Folder>[\<path>]")  files in that folder and in every folder below it
  FILL_COUNT(<Sheet>!<range>, #RRGGBB)   cells in the range with that fill colour
  VISIBLE_COUNT(<Sheet>!<range>)   non-empty cells of the range in rows that are not hidden
  FORMULA(<Sheet>!<cell>)          the formula text as SET_FORMULA wrote it (A1 or R1C1), or ""
  FILE_TEXT("<Folder>\<path>")     the exact text of that file, or "" when there is none
  RESULT         the Function's return value (from the last run)
  ERROR_COUNT    errors written to ForgeLog during the run
  PRE_<measure>  the same measurement taken before the first run, e.g. PRE_COUNTA(Data!A:A)
                 or PRE_COUNTIFS(Data![Status], "Open", Data![Owner], "QA")
COUNTIF, COUNTIFS and SUMIFS work as in Excel 2016: criteria such as "Ready", "<>" (not
blank), "<>B-109" or ">=5"; text matching ignores case. COUNTIFS and SUMIFS ranges must be
the same size, so use [Header] ranges from the same sheet.
Ranges by header: <Sheet>![Header] is that header's whole column from row 1, the header
cell included (5 values under a header: COUNTA, COUNTIF "<>" and VISIBLE_COUNT give 6);
<Sheet>![Header]<n> is that column's cell in row n. Prefer them to column letters.
An assertion names a sheet as one word, so it can read only sheets without spaces in
their names; SHEET_EXISTS("<name>") takes any name.
runs=<n> calls the procedure n times on the same test workbook before checking
(default 1).
Combine assertions with AND; there is no TRUE, FALSE or NOT. SHEET_EXISTS(...) and
FILE_EXISTS(...) stand alone; write "is not there" as SHEET_EXISTS("<name>") = 0 or
FILE_EXISTS("<path>") = 0. When every run stops at PRE, the automatic check already proves
that nothing changed. In TEST call arguments, a bare sheet name (Orders) means that worksheet
in the test workbook, a bare folder name (Dump) means that FIXTURE folder, the TEST's fixture
file name (orders_01.xlsx) means the test workbook, for a Workbook argument, and text and
numbers are passed as written ("P-104", 5).
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
  FIXTURE <name>.xlsx folder=<Name> [files=<path>[:<bytes>|:"<text>"],...] [locked=<path>,...]
  adds a synthetic folder to the same test file: each file below it (\ between folders,
  each part a safe name of letters, digits, _ . and -), with that size or one from its own
  seed, or holding exactly "<text>" (one line of printable ASCII, commas and colons too;
  write a quote as ""). locked= files act as if another program has them open. A folder
  exists when a file is below it. <Name> is one word and not a sheet's name. Nothing real
  is read or written: the folders sit in a sandbox (C:\KRONOS_SANDBOX in the interpreter, a
  new folder below %TEMP% in Excel, whose root may be up to 120 characters), so a TEST fails
  when a file op path could be Too long in Excel but not in the interpreter: keep paths far
  over or well under the limits.
  Sheet names have no spaces, so TEST calls can use them bare. They may be short forms of
  the real sheet names (WelderPerf for "Welder Performance"); ASSUME lines describe the
  real workbook.
  Each column has its own fixed seed, made from file name + sheet + column. The builder
  uses it for everything random in that column: the shuffled order of values(...), the
  int, num, dates, reldates and text values, and which cells +blank(<n>) leaves blank;
  +dupes(<n>) and stale=<n> have seeds of their own. Columns are not correlated,
  and the same spec always builds the same file.
```

## Appendix B: the rest of the Orchestrator prompt

The parts of `vcl/orchestrator_prompt.txt` that are not in Appendix A: its role, its examples
and the final check it runs before every reply.

```text
VBA FORGE ORCHESTRATOR · SYSTEM PROMPT · VCL 0.4

ROLE
You are the Orchestrator for VBA Forge, an offline VBA code generator. You translate the
user's Excel automation request into exactly one VCL 0.4 spec. The VBA Forge compiler turns
your spec into VBA. A reference interpreter runs your TEST lines on synthetic test workbooks
and folders built from your FIXTURE lines. A Python validator rejects any spec that breaks
the rules below.
Precision matters more than anything else.
```

```text
EXAMPLE 1
User: "In the Orders sheet, delete every row whose Status says Cancelled, then sort what's
left by Order Date, newest first."
Reply (inside one vcl code block):
VCL 0.4
TASK
# ASSUME: Orders headers in row 1 include "Status" and "Order Date"
# ASSUME: "Cancelled" is matched ignoring case and extra spaces
# ASSUME: Order Date holds real dates; rows with a blank date sort to the bottom
PROC     RemoveCancelledAndSortByDate
KIND     Function
MODULE   modCleanup
ARGS     ws As Worksheet
RET      Long
VAR      cStatus As Long
VAR      cDate As Long
VAR      removed As Long
PRE      COL($ws, "Status") > 0; COL($ws, "Order Date") > 0
MODIFIES $ws
OP       LET(cStatus, COL($ws, "Status"))
OP       LET(cDate, COL($ws, "Order Date"))
OP       LET(removed, 0)
OP       LOOP_ROWS dir=up from=LASTROW($ws, cStatus) to=2 var=r
OP         IF SAME(CELL($ws, r, cStatus), "CANCELLED")
OP           DELETE_ROW($ws, r)
OP           LET(removed, removed + 1)
OP       SORT(USED_RANGE($ws), key=cDate, order=desc, header=yes)
OP       RETURN(removed)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  orders_01.xlsx sheet=Orders cols=A:Order No,B:Order Date,C:Customer,D:Status rows=40 data=A:seq(1,1);B:dates(2026-01-01,2026-09-30)+blank(2);C:text(12);D:values("Open"x20,"Cancelled"x8,"Closed"x12)
TEST     fixture=orders_01.xlsx call=RemoveCancelledAndSortByDate(Orders) runs=2 assert=COUNTIF(Orders![Status], "Cancelled") = 0 AND ERROR_COUNT = 0
TEST     fixture=orders_01.xlsx call=RemoveCancelledAndSortByDate(Orders) assert=RESULT >= PRE_COUNTIF(Orders![Status], "Cancelled") AND COUNTA(Orders![Status]) = PRE_COUNTA(Orders![Status]) - RESULT
TEST     fixture=orders_01.xlsx call=RemoveCancelledAndSortByDate(Orders) assert=VALUE(Orders![Order Date]2) >= VALUE(Orders![Order Date]3)

EXAMPLE 2
User: "Delete the old files."
Reply (inside one vcl code block):
# QUESTION: Which folder should I clean, and how many days old counts as old?

EXAMPLE 3
User: "Each row of the Tracker sheet has a Job ID and a Report File. Copy each report from
the dump folder into a folder for its job inside the jobs folder, and write what happened
in Status."
Reply (inside one vcl code block):
VCL 0.4
TASK
# ASSUME: Tracker headers in row 1 include "Job ID", "Report File" and "Status"
# ASSUME: both folders exist; each report is copied, never moved, and nothing is overwritten
# ASSUME: a blank Job ID is skipped; an unsafe Job ID or Report File gets Status "Bad name"
PROC     FileJobReports
KIND     Function
MODULE   modReports
ARGS     tracker As Worksheet, dump As String, jobs As String
RET      Long
VAR      cJob As Long
VAR      cFile As Long
VAR      cStatus As Long
VAR      status As String
VAR      filed As Long
PRE      COL($tracker, "Job ID") > 0; COL($tracker, "Report File") > 0; COL($tracker, "Status") > 0; FOLDER_EXISTS($dump); FOLDER_EXISTS($jobs)
MODIFIES $tracker["Status"]
MODIFIES $jobs files
OP       LET(cJob, COL($tracker, "Job ID"))
OP       LET(cFile, COL($tracker, "Report File"))
OP       LET(cStatus, COL($tracker, "Status"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($tracker, cJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($tracker, r, cJob))
OP         IF SAFE_NAME(CELL($tracker, r, cJob)) = "" OR SAFE_NAME(CELL($tracker, r, cFile)) = ""
OP           SET_VALUE(CELL($tracker, r, cStatus), "Bad name")
OP         ELSE
OP           MAKE_FOLDER(PATH_JOIN($jobs, SAFE_NAME(CELL($tracker, r, cJob))), into=status)
OP           IF status = "Created" OR status = "Exists"
OP             COPY_FILE(PATH_JOIN($dump, SAFE_NAME(CELL($tracker, r, cFile))), PATH_JOIN($jobs, SAFE_NAME(CELL($tracker, r, cJob)), SAFE_NAME(CELL($tracker, r, cFile))), into=status)
OP           SET_VALUE(CELL($tracker, r, cStatus), status)
OP           IF status = "Copied" OR status = "Already there" THEN LET(filed, filed + 1)
OP       RETURN(filed)
PERF     screen_off
ERR      standard
FIXTURE  reports_01.xlsx sheet=Tracker cols=A:Job ID,B:Report File,C:Status rows=7 data=A:values("J-1"x2,"J-2"x1,"..\x"x1,""x1,"J-3"x1,"J-4"x1)+boundary("J-1","J-2","..\x","","J-3","J-1","J-4");B:values("a.pdf"x2,"b.pdf"x2,"c.pdf"x1,"d.pdf"x1,"CON.pdf"x1)+boundary("a.pdf","b.pdf","a.pdf","b.pdf","c.pdf","d.pdf","CON.pdf");C:values("old"x7)
FIXTURE  reports_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  reports_01.xlsx folder=Dump files=a.pdf:100,b.pdf:80,c.pdf locked=c.pdf
FIXTURE  reports_01.xlsx folder=Jobs files=J-2\b.pdf:99
TEST     fixture=reports_01.xlsx call=FileJobReports(Tracker, Dump, Jobs) runs=2 assert=RESULT = 1 AND ERROR_COUNT = 0 AND COUNTIF(Tracker![Status], "Already there") = 1 AND COUNTIF(Tracker![Status], "Conflict") = 1 AND COUNTIF(Tracker![Status], "Locked") = 1 AND COUNTIF(Tracker![Status], "Missing") = 1 AND COUNTIF(Tracker![Status], "Bad name") = 2 AND FILE_EXISTS("Jobs\J-1\a.pdf") AND FILE_COUNT("Jobs") = PRE_FILE_COUNT("Jobs") + 1
TEST     fixture=reports_01.xlsx call=FileJobReports(Tracker, Dump, Jobs) assert=COUNTIF(Tracker![Status], "Copied") = 1 AND COUNTIF(Tracker![Status], "old") = 1 AND FILE_COUNT("Dump") = PRE_FILE_COUNT("Dump")
TEST     fixture=reports_01.xlsx call=FileJobReports(Notes, Dump, Jobs) assert=RESULT = -1 AND ERROR_COUNT = 1 AND FILE_COUNT("Jobs") = PRE_FILE_COUNT("Jobs")

FINAL CHECK  (verify silently before every reply)
- Only keys, ops and functions defined here are used
- One TASK of up to about 40 OP lines; no BUILD and no CALL yet (rule 9)
- Every name is an argument ($name), a VAR, or a loop variable
- Every input column comes from COL() and its header is checked in PRE
- Calculations and date comparisons on workbook values are guarded by IS_NUMBER or IS_DATE
  in a nested IF on the value itself, never through a flag
- No option word (dir, from, to, var, in, mode, into, key, order, header, ...) is used as a name
- Output sheets are rebuilt, never appended to; the goal TEST uses runs=2
- Every sheet the macro changes has a MODIFIES line that allows that change (rule 20)
- SKIP_IF and EXIT_LOOP_IF only inside a loop; EXIT_IF leaves the whole procedure
- Text compared with SAME; PARSE_DATE always has a format and its result an IS_DATE guard
- At least two TESTs: one compares a PRE_ measure to guard against collateral changes, one
  has ERROR_COUNT = 0
- Every TEST uses a named synthetic FIXTURE: no fixture=target, no real workbooks
- Rebuilt sheets have stale=<n>; comparisons have +boundary values; dictionary keys repeat;
  TODAY means reldates; SKIP_IF IS_EMPTY columns have a blank cell; copies and reads have a
  locked= file (rule 19)
- Files only through MAKE_FOLDER, COPY_FILE, LIST_FILES, READ_TEXT and WRITE_TEXT, on
  PATH_JOIN($root, SAFE_NAME(...)) paths; every root checked with FOLDER_EXISTS in PRE;
  MODIFIES $root files for each root that gets new folders or files (rules 20 and 21)
- Sheets added only with ADD_SHEET or COPY_SHEET, with name=, skipped when SHEET_EXISTS,
  and MODIFIES $book sheets; formulas only with SET_FORMULA and Excel 2016 functions,
  tested with FORMULA, never beside INSERT_ROW, DELETE_ROW, DEDUPE or SORT (rule 22)
- Rows are deleted only inside dir=up loops or by DEDUPE, and never just because data is
  missing
- Risky actions have ALLOW, and the user asked for them
- No VBA syntax anywhere: no Dim, Set, End If, .Value, Cells(, Range(
```

## Appendix C: the specs, each with the VBA the compiler writes for it

Every spec in `specs/`, then the procedure `vcl_codegen.py` makes from it. In Excel each
procedure calls ForgeLib (Appendix F); `dist/forge_module.bas` is ForgeLib plus these.

### task_01.vcl: BuildShortageList

```vcl
VCL 0.2
TASK
# ASSUME: BOM headers (row 1) include "Part No" and "Required Qty"; GRN headers include "Part No" and "Received Qty"; column order does not matter
# ASSUME: a part can appear on several GRN rows; its received quantities are added
# ASSUME: the Shortages sheet already exists; it is cleared and rebuilt on every run
# ASSUME: a blank or non-numeric Received Qty is ignored; a part with no valid GRN rows counts as 0 received
# ASSUME: a blank or non-numeric Required Qty is listed as "Required qty missing" instead of being skipped
PROC     BuildShortageList
KIND     Function
MODULE   modMaterials
ARGS     bom As Worksheet, grn As Worksheet, out As Worksheet
RET      Long
VAR      received As Dictionary
VAR      bomPart As Long
VAR      bomReq As Long
VAR      grnPart As Long
VAR      grnQty As Long
VAR      partNo As String
VAR      got As Double
VAR      outRow As Long
PRE      COL($bom, "Part No") > 0; COL($bom, "Required Qty") > 0; COL($grn, "Part No") > 0; COL($grn, "Received Qty") > 0
MODIFIES $out
OP       LET(bomPart, COL($bom, "Part No"))
OP       LET(bomReq, COL($bom, "Required Qty"))
OP       LET(grnPart, COL($grn, "Part No"))
OP       LET(grnQty, COL($grn, "Received Qty"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Part No", "Required", "Received", "Shortage", "Remark")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($grn, grnPart) var=r
OP         IF NOT IS_EMPTY(CELL($grn, r, grnPart)) AND IS_NUMBER(CELL($grn, r, grnQty)) THEN DICT_ADD(received, TRIM(CELL($grn, r, grnPart)), CELL($grn, r, grnQty), mode=sum)
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($bom, bomPart) var=r
OP         SKIP_IF IS_EMPTY(CELL($bom, r, bomPart))
OP         LET(partNo, TRIM(CELL($bom, r, bomPart)))
OP         LET(got, 0)
OP         IF EXISTS(received, partNo) THEN DICT_GET(received, partNo, into=got)
OP         IF NOT IS_NUMBER(CELL($bom, r, bomReq))
OP           SET_VALUE(CELL($out, outRow, 1), partNo)
OP           SET_VALUE(CELL($out, outRow, 3), got)
OP           SET_VALUE(CELL($out, outRow, 5), "Required qty missing")
OP           LET(outRow, outRow + 1)
OP         ELSE
OP           IF got < CELL($bom, r, bomReq)
OP             WRITE_ROW($out, outRow, partNo, CELL($bom, r, bomReq), got, CELL($bom, r, bomReq) - got)
OP             IF got = 0 THEN SET_VALUE(CELL($out, outRow, 5), "Not received")
OP             IF got > 0 THEN SET_VALUE(CELL($out, outRow, 5), "Short")
OP             LET(outRow, outRow + 1)
OP       RETURN(outRow - 2)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  bom_grn_01.xlsx sheet=BOM cols=A:Part No,B:Description,C:Required Qty rows=40 data=A:seq(1001,1)+blank(2);B:text(12);C:int(1,50)+boundary(30,7)+blank(3)
FIXTURE  bom_grn_01.xlsx sheet=GRN cols=A:Part No,B:Received Qty rows=35 data=A:seq(1001,1)+boundary(1001,1002)+dupes(4);B:int(1,50)+boundary(1,7)+blank(4)
FIXTURE  bom_grn_01.xlsx sheet=Shortages cols=A:Part No,B:Required,C:Received,D:Shortage,E:Remark rows=0 stale=40
TEST     fixture=bom_grn_01.xlsx call=BuildShortageList(BOM, GRN, Shortages) runs=2 assert=RESULT >= 5 AND COUNTA(Shortages![Part No]) = RESULT + 1 AND COUNTIF(Shortages![Part No], "stale") = 0 AND COUNTIF(Shortages![Shortage], "<0") = 0 AND ERROR_COUNT = 0
TEST     fixture=bom_grn_01.xlsx call=BuildShortageList(BOM, GRN, Shortages) assert=COUNTIF(Shortages![Remark], "Required qty missing") = 3 AND COUNTIFS(Shortages![Part No], 1001, Shortages![Remark], "Short") = 1 AND COUNTIF(Shortages![Part No], 1002) = 0
TEST     fixture=bom_grn_01.xlsx call=BuildShortageList(BOM, GRN, Shortages) assert=SUM(BOM![Required Qty]) = PRE_SUM(BOM![Required Qty]) AND SUM(GRN![Received Qty]) = PRE_SUM(GRN![Received Qty])
TEST     fixture=bom_grn_01.xlsx call=BuildShortageList(Shortages, GRN, Shortages) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(Shortages![Part No]) = PRE_COUNTA(Shortages![Part No])
```

```vba
'@VCL PROC=BuildShortageList  source: task_01.vcl
Public Function BuildShortageList(ByVal bom As Worksheet, ByVal grn As Worksheet, ByVal out As Worksheet) As Long
          Dim received As Object
          Dim bomPart As Long
          Dim bomReq As Long
          Dim grnPart As Long
          Dim grnQty As Long
          Dim partNo As String
          Dim got As Double
          Dim outRow As Long
          Dim r As Long
          Dim forgeScreen As Boolean
          Dim forgeCalc As Long
          Dim forgeEvents As Boolean
          forgeScreen = Application.ScreenUpdating
          forgeCalc = Application.Calculation
          forgeEvents = Application.EnableEvents
          On Error GoTo ErrHandler
          Set received = ForgeNewDict()
210       If Not (ForgeCol(bom, "Part No") > 0) Then ForgeLog "BuildShortageList", 0, "PRE failed: (COL($bom,""Part No"")>0)", 210: BuildShortageList = -1: GoTo CleanExit
211       If Not (ForgeCol(bom, "Required Qty") > 0) Then ForgeLog "BuildShortageList", 0, "PRE failed: (COL($bom,""Required Qty"")>0)", 210: BuildShortageList = -1: GoTo CleanExit
212       If Not (ForgeCol(grn, "Part No") > 0) Then ForgeLog "BuildShortageList", 0, "PRE failed: (COL($grn,""Part No"")>0)", 210: BuildShortageList = -1: GoTo CleanExit
213       If Not (ForgeCol(grn, "Received Qty") > 0) Then ForgeLog "BuildShortageList", 0, "PRE failed: (COL($grn,""Received Qty"")>0)", 210: BuildShortageList = -1: GoTo CleanExit
          Application.ScreenUpdating = False
          Application.Calculation = xlCalculationManual
230       bomPart = ForgeCol(bom, "Part No")
240       bomReq = ForgeCol(bom, "Required Qty")
250       grnPart = ForgeCol(grn, "Part No")
260       grnQty = ForgeCol(grn, "Received Qty")
270       ForgeUsedRange(out).Clear
280       ForgeWriteRow out, 1, 1, Array("Part No", "Required", "Received", "Shortage", "Remark")
290       For r = 2 To ForgeLastRow(grn, grnPart)
300           If ((Not ForgeIsEmpty(grn.Cells(r, grnPart).Value)) And ForgeIsNumber(grn.Cells(r, grnQty).Value)) Then ForgeDictAdd received, ForgeTrim(grn.Cells(r, grnPart).Value), grn.Cells(r, grnQty).Value, "sum"
          Next r
310       outRow = 2
320       For r = 2 To ForgeLastRow(bom, bomPart)
330           If ForgeIsEmpty(bom.Cells(r, bomPart).Value) Then GoTo Next_32
340           partNo = ForgeTrim(bom.Cells(r, bomPart).Value)
350           got = 0
360           If ForgeDictExists(received, partNo) Then got = ForgeNum(ForgeDictGet(received, partNo))
370           If (Not ForgeIsNumber(bom.Cells(r, bomReq).Value)) Then
380               ForgeSetValue out.Cells(outRow, 1), partNo
390               ForgeSetValue out.Cells(outRow, 3), got
400               ForgeSetValue out.Cells(outRow, 5), "Required qty missing"
410               outRow = (outRow + 1)
              Else
430               If (got < ForgeNum(bom.Cells(r, bomReq).Value)) Then
440                   ForgeWriteRow out, outRow, 1, Array(partNo, bom.Cells(r, bomReq).Value, got, (ForgeNum(bom.Cells(r, bomReq).Value) - got))
450                   If (got = 0) Then ForgeSetValue out.Cells(outRow, 5), "Not received"
460                   If (got > 0) Then ForgeSetValue out.Cells(outRow, 5), "Short"
470                   outRow = (outRow + 1)
                  End If
              End If
Next_32:
          Next r
480       BuildShortageList = (outRow - 2)
481       GoTo CleanExit
CleanExit:
          ForgeRestore forgeScreen, forgeCalc, forgeEvents
          Exit Function
ErrHandler:
          ForgeLog "BuildShortageList", Err.Number, Err.Description, Erl
          BuildShortageList = -2
          Resume CleanExit
End Function
```

### task_02.vcl: FlagMissingTraceability

```vcl
VCL 0.2
TASK
# ASSUME: Material Traceability headers (row 1) include "Part No", "Plate/Heat No" and "MTC No"; column order does not matter
# ASSUME: rows with a blank Part No are skipped
# ASSUME: the macro owns the fill colour of listed rows: it clears the fill first, so fixed rows lose their red
# ASSUME: the whole row is highlighted, from column A to the last header column
PROC     FlagMissingTraceability
KIND     Function
MODULE   modMaterials
ARGS     ws As Worksheet
RET      Long
VAR      cPart As Long
VAR      cHeat As Long
VAR      cMtc As Long
VAR      lastCol As Long
VAR      flagged As Long
PRE      COL($ws, "Part No") > 0; COL($ws, "Plate/Heat No") > 0; COL($ws, "MTC No") > 0
MODIFIES $ws fills
OP       LET(cPart, COL($ws, "Part No"))
OP       LET(cHeat, COL($ws, "Plate/Heat No"))
OP       LET(cMtc, COL($ws, "MTC No"))
OP       LET(lastCol, LASTCOL($ws, 1))
OP       LET(flagged, 0)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cPart) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cPart))
OP         FORMAT(ROW_RANGE($ws, r, 1, lastCol), fill=none)
OP         IF IS_EMPTY(CELL($ws, r, cHeat)) OR IS_EMPTY(CELL($ws, r, cMtc))
OP           FORMAT(ROW_RANGE($ws, r, 1, lastCol), fill=#FF0000)
OP           LET(flagged, flagged + 1)
OP       RETURN(flagged)
PERF     screen_off
ERR      standard
FIXTURE  traceability_01.xlsx sheet=Traceability cols=A:Part No,B:Plate/Heat No,C:MTC No,D:Remarks rows=40 stale=6 data=A:seq(1,1)+blank(2);B:text(8)+boundary("")+blank(4);C:text(10)+blank(3);D:text(10)+boundary("")
FIXTURE  traceability_01.xlsx sheet=Notes cols=A:Note rows=3 data=A:text(20)
TEST     fixture=traceability_01.xlsx call=FlagMissingTraceability(Traceability) runs=2 assert=RESULT >= 5 AND RESULT <= 8 AND FILL_COUNT(Traceability![Part No], #FF0000) = RESULT + 1 AND FILL_COUNT(Traceability![Part No]2, #FF0000) = 1 AND ERROR_COUNT = 0
TEST     fixture=traceability_01.xlsx call=FlagMissingTraceability(Traceability) assert=COUNTA(Traceability![Plate/Heat No]) = PRE_COUNTA(Traceability![Plate/Heat No]) AND COUNTA(Traceability![MTC No]) = PRE_COUNTA(Traceability![MTC No])
TEST     fixture=traceability_01.xlsx call=FlagMissingTraceability(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1 AND FILL_COUNT(Notes![Note], #FF0000) = 0
```

```vba
'@VCL PROC=FlagMissingTraceability  source: task_02.vcl
Public Function FlagMissingTraceability(ByVal ws As Worksheet) As Long
          Dim cPart As Long
          Dim cHeat As Long
          Dim cMtc As Long
          Dim lastCol As Long
          Dim flagged As Long
          Dim r As Long
          Dim forgeScreen As Boolean
          Dim forgeCalc As Long
          Dim forgeEvents As Boolean
          forgeScreen = Application.ScreenUpdating
          forgeCalc = Application.Calculation
          forgeEvents = Application.EnableEvents
          On Error GoTo ErrHandler
170       If Not (ForgeCol(ws, "Part No") > 0) Then ForgeLog "FlagMissingTraceability", 0, "PRE failed: (COL($ws,""Part No"")>0)", 170: FlagMissingTraceability = -1: GoTo CleanExit
171       If Not (ForgeCol(ws, "Plate/Heat No") > 0) Then ForgeLog "FlagMissingTraceability", 0, "PRE failed: (COL($ws,""Plate/Heat No"")>0)", 170: FlagMissingTraceability = -1: GoTo CleanExit
172       If Not (ForgeCol(ws, "MTC No") > 0) Then ForgeLog "FlagMissingTraceability", 0, "PRE failed: (COL($ws,""MTC No"")>0)", 170: FlagMissingTraceability = -1: GoTo CleanExit
          Application.ScreenUpdating = False
190       cPart = ForgeCol(ws, "Part No")
200       cHeat = ForgeCol(ws, "Plate/Heat No")
210       cMtc = ForgeCol(ws, "MTC No")
220       lastCol = ForgeLastCol(ws, 1)
230       flagged = 0
240       For r = 2 To ForgeLastRow(ws, cPart)
250           If ForgeIsEmpty(ws.Cells(r, cPart).Value) Then GoTo Next_24
260           ForgeNoFill ForgeRowRange(ws, r, 1, lastCol)
270           If (ForgeIsEmpty(ws.Cells(r, cHeat).Value) Or ForgeIsEmpty(ws.Cells(r, cMtc).Value)) Then
280               ForgeRowRange(ws, r, 1, lastCol).Interior.Color = RGB(255, 0, 0)
290               flagged = (flagged + 1)
              End If
Next_24:
          Next r
300       FlagMissingTraceability = flagged
301       GoTo CleanExit
CleanExit:
          ForgeRestore forgeScreen, forgeCalc, forgeEvents
          Exit Function
ErrHandler:
          ForgeLog "FlagMissingTraceability", Err.Number, Err.Description, Erl
          FlagMissingTraceability = -2
          Resume CleanExit
End Function
```

### task_03.vcl: BuildWelderRepairRate

```vcl
VCL 0.2
TASK
# ASSUME: NDE Report headers (row 1) include "Weld Seam No", "Welder ID" and "Result" (Accepted / Repair); one row per seam
# ASSUME: rows with a blank Welder ID or a blank Result are skipped; a blank Result means not inspected yet
# ASSUME: "Repair" is matched ignoring case and extra spaces
# ASSUME: repair rate = repaired seams / inspected seams per welder; Welder Performance is rebuilt every run, highest rate first
PROC     BuildWelderRepairRate
KIND     Function
MODULE   modNde
ARGS     nde As Worksheet, out As Worksheet
RET      Long
VAR      total As Dictionary
VAR      repairs As Dictionary
VAR      cSeam As Long
VAR      cWelder As Long
VAR      cResult As Long
VAR      welderId As String
VAR      inspected As Long
VAR      repaired As Long
VAR      outRow As Long
PRE      COL($nde, "Weld Seam No") > 0; COL($nde, "Welder ID") > 0; COL($nde, "Result") > 0
MODIFIES $out
OP       LET(cSeam, COL($nde, "Weld Seam No"))
OP       LET(cWelder, COL($nde, "Welder ID"))
OP       LET(cResult, COL($nde, "Result"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Welder ID", "Seams Inspected", "Repairs", "Repair Rate")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($nde, cSeam) var=r
OP         IF NOT IS_EMPTY(CELL($nde, r, cWelder)) AND NOT IS_EMPTY(CELL($nde, r, cResult))
OP           LET(welderId, TRIM(CELL($nde, r, cWelder)))
OP           DICT_ADD(total, welderId, 1, mode=count)
OP           IF SAME(CELL($nde, r, cResult), "REPAIR") THEN DICT_ADD(repairs, welderId, 1, mode=count)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(total)
OP         DICT_GET(total, k, into=inspected)
OP         LET(repaired, 0)
OP         IF EXISTS(repairs, k) THEN DICT_GET(repairs, k, into=repaired)
OP         WRITE_ROW($out, outRow, k, inspected, repaired)
OP         SET_VALUE(CELL($out, outRow, 4), repaired / inspected)
OP         FORMAT(CELL($out, outRow, 4), numfmt="0.0%")
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=4, order=desc, header=yes)
OP       RETURN(outRow - 2)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  nde_repair_01.xlsx sheet=NDE cols=A:Weld Seam No,B:Welder ID,C:Result rows=60 data=A:seq(1,1);B:values("W-01"x20,"W-02"x19,"W-03"x15,"W-04"x1,""x5)+boundary("W-01","W-04");C:values("Accepted"x48,"Repair"x8,""x4)+boundary("Repair","Accepted")
FIXTURE  nde_repair_01.xlsx sheet=WelderPerf cols=A:Welder ID,B:Seams Inspected,C:Repairs,D:Repair Rate rows=0 stale=10
TEST     fixture=nde_repair_01.xlsx call=BuildWelderRepairRate(NDE, WelderPerf) runs=2 assert=RESULT = 4 AND COUNTA(WelderPerf![Welder ID]) = RESULT + 1 AND ERROR_COUNT = 0
TEST     fixture=nde_repair_01.xlsx call=BuildWelderRepairRate(NDE, WelderPerf) assert=SUM(WelderPerf![Repairs]) = PRE_COUNTIFS(NDE![Result], "Repair", NDE![Welder ID], "<>") AND SUM(WelderPerf![Seams Inspected]) = PRE_COUNTIFS(NDE![Result], "<>", NDE![Welder ID], "<>") - 1 AND COUNTIFS(WelderPerf![Welder ID], "W-04", WelderPerf![Repairs], 0) = 1 AND VALUE(WelderPerf![Repair Rate]2) >= VALUE(WelderPerf![Repair Rate]3)
TEST     fixture=nde_repair_01.xlsx call=BuildWelderRepairRate(NDE, WelderPerf) assert=COUNTA(NDE![Welder ID]) = PRE_COUNTA(NDE![Welder ID]) AND COUNTA(NDE![Result]) = PRE_COUNTA(NDE![Result])
TEST     fixture=nde_repair_01.xlsx call=BuildWelderRepairRate(WelderPerf, WelderPerf) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(WelderPerf![Welder ID]) = PRE_COUNTA(WelderPerf![Welder ID])
```

```vba
'@VCL PROC=BuildWelderRepairRate  source: task_03.vcl
Public Function BuildWelderRepairRate(ByVal nde As Worksheet, ByVal out As Worksheet) As Long
          Dim total As Object
          Dim repairs As Object
          Dim cSeam As Long
          Dim cWelder As Long
          Dim cResult As Long
          Dim welderId As String
          Dim inspected As Long
          Dim repaired As Long
          Dim outRow As Long
          Dim r As Long
          Dim k As Variant
          Dim forgeScreen As Boolean
          Dim forgeCalc As Long
          Dim forgeEvents As Boolean
          forgeScreen = Application.ScreenUpdating
          forgeCalc = Application.Calculation
          forgeEvents = Application.EnableEvents
          On Error GoTo ErrHandler
          Set total = ForgeNewDict()
          Set repairs = ForgeNewDict()
210       If Not (ForgeCol(nde, "Weld Seam No") > 0) Then ForgeLog "BuildWelderRepairRate", 0, "PRE failed: (COL($nde,""Weld Seam No"")>0)", 210: BuildWelderRepairRate = -1: GoTo CleanExit
211       If Not (ForgeCol(nde, "Welder ID") > 0) Then ForgeLog "BuildWelderRepairRate", 0, "PRE failed: (COL($nde,""Welder ID"")>0)", 210: BuildWelderRepairRate = -1: GoTo CleanExit
212       If Not (ForgeCol(nde, "Result") > 0) Then ForgeLog "BuildWelderRepairRate", 0, "PRE failed: (COL($nde,""Result"")>0)", 210: BuildWelderRepairRate = -1: GoTo CleanExit
          Application.ScreenUpdating = False
          Application.Calculation = xlCalculationManual
230       cSeam = ForgeCol(nde, "Weld Seam No")
240       cWelder = ForgeCol(nde, "Welder ID")
250       cResult = ForgeCol(nde, "Result")
260       ForgeUsedRange(out).Clear
270       ForgeWriteRow out, 1, 1, Array("Welder ID", "Seams Inspected", "Repairs", "Repair Rate")
280       For r = 2 To ForgeLastRow(nde, cSeam)
290           If ((Not ForgeIsEmpty(nde.Cells(r, cWelder).Value)) And (Not ForgeIsEmpty(nde.Cells(r, cResult).Value))) Then
300               welderId = ForgeTrim(nde.Cells(r, cWelder).Value)
310               ForgeDictAdd total, welderId, 1, "count"
320               If ForgeSame(nde.Cells(r, cResult).Value, "REPAIR") Then ForgeDictAdd repairs, welderId, 1, "count"
              End If
          Next r
330       outRow = 2
340       For Each k In total.Keys
350           inspected = ForgeNum(ForgeDictGet(total, k))
360           repaired = 0
370           If ForgeDictExists(repairs, k) Then repaired = ForgeNum(ForgeDictGet(repairs, k))
380           ForgeWriteRow out, outRow, 1, Array(k, inspected, repaired)
390           ForgeSetValue out.Cells(outRow, 4), (repaired / inspected)
400           out.Cells(outRow, 4).NumberFormat = "0.0%"
410           outRow = (outRow + 1)
          Next k
420       ForgeSort ForgeUsedRange(out), 4, False, True
430       BuildWelderRepairRate = (outRow - 2)
431       GoTo CleanExit
CleanExit:
          ForgeRestore forgeScreen, forgeCalc, forgeEvents
          Exit Function
ErrHandler:
          ForgeLog "BuildWelderRepairRate", Err.Number, Err.Description, Erl
          BuildWelderRepairRate = -2
          Resume CleanExit
End Function
```

### task_04.vcl: BuildNdeSummary

```vcl
VCL 0.2
TASK
# ASSUME: the RT, UT, MPT and DPT sheets each have headers "Weld Seam No" and "Result" in row 1
# ASSUME: seam numbers are compared as trimmed text, ignoring case
# ASSUME: if a seam appears more than once in a sheet (re-shoot), the lowest row is the latest result
# ASSUME: Summary has one row per seam found in any of the four sheets; it is rebuilt every run, sorted by seam
PROC     BuildNdeSummary
KIND     Function
MODULE   modNde
ARGS     rt As Worksheet, ut As Worksheet, mpt As Worksheet, dpt As Worksheet, out As Worksheet
RET      Long
VAR      seams As Dictionary
VAR      rtResults As Dictionary
VAR      utResults As Dictionary
VAR      mptResults As Dictionary
VAR      dptResults As Dictionary
VAR      cRtSeam As Long
VAR      cRtRes As Long
VAR      cUtSeam As Long
VAR      cUtRes As Long
VAR      cMptSeam As Long
VAR      cMptRes As Long
VAR      cDptSeam As Long
VAR      cDptRes As Long
VAR      v As Variant
VAR      outRow As Long
PRE      COL($rt, "Weld Seam No") > 0; COL($rt, "Result") > 0; COL($ut, "Weld Seam No") > 0; COL($ut, "Result") > 0; COL($mpt, "Weld Seam No") > 0; COL($mpt, "Result") > 0; COL($dpt, "Weld Seam No") > 0; COL($dpt, "Result") > 0
MODIFIES $out
OP       LET(cRtSeam, COL($rt, "Weld Seam No"))
OP       LET(cRtRes, COL($rt, "Result"))
OP       LET(cUtSeam, COL($ut, "Weld Seam No"))
OP       LET(cUtRes, COL($ut, "Result"))
OP       LET(cMptSeam, COL($mpt, "Weld Seam No"))
OP       LET(cMptRes, COL($mpt, "Result"))
OP       LET(cDptSeam, COL($dpt, "Weld Seam No"))
OP       LET(cDptRes, COL($dpt, "Result"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Weld Seam No", "RT", "UT", "MPT", "DPT")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($rt, cRtSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($rt, r, cRtSeam))
OP         DICT_ADD(seams, TRIM(CELL($rt, r, cRtSeam)), 1, mode=keep)
OP         DICT_ADD(rtResults, TRIM(CELL($rt, r, cRtSeam)), CELL($rt, r, cRtRes), mode=overwrite)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ut, cUtSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($ut, r, cUtSeam))
OP         DICT_ADD(seams, TRIM(CELL($ut, r, cUtSeam)), 1, mode=keep)
OP         DICT_ADD(utResults, TRIM(CELL($ut, r, cUtSeam)), CELL($ut, r, cUtRes), mode=overwrite)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($mpt, cMptSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($mpt, r, cMptSeam))
OP         DICT_ADD(seams, TRIM(CELL($mpt, r, cMptSeam)), 1, mode=keep)
OP         DICT_ADD(mptResults, TRIM(CELL($mpt, r, cMptSeam)), CELL($mpt, r, cMptRes), mode=overwrite)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($dpt, cDptSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($dpt, r, cDptSeam))
OP         DICT_ADD(seams, TRIM(CELL($dpt, r, cDptSeam)), 1, mode=keep)
OP         DICT_ADD(dptResults, TRIM(CELL($dpt, r, cDptSeam)), CELL($dpt, r, cDptRes), mode=overwrite)
OP       LET(outRow, 2)
OP       FOR_EACH var=s in=KEYS(seams)
OP         SET_VALUE(CELL($out, outRow, 1), s)
OP         LET(v, "")
OP         IF EXISTS(rtResults, s) THEN DICT_GET(rtResults, s, into=v)
OP         IF IS_EMPTY(v) THEN LET(v, "Pending")
OP         SET_VALUE(CELL($out, outRow, 2), v)
OP         LET(v, "")
OP         IF EXISTS(utResults, s) THEN DICT_GET(utResults, s, into=v)
OP         IF IS_EMPTY(v) THEN LET(v, "Pending")
OP         SET_VALUE(CELL($out, outRow, 3), v)
OP         LET(v, "")
OP         IF EXISTS(mptResults, s) THEN DICT_GET(mptResults, s, into=v)
OP         IF IS_EMPTY(v) THEN LET(v, "Pending")
OP         SET_VALUE(CELL($out, outRow, 4), v)
OP         LET(v, "")
OP         IF EXISTS(dptResults, s) THEN DICT_GET(dptResults, s, into=v)
OP         IF IS_EMPTY(v) THEN LET(v, "Pending")
OP         SET_VALUE(CELL($out, outRow, 5), v)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=1, order=asc, header=yes)
OP       RETURN(outRow - 2)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  nde_summary_01.xlsx sheet=RT cols=A:Weld Seam No,B:Result rows=30 data=A:seq(33,-1)+boundary(101,111,111)+blank(2);B:values("Accepted"x25,"Repair"x2,""x3)+boundary("Accepted","Accepted","Repair")
FIXTURE  nde_summary_01.xlsx sheet=UT cols=A:Weld Seam No,B:Result rows=20 data=A:seq(1,1)+boundary(102,112,112)+blank(2);B:values("Accepted"x17,"Repair"x1,""x2)+boundary("Accepted","Repair","Accepted")
FIXTURE  nde_summary_01.xlsx sheet=MPT cols=A:Weld Seam No,B:Result rows=25 data=A:seq(6,1)+boundary(103,113,113)+blank(2);B:values("Accepted"x22,"Repair"x1,""x2)+boundary("Accepted","Accepted","Repair")
FIXTURE  nde_summary_01.xlsx sheet=DPT cols=A:Weld Seam No,B:Result rows=10 data=A:seq(1,1)+boundary(104,114,114)+blank(1);B:values("Accepted"x9,"Repair"x1)+boundary("Accepted","Repair","Accepted")
FIXTURE  nde_summary_01.xlsx sheet=Summary cols=A:Weld Seam No,B:RT,C:UT,D:MPT,E:DPT rows=0 stale=40
TEST     fixture=nde_summary_01.xlsx call=BuildNdeSummary(RT, UT, MPT, DPT, Summary) runs=2 assert=RESULT = 35 AND COUNTA(Summary![Weld Seam No]) = 36 AND COUNTIF(Summary![RT], "stale") = 0 AND ERROR_COUNT = 0
TEST     fixture=nde_summary_01.xlsx call=BuildNdeSummary(RT, UT, MPT, DPT, Summary) assert=COUNTIF(Summary![RT], "Pending") = 11 AND COUNTIF(Summary![UT], "Pending") = 20 AND COUNTIF(Summary![MPT], "Pending") = 15 AND COUNTIF(Summary![DPT], "Pending") = 27 AND COUNTIFS(Summary![Weld Seam No], 111, Summary![RT], "Repair") = 1 AND COUNTIFS(Summary![Weld Seam No], 112, Summary![UT], "Accepted") = 1 AND COUNTIFS(Summary![Weld Seam No], 113, Summary![MPT], "Repair") = 1 AND COUNTIFS(Summary![Weld Seam No], 114, Summary![DPT], "Accepted") = 1
TEST     fixture=nde_summary_01.xlsx call=BuildNdeSummary(RT, UT, MPT, DPT, Summary) assert=COUNTA(RT![Result]) = PRE_COUNTA(RT![Result]) AND COUNTA(UT![Result]) = PRE_COUNTA(UT![Result]) AND COUNTA(MPT![Result]) = PRE_COUNTA(MPT![Result]) AND COUNTA(DPT![Result]) = PRE_COUNTA(DPT![Result])
TEST     fixture=nde_summary_01.xlsx call=BuildNdeSummary(RT, Summary, MPT, DPT, Summary) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(Summary![Weld Seam No]) = PRE_COUNTA(Summary![Weld Seam No])
```

```vba
'@VCL PROC=BuildNdeSummary  source: task_04.vcl
Public Function BuildNdeSummary(ByVal rt As Worksheet, ByVal ut As Worksheet, ByVal mpt As Worksheet, ByVal dpt As Worksheet, ByVal out As Worksheet) As Long
          Dim seams As Object
          Dim rtResults As Object
          Dim utResults As Object
          Dim mptResults As Object
          Dim dptResults As Object
          Dim cRtSeam As Long
          Dim cRtRes As Long
          Dim cUtSeam As Long
          Dim cUtRes As Long
          Dim cMptSeam As Long
          Dim cMptRes As Long
          Dim cDptSeam As Long
          Dim cDptRes As Long
          Dim v As Variant
          Dim outRow As Long
          Dim r As Long
          Dim s As Variant
          Dim forgeScreen As Boolean
          Dim forgeCalc As Long
          Dim forgeEvents As Boolean
          forgeScreen = Application.ScreenUpdating
          forgeCalc = Application.Calculation
          forgeEvents = Application.EnableEvents
          On Error GoTo ErrHandler
          Set seams = ForgeNewDict()
          Set rtResults = ForgeNewDict()
          Set utResults = ForgeNewDict()
          Set mptResults = ForgeNewDict()
          Set dptResults = ForgeNewDict()
270       If Not (ForgeCol(rt, "Weld Seam No") > 0) Then ForgeLog "BuildNdeSummary", 0, "PRE failed: (COL($rt,""Weld Seam No"")>0)", 270: BuildNdeSummary = -1: GoTo CleanExit
271       If Not (ForgeCol(rt, "Result") > 0) Then ForgeLog "BuildNdeSummary", 0, "PRE failed: (COL($rt,""Result"")>0)", 270: BuildNdeSummary = -1: GoTo CleanExit
272       If Not (ForgeCol(ut, "Weld Seam No") > 0) Then ForgeLog "BuildNdeSummary", 0, "PRE failed: (COL($ut,""Weld Seam No"")>0)", 270: BuildNdeSummary = -1: GoTo CleanExit
273       If Not (ForgeCol(ut, "Result") > 0) Then ForgeLog "BuildNdeSummary", 0, "PRE failed: (COL($ut,""Result"")>0)", 270: BuildNdeSummary = -1: GoTo CleanExit
274       If Not (ForgeCol(mpt, "Weld Seam No") > 0) Then ForgeLog "BuildNdeSummary", 0, "PRE failed: (COL($mpt,""Weld Seam No"")>0)", 270: BuildNdeSummary = -1: GoTo CleanExit
275       If Not (ForgeCol(mpt, "Result") > 0) Then ForgeLog "BuildNdeSummary", 0, "PRE failed: (COL($mpt,""Result"")>0)", 270: BuildNdeSummary = -1: GoTo CleanExit
276       If Not (ForgeCol(dpt, "Weld Seam No") > 0) Then ForgeLog "BuildNdeSummary", 0, "PRE failed: (COL($dpt,""Weld Seam No"")>0)", 270: BuildNdeSummary = -1: GoTo CleanExit
277       If Not (ForgeCol(dpt, "Result") > 0) Then ForgeLog "BuildNdeSummary", 0, "PRE failed: (COL($dpt,""Result"")>0)", 270: BuildNdeSummary = -1: GoTo CleanExit
          Application.ScreenUpdating = False
          Application.Calculation = xlCalculationManual
290       cRtSeam = ForgeCol(rt, "Weld Seam No")
300       cRtRes = ForgeCol(rt, "Result")
310       cUtSeam = ForgeCol(ut, "Weld Seam No")
320       cUtRes = ForgeCol(ut, "Result")
330       cMptSeam = ForgeCol(mpt, "Weld Seam No")
340       cMptRes = ForgeCol(mpt, "Result")
350       cDptSeam = ForgeCol(dpt, "Weld Seam No")
360       cDptRes = ForgeCol(dpt, "Result")
370       ForgeUsedRange(out).Clear
380       ForgeWriteRow out, 1, 1, Array("Weld Seam No", "RT", "UT", "MPT", "DPT")
390       For r = 2 To ForgeLastRow(rt, cRtSeam)
400           If ForgeIsEmpty(rt.Cells(r, cRtSeam).Value) Then GoTo Next_39
410           ForgeDictAdd seams, ForgeTrim(rt.Cells(r, cRtSeam).Value), 1, "keep"
420           ForgeDictAdd rtResults, ForgeTrim(rt.Cells(r, cRtSeam).Value), rt.Cells(r, cRtRes).Value, "overwrite"
Next_39:
          Next r
430       For r = 2 To ForgeLastRow(ut, cUtSeam)
440           If ForgeIsEmpty(ut.Cells(r, cUtSeam).Value) Then GoTo Next_43
450           ForgeDictAdd seams, ForgeTrim(ut.Cells(r, cUtSeam).Value), 1, "keep"
460           ForgeDictAdd utResults, ForgeTrim(ut.Cells(r, cUtSeam).Value), ut.Cells(r, cUtRes).Value, "overwrite"
Next_43:
          Next r
470       For r = 2 To ForgeLastRow(mpt, cMptSeam)
480           If ForgeIsEmpty(mpt.Cells(r, cMptSeam).Value) Then GoTo Next_47
490           ForgeDictAdd seams, ForgeTrim(mpt.Cells(r, cMptSeam).Value), 1, "keep"
500           ForgeDictAdd mptResults, ForgeTrim(mpt.Cells(r, cMptSeam).Value), mpt.Cells(r, cMptRes).Value, "overwrite"
Next_47:
          Next r
510       For r = 2 To ForgeLastRow(dpt, cDptSeam)
520           If ForgeIsEmpty(dpt.Cells(r, cDptSeam).Value) Then GoTo Next_51
530           ForgeDictAdd seams, ForgeTrim(dpt.Cells(r, cDptSeam).Value), 1, "keep"
540           ForgeDictAdd dptResults, ForgeTrim(dpt.Cells(r, cDptSeam).Value), dpt.Cells(r, cDptRes).Value, "overwrite"
Next_51:
          Next r
550       outRow = 2
560       For Each s In seams.Keys
570           ForgeSetValue out.Cells(outRow, 1), s
580           v = ""
590           If ForgeDictExists(rtResults, s) Then v = ForgeDictGet(rtResults, s)
600           If ForgeIsEmpty(v) Then v = "Pending"
610           ForgeSetValue out.Cells(outRow, 2), v
620           v = ""
630           If ForgeDictExists(utResults, s) Then v = ForgeDictGet(utResults, s)
640           If ForgeIsEmpty(v) Then v = "Pending"
650           ForgeSetValue out.Cells(outRow, 3), v
660           v = ""
670           If ForgeDictExists(mptResults, s) Then v = ForgeDictGet(mptResults, s)
680           If ForgeIsEmpty(v) Then v = "Pending"
690           ForgeSetValue out.Cells(outRow, 4), v
700           v = ""
710           If ForgeDictExists(dptResults, s) Then v = ForgeDictGet(dptResults, s)
720           If ForgeIsEmpty(v) Then v = "Pending"
730           ForgeSetValue out.Cells(outRow, 5), v
740           outRow = (outRow + 1)
          Next s
750       ForgeSort ForgeUsedRange(out), 1, True, True
760       BuildNdeSummary = (outRow - 2)
761       GoTo CleanExit
CleanExit:
          ForgeRestore forgeScreen, forgeCalc, forgeEvents
          Exit Function
ErrHandler:
          ForgeLog "BuildNdeSummary", Err.Number, Err.Description, Erl
          BuildNdeSummary = -2
          Resume CleanExit
End Function
```

### task_05.vcl: ListWeldLogExceptions

```vcl
VCL 0.2
TASK
# ASSUME: Weld Log headers (row 1) include "Weld Seam No", "Fit-up" and "Visual"; column order does not matter
# ASSUME: rows are NOT deleted; seams not yet accepted are copied to the Weld Log Exceptions sheet instead
# ASSUME: a seam is an exception when Fit-up or Visual is blank or anything other than "Accepted" (case and spaces ignored)
# ASSUME: the Weld Log Exceptions sheet is rebuilt every run with the Weld Log's header row
PROC     ListWeldLogExceptions
KIND     Function
MODULE   modWelding
ARGS     weldLog As Worksheet, out As Worksheet
RET      Long
VAR      cSeam As Long
VAR      cFit As Long
VAR      cVisual As Long
VAR      lastCol As Long
VAR      outRow As Long
PRE      COL($weldLog, "Weld Seam No") > 0; COL($weldLog, "Fit-up") > 0; COL($weldLog, "Visual") > 0
MODIFIES $out
OP       LET(cSeam, COL($weldLog, "Weld Seam No"))
OP       LET(cFit, COL($weldLog, "Fit-up"))
OP       LET(cVisual, COL($weldLog, "Visual"))
OP       LET(lastCol, LASTCOL($weldLog, 1))
OP       CLEAR(USED_RANGE($out), what=all)
OP       COPY_ROW($weldLog, 1, $out, 1, cols=lastCol)
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($weldLog, cSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($weldLog, r, cSeam))
OP         IF NOT SAME(CELL($weldLog, r, cFit), "ACCEPTED") OR NOT SAME(CELL($weldLog, r, cVisual), "ACCEPTED")
OP           COPY_ROW($weldLog, r, $out, outRow, cols=lastCol)
OP           LET(outRow, outRow + 1)
OP       RETURN(outRow - 2)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  weld_log_01.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Welder ID,C:Fit-up,D:Visual,E:Weld Date rows=50 data=A:seq(1,1)+boundary(1,"")+blank(2);B:values("W-01"x25,"W-02"x25);C:values("Accepted"x44,"Rejected"x2,""x4)+boundary("Rejected","Rejected");D:values("Accepted"x43,"Hold"x3,""x4);E:dates(2026-01-01,2026-09-30)+boundary("")
FIXTURE  weld_log_01.xlsx sheet=Exceptions cols=A:Weld Seam No rows=0 stale=10
TEST     fixture=weld_log_01.xlsx call=ListWeldLogExceptions(WeldLog, Exceptions) runs=2 assert=RESULT = 12 AND COUNTA(Exceptions![Weld Seam No]) = 13 AND COUNTIF(Exceptions![Weld Seam No], "stale") = 0 AND COUNTA(Exceptions![Weld Date]) = 12 AND ERROR_COUNT = 0
TEST     fixture=weld_log_01.xlsx call=ListWeldLogExceptions(WeldLog, Exceptions) assert=COUNTA(WeldLog![Weld Seam No]) = PRE_COUNTA(WeldLog![Weld Seam No]) AND COUNTA(WeldLog![Fit-up]) = PRE_COUNTA(WeldLog![Fit-up])
TEST     fixture=weld_log_01.xlsx call=ListWeldLogExceptions(Exceptions, Exceptions) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(Exceptions![Weld Seam No]) = PRE_COUNTA(Exceptions![Weld Seam No])
```

```vba
'@VCL PROC=ListWeldLogExceptions  source: task_05.vcl
Public Function ListWeldLogExceptions(ByVal weldLog As Worksheet, ByVal out As Worksheet) As Long
          Dim cSeam As Long
          Dim cFit As Long
          Dim cVisual As Long
          Dim lastCol As Long
          Dim outRow As Long
          Dim r As Long
          Dim forgeScreen As Boolean
          Dim forgeCalc As Long
          Dim forgeEvents As Boolean
          forgeScreen = Application.ScreenUpdating
          forgeCalc = Application.Calculation
          forgeEvents = Application.EnableEvents
          On Error GoTo ErrHandler
170       If Not (ForgeCol(weldLog, "Weld Seam No") > 0) Then ForgeLog "ListWeldLogExceptions", 0, "PRE failed: (COL($weldLog,""Weld Seam No"")>0)", 170: ListWeldLogExceptions = -1: GoTo CleanExit
171       If Not (ForgeCol(weldLog, "Fit-up") > 0) Then ForgeLog "ListWeldLogExceptions", 0, "PRE failed: (COL($weldLog,""Fit-up"")>0)", 170: ListWeldLogExceptions = -1: GoTo CleanExit
172       If Not (ForgeCol(weldLog, "Visual") > 0) Then ForgeLog "ListWeldLogExceptions", 0, "PRE failed: (COL($weldLog,""Visual"")>0)", 170: ListWeldLogExceptions = -1: GoTo CleanExit
          Application.ScreenUpdating = False
          Application.Calculation = xlCalculationManual
190       cSeam = ForgeCol(weldLog, "Weld Seam No")
200       cFit = ForgeCol(weldLog, "Fit-up")
210       cVisual = ForgeCol(weldLog, "Visual")
220       lastCol = ForgeLastCol(weldLog, 1)
230       ForgeUsedRange(out).Clear
240       ForgeCopyRow weldLog, 1, out, 1, lastCol
250       outRow = 2
260       For r = 2 To ForgeLastRow(weldLog, cSeam)
270           If ForgeIsEmpty(weldLog.Cells(r, cSeam).Value) Then GoTo Next_26
280           If ((Not ForgeSame(weldLog.Cells(r, cFit).Value, "ACCEPTED")) Or (Not ForgeSame(weldLog.Cells(r, cVisual).Value, "ACCEPTED"))) Then
290               ForgeCopyRow weldLog, r, out, outRow, lastCol
300               outRow = (outRow + 1)
              End If
Next_26:
          Next r
310       ListWeldLogExceptions = (outRow - 2)
311       GoTo CleanExit
CleanExit:
          ForgeRestore forgeScreen, forgeCalc, forgeEvents
          Exit Function
ErrHandler:
          ForgeLog "ListWeldLogExceptions", Err.Number, Err.Description, Erl
          ListWeldLogExceptions = -2
          Resume CleanExit
End Function
```

### task_06.vcl: FlagExpiredWelderQualifications

```vcl
VCL 0.2
TASK
# ASSUME: Welder Qualification headers (row 1) include "Welder ID" and "Expiry Date"; one row per welder
# ASSUME: Weld Log headers include "Weld Seam No" and "Welder ID"; the Weld Log row is the one highlighted
# ASSUME: a Weld Log row is red when its Welder ID is blank, not in the qualification list, has no valid expiry date, or the expiry date is before today
# ASSUME: only the expiry date is checked, not whether the welder is qualified for that specific WPS
# ASSUME: the macro owns the fill colour of Weld Log rows: it clears the fill first, so fixed rows lose their red
PROC     FlagExpiredWelderQualifications
KIND     Function
MODULE   modWelding
ARGS     weldLog As Worksheet, qual As Worksheet
RET      Long
VAR      expiry As Dictionary
VAR      cSeam As Long
VAR      cWelder As Long
VAR      cQualId As Long
VAR      cExpiry As Long
VAR      lastCol As Long
VAR      flagged As Long
VAR      welderId As String
VAR      expDate As Variant
VAR      bad As Boolean
PRE      COL($weldLog, "Weld Seam No") > 0; COL($weldLog, "Welder ID") > 0; COL($qual, "Welder ID") > 0; COL($qual, "Expiry Date") > 0
MODIFIES $weldLog fills
OP       LET(cSeam, COL($weldLog, "Weld Seam No"))
OP       LET(cWelder, COL($weldLog, "Welder ID"))
OP       LET(cQualId, COL($qual, "Welder ID"))
OP       LET(cExpiry, COL($qual, "Expiry Date"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($qual, cQualId) var=r
OP         IF NOT IS_EMPTY(CELL($qual, r, cQualId)) THEN DICT_ADD(expiry, TRIM(CELL($qual, r, cQualId)), CELL($qual, r, cExpiry), mode=overwrite)
OP       LET(lastCol, LASTCOL($weldLog, 1))
OP       LET(flagged, 0)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($weldLog, cSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($weldLog, r, cSeam))
OP         FORMAT(ROW_RANGE($weldLog, r, 1, lastCol), fill=none)
OP         LET(bad, FALSE)
OP         LET(expDate, "")
OP         LET(welderId, TRIM(CELL($weldLog, r, cWelder)))
OP         IF IS_EMPTY(welderId) THEN LET(bad, TRUE)
OP         IF EXISTS(expiry, welderId) THEN DICT_GET(expiry, welderId, into=expDate)
OP         IF NOT IS_DATE(expDate) THEN LET(bad, TRUE)
OP         IF IS_DATE(expDate)
OP           IF expDate < TODAY THEN LET(bad, TRUE)
OP         IF bad
OP           FORMAT(ROW_RANGE($weldLog, r, 1, lastCol), fill=#FF0000)
OP           LET(flagged, flagged + 1)
OP       RETURN(flagged)
PERF     screen_off
ERR      standard
FIXTURE  welder_qual_01.xlsx sheet=WelderQual cols=A:Welder ID,B:Name,C:Expiry Date rows=7 data=A:values("W-01"x2,"W-02"x1,"W-03"x1,"W-04"x1,"W-05"x1,"W-06"x1)+boundary("W-02","W-01","W-03","W-01","W-04");B:text(10);C:reldates(-400,800)+boundary(0,-1,1,30,"")
FIXTURE  welder_qual_01.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Welder ID,C:Fit-up,D:Visual rows=40 stale=5 data=A:seq(1,1)+blank(2);B:values("W-01"x8,"W-02"x8,"W-03"x8,"W-04"x8,"W-07"x6,""x2)+boundary("W-07");C:values("Accepted"x40);D:values("Accepted"x39,""x1)+boundary("")
TEST     fixture=welder_qual_01.xlsx call=FlagExpiredWelderQualifications(WeldLog, WelderQual) runs=2 assert=RESULT = 15 AND FILL_COUNT(WeldLog![Weld Seam No], #FF0000) = 15 AND FILL_COUNT(WeldLog![Visual], #FF0000) = 15 AND FILL_COUNT(WeldLog![Welder ID]2, #FF0000) = 1 AND ERROR_COUNT = 0
TEST     fixture=welder_qual_01.xlsx call=FlagExpiredWelderQualifications(WeldLog, WelderQual) assert=COUNTA(WeldLog![Welder ID]) = PRE_COUNTA(WeldLog![Welder ID]) AND COUNTA(WelderQual![Expiry Date]) = PRE_COUNTA(WelderQual![Expiry Date])
TEST     fixture=welder_qual_01.xlsx call=FlagExpiredWelderQualifications(WelderQual, WelderQual) assert=RESULT = -1 AND ERROR_COUNT = 1 AND FILL_COUNT(WelderQual![Welder ID], #FF0000) = 0
```

```vba
'@VCL PROC=FlagExpiredWelderQualifications  source: task_06.vcl
Public Function FlagExpiredWelderQualifications(ByVal weldLog As Worksheet, ByVal qual As Worksheet) As Long
          Dim expiry As Object
          Dim cSeam As Long
          Dim cWelder As Long
          Dim cQualId As Long
          Dim cExpiry As Long
          Dim lastCol As Long
          Dim flagged As Long
          Dim welderId As String
          Dim expDate As Variant
          Dim bad As Boolean
          Dim r As Long
          Dim forgeScreen As Boolean
          Dim forgeCalc As Long
          Dim forgeEvents As Boolean
          forgeScreen = Application.ScreenUpdating
          forgeCalc = Application.Calculation
          forgeEvents = Application.EnableEvents
          On Error GoTo ErrHandler
          Set expiry = ForgeNewDict()
230       If Not (ForgeCol(weldLog, "Weld Seam No") > 0) Then ForgeLog "FlagExpiredWelderQualifications", 0, "PRE failed: (COL($weldLog,""Weld Seam No"")>0)", 230: FlagExpiredWelderQualifications = -1: GoTo CleanExit
231       If Not (ForgeCol(weldLog, "Welder ID") > 0) Then ForgeLog "FlagExpiredWelderQualifications", 0, "PRE failed: (COL($weldLog,""Welder ID"")>0)", 230: FlagExpiredWelderQualifications = -1: GoTo CleanExit
232       If Not (ForgeCol(qual, "Welder ID") > 0) Then ForgeLog "FlagExpiredWelderQualifications", 0, "PRE failed: (COL($qual,""Welder ID"")>0)", 230: FlagExpiredWelderQualifications = -1: GoTo CleanExit
233       If Not (ForgeCol(qual, "Expiry Date") > 0) Then ForgeLog "FlagExpiredWelderQualifications", 0, "PRE failed: (COL($qual,""Expiry Date"")>0)", 230: FlagExpiredWelderQualifications = -1: GoTo CleanExit
          Application.ScreenUpdating = False
250       cSeam = ForgeCol(weldLog, "Weld Seam No")
260       cWelder = ForgeCol(weldLog, "Welder ID")
270       cQualId = ForgeCol(qual, "Welder ID")
280       cExpiry = ForgeCol(qual, "Expiry Date")
290       For r = 2 To ForgeLastRow(qual, cQualId)
300           If (Not ForgeIsEmpty(qual.Cells(r, cQualId).Value)) Then ForgeDictAdd expiry, ForgeTrim(qual.Cells(r, cQualId).Value), qual.Cells(r, cExpiry).Value, "overwrite"
          Next r
310       lastCol = ForgeLastCol(weldLog, 1)
320       flagged = 0
330       For r = 2 To ForgeLastRow(weldLog, cSeam)
340           If ForgeIsEmpty(weldLog.Cells(r, cSeam).Value) Then GoTo Next_33
350           ForgeNoFill ForgeRowRange(weldLog, r, 1, lastCol)
360           bad = False
370           expDate = ""
380           welderId = ForgeTrim(weldLog.Cells(r, cWelder).Value)
390           If ForgeIsEmpty(welderId) Then bad = True
400           If ForgeDictExists(expiry, welderId) Then expDate = ForgeDictGet(expiry, welderId)
410           If (Not ForgeIsDate(expDate)) Then bad = True
420           If ForgeIsDate(expDate) Then
430               If (expDate < Date) Then bad = True
              End If
440           If bad Then
450               ForgeRowRange(weldLog, r, 1, lastCol).Interior.Color = RGB(255, 0, 0)
460               flagged = (flagged + 1)
              End If
Next_33:
          Next r
470       FlagExpiredWelderQualifications = flagged
471       GoTo CleanExit
CleanExit:
          ForgeRestore forgeScreen, forgeCalc, forgeEvents
          Exit Function
ErrHandler:
          ForgeLog "FlagExpiredWelderQualifications", Err.Number, Err.Description, Erl
          FlagExpiredWelderQualifications = -2
          Resume CleanExit
End Function
```

### task_07.vcl: BuildCallNoticeList

```vcl
VCL 0.2
TASK
# ASSUME: ITP Tracker headers (row 1) include "Point Type" (Hold / Witness / Review) and "Inspection Status"
# ASSUME: a row is copied when Inspection Status is "Ready" and Point Type is "Hold" or "Witness" (case and spaces ignored)
# ASSUME: rows with a blank Inspection Status are ignored
# ASSUME: the Call Notice sheet is rebuilt every run with the ITP Tracker's header row
PROC     BuildCallNoticeList
KIND     Function
MODULE   modQuality
ARGS     itp As Worksheet, out As Worksheet
RET      Long
VAR      cType As Long
VAR      cStatus As Long
VAR      lastCol As Long
VAR      outRow As Long
VAR      pointType As String
VAR      status As String
PRE      COL($itp, "Point Type") > 0; COL($itp, "Inspection Status") > 0
MODIFIES $out
OP       LET(cType, COL($itp, "Point Type"))
OP       LET(cStatus, COL($itp, "Inspection Status"))
OP       LET(lastCol, LASTCOL($itp, 1))
OP       CLEAR(USED_RANGE($out), what=all)
OP       COPY_ROW($itp, 1, $out, 1, cols=lastCol)
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($itp, cStatus) var=r
OP         LET(status, UPPER(TRIM(CELL($itp, r, cStatus))))
OP         LET(pointType, UPPER(TRIM(CELL($itp, r, cType))))
OP         IF status = "READY"
OP           IF pointType = "HOLD" OR pointType = "WITNESS"
OP             COPY_ROW($itp, r, $out, outRow, cols=lastCol)
OP             LET(outRow, outRow + 1)
OP       RETURN(outRow - 2)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  itp_01.xlsx sheet=ITP cols=A:Activity,B:Point Type,C:Inspection Status,D:Planned Date rows=40 data=A:text(15);B:values("Hold"x12,"Witness"x14,"Review"x10,""x4)+boundary("Hold");C:values("Ready"x10,"Done"x18,"Pending"x8,""x4)+boundary("Ready");D:dates(2026-10-01,2026-12-31)+boundary("")
FIXTURE  itp_01.xlsx sheet=CallNotice cols=A:Activity rows=0 stale=10
TEST     fixture=itp_01.xlsx call=BuildCallNoticeList(ITP, CallNotice) runs=2 assert=RESULT = 8 AND RESULT = PRE_COUNTIFS(ITP![Inspection Status], "Ready", ITP![Point Type], "Hold") + PRE_COUNTIFS(ITP![Inspection Status], "Ready", ITP![Point Type], "Witness") AND COUNTA(CallNotice![Activity]) = 9 AND COUNTIF(CallNotice![Activity], "stale") = 0 AND COUNTA(CallNotice![Planned Date]) = 8 AND ERROR_COUNT = 0
TEST     fixture=itp_01.xlsx call=BuildCallNoticeList(ITP, CallNotice) assert=COUNTIF(CallNotice![Inspection Status], "Ready") = RESULT AND COUNTIF(CallNotice![Point Type], "Review") = 0
TEST     fixture=itp_01.xlsx call=BuildCallNoticeList(ITP, CallNotice) assert=COUNTA(ITP![Inspection Status]) = PRE_COUNTA(ITP![Inspection Status]) AND COUNTA(ITP![Activity]) = PRE_COUNTA(ITP![Activity])
TEST     fixture=itp_01.xlsx call=BuildCallNoticeList(CallNotice, CallNotice) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(CallNotice![Activity]) = PRE_COUNTA(CallNotice![Activity])
```

```vba
'@VCL PROC=BuildCallNoticeList  source: task_07.vcl
Public Function BuildCallNoticeList(ByVal itp As Worksheet, ByVal out As Worksheet) As Long
          Dim cType As Long
          Dim cStatus As Long
          Dim lastCol As Long
          Dim outRow As Long
          Dim pointType As String
          Dim status As String
          Dim r As Long
          Dim forgeScreen As Boolean
          Dim forgeCalc As Long
          Dim forgeEvents As Boolean
          forgeScreen = Application.ScreenUpdating
          forgeCalc = Application.Calculation
          forgeEvents = Application.EnableEvents
          On Error GoTo ErrHandler
180       If Not (ForgeCol(itp, "Point Type") > 0) Then ForgeLog "BuildCallNoticeList", 0, "PRE failed: (COL($itp,""Point Type"")>0)", 180: BuildCallNoticeList = -1: GoTo CleanExit
181       If Not (ForgeCol(itp, "Inspection Status") > 0) Then ForgeLog "BuildCallNoticeList", 0, "PRE failed: (COL($itp,""Inspection Status"")>0)", 180: BuildCallNoticeList = -1: GoTo CleanExit
          Application.ScreenUpdating = False
          Application.Calculation = xlCalculationManual
200       cType = ForgeCol(itp, "Point Type")
210       cStatus = ForgeCol(itp, "Inspection Status")
220       lastCol = ForgeLastCol(itp, 1)
230       ForgeUsedRange(out).Clear
240       ForgeCopyRow itp, 1, out, 1, lastCol
250       outRow = 2
260       For r = 2 To ForgeLastRow(itp, cStatus)
270           status = UCase$(ForgeText(ForgeTrim(itp.Cells(r, cStatus).Value)))
280           pointType = UCase$(ForgeText(ForgeTrim(itp.Cells(r, cType).Value)))
290           If (status = "READY") Then
300               If ((pointType = "HOLD") Or (pointType = "WITNESS")) Then
310                   ForgeCopyRow itp, r, out, outRow, lastCol
320                   outRow = (outRow + 1)
                  End If
              End If
          Next r
330       BuildCallNoticeList = (outRow - 2)
331       GoTo CleanExit
CleanExit:
          ForgeRestore forgeScreen, forgeCalc, forgeEvents
          Exit Function
ErrHandler:
          ForgeLog "BuildCallNoticeList", Err.Number, Err.Description, Erl
          BuildCallNoticeList = -2
          Resume CleanExit
End Function
```

### task_08.vcl: ReconcileConsumableStock

```vcl
VCL 0.2
TASK
# ASSUME: Consumable Issue headers (row 1) include "Batch No" and "Qty Issued"
# ASSUME: Inventory headers include "Batch No", "Opening Stock", "Issued to Date", "Balance" and "Remark"
# ASSUME: Issued to Date, Balance and Remark are owned by this macro and recalculated from the whole issue log on every run, so running it twice never double-counts
# ASSUME: issue rows with a blank or non-numeric Qty Issued, or a blank Batch No, are ignored
# ASSUME: the function returns how many issue rows use a batch number that is not in Inventory
PROC     ReconcileConsumableStock
KIND     Function
MODULE   modWelding
ARGS     issues As Worksheet, inv As Worksheet
RET      Long
VAR      known As Dictionary
VAR      issued As Dictionary
VAR      iBatch As Long
VAR      iQty As Long
VAR      vBatch As Long
VAR      vOpen As Long
VAR      vIssued As Long
VAR      vBal As Long
VAR      vRemark As Long
VAR      batch As String
VAR      used As Double
VAR      unknown As Long
PRE      COL($issues, "Batch No") > 0; COL($issues, "Qty Issued") > 0; COL($inv, "Batch No") > 0; COL($inv, "Opening Stock") > 0; COL($inv, "Issued to Date") > 0; COL($inv, "Balance") > 0; COL($inv, "Remark") > 0
MODIFIES $inv["Issued to Date", "Remark", "Balance"]
OP       LET(iBatch, COL($issues, "Batch No"))
OP       LET(iQty, COL($issues, "Qty Issued"))
OP       LET(vBatch, COL($inv, "Batch No"))
OP       LET(vOpen, COL($inv, "Opening Stock"))
OP       LET(vIssued, COL($inv, "Issued to Date"))
OP       LET(vBal, COL($inv, "Balance"))
OP       LET(vRemark, COL($inv, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($inv, vBatch) var=r
OP         IF NOT IS_EMPTY(CELL($inv, r, vBatch)) THEN DICT_ADD(known, TRIM(CELL($inv, r, vBatch)), r, mode=keep)
OP       LET(unknown, 0)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($issues, iBatch) var=r
OP         LET(batch, TRIM(CELL($issues, r, iBatch)))
OP         IF NOT IS_EMPTY(batch) AND IS_NUMBER(CELL($issues, r, iQty))
OP           IF EXISTS(known, batch) THEN DICT_ADD(issued, batch, CELL($issues, r, iQty), mode=sum)
OP           IF NOT EXISTS(known, batch) THEN LET(unknown, unknown + 1)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($inv, vBatch) var=r
OP         SKIP_IF IS_EMPTY(CELL($inv, r, vBatch))
OP         LET(batch, TRIM(CELL($inv, r, vBatch)))
OP         LET(used, 0)
OP         IF EXISTS(issued, batch) THEN DICT_GET(issued, batch, into=used)
OP         SET_VALUE(CELL($inv, r, vIssued), used)
OP         CLEAR(CELL($inv, r, vRemark), what=contents)
OP         IF IS_NUMBER(CELL($inv, r, vOpen))
OP           SET_VALUE(CELL($inv, r, vBal), CELL($inv, r, vOpen) - used)
OP           IF CELL($inv, r, vOpen) - used < 0 THEN SET_VALUE(CELL($inv, r, vRemark), "Negative balance")
OP         ELSE
OP           CLEAR(CELL($inv, r, vBal), what=contents)
OP           SET_VALUE(CELL($inv, r, vRemark), "Opening stock missing")
OP       RETURN(unknown)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  consumables_01.xlsx sheet=Inventory cols=A:Batch No,B:Opening Stock,C:Issued to Date,D:Balance,E:Remark rows=9 data=A:values("B-101"x1,"B-102"x1,"B-103"x1,"B-104"x1,"B-105"x1,"B-106"x1,"B-107"x1,"B-108"x1,""x1)+boundary("B-101","B-106");B:int(200,500)+boundary(1,0)+blank(1);D:int(1,99);E:values("Old remark"x5,""x4)
FIXTURE  consumables_01.xlsx sheet=Issues cols=A:Issue Date,B:Batch No,C:Qty Issued rows=60 data=A:dates(2026-06-01,2026-09-30);B:values("B-101"x10,"B-102"x10,"B-103"x10,"B-104"x10,"B-105"x10,"B-109"x5,""x5);C:int(1,20)+blank(6)
TEST     fixture=consumables_01.xlsx call=ReconcileConsumableStock(Issues, Inventory) runs=2 assert=RESULT = PRE_COUNTIFS(Issues![Batch No], "B-109", Issues![Qty Issued], "<>") AND SUM(Inventory![Issued to Date]) = PRE_SUMIFS(Issues![Qty Issued], Issues![Batch No], "<>B-109", Issues![Batch No], "<>") AND ERROR_COUNT = 0
TEST     fixture=consumables_01.xlsx call=ReconcileConsumableStock(Issues, Inventory) assert=COUNTIF(Inventory![Remark], "Opening stock missing") = 1 AND COUNTIF(Inventory![Remark], "Negative balance") = 1 AND COUNTIF(Inventory![Remark], "Old remark") = 1 AND VALUE(Inventory![Balance]3) = 0 AND COUNTA(Inventory![Balance]) = 9
TEST     fixture=consumables_01.xlsx call=ReconcileConsumableStock(Issues, Inventory) assert=SUM(Inventory![Opening Stock]) = PRE_SUM(Inventory![Opening Stock]) AND SUM(Issues![Qty Issued]) = PRE_SUM(Issues![Qty Issued])
TEST     fixture=consumables_01.xlsx call=ReconcileConsumableStock(Inventory, Inventory) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(Inventory![Balance]) = PRE_COUNTA(Inventory![Balance])
```

```vba
'@VCL PROC=ReconcileConsumableStock  source: task_08.vcl
Public Function ReconcileConsumableStock(ByVal issues As Worksheet, ByVal inv As Worksheet) As Long
          Dim known As Object
          Dim issued As Object
          Dim iBatch As Long
          Dim iQty As Long
          Dim vBatch As Long
          Dim vOpen As Long
          Dim vIssued As Long
          Dim vBal As Long
          Dim vRemark As Long
          Dim batch As String
          Dim used As Double
          Dim unknown As Long
          Dim r As Long
          Dim forgeScreen As Boolean
          Dim forgeCalc As Long
          Dim forgeEvents As Boolean
          forgeScreen = Application.ScreenUpdating
          forgeCalc = Application.Calculation
          forgeEvents = Application.EnableEvents
          On Error GoTo ErrHandler
          Set known = ForgeNewDict()
          Set issued = ForgeNewDict()
250       If Not (ForgeCol(issues, "Batch No") > 0) Then ForgeLog "ReconcileConsumableStock", 0, "PRE failed: (COL($issues,""Batch No"")>0)", 250: ReconcileConsumableStock = -1: GoTo CleanExit
251       If Not (ForgeCol(issues, "Qty Issued") > 0) Then ForgeLog "ReconcileConsumableStock", 0, "PRE failed: (COL($issues,""Qty Issued"")>0)", 250: ReconcileConsumableStock = -1: GoTo CleanExit
252       If Not (ForgeCol(inv, "Batch No") > 0) Then ForgeLog "ReconcileConsumableStock", 0, "PRE failed: (COL($inv,""Batch No"")>0)", 250: ReconcileConsumableStock = -1: GoTo CleanExit
253       If Not (ForgeCol(inv, "Opening Stock") > 0) Then ForgeLog "ReconcileConsumableStock", 0, "PRE failed: (COL($inv,""Opening Stock"")>0)", 250: ReconcileConsumableStock = -1: GoTo CleanExit
254       If Not (ForgeCol(inv, "Issued to Date") > 0) Then ForgeLog "ReconcileConsumableStock", 0, "PRE failed: (COL($inv,""Issued to Date"")>0)", 250: ReconcileConsumableStock = -1: GoTo CleanExit
255       If Not (ForgeCol(inv, "Balance") > 0) Then ForgeLog "ReconcileConsumableStock", 0, "PRE failed: (COL($inv,""Balance"")>0)", 250: ReconcileConsumableStock = -1: GoTo CleanExit
256       If Not (ForgeCol(inv, "Remark") > 0) Then ForgeLog "ReconcileConsumableStock", 0, "PRE failed: (COL($inv,""Remark"")>0)", 250: ReconcileConsumableStock = -1: GoTo CleanExit
          Application.ScreenUpdating = False
          Application.Calculation = xlCalculationManual
270       iBatch = ForgeCol(issues, "Batch No")
280       iQty = ForgeCol(issues, "Qty Issued")
290       vBatch = ForgeCol(inv, "Batch No")
300       vOpen = ForgeCol(inv, "Opening Stock")
310       vIssued = ForgeCol(inv, "Issued to Date")
320       vBal = ForgeCol(inv, "Balance")
330       vRemark = ForgeCol(inv, "Remark")
340       For r = 2 To ForgeLastRow(inv, vBatch)
350           If (Not ForgeIsEmpty(inv.Cells(r, vBatch).Value)) Then ForgeDictAdd known, ForgeTrim(inv.Cells(r, vBatch).Value), r, "keep"
          Next r
360       unknown = 0
370       For r = 2 To ForgeLastRow(issues, iBatch)
380           batch = ForgeTrim(issues.Cells(r, iBatch).Value)
390           If ((Not ForgeIsEmpty(batch)) And ForgeIsNumber(issues.Cells(r, iQty).Value)) Then
400               If ForgeDictExists(known, batch) Then ForgeDictAdd issued, batch, issues.Cells(r, iQty).Value, "sum"
410               If (Not ForgeDictExists(known, batch)) Then unknown = (unknown + 1)
              End If
          Next r
420       For r = 2 To ForgeLastRow(inv, vBatch)
430           If ForgeIsEmpty(inv.Cells(r, vBatch).Value) Then GoTo Next_42
440           batch = ForgeTrim(inv.Cells(r, vBatch).Value)
450           used = 0
460           If ForgeDictExists(issued, batch) Then used = ForgeNum(ForgeDictGet(issued, batch))
470           ForgeSetValue inv.Cells(r, vIssued), used
480           inv.Cells(r, vRemark).ClearContents
490           If ForgeIsNumber(inv.Cells(r, vOpen).Value) Then
500               ForgeSetValue inv.Cells(r, vBal), (ForgeNum(inv.Cells(r, vOpen).Value) - used)
510               If ((ForgeNum(inv.Cells(r, vOpen).Value) - used) < 0) Then ForgeSetValue inv.Cells(r, vRemark), "Negative balance"
              Else
530               inv.Cells(r, vBal).ClearContents
540               ForgeSetValue inv.Cells(r, vRemark), "Opening stock missing"
              End If
Next_42:
          Next r
550       ReconcileConsumableStock = unknown
551       GoTo CleanExit
CleanExit:
          ForgeRestore forgeScreen, forgeCalc, forgeEvents
          Exit Function
ErrHandler:
          ForgeLog "ReconcileConsumableStock", Err.Number, Err.Description, Erl
          ReconcileConsumableStock = -2
          Resume CleanExit
End Function
```

### task_09.vcl: BuildOpenNcrSummary

```vcl
VCL 0.2
TASK
# ASSUME: NCR Register headers (row 1) include "NCR No", "Issue Date" and "Status"
# ASSUME: a row is copied when Status is "Open" (case and spaces ignored); blank Status rows are ignored
# ASSUME: the Summary sheet is rebuilt every run with the register's header row plus a Days Open column, oldest first
# ASSUME: rows with a missing or invalid Issue Date are still copied, marked "Issue date missing", and sort to the bottom
PROC     BuildOpenNcrSummary
KIND     Function
MODULE   modQuality
ARGS     ncr As Worksheet, out As Worksheet
RET      Long
VAR      cNcr As Long
VAR      cDate As Long
VAR      cStatus As Long
VAR      lastCol As Long
VAR      outRow As Long
PRE      COL($ncr, "NCR No") > 0; COL($ncr, "Issue Date") > 0; COL($ncr, "Status") > 0
MODIFIES $out
OP       LET(cNcr, COL($ncr, "NCR No"))
OP       LET(cDate, COL($ncr, "Issue Date"))
OP       LET(cStatus, COL($ncr, "Status"))
OP       LET(lastCol, LASTCOL($ncr, 1))
OP       CLEAR(USED_RANGE($out), what=all)
OP       COPY_ROW($ncr, 1, $out, 1, cols=lastCol)
OP       SET_VALUE(CELL($out, 1, lastCol + 1), "Days Open")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ncr, cNcr) var=r
OP         IF SAME(CELL($ncr, r, cStatus), "OPEN")
OP           COPY_ROW($ncr, r, $out, outRow, cols=lastCol)
OP           IF IS_DATE(CELL($ncr, r, cDate)) THEN SET_VALUE(CELL($out, outRow, lastCol + 1), TODAY - CELL($ncr, r, cDate))
OP           IF NOT IS_DATE(CELL($ncr, r, cDate)) THEN SET_VALUE(CELL($out, outRow, lastCol + 1), "Issue date missing")
OP           LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=cDate, order=asc, header=yes)
OP       RETURN(outRow - 2)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  ncr_01.xlsx sheet=NCR cols=A:NCR No,B:Issue Date,C:Description,D:Owner,E:Status,F:Remarks rows=30 data=A:seq(1,1);B:reldates(-290,-1)+boundary(0,"")+blank(1);C:text(20);D:values("QA"x10,"Production"x10,"Stores"x10);E:values("Open"x12,"Closed"x15,""x3)+boundary("Open","Open");F:values("See report"x6,""x24)+boundary("")
FIXTURE  ncr_01.xlsx sheet=NcrSummary cols=A:NCR No rows=0 stale=10
TEST     fixture=ncr_01.xlsx call=BuildOpenNcrSummary(NCR, NcrSummary) runs=2 assert=RESULT = 12 AND COUNTA(NcrSummary![NCR No]) = 13 AND COUNTIF(NcrSummary![NCR No], "stale") = 0 AND COUNTIF(NcrSummary![Days Open], "Issue date missing") = 1 AND COUNTIF(NcrSummary![Days Open], 0) = 1 AND COUNTA(NcrSummary![Remarks]) = 3 AND ERROR_COUNT = 0
TEST     fixture=ncr_01.xlsx call=BuildOpenNcrSummary(NCR, NcrSummary) assert=COUNTIF(NcrSummary![Status], "Closed") = 0 AND VALUE(NcrSummary![Issue Date]2) <= VALUE(NcrSummary![Issue Date]3)
TEST     fixture=ncr_01.xlsx call=BuildOpenNcrSummary(NCR, NcrSummary) assert=COUNTA(NCR![Status]) = PRE_COUNTA(NCR![Status]) AND COUNTA(NCR![Issue Date]) = PRE_COUNTA(NCR![Issue Date])
TEST     fixture=ncr_01.xlsx call=BuildOpenNcrSummary(NcrSummary, NcrSummary) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(NcrSummary![NCR No]) = PRE_COUNTA(NcrSummary![NCR No])
```

```vba
'@VCL PROC=BuildOpenNcrSummary  source: task_09.vcl
Public Function BuildOpenNcrSummary(ByVal ncr As Worksheet, ByVal out As Worksheet) As Long
          Dim cNcr As Long
          Dim cDate As Long
          Dim cStatus As Long
          Dim lastCol As Long
          Dim outRow As Long
          Dim r As Long
          Dim forgeScreen As Boolean
          Dim forgeCalc As Long
          Dim forgeEvents As Boolean
          forgeScreen = Application.ScreenUpdating
          forgeCalc = Application.Calculation
          forgeEvents = Application.EnableEvents
          On Error GoTo ErrHandler
170       If Not (ForgeCol(ncr, "NCR No") > 0) Then ForgeLog "BuildOpenNcrSummary", 0, "PRE failed: (COL($ncr,""NCR No"")>0)", 170: BuildOpenNcrSummary = -1: GoTo CleanExit
171       If Not (ForgeCol(ncr, "Issue Date") > 0) Then ForgeLog "BuildOpenNcrSummary", 0, "PRE failed: (COL($ncr,""Issue Date"")>0)", 170: BuildOpenNcrSummary = -1: GoTo CleanExit
172       If Not (ForgeCol(ncr, "Status") > 0) Then ForgeLog "BuildOpenNcrSummary", 0, "PRE failed: (COL($ncr,""Status"")>0)", 170: BuildOpenNcrSummary = -1: GoTo CleanExit
          Application.ScreenUpdating = False
          Application.Calculation = xlCalculationManual
190       cNcr = ForgeCol(ncr, "NCR No")
200       cDate = ForgeCol(ncr, "Issue Date")
210       cStatus = ForgeCol(ncr, "Status")
220       lastCol = ForgeLastCol(ncr, 1)
230       ForgeUsedRange(out).Clear
240       ForgeCopyRow ncr, 1, out, 1, lastCol
250       ForgeSetValue out.Cells(1, (lastCol + 1)), "Days Open"
260       outRow = 2
270       For r = 2 To ForgeLastRow(ncr, cNcr)
280           If ForgeSame(ncr.Cells(r, cStatus).Value, "OPEN") Then
290               ForgeCopyRow ncr, r, out, outRow, lastCol
300               If ForgeIsDate(ncr.Cells(r, cDate).Value) Then ForgeSetValue out.Cells(outRow, (lastCol + 1)), (Date - ForgeNum(ncr.Cells(r, cDate).Value))
310               If (Not ForgeIsDate(ncr.Cells(r, cDate).Value)) Then ForgeSetValue out.Cells(outRow, (lastCol + 1)), "Issue date missing"
320               outRow = (outRow + 1)
              End If
          Next r
330       ForgeSort ForgeUsedRange(out), cDate, True, True
340       BuildOpenNcrSummary = (outRow - 2)
341       GoTo CleanExit
CleanExit:
          ForgeRestore forgeScreen, forgeCalc, forgeEvents
          Exit Function
ErrHandler:
          ForgeLog "BuildOpenNcrSummary", Err.Number, Err.Description, Erl
          BuildOpenNcrSummary = -2
          Resume CleanExit
End Function
```

### task_10.vcl: BuildCalibrationDueList

```vcl
VCL 0.2
TASK
# ASSUME: Calibration Tracker headers (row 1) include "Instrument ID" and "Next Calibration Due"
# ASSUME: an instrument is listed when its due date is blank or invalid, already passed, or within the next 15 days
# ASSUME: the Overdue sheet is rebuilt every run with the tracker's header row plus a Status column, soonest due first
# ASSUME: rows with a blank Instrument ID are skipped
PROC     BuildCalibrationDueList
KIND     Function
MODULE   modQuality
ARGS     cal As Worksheet, out As Worksheet
RET      Long
VAR      cId As Long
VAR      cDue As Long
VAR      lastCol As Long
VAR      outRow As Long
VAR      status As String
PRE      COL($cal, "Instrument ID") > 0; COL($cal, "Next Calibration Due") > 0
MODIFIES $out
OP       LET(cId, COL($cal, "Instrument ID"))
OP       LET(cDue, COL($cal, "Next Calibration Due"))
OP       LET(lastCol, LASTCOL($cal, 1))
OP       CLEAR(USED_RANGE($out), what=all)
OP       COPY_ROW($cal, 1, $out, 1, cols=lastCol)
OP       SET_VALUE(CELL($out, 1, lastCol + 1), "Status")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($cal, cId) var=r
OP         SKIP_IF IS_EMPTY(CELL($cal, r, cId))
OP         LET(status, "")
OP         IF NOT IS_DATE(CELL($cal, r, cDue)) THEN LET(status, "Due date missing")
OP         IF IS_DATE(CELL($cal, r, cDue))
OP           IF CELL($cal, r, cDue) < TODAY THEN LET(status, "Overdue")
OP           IF CELL($cal, r, cDue) >= TODAY AND CELL($cal, r, cDue) <= TODAY + 15 THEN LET(status, "Due within 15 days")
OP         IF NOT IS_EMPTY(status)
OP           COPY_ROW($cal, r, $out, outRow, cols=lastCol)
OP           SET_VALUE(CELL($out, outRow, lastCol + 1), status)
OP           LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=cDue, order=asc, header=yes)
OP       RETURN(outRow - 2)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  calibration_01.xlsx sheet=Calibration cols=A:Instrument ID,B:Description,C:Last Calibrated,D:Next Calibration Due rows=30 data=A:seq(1,1)+blank(2);B:text(15);C:dates(2025-01-01,2026-09-30);D:reldates(-500,990)+boundary("",0,15,16,-1)+blank(2)
FIXTURE  calibration_01.xlsx sheet=Overdue cols=A:Instrument ID rows=0 stale=10
TEST     fixture=calibration_01.xlsx call=BuildCalibrationDueList(Calibration, Overdue) runs=2 assert=RESULT = 11 AND COUNTA(Overdue![Instrument ID]) = 12 AND COUNTIF(Overdue![Instrument ID], "stale") = 0 AND COUNTIF(Overdue![Status], "Overdue") = 7 AND COUNTIF(Overdue![Status], "Due within 15 days") = 2 AND ERROR_COUNT = 0
TEST     fixture=calibration_01.xlsx call=BuildCalibrationDueList(Calibration, Overdue) assert=COUNTIF(Overdue![Status], "Due date missing") = 2
TEST     fixture=calibration_01.xlsx call=BuildCalibrationDueList(Calibration, Overdue) assert=COUNTA(Calibration![Next Calibration Due]) = PRE_COUNTA(Calibration![Next Calibration Due]) AND COUNTA(Calibration![Instrument ID]) = PRE_COUNTA(Calibration![Instrument ID])
TEST     fixture=calibration_01.xlsx call=BuildCalibrationDueList(Overdue, Overdue) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(Overdue![Instrument ID]) = PRE_COUNTA(Overdue![Instrument ID])
```

```vba
'@VCL PROC=BuildCalibrationDueList  source: task_10.vcl
Public Function BuildCalibrationDueList(ByVal cal As Worksheet, ByVal out As Worksheet) As Long
          Dim cId As Long
          Dim cDue As Long
          Dim lastCol As Long
          Dim outRow As Long
          Dim status As String
          Dim r As Long
          Dim forgeScreen As Boolean
          Dim forgeCalc As Long
          Dim forgeEvents As Boolean
          forgeScreen = Application.ScreenUpdating
          forgeCalc = Application.Calculation
          forgeEvents = Application.EnableEvents
          On Error GoTo ErrHandler
170       If Not (ForgeCol(cal, "Instrument ID") > 0) Then ForgeLog "BuildCalibrationDueList", 0, "PRE failed: (COL($cal,""Instrument ID"")>0)", 170: BuildCalibrationDueList = -1: GoTo CleanExit
171       If Not (ForgeCol(cal, "Next Calibration Due") > 0) Then ForgeLog "BuildCalibrationDueList", 0, "PRE failed: (COL($cal,""Next Calibration Due"")>0)", 170: BuildCalibrationDueList = -1: GoTo CleanExit
          Application.ScreenUpdating = False
          Application.Calculation = xlCalculationManual
190       cId = ForgeCol(cal, "Instrument ID")
200       cDue = ForgeCol(cal, "Next Calibration Due")
210       lastCol = ForgeLastCol(cal, 1)
220       ForgeUsedRange(out).Clear
230       ForgeCopyRow cal, 1, out, 1, lastCol
240       ForgeSetValue out.Cells(1, (lastCol + 1)), "Status"
250       outRow = 2
260       For r = 2 To ForgeLastRow(cal, cId)
270           If ForgeIsEmpty(cal.Cells(r, cId).Value) Then GoTo Next_26
280           status = ""
290           If (Not ForgeIsDate(cal.Cells(r, cDue).Value)) Then status = "Due date missing"
300           If ForgeIsDate(cal.Cells(r, cDue).Value) Then
310               If (ForgeNum(cal.Cells(r, cDue).Value) < Date) Then status = "Overdue"
320               If ((ForgeNum(cal.Cells(r, cDue).Value) >= Date) And (ForgeNum(cal.Cells(r, cDue).Value) <= (Date + 15))) Then status = "Due within 15 days"
              End If
330           If (Not ForgeIsEmpty(status)) Then
340               ForgeCopyRow cal, r, out, outRow, lastCol
350               ForgeSetValue out.Cells(outRow, (lastCol + 1)), status
360               outRow = (outRow + 1)
              End If
Next_26:
          Next r
370       ForgeSort ForgeUsedRange(out), cDue, True, True
380       BuildCalibrationDueList = (outRow - 2)
381       GoTo CleanExit
CleanExit:
          ForgeRestore forgeScreen, forgeCalc, forgeEvents
          Exit Function
ErrHandler:
          ForgeLog "BuildCalibrationDueList", Err.Number, Err.Description, Erl
          BuildCalibrationDueList = -2
          Resume CleanExit
End Function
```

### task_11.vcl: HighlightRejectedWelds

```vcl
VCL 0.2
TASK
# ASSUME: WeldLog headers are in row 1 and include "Weld Seam No" and "Visual"; data starts in row 2
# ASSUME: "Accepted" is matched ignoring case and extra spaces; blank, error or any other Visual value counts as not Accepted
# ASSUME: a Weld Seam No made only of spaces counts as blank
# ASSUME: the macro owns the row fills on WeldLog: every data row is reset to no fill, then flagged rows get red #FF0000 across columns 1 to the last header column
# ASSUME: rows with a blank Weld Seam No are never highlighted and are left with no fill
# ASSUME: the last data row is the lowest row holding a value in any header column
PROC     HighlightRejectedWelds
KIND     Function
MODULE   modWeldCheck
ARGS     ws As Worksheet
RET      Long
VAR      cSeam As Long
VAR      cVisual As Long
VAR      lastCol As Long
VAR      lastRow As Long
VAR      n As Long
VAR      flagged As Long
PRE      COL($ws, "Weld Seam No") > 0; COL($ws, "Visual") > 0
MODIFIES $ws fills
OP       LET(cSeam, COL($ws, "Weld Seam No"))
OP       LET(cVisual, COL($ws, "Visual"))
OP       LET(lastCol, LASTCOL($ws, 1))
OP       LET(lastRow, 1)
OP       LOOP_COLS dir=right from=1 to=lastCol var=c
OP         LET(n, LASTROW($ws, c))
OP         IF n > lastRow THEN LET(lastRow, n)
OP       LET(flagged, 0)
OP       LOOP_ROWS dir=down from=2 to=lastRow var=r
OP         IF IS_EMPTY(CELL($ws, r, cSeam))
OP           FORMAT(ROW_RANGE($ws, r, 1, lastCol), fill=none)
OP         ELSE
OP           IF SAME(CELL($ws, r, cVisual), "Accepted")
OP             FORMAT(ROW_RANGE($ws, r, 1, lastCol), fill=none)
OP           ELSE
OP             FORMAT(ROW_RANGE($ws, r, 1, lastCol), fill=red)
OP             LET(flagged, flagged + 1)
OP       RETURN(flagged)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  weldlog_01.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Joint Type,C:Welder ID,D:Visual,E:Remarks rows=30 stale=10 data=A:seq(1001,1)+boundary(1001,"")+blank(2);B:values("Butt"x15,"Fillet"x15);C:values("W-01"x10,"W-02"x10,"W-03"x10);D:values("Accepted"x14," accepted "x2,"ACCEPTED"x2,"Rejected"x6,"Repair"x6)+boundary("Rejected","Repair")+blank(2);E:text(15)+blank(3)
FIXTURE  weldlog_01.xlsx sheet=Notes cols=A:Note rows=3 data=A:text(10)
FIXTURE  weldlog_02.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Joint Type,C:Welder ID,D:Visual,E:Remarks rows=30 stale=30 data=A:seq(1001,1)+boundary(1001,"")+blank(2);B:values("Butt"x15,"Fillet"x15);C:values("W-01"x10,"W-02"x10,"W-03"x10);D:values("Accepted"x14," accepted "x2,"ACCEPTED"x2,"Rejected"x6,"Repair"x6)+boundary("Rejected","Repair")+blank(2);E:text(15)+blank(3)
TEST     fixture=weldlog_01.xlsx call=HighlightRejectedWelds(WeldLog) runs=2 assert=FILL_COUNT(WeldLog![Visual], #FF0000) = RESULT AND RESULT = COUNTIFS(WeldLog![Weld Seam No], "<>", WeldLog![Visual], "Rejected") + COUNTIFS(WeldLog![Weld Seam No], "<>", WeldLog![Visual], "Repair") + COUNTIFS(WeldLog![Weld Seam No], "<>", WeldLog![Visual], "=") AND ERROR_COUNT = 0
TEST     fixture=weldlog_01.xlsx call=HighlightRejectedWelds(WeldLog) runs=2 assert=FILL_COUNT(WeldLog![Weld Seam No]2, #FF0000) = 1 AND FILL_COUNT(WeldLog![Visual]2, #FF0000) = 1 AND FILL_COUNT(WeldLog![Remarks]2, #FF0000) = 1 AND FILL_COUNT(WeldLog![Visual]3, #FF0000) = 0 AND ERROR_COUNT = 0
TEST     fixture=weldlog_02.xlsx call=HighlightRejectedWelds(WeldLog) runs=2 assert=FILL_COUNT(WeldLog![Visual], #FF0000) = RESULT AND RESULT = COUNTIFS(WeldLog![Weld Seam No], "<>", WeldLog![Visual], "Rejected") + COUNTIFS(WeldLog![Weld Seam No], "<>", WeldLog![Visual], "Repair") + COUNTIFS(WeldLog![Weld Seam No], "<>", WeldLog![Visual], "=") AND FILL_COUNT(WeldLog![Weld Seam No]3, #FF0000) = 0 AND FILL_COUNT(WeldLog![Visual]3, #FF0000) = 0 AND FILL_COUNT(WeldLog![Remarks]3, #FF0000) = 0 AND ERROR_COUNT = 0
TEST     fixture=weldlog_01.xlsx call=HighlightRejectedWelds(WeldLog) assert=FILL_COUNT(WeldLog![Weld Seam No], #FF0000) = RESULT AND COUNTA(WeldLog![Visual]) = PRE_COUNTA(WeldLog![Visual]) AND COUNTIF(WeldLog![Visual], "Accepted") = PRE_COUNTIF(WeldLog![Visual], "Accepted") AND COUNTA(WeldLog![Weld Seam No]) = PRE_COUNTA(WeldLog![Weld Seam No]) AND ERROR_COUNT = 0
TEST     fixture=weldlog_01.xlsx call=HighlightRejectedWelds(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
```

```vba
'@VCL PROC=HighlightRejectedWelds  source: task_11.vcl
Public Function HighlightRejectedWelds(ByVal ws As Worksheet) As Long
          Dim cSeam As Long
          Dim cVisual As Long
          Dim lastCol As Long
          Dim lastRow As Long
          Dim n As Long
          Dim flagged As Long
          Dim c As Long
          Dim r As Long
          Dim forgeScreen As Boolean
          Dim forgeCalc As Long
          Dim forgeEvents As Boolean
          forgeScreen = Application.ScreenUpdating
          forgeCalc = Application.Calculation
          forgeEvents = Application.EnableEvents
          On Error GoTo ErrHandler
200       If Not (ForgeCol(ws, "Weld Seam No") > 0) Then ForgeLog "HighlightRejectedWelds", 0, "PRE failed: (COL($ws,""Weld Seam No"")>0)", 200: HighlightRejectedWelds = -1: GoTo CleanExit
201       If Not (ForgeCol(ws, "Visual") > 0) Then ForgeLog "HighlightRejectedWelds", 0, "PRE failed: (COL($ws,""Visual"")>0)", 200: HighlightRejectedWelds = -1: GoTo CleanExit
          Application.ScreenUpdating = False
          Application.Calculation = xlCalculationManual
220       cSeam = ForgeCol(ws, "Weld Seam No")
230       cVisual = ForgeCol(ws, "Visual")
240       lastCol = ForgeLastCol(ws, 1)
250       lastRow = 1
260       For c = 1 To lastCol
270           n = ForgeLastRow(ws, c)
280           If (n > lastRow) Then lastRow = n
          Next c
290       flagged = 0
300       For r = 2 To lastRow
310           If ForgeIsEmpty(ws.Cells(r, cSeam).Value) Then
320               ForgeNoFill ForgeRowRange(ws, r, 1, lastCol)
              Else
340               If ForgeSame(ws.Cells(r, cVisual).Value, "Accepted") Then
350                   ForgeNoFill ForgeRowRange(ws, r, 1, lastCol)
                  Else
370                   ForgeRowRange(ws, r, 1, lastCol).Interior.Color = RGB(255, 0, 0)
380                   flagged = (flagged + 1)
                  End If
              End If
          Next r
390       HighlightRejectedWelds = flagged
391       GoTo CleanExit
CleanExit:
          ForgeRestore forgeScreen, forgeCalc, forgeEvents
          Exit Function
ErrHandler:
          ForgeLog "HighlightRejectedWelds", Err.Number, Err.Description, Erl
          HighlightRejectedWelds = -2
          Resume CleanExit
End Function
```

### task_12.vcl: RouteJobPacks

```vcl
VCL 0.3
TASK
# ASSUME: Tracker headers are in row 1 and include "Job ID", "Report File" and "Status"; column order does not matter
# ASSUME: $dump is the folder the test reports arrive in, and $jobs holds one folder per job; both must exist (sandbox folders only until March 2027)
# ASSUME: each report is copied into its job's folder, never moved: the file in $dump always stays, and nothing is ever overwritten
# ASSUME: a row with a blank Job ID is skipped and its Status is left as it is
# ASSUME: a Job ID or Report File that is not a safe file name gets Status "Bad name", and nothing is made or copied for that row
# ASSUME: Status gets the file op's status word (Copied, Already there, Conflict, Locked, Missing, Blocked by file, ...)
# ASSUME: the result counts the rows whose report is now in its job's folder (Copied or Already there), so a re-run gives the same result
PROC     RouteJobPacks
KIND     Function
MODULE   modRouter
ARGS     tracker As Worksheet, dump As String, jobs As String
RET      Long
VAR      cJob As Long
VAR      cFile As Long
VAR      cStatus As Long
VAR      status As String
VAR      routed As Long
PRE      COL($tracker, "Job ID") > 0; COL($tracker, "Report File") > 0; COL($tracker, "Status") > 0; FOLDER_EXISTS($dump); FOLDER_EXISTS($jobs)
MODIFIES $tracker["Status"]
MODIFIES $jobs files
OP       LET(cJob, COL($tracker, "Job ID"))
OP       LET(cFile, COL($tracker, "Report File"))
OP       LET(cStatus, COL($tracker, "Status"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($tracker, cJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($tracker, r, cJob))
OP         IF SAFE_NAME(CELL($tracker, r, cJob)) = "" OR SAFE_NAME(CELL($tracker, r, cFile)) = ""
OP           SET_VALUE(CELL($tracker, r, cStatus), "Bad name")
OP         ELSE
OP           MAKE_FOLDER(PATH_JOIN($jobs, SAFE_NAME(CELL($tracker, r, cJob))), into=status)
OP           IF status = "Created" OR status = "Exists"
OP             COPY_FILE(PATH_JOIN($dump, SAFE_NAME(CELL($tracker, r, cFile))), PATH_JOIN($jobs, SAFE_NAME(CELL($tracker, r, cJob)), SAFE_NAME(CELL($tracker, r, cFile))), into=status)
OP           SET_VALUE(CELL($tracker, r, cStatus), status)
OP           IF status = "Copied" OR status = "Already there" THEN LET(routed, routed + 1)
OP       RETURN(routed)
PERF     screen_off
ERR      standard
FIXTURE  jobpacks_01.xlsx sheet=Tracker cols=A:Job ID,B:Report File,C:Status rows=10 data=A:values("J-101"x2,"J-102"x1,"J-103"x1,"..\x"x1,"J-104"x1,""x1,"J-105"x1,"J-106"x1,"J-107"x1)+boundary("J-101","J-102","J-103","..\x","J-104","","J-105","J-101","J-106","J-107");B:values("r1.pdf"x2,"r2.pdf"x1,"r3.pdf"x1,"r4.pdf"x2,"r5.pdf"x1,"r9.pdf"x1,"CON.pdf"x1,"r6.pdf"x1)+boundary("r1.pdf","r2.pdf","r3.pdf","r4.pdf","r5.pdf","r1.pdf","r9.pdf","r4.pdf","CON.pdf","r6.pdf");C:values("old"x10)
FIXTURE  jobpacks_01.xlsx sheet=Notes cols=A:Note rows=3 data=A:text(10)
FIXTURE  jobpacks_01.xlsx folder=Dump files=r1.pdf:120,r2.pdf:80,r3.pdf,r4.pdf:64,r5.pdf,r6.pdf:70 locked=r3.pdf
FIXTURE  jobpacks_01.xlsx folder=Jobs files=J-102\r2.pdf:99,J-104:5,J-107\r6.pdf:70
TEST     fixture=jobpacks_01.xlsx call=RouteJobPacks(Tracker, Dump, Jobs) assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(Tracker![Status], "Copied") = 2 AND COUNTIF(Tracker![Status], "Already there") = 1 AND COUNTIF(Tracker![Status], "Conflict") = 1 AND COUNTIF(Tracker![Status], "Locked") = 1 AND COUNTIF(Tracker![Status], "Missing") = 1 AND COUNTIF(Tracker![Status], "Blocked by file") = 1 AND COUNTIF(Tracker![Status], "Bad name") = 2 AND COUNTIF(Tracker![Status], "old") = 1 AND FILE_EXISTS("Jobs\J-101\r1.pdf") AND FILE_EXISTS("Jobs\J-101\r4.pdf") AND FILE_COUNT("Jobs") = PRE_FILE_COUNT("Jobs") + 2 AND FILE_COUNT("Dump") = PRE_FILE_COUNT("Dump")
TEST     fixture=jobpacks_01.xlsx call=RouteJobPacks(Tracker, Dump, Jobs) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(Tracker![Status], "Already there") = 3 AND COUNTIF(Tracker![Status], "Copied") = 0 AND COUNTIF(Tracker![Status], "Locked") = 1 AND COUNTIF(Tracker![Status], "Missing") = 1 AND FILE_COUNT("Jobs") = PRE_FILE_COUNT("Jobs") + 2
TEST     fixture=jobpacks_01.xlsx call=RouteJobPacks(Tracker, Dump, "X:\KRONOS_NO_SUCH_FOLDER") assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTIF(Tracker![Status], "old") = 10 AND FILE_COUNT("Jobs") = PRE_FILE_COUNT("Jobs")
TEST     fixture=jobpacks_01.xlsx call=RouteJobPacks(Notes, Dump, Jobs) assert=RESULT = -1 AND ERROR_COUNT = 1 AND FILE_COUNT("Jobs") = PRE_FILE_COUNT("Jobs") AND FILE_COUNT("Dump") = PRE_FILE_COUNT("Dump")
```

```vba
'@VCL PROC=RouteJobPacks  source: task_12.vcl
Public Function RouteJobPacks(ByVal tracker As Worksheet, ByVal dump As String, ByVal jobs As String) As Long
          Dim cJob As Long
          Dim cFile As Long
          Dim cStatus As Long
          Dim status As String
          Dim routed As Long
          Dim r As Long
          Dim forgeScreen As Boolean
          Dim forgeCalc As Long
          Dim forgeEvents As Boolean
          forgeScreen = Application.ScreenUpdating
          forgeCalc = Application.Calculation
          forgeEvents = Application.EnableEvents
          On Error GoTo ErrHandler
200       If Not (ForgeCol(tracker, "Job ID") > 0) Then ForgeLog "RouteJobPacks", 0, "PRE failed: (COL($tracker,""Job ID"")>0)", 200: RouteJobPacks = -1: GoTo CleanExit
201       If Not (ForgeCol(tracker, "Report File") > 0) Then ForgeLog "RouteJobPacks", 0, "PRE failed: (COL($tracker,""Report File"")>0)", 200: RouteJobPacks = -1: GoTo CleanExit
202       If Not (ForgeCol(tracker, "Status") > 0) Then ForgeLog "RouteJobPacks", 0, "PRE failed: (COL($tracker,""Status"")>0)", 200: RouteJobPacks = -1: GoTo CleanExit
203       If Not ForgeFolderExists(dump) Then ForgeLog "RouteJobPacks", 0, "PRE failed: FOLDER_EXISTS($dump)", 200: RouteJobPacks = -1: GoTo CleanExit
204       If Not ForgeFolderExists(jobs) Then ForgeLog "RouteJobPacks", 0, "PRE failed: FOLDER_EXISTS($jobs)", 200: RouteJobPacks = -1: GoTo CleanExit
          Application.ScreenUpdating = False
230       cJob = ForgeCol(tracker, "Job ID")
240       cFile = ForgeCol(tracker, "Report File")
250       cStatus = ForgeCol(tracker, "Status")
260       For r = 2 To ForgeLastRow(tracker, cJob)
270           If ForgeIsEmpty(tracker.Cells(r, cJob).Value) Then GoTo Next_26
280           If ((ForgeSafeName(tracker.Cells(r, cJob).Value) = "") Or (ForgeSafeName(tracker.Cells(r, cFile).Value) = "")) Then
290               ForgeSetValue tracker.Cells(r, cStatus), "Bad name"
              Else
310               status = ForgeMakeFolder(jobs, ForgePathJoin(jobs, Array(ForgeSafeName(tracker.Cells(r, cJob).Value))))
320               If ((status = "Created") Or (status = "Exists")) Then
330                   status = ForgeCopyFile(dump, ForgePathJoin(dump, Array(ForgeSafeName(tracker.Cells(r, cFile).Value))), jobs, ForgePathJoin(jobs, Array(ForgeSafeName(tracker.Cells(r, cJob).Value), ForgeSafeName(tracker.Cells(r, cFile).Value))))
                  End If
340               ForgeSetValue tracker.Cells(r, cStatus), status
350               If ((status = "Copied") Or (status = "Already there")) Then routed = (routed + 1)
              End If
Next_26:
          Next r
360       RouteJobPacks = routed
361       GoTo CleanExit
CleanExit:
          ForgeRestore forgeScreen, forgeCalc, forgeEvents
          Exit Function
ErrHandler:
          ForgeLog "RouteJobPacks", Err.Number, Err.Description, Erl
          RouteJobPacks = -2
          Resume CleanExit
End Function
```

## Appendix D: DECISIONS.md, in full

### VBA Forge decisions

One line per decision: what was decided, why, and where it is enforced. Boss can overturn any
of them. Ids: **G** grammar, **S** semantics, **T** types (from `docs/vcl-m2-findings.md`),
**E** language rules, **F** fixtures, **D** decisions Claude made on its own under the rule "pick the more
conservative option, record it here, and carry on".

#### Language (VCL 0.1 revision 4, 16 Oct 2026)

These clarify revision 3; they don't change any valid spec except task_06 (T2). spec.md and the
Orchestrator prompt state them in the reference text.

| Id | Decision | Why | Enforced by |
| --- | --- | --- | --- |
| G2 | Option names are reserved words: dir, from, to, var, in, mode, into, key, order, header, what, fill, numfmt, bold, width, style, cols, col, crit, readonly, format, save, after, name, match. Only these make `name=value` options (no space before `=`); any other `name=value` is a comparison. They can't be argument, variable or loop variable names. | `DICT_ADD(d, k, v, mode=sum)` and `SET_VALUE(c, got = 0)` look alike; reserved keys make the grammar unambiguous (it builds as LALR). | `vcl/grammar.lark` (OPTKEY), validator (reserved names, option splitting) |
| G6 | A quote inside text is written as two quotes, as in VBA: `"Say ""hi"""`. | Revision 3 never said how. | grammar (STRING), `compiler/parser.py` tokenizer and `values(...)` |
| G7 | Operator precedence matches VBA, tightest first: `-` (negation), `* /`, `+ -`, `&`, comparisons, NOT, AND, OR. Comparisons don't chain. | `a < b < c` would make VBA compare a Boolean with c, silently. | grammar |
| S1 | IS_EMPTY: empty cell, empty text, or text made only of spaces, non-breaking spaces (character 160), tabs or line breaks. An error cell is not empty. | Copy-pasted "blank" cells often hold character 160. | `vcl/semantics.py`, ForgeLib `ForgeIsEmpty` |
| S2 | TRIM removes spaces, non-breaking spaces, tabs and line breaks at both ends; inner spaces stay. | VBA's Trim removes only ordinary spaces. | semantics, `ForgeTrim` |
| S3 | IS_NUMBER: numbers, and text that is a plain number after TRIM (optional sign, digits, optional `.` and digits). False for blanks, TRUE/FALSE, dates, errors, "1,250", "1e3", "25%", "(25)". | Regional settings decide what "1,250" means. | semantics, `ForgeIsNumber` |
| S4 | IS_DATE is true only for real date values; text that looks like a date is false. | "03/04/2026" could be March or April. | semantics, `ForgeIsDate` |
| S5 | IS_ERROR (true for #N/A, #DIV/0!, ...) is decided, and arrives with VCL 0.2. In 0.1 an error is not empty, not a number and not a date, so most specs already treat it as invalid. | Reports need to detect error cells. | semantics (`is_error`), VCL 0.2 |
| S6 | Every value becomes text the same way on every PC: empty is "", whole numbers without ".0", "." as the decimal point, dates as yyyy-mm-dd, TRUE/FALSE, errors as their Excel text (#N/A). | VBA's CStr on a date follows regional settings, and Trim on a #N/A cell crashes. | semantics, `ForgeText` |
| S7 | `=` and `<>` with text on either side compare the canonical text exactly (case and spaces count), so a cell holding 25 equals "25". | In VBA `cell = "25"` is False when the cell holds the number 25. | semantics, `ForgeEq` |
| S8 | LEFT and RIGHT: a count below 0 counts as 0; a count beyond the length gives the whole text. | VBA raises an error for a negative count. | semantics, `ForgeLeft`, `ForgeRight` |
| S9 | COL matches a header by its canonical text with line breaks and other whitespace turned into single spaces, ends trimmed, case ignored; the first match wins; 0 if none. | Headers typed with Alt+Enter or double spaces must still match. | semantics, `ForgeCol` |
| S10 | Dictionary keys compare as canonical text, ignoring case: 1001 and "1001" are one key. | VBA's Dictionary treats 1001 and "1001" as different keys. | semantics, `ForgeDictAdd` and friends |
| S11 | LASTROW and LASTCOL work like End(xlUp) and End(xlToLeft): the last cell holding any value, spaces-only text included; an empty column gives 1. | Matches what the generated VBA does. | semantics, `ForgeLastRow`, `ForgeLastCol` |
| T2 | Guard a calculation or date comparison with the check itself (`IF IS_DATE(x)`), never with a flag set earlier (`IF NOT bad`). task_06 was rewritten this way; the rewrite was proved equivalent over all 20 input combinations. | No checker can follow a flag. | rule 15, FINAL CHECK, type checker |
| E1 | `EXIT_IF <condition>` leaves the procedure when the condition is true. A Function then returns its RET type's default (0, "" or FALSE), as VBA does; to return something else write `IF <condition> THEN RETURN(<value>)`. VCL 0.2 adds EXIT_LOOP_IF to leave only a loop. | Revision 3 never defined EXIT_IF. | interpreter, compiler (`GoTo CleanExit`) |

#### Language (VCL 0.1 revision 5, 17 Oct 2026)

Fixture syntax from the mutation evidence (`docs/vcl-mvp-plan.md`). A FIXTURE line without the
new parts builds exactly the same data as before. spec.md and the Orchestrator prompt state
them in the reference text; spec.md has the exact algorithm.

| Id | Decision | Why | Enforced by |
| --- | --- | --- | --- |
| F1 | `stale=<n>` after `rows=<n>`: on a rows=0 sheet, n rows of the text "stale" with a red fill; on a sheet with data, a red fill on n data rows picked by their own PRNG ("<file>\|<sheet>\|stale"), values unchanged. stale may not exceed rows when rows > 0. | Without leftovers, a macro that forgets to clear its output sheet or old fills passes every TEST. | grammar (`stale_opt`), `vcl_fixtures.py`, validator |
| F2 | `+boundary(v1,...)` puts exactly these values in the first data rows. For `values()` the values are moved up, so the counts in the rule stay true; `""` is a blank cell; values must fit the rule (text, numbers, yyyy-mm-dd dates, whole days for reldates). No PRNG draws. | Random data almost never sits exactly on a `<` or `<=` edge, or puts the interesting row in row 2 where `from=2` vs `from=3` shows. | grammar, `vcl_fixtures.py`, validator (`check_modifiers`) |
| F3 | `+dupes(<n>)` gives n non-blank rows below the boundary rows the value of another such row, with its own PRNG ("<file>\|<sheet>\|<col>\|dupes"). | `mode=keep` and `mode=overwrite` only differ when a key repeats. | grammar, `vcl_fixtures.py`, validator |
| F4 | `reldates(lo, hi)`: TODAY + lo to TODAY + hi days. The generator keeps the offset; the interpreter resolves it with the run's TODAY and the fixture-builder VBA writes `Date + n`. | Fixed dates drift past TODAY, so a TEST that is exact today fails next month. With offsets, every result is the same on any day. | `vcl_fixtures.py` (`RelDate`), interpreter, `vba_literal` |
| F5 | Modifiers go in the order +boundary, +blank, +dupes, each at most once. +blank and +dupes only touch rows below the boundary rows. | One order means one algorithm, and the boundary rows stay exactly as written. | validator |
| E2 | Rule 19: a FIXTURE should be able to catch a wrong macro. The validator warns (not errors) when a sheet the macro rebuilds or whose fills it owns has no stale rows in a TEST that gets past PRE, a spec with TODAY has no reldates, a spec with DICT_ADD never repeats a key, or a spec with `<`, `<=`, `>` or `>=` has no +boundary. | Warnings, so every revision 4 spec stays valid; the 10 specs now have none. | validator (`check_fixtures_can_catch_mistakes`) |

#### Language (VCL 0.2 value functions, PR A, 17 Oct 2026)

The first of the four VCL 0.2 PRs (`docs/vep-0.2-core.md`). spec.md lists them in its
"VCL 0.2 additions" section, outside the text it shares with the Orchestrator prompt, until the
migration PR.

| Id | Decision | Why | Enforced by |
| --- | --- | --- | --- |
| E3 | The first line may be `VCL 0.1` or `VCL 0.2`. The VCL 0.2 additions are errors under `VCL 0.1`. Every TASK inside a BUILD repeats the BUILD's own version line. | One spec, one language version; a 0.1 spec never silently uses 0.2. | `parser.py`, `vcl_parse.py`, validator |
| S12 | SAME(a, b): TRIM(a) and TRIM(b), as canonical text, are equal ignoring case, letter by letter (as LOWER). Never compared as numbers: SAME(25, "25.0") is FALSE. | Replaces `UPPER(TRIM(x)) = "READY"`, the most repeated pattern in the 10 specs. | semantics (`same`), `ForgeSame` |
| S13 | CLEAN_TEXT: characters below code 32 (other than tabs and line breaks) are removed, every run of spaces, non-breaking spaces, tabs and line breaks becomes one space, and the ends are trimmed. Case is kept. | Typed and pasted text needs the same tidy-up as COL's header match (S9). | semantics (`clean_text`), `ForgeCleanText` |
| S14 | PARSE_DATE(x, "<format>") with one of five formats: a real date passes through without its time; text is TRIMmed and must match exactly (day and month 1 or 2 digits, year 4 digits from 1900, the date must exist); anything else gives empty. | `CDate` and `DateValue` read "03/04/2026" by the PC's regional settings; a stated format never guesses. | semantics (`parse_date`), validator (format list), `ForgeParseDate` (DateSerial only) |
| S15 | IIF(cond, a, b) gives a or b; both are always evaluated, as VBA's IIf. | Matches the generated VBA exactly. | semantics (`iif`), interpreter, `ForgeIif` |
| S16 | MOD(a, b) = a - b * INT(a / b), as Excel's MOD: sign of b, decimals kept; b = 0 is a run-time error. | VBA's `Mod` rounds both sides (banker's rounding) and gives the sign of a. | semantics (`mod`), `ForgeMod` |
| T3 | IIF's condition never guards its values, so `IIF(IS_NUMBER(x), x * 2, 0)` is rejected. IIF of two values of one type has that type; of two different types it is a raw value that needs a guard. MOD takes numbers only: a raw cell needs IS_NUMBER, and a date is rejected. PARSE_DATE gives a raw value that needs IS_DATE. | Both IIF values run, as both sides of AND / OR do; MOD of a date has no clear meaning. | type checker |

#### Language (VCL 0.2 loop and write ops, PR B, 17 Oct 2026)

| Id | Decision | Why | Enforced by |
| --- | --- | --- | --- |
| E4 | `SKIP_IF <condition>` goes on with the next pass and `EXIT_LOOP_IF <condition>` leaves the loop, both for the innermost LOOP_ROWS, LOOP_COLS or FOR_EACH, and only inside one. They are structural ops, so they can't follow `IF ... THEN` and can't have nested steps. VBA: `If <cond> Then GoTo Next_<loop line>` with that label just before the loop's `Next`, and `If <cond> Then Exit For`. Past either line the type checker knows the condition is FALSE for the rest of the block. | Removes the "whole body inside IF NOT IS_EMPTY" nesting; VBA has no Continue, and every VCL loop is a `For` or `For Each`. | grammar, validator, type checker, interpreter, compiler |
| E5 | `WRITE_ROW(<ws>, <row>, v1, v2, ..., col=<n>)` writes the values to columns n, n+1, ... (n = 1 by default), each as SET_VALUE writes it. It needs at least one value; the values are scalars. Row 1 text values count as written headers for TEST `[Header]` ranges. | Replaces runs of SET_VALUE on one row (task_04's header: 5 lines to 1). | validator, type checker, interpreter, `ForgeWriteRow` |
| S17 | SET_VALUE, WRITE_ROW and COPY_ROW store text as exactly that text. Text Excel could convert (a digit or a character above code 127, a first character of = + - @ ' or #, or TRUE / FALSE after TRIM) is written with a hidden ' prefix; "" empties the cell. The fixture builder uses the same rule. | A plain `.Value = "1001"` stores the number 1001 in Excel while the interpreter keeps the text, so the two could disagree. | semantics (`needs_text_guard`), `ForgeSetValue`, `vcl_fixtures.py` |

#### Language (VCL 0.2 MODIFIES and UNCHANGED, PR C, 17 Oct 2026)

| Id | Decision | Why | Enforced by |
| --- | --- | --- | --- |
| E6 | `MODIFIES $<arg>` (the whole sheet), `MODIFIES $<arg>["<header>", ...]` (values in those columns, row 2 down) or `MODIFIES $<arg> fills`: a TASK key after PRE, one line per sheet or kind. The argument must be a Worksheet. A whole-sheet line can't be combined with others for that argument; each other kind appears at most once. Every write op must go through a `$sheet` argument whose lines allow that kind of change (values, fills, or formats and rows, which need the whole sheet). A declared sheet that nothing writes is a warning. | Each spec says what it may change, so a macro that touches anything else is caught before it runs. | grammar (`modifies_value`), validator (`check_modifies`, `check_writes_are_declared`) |
| E7 | UNCHANGED: for every TEST of a spec with MODIFIES, the interpreter compares the workbook before the first run with the workbook after the last run. Outside what MODIFIES allows, values (with their kind: a number turned into text is a change), fills and number formats must be identical; the header row is always compared; when every run stopped at PRE, nothing may change. Each difference fails the TEST with `UNCHANGED: <Sheet>!<cell> was <old>, now <new>` (at most 5 listed, then a count). | Stronger than the hand-written collateral TESTs, which only count cells; those stay in the specs. | interpreter (`unchanged_problems`), `vclc test` prints the lines |

#### Language (VCL 0.3 router Phase 1a, 18 Oct 2026)

The first VCL 0.3 PR (`docs/vep-0.3-file-ops.md`, the Job-Pack Router). spec.md lists the
additions in its "VCL 0.3 additions" section, outside the text it shares with the Orchestrator
prompt, until router Phase 4.

| Id | Decision | Why | Enforced by |
| --- | --- | --- | --- |
| E8 | The first line may be `VCL 0.1`, `VCL 0.2` or `VCL 0.3`. The VCL 0.3 additions are errors under 0.1 and 0.2, with "X is VCL 0.3: start the spec with 'VCL 0.3' to use it". Everything 0.2 allows, 0.3 allows too (MODIFIES, the automatic UNCHANGED check, SKIP_IF ...): every committed spec, with its first line changed to 0.3, gives the same problems and the same TEST results. The shared reference text and the prompt move to 0.3 in one step, at router Phase 4. | Same rule as E3: one spec, one language version, so a 0.2 spec never silently uses file ops. One switch means Boss updates the Orchestrator Project once. | `parser.py`, `vcl_parse.py`, validator (`at_least`), interpreter, `tests/test_vcl03.py` |
| S18 | SAFE_NAME(x) is CLEAN_TEXT(x), or "" when that is not safe as one file or folder name: empty, over 100 characters, any of `\ / : * ? " < > \|`, `.` or `..`, a trailing dot, or a reserved device name (CON, PRN, AUX, NUL, COM1 to COM9, LPT1 to LPT9) in any case and with any extension (`con.txt`, `COM1 .pdf`). CONSOLE and COM10 are fine. Control characters are already gone (CLEAN_TEXT); text above code 127 is kept. | CLEAN_TEXT is not path-safe: a typed Job ID like `..\..\Windows`, `A/B`, `CON`, `J-1.` or `J*1` would leave the folder, fail in FSO, or act as a wildcard. 100 characters leaves room under Windows' 259-character path limit. | semantics (`safe_name`), conformance cases; `ForgeSafeName` from router Phase 3 |
| S19 | PATH_JOIN(root, part, ...) is the root without its trailing backslashes, then each part, joined by single backslashes. It gives "" when the root is empty or any part is empty or not exactly its own SAFE_NAME. A UNC root (`\\server\share`) works the same; there is no drive-letter rule. | The only way to build a path, so the type checker can see every part. "" instead of a run-time error, so one bad tracker row becomes a status (`Bad name`, router Phase 2), not a stopped run. | semantics (`path_join`), conformance cases; `ForgePathJoin` from router Phase 3 |
| T4 | Two text types, Name (a SAFE_NAME result) and Path. A path is a String `$argument` (the root) or `PATH_JOIN(...)`. PATH_JOIN's first argument must be a String `$argument`; every later part must be `SAFE_NAME(...)` or a quoted literal that is its own safe name, written in place (a variable, a cell or any other function is rejected). FOLDER_EXISTS takes only a path. Names and paths are text everywhere text is read (comparisons, `&`, LEN, TRIM, COL, FILE_NAME, a String VAR), but never numbers. | A check on the shape of each call can't be fooled by the flow of a spec: a variable that held a safe name could be given something else later, so variables are never trusted as parts. | type checker (`check_path_call`) |

#### Language (VCL 0.3 router Phase 1b, 18 Oct 2026)

| Id | Decision | Why | Enforced by |
| --- | --- | --- | --- |
| E9 | `MAKE_FOLDER(<path>, into=<status>)` and `COPY_FILE(<from>, <to>, into=<status>)`, VCL 0.3 only. Every path is `PATH_JOIN($root, ...)` (T4), never a bare root: the root must already exist. `into=` is required and names a String VAR, which gets a status text instead of a run-time error. Every root a file op uses (either side of a copy) must be checked with `FOLDER_EXISTS($root)` in PRE; that is an error, as an unchecked header is (rule 14). A FOLDER_EXISTS($root) PRE condition keeps PRE evaluable, so a TEST whose root is not one of its fixture's folders fails PRE and must expect RESULT = -1. | FSO raises errors for ordinary cases (a locked, missing or existing file); under the standard handler one of them would stop the whole run halfway. A missing root is a wrong argument, so it belongs in PRE. | validator (`check_file_ops`, `check_pre`), type checker |
| E10 | `MODIFIES $<root> files`: folders and files may be added below that String argument, and nothing else on disk may change (checked from router Phase 2). It is the only MODIFIES form for a String argument, and `files` is only for one. MAKE_FOLDER's folder and COPY_FILE's `<to>` must be below a declared root; COPY_FILE's `<from>` may be below any String argument. A declared root that no op adds to is a warning. | Each spec says where it may write on disk, as MODIFIES does for sheets. | grammar (`FILES_WORD`), validator (`check_modifies`, `check_file_ops`), interpreter (`modifies_of`) |
| F6 | Folder FIXTUREs: `FIXTURE <file>.xlsx folder=<Name> [files=<path>[:<bytes>],...] [locked=<path>,...]`, VCL 0.3 only. They share the fixture file name with its sheet lines, so one TEST `fixture=` gets both. `<Name>` is one word, unique in the file and not a sheet name there; TESTs pass it bare for a String argument. `files=` paths are below the folder with `\` between folders; each part is letters, digits, `_ . -` and must be its own SAFE_NAME; no file twice (ignoring case), and no file with something below it. `locked=` files must be in `files=`. A folder exists when a file is below it. Rule 19: a spec that copies files gets a warning when no FIXTURE folder has a `locked=` file. | Synthetic folders only (CLAUDE.md), built the same way every time; locked and conflicting files are the cases a router most often gets wrong. | grammar (`folder_fixture`), `parser.py`, validator (`check_folders`, `check_folders_can_catch_mistakes`) |

#### Language (VCL 0.3 router Phase 2, 18 Oct 2026)

| Id | Decision | Why | Enforced by |
| --- | --- | --- | --- |
| E11 | The file op statuses, checked in this order. MAKE_FOLDER: `Bad name` (path "" or with `*` `?`), `Too long` (over 247 characters), `Exists`, `Blocked by file` (a file at the path or at any folder on the way; then nothing is made), `No folder` (the root is missing), `Created` (every missing level below the root, never the root). COPY_FILE: `Bad name`, `Too long` (either path over 258 characters), `Missing` (no file at `<from>`), `No folder` (the folder `<to>` goes in is missing), `Already there` (same size), `Conflict` (another size, or a folder), `Locked`, `Copied` (same size; `<from>` stays). Excel only, from router Phase 3: `Check copy` (sizes differ after the copy; never deleted) and `Failed` (any other error, logged). FOLDER_EXISTS is TRUE only for a folder. | One attempt, then a status: no overwrite, no delete, no retry loop (the VEP's critique table). Every check happens before anything is made, so a failed op changes nothing. | interpreter (`make_folder`, `copy_file`), `tests/test_vcl03.py` (one case per status); ForgeLib from router Phase 3 |
| F7 | The mock file system. Each FIXTURE folder sits right below the sandbox root `C:\KRONOS_SANDBOX`, and a TEST passing `Dump` passes `C:\KRONOS_SANDBOX\Dump`; paths compare ignoring case, as on Windows. A file without `:<bytes>` gets 100 + below(900) bytes from the PRNG seeded with `"<file>\|<folder>\|<path>\|size"`. TEST assertions `FILE_EXISTS("<Folder>\<path>")` (a file, not a folder) and `FILE_COUNT("<Folder>[\<path>]")` (files at any depth below it; VCL 0.3), with `PRE_FILE_COUNT`; their paths start with a FIXTURE folder of the TEST and every part is a safe name. The final mock file system is part of each TEST result: the snapshot mutation score and `dist/expected` compare it. UNCHANGED for files (E10): nothing removed or changed, additions only below a root the TEST passed for a `MODIFIES $root files` argument (none when every run stopped at PRE). | Synthetic folders only, in memory: nothing real is read or written. The same seed rule as the sheets, so a fixture is the same on every run. | `vcl_fixtures.py` (`MockFS`, `build_folders`), interpreter (`Asserter`, `file_unchanged_problems`), validator (assertion paths), `vcl_mutate.py`, `vclc.py` |

#### Language (VCL 0.4, item 2 switch, 20 Oct 2026)

Item 2 (`docs/vep-0.4-rev3-rest.md`) recorded its meanings as S20 to S25 inside D37 to D45. The
switch moves spec.md and the Orchestrator prompt to 0.4 in one step (D46).

| Id | Decision | Why | Enforced by |
| --- | --- | --- | --- |
| E12 | The first line may be `VCL 0.1` to `VCL 0.4`. The VCL 0.4 additions (a list in DEDUPE `cols=[...]`, SHEET_EXISTS, `MODIFIES $<book> sheets` and with it ADD_SHEET and COPY_SHEET, LIST_FILES, READ_TEXT, WRITE_TEXT, the VISIBLE_COUNT, FORMULA and FILE_TEXT assertions, and text files in folder FIXTUREs) are errors under 0.1 to 0.3, with "X is VCL 0.4: start the spec with 'VCL 0.4' to use it". Everything 0.3 allows, 0.4 allows too, with one exception: DEDUPE's old `cols=1,2`, which 0.4 refuses as ambiguous (G3) and writes `cols=[1, 2]`, with the same meaning (D38); no committed spec uses DEDUPE. Every committed spec and Examples 1 and 3, with their first line changed to 0.4, give the same problems and the same TEST results. | Same rule as E3 and E8: one spec, one language version, so a 0.3 spec never silently uses the new ops. One switch means Boss updates the Orchestrator Project once. | `parser.py`, `vcl_parse.py`, validator (`at_least`), `tests/test_vcl04.py` (VersionSwitch) |

#### Language (VCL 0.5 Phase 1: CALL and BUILD, 21 Oct 2026)

The first VCL 0.5 PR (`docs/vep-0.5-build.md`, D49). spec.md and the prompt move to 0.5 in one step
after the last phase, so Boss updates the Orchestrator once.

| Id | Decision | Why | Enforced by |
| --- | --- | --- | --- |
| E13 | The first line may be `VCL 0.1` to `VCL 0.5`. The 0.5 additions (CALL `into=` and the BUILD rules of E14) are errors under 0.1 to 0.4, or don't apply there. Every committed spec and Examples 1 and 3, with the first line changed to 0.5, give the same problems and the same TEST results. | Same rule as E3, E8 and E12: one spec, one language version. | `parser.py`, `vcl_parse.py`, validator (`at_least`), `tests/test_vcl05.py` |
| E14 | In a VCL 0.5 BUILD: `CALL(<module>.<Proc>, <args>, into=<VAR>)` runs a procedure of the same BUILD, and `into=` keeps a Function's result in a VAR of its RET type (a Sub gives none). Each value fits its SIG parameter, as LET would (a raw cell is guarded first). A Worksheet, Workbook or Range parameter, and any parameter the callee changes, is passed as the caller's own `$argument` of that type. A caller declares, on what it passes, every change the callee's MODIFIES lines allow (the whole sheet covers columns and fills; columns by name; fills, files and sheets as they are). Every TASK that writes declares MODIFIES. Procedure names are unique in the BUILD; the CALL graph has no cycle; every FLOW arrow is a CALL and every CALL is in a FLOW line. `modMain.Run()` takes nothing and gives nothing, only CALLs modMain.RunHeadless (passing sheets as `SHEET(THIS_BOOK, "<real name>")`, a workbook as THIS_BOOK), LETs, IFs and MSGs. FORM and TARGET xlam are errors (D49). In a 0.5 TASK spec, CALL is an error. | RunHeadless's MODIFIES lines then cover the whole tool, so a BUILD TEST gets the same UNCHANGED check as a TASK TEST, and nothing in the tool goes untested. | validator (`check_build_05`), type checker (CALL), `tests/test_vcl05.py` |

#### Decisions Claude made on its own

| Id | Date | Decision | Why it is the conservative choice |
| --- | --- | --- | --- |
| D1 | 16 Oct 2026 | The one fixture generator (`compiler/vcl_fixtures.py`) uses the pilot's SplitMix64 algorithm from `harness/fixtures.py` (PR #4), not the MVP's Python `random` module. It builds bit-for-bit the same data as the pilot. | The algorithm is written down in full (spec.md), so the data never depends on a Python version, and the pilot's expected results stay valid if PR #4 is revived. All specs still pass their TESTs on this data. |
| D2 | 16 Oct 2026 | The batch builder is `vclc build`: besides `forge_module.bas` it writes `forge_fixtures.bas`, one `ForgeFixture_<file>(wb)` builder per fixture file, from the same generator. | Boss asked for one generator used by the interpreter, the batch builder and the expected results; PR #4's runner is deferred, so nothing in main may depend on it. |
| D3 | 16 Oct 2026 | Only `dist/forge_module.bas` is committed (CI fails if it is out of date). The rest of `dist/` is rebuilt by `vclc build` and git-ignored. | The expected results depend on TODAY (task_06, task_09, task_10), so a committed copy would go stale every day. |
| D4 | 16 Oct 2026 | `compiler/validator.py` keeps its hand-written parser and rules and adds two stages: the grammar and the type checker. `vclc check` runs this validator. | Keeps every existing rule and message; a differential run (see the PR) showed where the checkers disagreed, and each disagreement was fixed or is now reported. |
| D5 | 16 Oct 2026 | Where the parsers disagreed, they were aligned on what spec.md already allows: a UTF-8 byte order mark and blank lines before `VCL 0.1` are fine; `IF ... THEN NOTE "..."` is fine. | Accepts only what the validator already accepted; no valid spec is rejected. |
| D6 | 16 Oct 2026 | Assertions allow `+` and `-` only, not `*` and `/`; `NOT` right after a comparison operator needs brackets. | That is what the grammar and the interpreter support; spec.md never listed more. Stricter is safer. |
| D7 | 16 Oct 2026 | ForgeLib stays in `vba/ForgeLib.bas` (the MVP's place), not in `library/`. The conformance cases stay in `tests/`. | `library/` is for approved code, and ForgeLib has never run in Excel. Fewer moves, fewer surprises. |
| D8 | 16 Oct 2026 | These clarifications are VCL 0.1 revision 4; the version line in specs stays `VCL 0.1`. | No valid spec's syntax changes. |
| D9 | 16 Oct 2026 | PRs #5 and #6 (Phase 1 and Phase 2 plans) stay open and untouched. No M0, M1 or M2-research PRs exist; the M2 research arrived inside the MVP zip. | They are proposals waiting for Boss, and nothing depends on them. |
| D10 | 16 Oct 2026 | A RETURN value is converted to the Function's RET type, as VBA does (a Long rounds half to even). | Matches the generated VBA; no spec's results change. |
| D11 | 16 Oct 2026 | Removed `pyflakes-4.0.3-py2.py3-none-any.whl` from the repo root. | It was committed by accident in 7c8e5fe; nothing uses it. |
| D12 | 17 Oct 2026 | task_01 as the pilot ran it is frozen in `tests/data/task_01_pilot.vcl`. The seed-stability test and the pilot-model test (26 shortages, sums 986 and 661) run on that copy instead of being rewritten for the new task_01. | Keeps both pinned tests exactly as strong as before; changing their numbers would have weakened them. |
| D13 | 17 Oct 2026 | Only the FIXTURE and TEST lines of the 10 specs changed for item 2 of the MVP plan, never an OP line. Where a column had to be told apart from its neighbour, a column was added (task_09 NCR gets `F:Remarks`) or given old values (task_08 Inventory `Balance` and `Remark`). Goal TESTs now assert exact results, worked out by the reference interpreter. | The specs' logic is untouched; only the test data and assertions got stronger. |
| D14 | 17 Oct 2026 | The 15 mutants that still survive are equivalent: they can't change any result. They are `LET(flagged, 0)` or `LET(unknown, 0)` on a Long that starts at 0 (task_02, task_06, task_08); the value or mode of a dictionary only used through KEYS or EXISTS (task_04 `seams` x8, task_08 `known`); the value of `mode=count`, which ignores it (task_03 x2); and task_10's `<` vs `<=` TODAY, which the next line overrides. They still count as survivors in the score. | Excluding them would make the score look better than the evidence; the score is 96.4% with them. |
| D15 | 17 Oct 2026 | The interpreter now moves fills and number formats with their rows on SORT and DELETE_ROW, and the expected results include number formats. | That is what Excel does; with stale fills on data sheets the old behaviour would have hidden real differences. |
| D16 | 17 Oct 2026 | The VCL 0.2 core is designed in `docs/vep-0.2-core.md` before any code, and lands in four PRs (value functions; loop and write ops; MODIFIES with UNCHANGED; migration). The key choices: SAME compares trimmed text ignoring case, never as numbers; PARSE_DATE needs an explicit format and never guesses; IIF evaluates both branches, so it is never a guard; MOD follows Excel, not VBA's `Mod`; MODIFIES is checked statically per sheet and at run time per cell; specs and the prompt stay `VCL 0.1` until the migration PR switches them all at once. | Each choice is the stricter or more predictable option, and nothing reaches the Orchestrator until all of 0.2 is proven on the 10 specs. |
| D17 | 17 Oct 2026 | Choices inside PR A that the VEP left open: SAME compares lower-case text with a binary compare, not VBA's `vbTextCompare` (which follows the PC's locale); PARSE_DATE rejects years before 1900, because Excel can't hold them and VBA's DateSerial reads a year below 100 as 19xx or 20xx; CLEAN_TEXT removes control characters before it joins spaces, so "a " & Chr(7) & " b" gives "a b"; a BUILD's TASK sections must use the BUILD's version line; spec.md's revision number and the Orchestrator prompt stay as they are, because nothing in the shared reference text changed. | Each is the stricter or more predictable option, and the Orchestrator Project needs no update until the migration PR. |
| D18 | 17 Oct 2026 | Choices inside PR B: WRITE_ROW compiles to one `ForgeWriteRow ws, row, col, Array(...)` line, not one assignment per value as the VEP sketched, because a spec line may expand to at most 10 numbered VBA lines; every cell still goes through `ForgeSetValue`. COPY_ROW now copies cell by cell through `ForgeSetValue` too, because assigning a block of values lets Excel convert text the same way. The text guard also covers a leading `#` (Excel turns the text "#N/A" into the error) and any character above code 127. The mutation scorer also deletes SKIP_IF and EXIT_LOOP_IF lines and swaps one for the other. The compiler now drops facts about a FOR_EACH variable at the start of the loop, as it already did for LOOP_ROWS. | Each makes Excel hold exactly what the interpreter computed, or makes the checks stricter. No VCL 0.1 spec's results change, and the mutation score stays 96.4%. |
| D19 | 17 Oct 2026 | Choices inside PR C that the VEP left open: SET_FORMULA, DEDUPE and FILTER also count as write ops (the VEP listed nine; these three change sheets too); the validator also checks the kind of change against the kind of MODIFIES line, so a FORMAT fill needs a `fills` line and SORT or DELETE_ROW need the whole sheet, instead of leaving that to the run-time check; a write through a Range variable or anything but a `$sheet` argument is an error, because MODIFIES can't follow it (all 10 specs write through arguments); the header row of a column-declared sheet is never declared, so it must stay identical. | Each is the stricter choice: problems are reported before any TEST runs, with the MODIFIES line that would allow them. |
| D20 | 17 Oct 2026 | PR D is split in two: D1 (`compiler/vcl_migrate.py` and the 10 specs moved to `VCL 0.2`) and D2 (the shared reference text in spec.md, the Orchestrator prompt and Example 1). The VCL 0.1 specs are frozen in `tests/data/specs_v01`; a test proves that migrating them gives exactly the committed specs with the same RESULTs and final workbooks, and the parser and type-checker tests that need the old shapes read those copies (as D12). The VEP's proof "the mutation score is not lower" was replaced with "the TESTs miss no more mutants than before": WRITE_ROW merges lines, which removes easy-to-catch mutants, so task_04 could never keep its percentage (its 8 missed mutants are the D14 equivalents). Every spec passes it (15 missed before, 15 after). Mutants of the new SKIP_IF lines have no VCL 0.1 counterpart; 22 survive because no FIXTURE has a blank key row for the SKIP_IF to skip. They are reported to Boss, not fixed by editing the specs. With them the overall snapshot score reads 90.9% (was 96.4%). From now on a `VCL 0.2` spec that writes but has no MODIFIES gets a warning, and one that writes nothing gets the UNCHANGED check without MODIFIES lines (it may change nothing). | The survivor count is the measure that matters (bugs the TESTs miss) and is strictly comparable; nothing is hidden, the new findings stay in the score. Keeping every VCL 0.1 spec valid under 0.2 (a warning, not an error) is what spec.md promises. |
| D21 | 17 Oct 2026 | PR D2 moved VCL 0.2 into the reference text that spec.md shares with the Orchestrator prompt: the version lines, the MODIFIES key in TASK FORMAT, SKIP_IF, EXIT_LOOP_IF, WRITE_ROW, the six functions with their SEMANTICS, a MODIFIES section, and hard rule 20 ("declare every sheet the macro changes"). Rule 15 now says to compare typed text with SAME. The prompt's FINAL CHECK gained three lines (MODIFIES, loop-only jumps, SAME and PARSE_DATE). spec.md's "VCL 0.2 additions" section became one changelog. Example 1 was migrated by the tool and keeps its file name. The validator still reads VCL 0.1 specs. | One step for the Orchestrator, after all of 0.2 was proven on the 10 specs (as D16 planned); old specs keep working. |
| D22 | 17 Oct 2026 | Phase 0 of `docs/vep-0.3-file-ops.md`, approved by Boss: the FIXTURE lines of task_01, 02, 04, 05, 06, 08 and 10 get blank cells in each column a SKIP_IF tests (`+blank(n)`; task_08 gets a 9th Inventory row with a blank batch, so no real batch is lost; task_05 puts a blank seam with a Rejected fit-up in row 3 and keeps a real exception in row 2). The TEST lines were set to the interpreter's exact new results; where a skipped row keeps its old fill or remark, the TEST now says so (task_02 `= RESULT + 1`, task_08 `"Old remark") = 1`). The same lines were copied into `tests/data/specs_v01`, so the migration proof still holds for every spec. Two fixture unit tests now pin the blanks. A new rule 19 warning asks for a blank cell in a column a SKIP_IF IS_EMPTY tests; the shared reference text gets that sentence with the VCL 0.3 switch (Phase 4), so the Orchestrator prompt doesn't change twice. | It closes the D20 gap: all 22 SKIP_IF mutants are caught, and the 15 survivors left are exactly the D14 equivalents (370 of 385, 96.1%). Only test data and assertions changed, never an OP line. |
| D23 | 17 Oct 2026 | FORMAT fill names (spec.md lists none, black, white, red, green, blue, yellow, magenta, cyan) now work in the interpreter and the compiler. Each name means the colour of the VBA constant of that name (red = #FF0000 = vbRed), defined once in `vcl/semantics.py` (`FILL_COLORS`). Before, the interpreter stored the word "RED", so `FILL_COUNT(..., #FF0000)` never counted it, and the compiler refused every name but none. Found with Boss's weld-log spec (task_11), which uses `fill=red`. | It is what spec.md already allows; the standard VBA colours are the most predictable meaning, and fill=red and fill=#FF0000 now give the same fill and the same VBA. No existing spec uses a fill name, so no result changes. |
| D24 | 17 Oct 2026 | Boss's weld-log spec (pasted from the Orchestrator) is `specs/task_11.vcl`, unchanged. The Job-Pack Router spec planned as task_11 in `docs/vep-0.3-file-ops.md` becomes `specs/task_12.vcl`. Two unit tests that pin the spec count now say 11. task_11 catches 19 of 24 mutants. 4 survivors are equivalent: `LET(lastRow, 1)` changed to 2 or deleted, and `n > lastRow` changed to `>=` (lastRow is always raised to the lowest data row, 31), and `LET(flagged, 0)` deleted (a Long starts at 0, as D14). 1 is a real but small gap: `LOOP_COLS from=1` changed to 2 survives because in both fixtures every column reaches the last row, so skipping column A never shortens lastRow. Catching it needs a fixture whose last row has a value only in column A; that is a spec change, so it is reported to Boss, not made. | Specs are inputs: the spec is kept exactly as written and its gap is reported. The next free number keeps the repo's naming, which was the plan's reason for task_11. |
| D25 | 18 Oct 2026 | **Original model and context data (Boss's standing rule).** The VBA Forge model (March 2027 on) is designed and trained from zero by us, and everything it learns from or reads is in-house. Banned: any pre-trained model or its weights, fine-tuning someone else's model, scraped or downloaded code, datasets or prompts (GitHub, Stack Overflow, public VBA collections), retrieval over outside sources, pre-packaged context files. Allowed: our own VCL-to-VBA datasets and VBA-Bench; the compiler and ForgeLib; general tools (Python, PyTorch) and standard published building blocks such as transformer layers; Claude as the design and build tool (Orchestrator, Claude Code). In-house means made by this repo's tools, by Boss, or by Claude through our pipeline, and verified by our toolchain. Design: compiler first, own model second; the model handles what the compiler can't (NOTE steps, repairs, later plain English to VCL) and ships only where it beats the compiler on VBA-Bench. This reverses the roadmap's 14 Oct choice ("fine-tune an open-weight base model"; "training from scratch on one GPU cannot reach usable VBA quality"): that is still true for general VBA, which is why the compiler stays the core. Updated: the roadmap .docx (key decisions, principles, model strategy, stack, roadmap Phases 3 and 4, VBA-Bench floor, risks; Appendix A kept as history with a note), CLAUDE.md, README, `docs/vcl-mvp-plan.md`, `docs/kgc-kronos-context.md`. The Orchestrator prompt's ROLE line ("A local fine-tuned model compiles your spec") changes with the VCL 0.3 prompt switch (router Phase 4), so Boss updates the Orchestrator Project only once. | Boss's directive, clarified by his answers of 18 Oct 2026: from zero, general libraries allowed, product only. No model work happens before March 2027, so nothing built so far changes. |
| D26 | 18 Oct 2026 | Router Phase 1 is split in two PRs. 1a: the VCL 0.3 line, SAFE_NAME, PATH_JOIN and FOLDER_EXISTS with the path types (T4). 1b: MAKE_FOLDER, COPY_FILE, `MODIFIES $<arg> files` and folder FIXTUREs. The meaning of SAFE_NAME and PATH_JOIN (S18, S19, `vcl/semantics.py`, conformance cases, `docs/vcl-semantics.md`) moved from Phase 2 into 1a, and the interpreter runs them already. FOLDER_EXISTS and the file ops run from Phase 2, with the mock file system. `docs/vep-0.3-file-ops.md` is updated to match. | The type checker needs `safe_name` to check a quoted part, so the rule has one definition, used by the checker and the meaning alike, instead of a copy until Phase 2. Smaller PRs, each verified on its own. |
| D27 | 18 Oct 2026 | Router Phase 1b choices beyond `docs/vep-0.3-file-ops.md`. (1) Adding below a root that has no `MODIFIES $root files` line is always an error, even in a spec with no other MODIFIES lines (sheets only get a warning there). (2) A bare name passed in a TEST for a String argument must be a FIXTURE folder, in every VCL version; before, the validator let it through and the interpreter crashed on it. (3) Folder FIXTURE paths allow only letters, digits, `_ . -` and `\`, with no spaces. (4) Until Phase 2 the interpreter refuses a spec with folder FIXTUREs (`VclError`), and until Phase 3 the fixture-builder VBA refuses one (`ValueError`), instead of skipping the folders. | Writing to disk is the riskiest thing a macro does, so the stricter check. Clear refusals instead of silent gaps: a TEST must never pass because its folders were left out. The small character set keeps fixture lines easy to read and every name valid on Windows. |
| D28 | 18 Oct 2026 | Router Phase 2 choices beyond `docs/vep-0.3-file-ops.md`. (1) A new status, `No folder`, for a missing root (MAKE_FOLDER) or a missing destination folder (COPY_FILE), where the VEP would have said `Failed`. (2) The VEP's "Too long (259 characters or more)" is kept for files (258 is the longest allowed) and tightened for folders to 247, Windows' limit for a new folder. (3) MAKE_FOLDER checks every level for a blocking file before it makes any folder. (4) COPY_FILE onto an existing folder is `Conflict`. (5) A TEST that passes a bare name which is neither a sheet nor a folder of its fixture stops with a VclError. Known gap for Phase 3: in Excel the sandbox sits in `%TEMP%`, a longer root than `C:\KRONOS_SANDBOX`, so a path near the length limit could be `Too long` in Excel and not in the interpreter; Phase 3 must make the sandbox report its root length, and a TEST of `Too long` should use a path far over the limit. | Each status names what the user must fix (`No folder` means "make the folder first"), instead of one vague `Failed`. The tighter folder limit and the check-first rule mean an op never fails halfway. The gap is recorded, not hidden. |
| D29 | 19 Oct 2026 | Router Phase 3 is split in two PRs: 3a, the compiler, ForgeLib's file helpers and the linter's file safety rules; 3b, the sandbox builder (the fixture-builder VBA for folder FIXTUREs). Choices in 3a beyond `docs/vep-0.3-file-ops.md`. (1) The helpers are Functions that return the status, `status = ForgeMakeFolder(root, path)` and `status = ForgeCopyFile(srcRoot, src, dstRoot, dst)`, not Subs with a ByRef status; the compiler takes each root from its path's PATH_JOIN. (2) Each helper runs the interpreter's checks in the same order with the same statuses (E11). Each has its own `On Error GoTo` handler (Excel restores the caller's handler when it returns), not the VEP's `On Error Resume Next`, so no error is skipped silently. Error 70 during the copy is `Locked`; any other error is `Failed`, logged to ForgeLog with the helper's name and the path (spec line 0). (3) FOLDER_EXISTS is TRUE only for a folder written in full (GetAbsolutePathName gives it back unchanged, ignoring case), as in the mock: a relative root, which would depend on Excel's current folder, or one with a trailing `\`, fails PRE. An FSO error there is logged and gives FALSE. (4) The run-time root check (VEP item 3) wants the root written in full and the path below it, both as written and as Windows resolves it; a generated path always passes, so the check only adds `Bad name` for a path built some other way. (5) ForgeLib's limits are Private Consts, and unit tests compare them with `semantics.py` and `vcl_interp.py`, and check that each helper gives exactly the interpreter's statuses plus `Check copy` and `Failed`. (6) The linter now bans, in any letter case, Kill, RmDir, MkDir, ChDir, ChDrive, Shell, GetObject, Declare, the Name statement, writing with Open, the FSO delete, move and text-file methods, the registry functions and SaveAs. CreateObject may make only Scripting.Dictionary (in ForgeNewDict) and Scripting.FileSystemObject (in ForgeFso). CopyFile, CreateFolder and ForgeFso are allowed only in their ForgeLib helpers, and a module-level declaration after a procedure is an error, as in Excel. Known gaps for Phase 5: ForgeLib's file helpers have never run in Excel; and VBA's Len counts a character outside Unicode's Basic Multilingual Plane (an emoji) as 2 where Python counts 1, as LEN already does, so the 100-character name limit and the path limits can differ for such names. | Every check mirrors the reference interpreter, so Excel and the interpreter give the same status; the extra checks only ever refuse. The linter now proves that generated code reaches the disk only through these helpers. Smaller PRs, each verified on its own. |
| D30 | 19 Oct 2026 | Router Phase 3b: the sandbox builder. (1) `vclc build` gives every fixture file with folder FIXTUREs a `ForgeSandbox_<file>()` in `forge_fixtures.bas`, next to `ForgeFixture_<file>(wb)` for its sheets. It makes a new folder `%TEMP%\KGC_KRONOS_SANDBOX_<yyyymmdd_hhnnss>_<n>`, with n from 1 to 99 so several builds a second each get their own (resolved with GetAbsolutePathName, so FOLDER_EXISTS accepts it), builds the FIXTURE folders and files in it with exactly the interpreter's sizes (each file "x" repeated, no line break), holds every `locked=` file open with no sharing until `ForgeSandboxUnlock`, prints the root and its length, and returns the root. A TEST that passes `Dump` to the interpreter passes `<root>\Dump` in Excel. (2) It never writes into anything that was there before: CreateFolder raises error 58 for an existing folder, and a file that is already there stops the build. Nothing is deleted. (3) D28's gap is closed. The builder refuses a root longer than 120 characters, and a TEST now fails (`SANDBOX: ...`) when a file op path below the sandbox is at most its limit here but could be over it in Excel, where the root is up to 120 characters instead of 17. A path must be Too long in both or in neither. `vclc build` also stops for a FIXTURE path too long to build below a 120-character root, or a file over 10,000,000 bytes. (4) The path limits moved from `vcl_interp.py` to `vcl_fixtures.py` beside the mock, which both use. (5) The linter allows CreateObject("Scripting.FileSystemObject") in `SandboxFso` too, CreateFolder in `SandboxFolder`, and writing with Open only in `SandboxFile` and `SandboxLock`. Since the allowlists go by procedure name, the linter now also reports a procedure name defined twice (Excel's "Ambiguous name detected"). | A new folder each time means the builder can't damage anything, even when run twice. The length rule makes every TEST's statuses the same in Excel as in the interpreter for any root up to 120 characters, instead of hoping no TEST sits near a limit. |
| D31 | 19 Oct 2026 | **The model writes VCL, never VBA (Boss approved).** Our own model (D25) writes VCL specs: from plain English, as an offline Orchestrator, and as repairs of specs that fail a check. The compiler stays the only thing that writes VBA. NOTE steps are not translated by the model: VCL grows to cover them, and until it does such a task goes to Claude or Boss. The model is used once it passes enough of the held-out benchmark (the bar is set in March), and Claude stays the Orchestrator until then. This replaces D25's "the model handles what the compiler can't and ships only where it beats the compiler on VBA-Bench": they no longer do the same job. Planned VCL changes for it (a `GOAL` line, a canonical formatter, error codes, fewer NOTE steps, a small language, a minimum mutation score for model-written TESTs) each come as their own VEP after the MVP. The general plan, with its stages, is `docs/ai-plan.md`. Updated: CLAUDE.md, README, `docs/vcl-mvp-plan.md`, `docs/kgc-kronos-context.md`, and the roadmap .docx (a note and its model-strategy lines; a full pass on it comes when Stage A starts). The Orchestrator prompt's ROLE line, which still says a local fine-tuned model compiles the spec, changes at the router Phase 4 prompt switch (as D25 planned) to say the compiler turns the spec into VBA, so Boss updates the Orchestrator once. | Every model output can then be proven by the toolchain without Excel, and every VBA line stays tested, lint-clean compiler output. A small model trained from zero has a much smaller, checkable job, and the data the repo already makes fits it. |
| D32 | 20 Oct 2026 | Router Phase 4 (Boss said "go Phase 4" on 20 Oct 2026) is split in two PRs: 4a, `specs/task_12.vcl` RouteJobPacks; 4b, the shared reference text and the Orchestrator prompt move to VCL 0.3. Choices in 4a. (1) Claude Code wrote task_12 from the VEP's Phase 4 description, as the VEP planned (not the Orchestrator). (2) The result counts the rows whose report is now in its job folder (`Copied` or `Already there`), so a re-run gives the same result. (3) A row whose Job ID or Report File is not a safe name gets `Bad name` before any file op, so no job folder is made for it; fixture row 10 (J-106 with `CON.pdf`) proves it. (4) No `LET(routed, 0)`: a Long starts at 0, and the line would only add an equivalent mutant (D14). (5) The fixture file is `jobpacks_01.xlsx`, not the unit-test router's `router_01.xlsx`. Its fixture covers every status on its own row, plus a blank Job ID whose old Status stays, a re-run (`runs=2`: `Already there` three times), and PRE failures for a missing folder (`"X:\KRONOS_NO_SUCH_FOLDER"`, only ever looked at, never made) and a sheet without the headers. (6) The two unit tests that pin the spec count now say 12, and the test that every VCL 0.2 spec stays valid under 0.3 now pins the 11 VCL 0.2 specs and skips task_12. task_12 catches 29 of 29 mutants with the snapshot (28 of 29 from its TEST assertions alone: `OR` -> `AND` on the Bad name check only adds an empty folder, which only the snapshot sees). The overall snapshot score is 95.4% (418 of 438), with the same 20 survivors. | The VEP's design, with every status proven on a row of its own, and nothing made for a row it refuses. No test was weakened: the count pins moved with the new spec, and the 0.2 -> 0.3 test still checks every VCL 0.2 spec. |
| D33 | 20 Oct 2026 | Router Phase 4b: the reference text that spec.md shares with the Orchestrator prompt (and the context file's Appendix A) moves to VCL 0.3 in one step, so Boss updates the Orchestrator Project once. (1) The version lines, the two file ops, SAFE_NAME, PATH_JOIN and FOLDER_EXISTS with their meanings, the path rules (T4), the statuses in check order (E11), `MODIFIES $<root> files` with its UNCHANGED check (E10), folder FIXTUREs with the sandbox length rule (F6, D30), and FILE_EXISTS and FILE_COUNT (F7). FOLDER_EXISTS's Excel rule (a path written in full, D29) is in its meaning too. (2) A new hard rule 21: files only through MAKE_FOLDER and COPY_FILE, on paths built in place from a String argument; every root checked with FOLDER_EXISTS in PRE; the status shown to the user; no op moves, renames, overwrites or deletes a file. Rule 20 names `files`, and rule 19 gets the two sentences its warnings already check: a blank cell in each column a SKIP_IF IS_EMPTY tests (D22) and a `locked=` file when a spec copies files. The FINAL CHECK gets the same two lines. (3) The ROLE line now says the VBA Forge compiler turns the spec into VBA and a reference interpreter runs the TESTs (D25, D31); the old line named a local fine-tuned model and an Excel harness, neither of which exists. (4) Example 1 moved to 0.3 by its first line only (E8); the migration test now compares the tool's 0.2 output with that line changed. (5) A new Example 3, `vcl/examples/example3_files.vcl`, copies files the way task_12 does, on its own synthetic fixture: no validator problems, its 3 TESTs pass, and they catch 29 of 29 mutants; a unit test keeps it equal to the prompt's copy. (6) The shared-text test's first range now ends at rule 21, so the new rule is checked too, and a test pins the prompt to 0.3 with no "VCL 0.2" or "fine-tuned" left. (7) The 11 VCL 0.2 specs stay 0.2: specs are inputs, and E8's test proves they give the same results under 0.3. | One switch, one Orchestrator update, as E8 and D25 planned. The new rule and examples teach the Orchestrator the strict path rules before it writes its first file spec, and every new line is checked by a test. |
| D34 | 20 Oct 2026 | Router Phase 5 (Boss said "go Phase 5" on 20 Oct 2026) is split in two PRs: 5a, the ForgeLib self-test; 5b, the task_12 pilot runner, the results checker and the guide. Choices in 5a. (1) `compiler/vcl_selftest.py` turns every case of `tests/conformance_cases.py` into two VBA lines in `dist/forge_selftest.bas`: name the case, then check the ForgeLib expression the compiler emits for that meaning against the expected value. COL and LASTROW run on a scratch sheet, and dictionary keys (S10) through a real ForgeNewDict and ForgeDictAdd. Numbers are Doubles, as Excel cells hold them. Text outside printable ASCII is written as ChrW(code), so the module stays ASCII. (2) A value matches only when it is the same kind and the same value (text compared binary, dates as serials), as Python's `==` does on the reference results. (3) Each part runs under `On Error Resume Next`; a case whose call raises an error stays FAIL with the error number and text, and the run goes on. (4) The results go to the sheet ForgeSelfTest, one row per case, and one line to paste back ("SELFTEST n/163 FAILED: ..."). (5) `dist/forge_selftest.bas` is committed next to `forge_module.bas`, an exception to D3: its cases use fixed dates, so it doesn't go stale by itself. `vclc build` writes and lints it, and CI and a unit test fail when it is out of date. (6) The linter takes the Public Forge* helpers of another module (`lint(text, helpers)`), because the self-test calls ForgeLib from its own module. | The self-test is the first check of ForgeLib in real Excel 2016, which nothing in the cloud can run; the conformance table stays the one contract between semantics.py and ForgeLib. Committing it means Boss copies it from GitHub with no build step, and CI keeps it equal to the cases. |
| D35 | 20 Oct 2026 | Router Phase 5b: the task_12 pilot and the check of its results. (1) `vclc build` writes `dist/forge_pilot.bas` for the specs in `vcl_pilot.PILOT_SPECS` (task_12, the router MVP). It is committed, like D34's self-test, and CI and a unit test check it is up to date. Its fixture data is fixed (a reldates value would be written as `Date + n`). (2) Its ForgePilotRun runs every TEST as `run_tests` does: fresh fixture sheets in the pilot workbook; a new sandbox per TEST (D30); the procedure called `runs` times, with ForgeLog cleared before each run, so ERROR_COUNT is the last run's; then the locked files released. Bare names become the fixture's sheets or `<root>\<Folder>`; a quoted string is passed as it is, without the ' prefix cells get. (3) It writes plain lines to the sheet ForgePilot: RESULT and ERROR_COUNT, every non-empty cell as its canonical text (S6), every fill as #RRGGBB, and every folder and file with its size (folders are read only). Line breaks and tabs are written as {LF}, {CR} and {TAB}. A run that stops writes `ABORTED` with the error. (4) The module is self-contained: private copies of the fixture builders and sandbox helpers, and ForgePilotRun is its only Public procedure, so only `forge_module.bas` is needed beside it. (5) The VEP's `dist/expected/task_12*.json` is replaced by `vclc pilot-check RESULTS SPEC...`. It works out the expected lines in the interpreter on the DATE the run printed, compares them TEST by TEST (order doesn't matter), and prints one line per difference. Number formats and the locked flag are not compared: task_12 sets no format, and Excel has already released the locks. (6) The linter lets PilotFiles use SandboxFso. (7) `docs/excel-pilot.md` gives Boss the steps: a new .xlsm, three pasted modules, a compile, ForgeSelfTest, ForgePilotRun, and what to paste back. | A real Excel run compared line by line with the reference, with nothing to install and nothing outside the sandbox touched. A run only happens when Boss chooses, and its expected results are worked out on the day it happens. |
| D36 | 20 Oct 2026 | **MVP plan item 2 (the rest of revision 3), Boss said "go item 2".** Designed in `docs/vep-0.4-rev3-rest.md` before any code (as D16 did for 0.2). Phase 1 brings the interpreter up to the compiler for INSERT_ROW, the array ops, SELECT_CASE, LOG, MSG, RANGE, SHEET, FILE_NAME and FOR_EACH over SHEETS and CELLS, with no language change. Phases 2 to 7 add list syntax with DEDUPE, FIND, the sheet ops, FILTER, SET_FORMULA and text files. Choices made the conservative way, which Boss can overrule: (1) OPEN_BOOK, SAVE_AS and CLOSE_BOOK are deferred, because they open and save real workbooks, which CLAUDE.md rules out until March 2027; the compiler refuses them with a clear message. (2) FILTER hides non-matching rows itself, matching with SAME, instead of AutoFilter criteria, which read numbers and dates through the PC's settings and treat `*` and `?` as wildcards; it never deletes. (3) SET_FORMULA is tested by its formula text only (a new FORMULA assertion), because the interpreter can't calculate Excel formulas; the validator refuses Microsoft 365 functions. (4) WRITE_TEXT writes new files only, below a `MODIFIES $root files` root, with a status like COPY_FILE; `mode=append` is refused because it changes an existing file. (5) DEDUPE and FIND use VCL's own comparisons (S10 keys; SAME), not Excel's RemoveDuplicates and Range.Find, whose rules are undocumented or carry over between calls. (6) WRITE_ARRAY keeps text as text (S17), and an inserted row has no fill and no number format, so neither depends on Excel's defaults. (7) SELECT_CASE compares as `=` does (S7), so the compiler emits `Select Case True` with ForgeEq-style tests instead of VBA's Select Case, which compares 25 with "25.0" as numbers. (8) CALL, BUILD specs and NOTE are not part of this item. | The safest order: first make what already compiles provable, then add meanings one phase at a time, each defined by VCL's own rules rather than Excel defaults that change with the PC. Everything that would touch a real workbook waits. |
| D37 | 20 Oct 2026 | Item 2, Phase 1 (VEP 0.4): the interpreter runs what already compiles, with no language change. (1) READ_ARRAY gives a 1-based 2-D array, also for a single cell (ForgeReadArray), and ARR outside it is VBA error 9 in the interpreter too. (2) WRITE_ARRAY writes the whole array from the top-left cell, and ForgeWriteArray guards each text with the ' prefix as ForgeSetValue does (S17). (3) INSERT_ROW shifts the cells, fills and formats below it down, and the new row is empty with no fill and no number format (ForgeInsertRow clears the formats Excel copies from the row above). (4) SELECT_CASE runs the first CASE whose value is `=` to the selector (S7), else CASE_ELSE; the compiler evaluates the selector once into a Variant, then emits `Select Case True` with one comparison per value, and the type checker checks each CASE value as that comparison. (5) FOR_EACH over CELLS gives each cell's value, row by row; over SHEETS the loop variable has the Sheet type. (6) THIS_BOOK and ACTIVE_BOOK are the TEST's fixture workbook, sheets in FIXTURE order; SHEET of a missing name is a run-time error; ACTIVE_SHEET stops the TEST, because which sheet is active depends on the PC. (7) RANGE takes a cell or a block inside A1:XFD1048576 only; anything else stops the TEST (the validator learns this with the next version line). (8) LOG and MSG are recorded in order; the snapshot mutation check compares them too. (9) FILE_NAME is the text after the last backslash (S20, ForgeFileName), with conformance cases. (10) Finding, not changed: a `[Header]` assertion range includes the header row, though spec.md says "under that row-1 header"; specs don't depend on it today, so it waits for the next language switch. | Each meaning follows VCL's own rules (S7, S17) instead of Excel defaults that change with the PC or the row above, and anything the interpreter can't model the same way in Excel stops the TEST instead of guessing. No spec, no language rule and no threshold changed. |
| D38 | 20 Oct 2026 | Item 2, Phase 2 (VEP 0.4): list syntax and DEDUPE. (1) The toolchain accepts the `VCL 0.4` line now; spec.md and the Orchestrator prompt move to 0.4 in one step after the last language phase, as D33 did for 0.3, so Boss updates the Orchestrator once. Specs in 0.1 to 0.3 stay valid. (2) A list `[a, b]` is allowed only as DEDUPE `cols=` and only in VCL 0.4; anywhere else it is an error. (3) The old form `cols=1,2` keeps its meaning before 0.4 (the parser folds it into `[1, 2]`, so every stage sees one shape); in a 0.4 spec it is an error that shows the list form (G3). (4) `cols=` are sheet column numbers, as SORT `key=` is and as COL gives them, not positions inside the range as in Excel's RemoveDuplicates; each must be 1 to 16384. (5) The key (S21) compares each key cell as a dictionary key does (S10: text, case ignored, no TRIM). A row with an empty or error key cell is never deleted. (6) The later rows of a key are deleted as whole sheet rows, bottom up, as DELETE_ROW deletes; the whole sheet must be in MODIFIES (as today); `header=` defaults to yes, as in SORT. | Fewer deletions whenever there is doubt: no TRIM, and never a row with missing or broken key data (CLAUDE.md: never delete source rows because data is missing). VCL's own key rule instead of RemoveDuplicates, whose comparison isn't documented. One Orchestrator update for the whole of 0.4. |
| D39 | 20 Oct 2026 | Item 2, Phase 3 (VEP 0.4): FIND. FIND is revision 3 (VCL 0.1), so no version line changes. (1) FIND (S22) looks row by row, left to right, and gives the first match or NOTHING. `match=whole` (the default, the stricter one) is SAME; `match=part` means the cell's text holds TRIM(x), case ignored; with `match=part` an empty TRIM(x) finds NOTHING, since every cell would hold it. (2) ForgeFind loops over the cells itself instead of calling Range.Find, which keeps LookIn, LookAt and MatchCase from the last search, even one made by hand. (3) NOTHING is a value in the interpreter; writing to it stops the TEST as VBA error 91 does. (4) MODIFIES: a write through a FIND written in place, FIND(<CELL, ROW_RANGE, USED_RANGE or RANGE of $ws>, ...), is a write to $ws. A write through a Range variable is still refused, because the validator can't tell its sheet, so a spec that changes the found cell repeats the FIND. (5) Finding, not changed: VCL has no ROW() or COLUMN() of a range, so a spec can test whether a cell is there or change the found cell itself, but not reach the rest of its row; a candidate for the next language step. | Every write stays traceable to a declared sheet, and FIND gives the same answer on every PC, whatever was searched before. When in doubt, it finds nothing rather than a guess. |
| D40 | 20 Oct 2026 | Item 2, Phase 4 (VEP 0.4) in two PRs: 4a ADD_SHEET with SHEET_EXISTS and `MODIFIES $book sheets`; 4b COPY_SHEET. Choices in 4a: (1) A Workbook argument is passed in a TEST as that TEST's own fixture file, `call=P(ncr_01.xlsx, Ncr)` (the grammar reads a file name as one token). ADD_SHEET's first argument must be such a `$book`, never THIS_BOOK or ACTIVE_BOOK, so MODIFIES names the workbook it changes and a run never adds sheets to the workbook that holds the macro. (2) `MODIFIES $book sheets` (VCL 0.4): new sheets may be added; every sheet that was there stays as it was unless its own MODIFIES line allows a change; a sheet is never removed. From VCL 0.2 on, ADD_SHEET needs this line. (3) ADD_SHEET needs `name=`, because Excel's own names (Sheet4) depend on the PC's language. The name must pass S23 (1 to 31 characters, no spaces at the ends, none of \ / ? * [ ] :, no apostrophe first or last, not History): a quoted name is checked by the validator, any other at run time. A bad name, or one already used (case ignored), stops the run before anything is added: never a replacement and never a stray sheet. (4) Without `after=` the new sheet goes after the last sheet, set explicitly (Worksheets.Add on its own puts it before the active sheet). (5) SHEET_EXISTS(<book>, <name>) (VCL 0.4) lets a spec skip ADD_SHEET on a re-run (rule 16); the assertion SHEET_EXISTS("<name>") checks the TEST's workbook after the run. Both ignore case, as Excel does. (6) Writes through the Worksheet VAR that ADD_SHEET's `into=` sets are writes to the new sheet, which `MODIFIES $book sheets` covers; only ADD_SHEET may set that VAR, so it never points at a sheet that was already there. (7) Known: Worksheets.Add makes the new sheet the active one. Activate is banned, so nothing undoes it, and nothing in VCL depends on the active sheet (ACTIVE_SHEET stops a TEST, D37). | The workbook a spec changes is always named and checked; a sheet is only ever added, never replaced or removed; a name never depends on the PC. |
| D41 | 20 Oct 2026 | Item 2, Phase 4b (VEP 0.4): COPY_SHEET. (1) VCL 0.4 gives COPY_SHEET a `name=`, which it needs, as ADD_SHEET does: Excel's own name for a copy (Template (2)) is not one a TEST can rely on. The name passes S23 and is not in use (case ignored), checked before anything is copied. (2) The copy goes into the workbook of the sheet it copies, after `after=` or else after the last sheet; the spec needs a Workbook argument with `MODIFIES $book sheets`, because in a TEST every sheet is in that one workbook. (3) A copy holds what the source holds; the interpreter copies values, fills and number formats, which is what it models. (4) Writes through the VAR that `into=` sets are writes to the new sheet, as for ADD_SHEET (D40). (5) ForgeCopySheet finds the copy as the sheet right after `after=` instead of trusting ActiveSheet. | The same guarantees as ADD_SHEET: a sheet is only ever added, under a name the spec chose and checked first. |
| D42 | 20 Oct 2026 | Item 2, Phase 5 (VEP 0.4): FILTER. (1) FILTER(<range>, col=<n>, crit=<x>) leaves the range's first row (its header) alone; every row below it is visible when its cell in sheet column n is SAME as x (S12), and hidden otherwise, so a second FILTER shows rows again. `col=` is a sheet column number, as for SORT and DEDUPE. It never deletes, and ForgeFilter sets each row's Hidden itself instead of using AutoFilter, whose criteria read numbers and dates through the PC's settings and treat * and ? as wildcards. (2) FILTER changes rows, so it needs the whole sheet in MODIFIES, as SORT and DEDUPE do (no new MODIFIES form). The interpreter keeps hidden rows: DELETE_ROW, DEDUPE and INSERT_ROW move them with their rows, COPY_SHEET copies them, the UNCHANGED check reports a row that was hidden or shown on any other sheet, and the mutation snapshot and the expected results (`hidden_rows`) include them. (3) VISIBLE_COUNT(<Sheet>!<range>) (VCL 0.4) is an assertion that counts the non-empty cells of the range in rows that are not hidden, as Excel's SUBTOTAL(103, ...) does. (4) ForgeInsertRow now also makes the new row visible, so it is never hidden like a filtered row next to it. (5) ForgeLastRow keeps End(xlUp) but also checks the rows below it, up to the end of the used range, because End(xlUp) skips hidden rows: LASTROW gives the same answer (S11) whether a FILTER hid rows or not. Every spec's VBA calls it; on a sheet with no hidden rows the answer is unchanged. (6) A spec may not both FILTER and SORT, because how Excel's Range.Sort treats rows a FILTER hid can't be proven by a TEST; a spec that needs both is two specs. | Hiding is the gentlest change a macro can make, done the same way on every PC. Where Excel's behaviour with hidden rows is uncertain (End(xlUp), Sort), the code no longer depends on it or the validator refuses the combination. |
| D43 | 20 Oct 2026 | Item 2, Phase 6 (VEP 0.4): SET_FORMULA. (1) The formula is quoted text, checked by the validator (`compiler/vcl_formula.py`): it starts with =; every function is on an Excel 2016 allowlist (an allowlist, not the VEP's blocklist, so XLOOKUP, FILTER, LET, TEXTJOIN, IFS and any function nobody listed are refused); no other workbook or table (`[` outside text, except R1C1 offsets such as RC[-2] with `style=r1c1`), no `@`, no spill reference (A1#) and no _xlfn. prefix. The list also leaves out INDIRECT, OFFSET, HYPERLINK, WEBSERVICE, CELL and INFO (they reach outside the formula's own cells or the workbook) and TEXT, VALUE, DATEVALUE, TIMEVALUE, FIXED and DOLLAR (they read or write numbers and dates through the PC's regional settings); adding a function is a decision for this table. (2) ForgeSetFormula writes Excel's own English form, `.Formula` or `.FormulaR1C1`, never `.FormulaLocal`, so the text means the same formula on every PC. (3) The interpreter can't calculate a formula: it keeps the formula's text, and reading a formula cell as a value (CELL, FIND, FILTER, DEDUPE, COPY_ROW, READ_ARRAY, SORT, VALUE, COUNTIF, SUM) stops the TEST; COUNTA and LASTROW count it as filled, as Excel does. A TEST checks it with the new assertion FORMULA(<Sheet>!<cell>) (VCL 0.4), the formula text as written, or "" when the cell holds none. Its value is proven only by an Excel run. (4) `style=a1` writes one cell only: on a block Excel shifts each relative reference, which the interpreter doesn't model; `style=r1c1` writes the same text to every cell, as Excel does. (5) A spec may not both SET_FORMULA and move rows (INSERT_ROW, DELETE_ROW, DEDUPE, SORT): Excel rewrites formula references when rows move, which a TEST can't prove. | Microsoft 365 functions stay out by construction, a formula never reaches outside the workbook, and nothing is claimed about a formula's value that only Excel can prove. |
| D44 | 20 Oct 2026 | Item 2, Phase 7 (VEP 0.4) in two PRs: 7a LIST_FILES and ITEMS; 7b READ_TEXT and WRITE_TEXT (planned in the VEP). Choices in 7a: (1) `LIST_FILES(<folder>, "<pattern>", into=<Collection VAR>)` is VCL 0.4. The folder is the String root argument itself or PATH_JOIN of it (T4), and PRE checks the root with FOLDER_EXISTS (rule 14); listing changes nothing, so it needs no MODIFIES. (2) S24: the pattern is quoted, `*` and `?` only (no \ / : " < > \|), matched against the whole name with case ignored, by ForgeNameMatches itself: not VBA's Like, which also reads [ ] and #, and not Dir, which also matches Windows short names (`*.pdf` finds `x.pdfa`). (3) Only the files directly in the folder, sorted by LOWER(name) and then by the name as written, compared by character code: the same order on every PC (not Explorer's number-aware order). (4) The list holds names, not paths, so a spec that copies a listed file still builds its path with PATH_JOIN and SAFE_NAME (T4). (5) A bad, missing or too long folder gives an empty list: listing never raises. (6) ITEMS gives the names as text; COUNT counts a list as it counts a dictionary. (7) For 7b no option word is added: `status` would become reserved (G2), and task_07, task_10 and task_12 have a VAR of that name; READ_TEXT takes its text VAR as an argument and its status into `into=`, as the file ops do (E11). | Reading a folder is safe, but its order and matching depend on the PC unless VCL defines them: S24 does, and ForgeLib follows it. Writing files waits for 7b, with a status for every outcome. |
| D45 | 20 Oct 2026 | Item 2, Phase 7b (VEP 0.4): READ_TEXT and WRITE_TEXT. (1) Both are VCL 0.4 and report a status into `into=` (a String VAR), as MAKE_FOLDER and COPY_FILE do (E11), so a missing or locked file never stops a run halfway. No option word is added (D44). (2) `READ_TEXT(<path>, <String VAR>, into=<status VAR>)`: statuses in check order `Bad name`, `Too long`, `Missing`, `Too big` (over 32,767 characters, what a cell holds), `Locked`, `Not text`, `Read`, and in Excel `Failed`. Locked comes before Not text because a locked file can't be read to tell. The VAR gets the text, or "" unless Read. (3) S25: text is printable ASCII (codes 32 to 126), tabs and line breaks; any other character reads and writes differently with each PC's code page, so it is `Not text`. READ_TEXT drops one final line break (CR LF or LF), which editors add. (4) `WRITE_TEXT(<path>, <text>, into=<status VAR>)` writes a **new** file holding exactly the text, with no line break added, below a `MODIFIES $root files` root: `Bad name`, `Too long`, `Not text`, `No folder`, `Conflict` (a folder there, or a different file), `Locked` (a file of the same size that can't be read to compare), `Already there` (the same text), `Written`; in Excel also `Check write` (the size differs after writing) and `Failed`. It never writes over a file: ForgeWriteText calls CreateTextFile with overwrite off, so a file that appears meanwhile is an error, not a replacement; a partly written file is kept, as COPY_FILE keeps a short copy. `mode=` is gone, because append changes a file that is there. (5) Folder FIXTUREs hold text, `files=<path>:"<text>"` (VCL 0.4): one line of printable ASCII, written by the new SandboxText; a file given by size alone holds that many "x", as before. The new assertion FILE_TEXT("<path>") (VCL 0.4) gives a file's exact text, "" when there is none. The UNCHANGED check compares file texts, and the mutation snapshot and the expected results (`final_texts`) include them. (6) Lint: CreateTextFile moves from banned to allowed in ForgeWriteText only; a Binary read is allowed in ForgeReadRaw only, which closes the file before passing an error on. | Reading and writing text is where code pages and locks make results differ by PC; S25 keeps every text the same everywhere, every outcome is a status, and a file is only ever added, never changed. |
| D46 | 20 Oct 2026 | Item 2, the switch: spec.md, the Orchestrator prompt and the context file's Appendix A move to VCL 0.4 in one step, after the last language phase (D38), so Boss updates the Orchestrator Project once. (1) The shared reference text teaches every item 2 change: the version lines; DEDUPE `cols=[...]`; ADD_SHEET and COPY_SHEET with `name=`; LIST_FILES, READ_TEXT and WRITE_TEXT in their 0.4 forms (`mode=` is gone); SHEET_EXISTS; FIND and COUNT of a list; one SEMANTICS entry for each op item 2 made provable (RANGE, THIS_BOOK, READ_ARRAY, SELECT_CASE, FOR_EACH, INSERT_ROW, LOG and MSG, FILE_NAME, DEDUPE, FIND, FILTER, SET_FORMULA, sheet names, LIST_FILES, the text rule, and READ_TEXT and WRITE_TEXT with their statuses in check order); `MODIFIES $<book> sheets` and the wider UNCHANGED check (hidden rows, removed or new sheets); the VISIBLE_COUNT, FORMULA and FILE_TEXT assertions; a Workbook TEST argument; text files in folder FIXTUREs. (2) It stops teaching what no TEST can prove: ACTIVE_BOOK and ACTIVE_SHEET leave the function list (the interpreter stops a TEST on ACTIVE_SHEET, D37, and THIS_BOOK is the TEST's workbook), and OPEN_BOOK, SAVE_AS and CLOSE_BOOK are listed as refused until March 2027 (D36). The validator still reads all five, so no older spec breaks; the compiler and the interpreter now refuse the three book ops with "<op> opens or saves a real workbook: refused until March 2027 (D36)" instead of "not supported". (3) Hard rules: 10 lets DEDUPE delete rows (never one with a blank key); 19 asks a spec that reads files for a `locked=` file too, and the validator's rule 19 warning now covers READ_TEXT as it covers COPY_FILE; 20 adds filtered sheets and `sheets`; 21 lists the five file ops (LIST_FILES may list the root itself); a new rule 22 covers sheets (ADD_SHEET or COPY_SHEET with `name=`, on a Workbook argument with `MODIFIES $<book> sheets`, skipped when SHEET_EXISTS says the sheet is there, so a re-run is safe; no op removes or replaces a sheet) and formulas (only SET_FORMULA, only Excel 2016 functions, and a TEST checks the text with FORMULA, never the value). The FINAL CHECK gets the same lines. (4) Examples 1 and 3 move to 0.4 by their first line only (E12); the migration test compares the tool's 0.2 output with that line changed. No new example: the samples in `tests/test_vcl04.py` prove each new op, and the prompt stays short. (5) The 12 specs stay in their versions (11 in 0.2, task_12 in 0.3): specs are inputs, and E12's test proves they give the same results under 0.4. (6) The shared-text tests now end at rule 22, and a test pins the prompt to 0.4 with no "VCL 0.3", "VCL 0.2", ACTIVE_SHEET, ACTIVE_BOOK or `mode=new` left. | One switch, one Orchestrator update, as D38 planned. Teaching only what the toolchain can prove keeps the promise item 2 was for: a valid spec is a provable spec. The book ops keep their place in the language (D36) but say plainly why they are refused. |
| D47 | 20 Oct 2026 | Review pass after item 2 (Boss asked for corrections and improvements), with two Orchestrator dry runs: helpers that knew only the prompt wrote nine specs (sheets, DEDUPE, FILTER, formulas, FIND, text files, LIST_FILES, fills) and ran them through the toolchain. (1) Where the reference text differed from the tools, the words changed, not the meaning, so every committed spec and result stays as it is: a `<Sheet>![Header]` range is the column from row 1, the header cell included, as `A:A` is (COUNTA, COUNTIF "<>" and VISIBLE_COUNT count it; task_01 and task_03 assert `COUNTA(...) = RESULT + 1` and task_11 counts fills in it, which closes the D37 finding); ROWS_USED counts row 1; DEDUPE and dictionary keys are not trimmed; assertions have no NOT, TRUE or FALSE ("not there" is `SHEET_EXISTS("X") = 0`); an assertion names a sheet as one word; TEST calls pass text and numbers as written. (2) What the dry runs had to guess is now written down: the VAR type each op fills (READ_ARRAY Variant, LIST_FILES Collection, ADD_SHEET and COPY_SHEET Worksheet, statuses and READ_TEXT's text String) and the start values of VARs; the fill names as #RRGGBB for FILL_COUNT; FORMULA gives the text as written; a formula's references are fixed columns, checked in PRE; FIND gives the cell only. (3) A failing TEST now prints each false assertion with what each side measured (it printed only RESULT and ERROR_COUNT, so a spec writer could not tell what to fix), and `vclc mutate` says how many survivors it did not list. (4) Two table rows in this file (F7 and D44) had an unescaped `\|` that split their cells; `docs/vcl-mvp-plan.md` still said the item 2 ops were "not yet covered", with numbers from 15 Oct (refreshed with measured figures: 12 specs, 46 TESTs, 89.0% assertions-only and 95.4% whole-workbook mutation scores, 1,819 lint lines); CLAUDE.md said all specs were VCL 0.2; the context file now lists the item 2 decisions. (5) A new test keeps the reference text and the validator's tables in step: every op, function and assertion the validator knows is taught (except ACTIVE_BOOK and ACTIVE_SHEET, D46), and nothing is taught that it does not know. (6) Found, not changed (language gaps, for the next version): no row number of a found cell, no count of the rows DEDUPE removed, no rule-14 form for a header the macro adds to an existing sheet, no way for an assertion to name a sheet with spaces. | A reference that says one thing while the tools do another costs a failed TEST for every spec that trusts it. Fixing the words keeps every committed spec and result as it is; the gaps that need new language wait for one version switch, so Boss updates the Orchestrator once more, not once per gap. |
| D48 | 21 Oct 2026 | A strict review of ForgeLib's item 2 helpers (the VBA has never run in Excel), each finding checked against the code before any change. It found nothing that stops the module compiling (lines, continuations, Set, ByRef, duplicate names all fine). Fixed: (1) **Safety:** COPY_SHEET with `after=` a sheet of another workbook put the copy in that workbook, and the rename that follows then renamed a sheet that was already in the source workbook. ForgeCopySheet and ForgeAddSheet now refuse an `after=` of another workbook (error 1004) before anything changes. A TEST has one workbook, so the interpreter can't reach this case. (2) ForgeSheetNameOk took every letter above U+7FFF for a control character (AscW is negative there), so Excel refused names the interpreter allowed; fixed as ForgeCleanText already does. (3) DEDUPE keys: VBA compared them with vbTextCompare, which follows the PC's locale and takes a full-width letter for its ASCII twin, and Python with casefold, where "ß" is "ss", so either side could delete a row the other kept. Both now lower-case letter by letter (LCase$, Python `lower`) and compare exactly (S21 amended, with conformance cases). (4) SET_FORMULA on a cell formatted as Text (@) leaves the formula as text in Excel; both sides now stop the run before writing. (5) VBA's `Open` reads a path through the PC's code page, so READ_TEXT and WRITE_TEXT now give Bad name for a path outside printable ASCII, on both sides (S25 amended). Known, not changed: dictionaries (S10) still compare keys with vbTextCompare in VBA and casefold in Python, so non-ASCII keys (ß, full-width letters) may group differently in Excel, while ASCII keys, which every ID in the registers is, agree; changing S10 touches every dictionary spec and KEYS' first-spelling rule, so it waits for an Excel run to show a need. The same holds for sheet-name case and for LCase$ in FIND, SAME and FILTER on non-ASCII text. LASTROW reads the cells below End(xlUp) up to the end of the used range, and LIST_FILES sorts by insertion: both are bounded, and slow only on a sheet whose used range reaches far below its data or a folder of thousands of files. | Each fixed case is one where Excel would have done what no TEST showed: renamed a sheet, refused a valid name, deleted a row, kept a formula as text. Both sides now refuse or agree, and the wider text-comparison change waits for evidence instead of touching every dictionary spec. |
| D49 | 21 Oct 2026 | **CALL and BUILD specs (VEP 0.5), the next part.** After item 2 Boss said "continue building the next part according to your recommendations"; the recommendation was CALL and BUILD specs before the AI plan's Stage A, because rule 9 already sends the Orchestrator to a BUILD for larger tasks and no BUILD can be proven today. Designed in `docs/vep-0.5-build.md` before any code. Choices: (1) A BUILD is one tool in one workbook: the compiler writes one module per BUILD (`dist/builds/<NAME>.bas`, ForgeLib plus all its procedures), so Boss pastes one module as today, and procedure names must be unique within a BUILD. (2) `CALL(..., into=<var>)` (VCL 0.5) keeps a Function's result; each procedure keeps its own PRE and ERR, as the compiled VBA does (a callee's failed PRE gives -1, its run-time error -2, and the caller goes on). (3) A Worksheet, Workbook or Range is passed as the caller's own `$argument`, written in place, so MODIFIES can follow it; values may be any expression of their type. No recursion: the CALL graph has no cycle. CALL is only for procedures of the same BUILD. FLOW must list exactly the CALL edges. (4) Every TASK of a 0.5 BUILD declares MODIFIES for its own writes, and a caller passing `$x` to a parameter the callee changes declares the same change on `$x`; so RunHeadless's MODIFIES cover the whole tool, and BUILD TESTs get the UNCHANGED check. (5) modMain.Run is a Sub with no ARGS that only CALLs RunHeadless (with `SHEET(THIS_BOOK, "<real name>")`), LETs, IFs and MSGs, so nothing in it is left untested. (6) Deferred: FORM and form events (a UserForm can't be added by pasting text, and building one from code needs a Trust Center change; no op shows a form or reads a control), TARGET xlam (THIS_BOOK would be the add-in, and installing one may need settings the work PC doesn't allow), and running BUILDs written before 0.5 (no MODIFIES in their TASKs). The language keeps them; the compiler refuses them with a clear message. (7) The language gaps of the item 2 review (D47) join 0.5, each decided in its phase from evidence, so Boss updates the Orchestrator once for all of it. | A BUILD becomes provable the same way every op of item 2 did, with the same safety check, before the corpus needs larger tasks. Everything that would need Excel settings, a manual UserForm or an untested path waits. |
| D50 | 21 Oct 2026 | VEP 0.5 Phase 1 (the language surface and the validator, E13, E14). Choices: (1) A caller's MODIFIES line that only a callee uses is not "unused": the warning counts the changes a CALL passes on. (2) After a CALL the type checker forgets every cell guard, since the callee may change any sheet it gets. (3) The MODIFIES rule compares header names as COL does (S9), and the message shows them as the callee wrote them. (4) A Function called without `into=` is allowed: its result is not needed. (5) The rules apply from VCL 0.5 only; a BUILD written before stays check-only, as D49 said. (6) A sample BUILD, NcrReview (mark old open NCRs Overdue, rebuild a status summary), validates with no problems, and each rule has a test with its exact message. Phase 2 runs its TESTs. | Every rule is one the TESTs need to prove the whole tool, checked before anything runs. Warnings that would push a writer to drop a needed MODIFIES line are wrong, so the warning learned about CALL. |
| D51 | 21 Oct 2026 | A second review pass after item 2, done at the same time by another session (its PR #48, which numbered it D47), folded in when Boss chose to combine both reviews. What it added beyond D47 and D48: (1) `docs/excel-pilot.md` said the self-test has 163 cases; it has more, and a test now keeps the guide's numbers equal to the conformance cases, because Boss compares the self-test's message with the guide. (2) Two more Orchestrator dry runs: B (job reports with LIST_FILES, COPY_FILE, WRITE_TEXT, READ_TEXT) passed every stage; A (an NDE register with duplicates, a formula summary and hidden rows) chose a BUILD, because a spec that SET_FORMULAs may not DEDUPE, and missed rule 5's PRE_ measure. Rewritten as one TASK by the new text, it passes every stage. (3) Wording, each checked against the code: rule 5 shows a PRE_ measure; rule 7 asks when one TASK can't do it all; rule 8 (no VCL op deletes, overwrites, runs a program or uses the web); rule 9 asks for one TASK of about 40 OP lines (the validator's limit) and no BUILD or CALL until they can be tested (VEP 0.5 makes them provable, and its last phase rewrites the rule); rule 15 (guards are for workbook values, not the spec's own counters); rule 22 (no row moves beside SET_FORMULA; PRE pins a formula's columns; a formula names only sheets every TEST workbook has); CLEAN_TEXT keeps tabs and line breaks as spaces (S13); Paths; LIST_FILES; the MODIFIES sheets line; the header row may change when the whole sheet is declared; FIXTURE text may hold commas, colons and doubled quotes. The FINAL CHECK matches, and the validator's size warning no longer suggests a BUILD. (4) ForgeLib: ForgeSort gives Orientation, OrderCustom, MatchCase and DataOption1, which Range.Sort otherwise takes from the sheet's last sort; ForgeListFiles leaves out hidden and system files (Excel's ~$ owner file, Thumbs.db), and S24 says so. The linter bans FileCopy, SetAttr, CallByName, ExecuteExcel4Macro, .Save, .Run, .OnTime, .OnKey, .AutoFilter, .FormulaLocal, .ExportAsFixedFormat, Workbooks.Open and SaveChanges, catches Open in every mode, a CreateObject split over lines, and keeps gForgeFso inside ForgeFso. The formula check refuses a quoted name that can only be a file, and the validator refuses a formula naming a sheet a TEST workbook lacks (Excel would open a file dialog). (5) The interpreter's SORT orders error values and date-times as Excel does; a BUILD said "spec doesn't parse: []" and now says it can't be tested yet (VEP 0.5 Phase 2 replaces this with running it). (6) Known gaps, recorded, not built: Excel's locale-aware text sort ignores hyphens; Excel leaves rows FILTER hid in place when it sorts; dictionary keys use vbTextCompare (on a Turkish PC "PIPE" and "pipe" are two keys; DEDUPE no longer does, D48); the self-test has no cases for the sheet and file helpers. (7) For item 3's expressibility report: no text search (INSTR, MID), no expression in a LIST_FILES pattern, FIXTURE text of one line only, a sheet ADD_SHEET made can't be rebuilt on a re-run, DEDUPE gives no count, FILTER hides blank results too. | Two reviews of the same work found different things; keeping both, with one number each, loses nothing and keeps the decision record readable. |
| D52 | 21 Oct 2026 | VEP 0.5 Phase 2 (the interpreter runs BUILD TESTs). (1) A BUILD TEST runs modMain.RunHeadless with the TEST's arguments; CALL runs the callee in a new machine that shares the workbook, the mock file system and ForgeLog, as the compiled procedures share Excel, so ERROR_COUNT counts every procedure's entries. (2) Each procedure keeps its own PRE and ERR, as each compiled procedure has its own handler: a callee's failed PRE gives -1 and its run-time error -2 (into `into=`), each logged once, and the caller goes on. (3) Each value is converted to its parameter's VBA type on the way in, as a VBA call does (a Long rounds half to even: 30.6 passed as a Long is 31). (4) The UNCHANGED check uses RunHeadless's MODIFIES lines, which E14 makes cover the whole tool. (5) A BUILD before VCL 0.5 can't run its TESTs, and says why (its TASKs don't declare MODIFIES); this replaces D51's "can't be tested yet". (6) The interpreter stops 20 calls deep; the validator already refuses cycles, so a real BUILD never gets near it. (7) The sample BUILD NcrReview now has exact assertions and a wrong-sheet TEST, and all three pass; tests show a callee's run-time error and failed PRE, the type conversion and the UNCHANGED check. | The interpreter does what the compiled VBA will do, procedure by procedure, so a BUILD's TESTs prove the tool the same way a TASK's TESTs prove one procedure. |
| D53 | 21 Oct 2026 | VEP 0.5 Phase 3 (the compiler and mutation). (1) `vclc build` writes a 0.5 BUILD as `dist/builds/<NAME>.bas`: one module with ForgeLib and every procedure of the tool (so Boss pastes one module, D49), lint-checked like `forge_module.bas`, with its expected results and fixture builders as for a TASK. It is not committed, as no committed spec is a BUILD yet. (2) CALL with `into=` compiles to `var = Proc(args)`, without it to `Call Proc(args)`; procedures are called by name, which E14's unique names allow. Each compiled procedure keeps its own error handler and restores the screen and calculation settings it found, so a callee nests inside its caller as the interpreter models (D52). (3) The compiler refuses FORM, TARGET xlam and a BUILD before 0.5 with the reason (D49), as the validator does. (4) The mutation scorer runs over every TASK of a BUILD and type-checks each mutant with the BUILD's SIGs, but skips modMain.Run: no TEST calls Run, and E14 keeps it to calling RunHeadless, LET, IF and MSG, so a mutant there could never be caught and would only lower the score. (5) The sample BUILD NcrReview: WriteSummary became a Sub (nothing used its count), and the 31-day row comes first so a loop that starts a row late is caught. Its TESTs catch 33 of 34 mutants; the survivor is DICT_ADD's value under mode=count, which that mode ignores (an equivalent mutant). | A BUILD now goes through the same check, test, build and mutation stages as a TASK, so a valid BUILD is a provable BUILD. Scoring only what a TEST can reach keeps the score honest. |
| D54 | 21 Oct 2026 | The AI context file (`docs/kgc-kronos-context.md`, the file Boss gives other Claude chats) grows from about 69 KB to about 400 KB, under Boss's 3 MB cap ("let the size be up to 3 mb according to your recommendations"). (1) Sections 1 to 12 stay hand-written. Below a marker, `compiler/gen_context.py` copies in seven appendices: A the full language reference, B the rest of the Orchestrator prompt, C every spec with the VBA the compiler writes for it, D DECISIONS.md, E the semantics tables, F ForgeLib, G CLAUDE.md, the plans and every VEP. (2) Not 3 MB: a chat with a 200,000-token window holds about 800 KB in all, and it needs room to think and answer; 400 KB is about half of that. The Python code (the interpreter, the validator) is left out: it is long, and the reference, the specs and the decisions say what it does. (3) A unit test fails when the file differs from what the script builds and when it passes 3 MB, and CI also runs `gen_context.py --check`; so a PR that changes a copied file reruns the script, as `vclc all` is rerun for `dist/`. (4) The test also checks that the hand-written core comes first and that every spec and every decision is in the file. | Copies made by a script can't go stale, so a chat that reads the file sees the same language, specs and decisions as the repo. |
| D55 | 21 Oct 2026 | VEP 0.5 Phase 4a, two of the review gaps (D47, D51). (1) `ROW_OF(<range>)` is the sheet row of the range's first cell (Range.Row), so a FIND result leads to another column of its row: `CELL($ws, ROW_OF(hit), cStatus)`. ROW_OF of NOTHING stops the run with a run-time error, as VBA's error 91 does, so the spec tests IS_NOTHING first; it compiles to plain `.Row`, as IS_NOTHING compiles to `Is Nothing`. (2) `DEDUPE(..., into=<var>)` sets the VAR to how many rows it deleted (0 when none); ForgeDedupe becomes a Function, and a DEDUPE without `into=` still calls it as a statement. (3) Both need `VCL 0.5`. (4) spec.md and the prompt teach VCL 0.4, which can't use them, so they are taught at the switch to 0.5 (Phase 5), and Boss updates the Orchestrator once. Until then the D47 drift guard lets a name of a version newer than spec.md's be untaught, but only if that version's VEP describes it; the moment spec.md moves up, the reference must teach it. (5) The work was started by the parallel session (its branch claude/tender-gates-vbdhjh, which called it D54) and finished here, after D54 went to the context file. A sample 0.5 TASK proves both: 4 TESTs pass and its TESTs catch 8 of 8 mutants. | A count and a row number were the two gaps that forced dry-run specs into workarounds; teaching them only at the switch keeps one Orchestrator update and never shows a 0.4 writer a function its specs can't use. |
| D56 | 21 Oct 2026 | VEP 0.5 Phase 4b, the rule-14 gap (D47). (1) In a VCL 0.5 spec, a header the spec itself writes to row 1 of a sheet argument, as quoted text (`SET_VALUE(CELL($ws, 1, <col>), "<header>")` or a WRITE_ROW to row 1), counts as checked for rule 14, so `COL($ws, "<header>")` needs no PRE line. PRE can't check it: the header isn't there before the first run, so a PRE check would stop that run. (2) The match is exact: the same sheet argument, the same header (as COL matches it, S9), row 1, quoted text. A 0.4 spec keeps the rule-14 error, now with a hint that this needs VCL 0.5. (3) Found while proving it: `MODIFIES $ws ["<header>"]` on such a header passed the validator, but every TEST then failed, because the UNCHANGED check finds a column by its header before the run and the column isn't there yet. In 0.5 the validator now refuses that line and asks for the whole sheet (`MODIFIES $ws`), which is also the only form that lets the header row change (D47). (4) Nothing else changes: the type checker, the interpreter and the compiler already ran such specs, and an assertion could already name the new header (the validator counts row-1 writes). spec.md's rule 14 says so at the switch to 0.5 (Phase 5); until then the VEP describes it. A sample spec proves it: 3 TESTs pass and catch 18 of 18 mutants. | A macro that adds its own status column to a register is a common request; rule 14 must not force a PRE line that can only fail, and the column-MODIFIES trap should fail at the validator with the reason, not at a TEST. |
| D57 | 21 Oct 2026 | VEP 0.5 Phase 4c, the last review gap (D47): an assertion couldn't name a sheet whose name has spaces, so a TEST couldn't look inside a sheet the macro adds as "Open NCRs". (1) In VCL 0.5 an assertion writes such a sheet in single quotes, as Excel writes it in a formula: `VALUE('Open NCRs'!B1)`, `COUNTA('Open NCRs'![NCR No])`, `ROWS_USED('Open NCRs')`; a quote in the name is written `''`. (2) One way to write each name: quotes only for a name that isn't one word (`'Ncr'!B2` is refused, write `Ncr!B2`), so the corpus never has two spellings of one reference. The name must be one ADD_SHEET could give (S23), so a TEST never names a sheet that can't exist. (3) FIXTURE sheets stay one word: a fixture is synthetic and its names are ours to choose, so only sheets the macro itself makes can have spaces. (4) Single quotes stay out of OP lines: text there is in double quotes, and a single-quoted text now gets a message that says so. (5) Changed: the tokenizer, the grammar, the validator's assertion checks and the interpreter's assertion reader; the compiler and ForgeLib don't read assertions. spec.md teaches it at the switch to 0.5 (Phase 5); until then the VEP describes it. A sample spec proves it: 3 TESTs pass and catch 13 of 13 mutants. | Excel's own quoting is the form a writer already knows, and it can name every sheet a macro may add. |

## Appendix E: the meaning of the value functions (`docs/vcl-semantics.md`)

### VCL core semantics (reference)

Generated from `vcl/semantics.py` and `tests/conformance_cases.py`. Do not edit by hand.
ForgeLib (VBA) must give exactly these results; the same table becomes the Excel conformance batch.
In the tables, ⍽ is a non-breaking space (character 160).

#### Decisions

| Id | Function | Meaning |
| --- | --- | --- |
| S1 | `is_empty` | empty cell, empty text, or text made only of spaces, non-breaking spaces, tabs or line breaks. |
| S2 | `trim` | text of v with whitespace (incl. non-breaking spaces) removed at both ends. VBA Trim only removes spaces. |
| S3 | `is_number` | true for numbers and for text that is a plain number after TRIM ("25", " -3.5 "). False for blanks, booleans, dates, errors, and text like "1,250", "1e3", "25%", "(25)". |
| S4 | `is_date` | true only for real date values. Text that looks like a date ("15/10/2026") is false. |
| S5 (VCL 0.2) | `is_error` | true for Excel error values (#N/A, #DIV/0!, ...); false for everything else, including the text "#N/A". |
| S6 | `text_of` | the canonical text of any value, independent of the PC's regional settings. empty -> ""; whole numbers without ".0"; "." as decimal point; dates -> yyyy-mm-dd; booleans -> TRUE/FALSE; errors -> their Excel text (#N/A). |
| S7 | `equals` | '=' between a cell and text compares the cell's canonical text, exactly (case and spaces count). So a cell holding the number 25 equals "25". Plain VBA would say False. |
| S8 | `left` | n below 0 counts as 0; n beyond the length returns the whole text. |
| S9 | `header_key` | how COL matches headers: canonical text, line breaks and other whitespace become one space, runs collapse, ends trimmed, case ignored. "Required\nQty" matches "required qty". |
| S10 | `dict_key` | dictionary keys compare as canonical text, ignoring case: 1001, "1001" and "1001 " (after TRIM) are the same key. Without this, VBA's Dictionary treats 1001 and "1001" as different keys. |
| S11 | `last_row` | like Excel's End(xlUp): the last row holding any value, including spaces or text from formulas. A column that is completely empty gives 1 (the header row). |
| S12 (VCL 0.2) | `same` | SAME(a, b) is true when TRIM(a) and TRIM(b) are the same text, ignoring case (letter by letter, as LOWER: "ß" is not "ss"). Numbers are compared as text, so 25 and "25" are the same but 25 and "25.0" are not. |
| S13 (VCL 0.2) | `clean_text` | CLEAN_TEXT removes characters below code 32 (other than tabs and line breaks), turns every run of spaces, non-breaking spaces, tabs and line breaks into one space, and trims both ends. Case is kept. |
| S14 (VCL 0.2) | `parse_date` | PARSE_DATE(x, format) gives a real date, or empty when x is not a date in that format. A real date passes through (time of day dropped). Text is TRIMmed and must match the format exactly: day and month 1 or 2 digits, year 4 digits from 1900 to 9999, and the date must exist (31/02/2026 gives empty). Numbers and everything else give empty. The format is required, because "03/04/2026" means a different day on different PCs. |
| S15 (VCL 0.2) | `iif` | IIF(cond, a, b) gives a when cond is TRUE, else b. Both a and b are always evaluated (as VBA's IIf), so an IIF branch is never a guard. |
| S16 (VCL 0.2) | `mod` | MOD(a, b) is a - b * INT(a / b), as Excel's MOD: the result has the sign of b, and decimals are kept. VBA's Mod operator rounds both sides first and gives the sign of a, so it is not used. MOD(a, 0) is a run-time error. |
| S17 (VCL 0.2) | `needs_text_guard` | SET_VALUE, WRITE_ROW and COPY_ROW store text as exactly that text. Excel would turn some text into something else when it is written to a cell: a number, date, time, TRUE/FALSE, error value or formula. Such text is written with a ' prefix, which Excel hides, so the cell holds the text unchanged. Being careful, the prefix goes on any text with a digit or a character above code 127, any text starting with = + - @ ' or #, and TRUE or FALSE after TRIM in any case. "" is never prefixed: it empties the cell. |
| S18 (VCL 0.3) | `safe_name` | SAFE_NAME(x) is CLEAN_TEXT(x) when that is a safe Windows file or folder name, else "" (empty). Unsafe: empty, longer than 100 characters, any of \ / : * ? " < > /, the names . and .., a trailing dot, or a reserved device name (CON, PRN, AUX, NUL, COM1 to COM9, LPT1 to LPT9), ignoring case and any extension (con.txt is reserved). A name can never climb out of its folder, because it can't contain a separator or be .. |
| S19 (VCL 0.3) | `path_join` | PATH_JOIN(base, name, ...) is base, then each name, separated by \ (a trailing \ characters on base are dropped). It is "" (empty) when base is empty or any name is not already a safe name (SAFE_NAME(name) = name), so an empty or unsafe part can never become a path; a file op given "" reports Bad name. |
| S20 (VEP 0.4, D36) | `file_name` | FILE_NAME(path) is the text after the last backslash, or all of the text when it has none; a path ending in a backslash gives "". Only \ separates: a / is part of the name. |
| S21 (VEP 0.4, D38) | `same_row_key` | DEDUPE treats two rows as the same when both have a key and the keys are equal. A row's key is the values of its key cells, each as its canonical text, lower-cased letter by letter as VBA's LCase$ does and compared exactly (no TRIM; "Straße" is not "STRASSE", and a full-width letter is not its ASCII twin), so Excel and the interpreter agree on every PC (D48). A row with an empty or error key cell has no key, so DEDUPE never deletes a row because data is missing. |
| S22 (VEP 0.4, D39) | `find` | FIND(range, x, match=...) looks at the cells row by row, left to right, and gives the first that matches, or NOTHING. match=whole (the default): the cell is SAME as x (S12). match=part: the cell's text holds TRIM(x), ignoring case (letter by letter, as LOWER); an empty TRIM(x) finds NOTHING, because every cell would hold it. Here: the 1-based position of that cell among values, or 0 for NOTHING. |
| S23 (VEP 0.4, D40) | `sheet_name_ok` | a name ADD_SHEET may give a new sheet: 1 to 31 characters, no spaces at the ends (TRIM leaves it as it is), none of \ / ? * [ ] :, no apostrophe first or last, no character below code 32, and not History, which Excel keeps for itself (case ignored). Stricter than Excel where Excel allows lookalikes, so a name never differs from what a TEST asserts. |
| S24 (VEP 0.4, D44) | `list_files` | LIST_FILES gives the names of the files directly in the folder that match the pattern, sorted by LOWER(name) and then by the name as written, so the order is the same on every PC. Hidden and system files are left out (D51; FIXTURE folders hold none). Here: from the names of the files in the folder. |
| S25 (VEP 0.4, D45) | `text_ok` | text that READ_TEXT and WRITE_TEXT handle: printable ASCII (codes 32 to 126), tabs and line breaks, at most 32,767 characters. Any other character would read or write differently with each PC's code page. |
| S25 (VEP 0.4, D45) | `file_text` | what READ_TEXT gives for a file's contents: the text with one final line break (CR LF, or LF) dropped, as editors add one; None when it isn't text (S25), which READ_TEXT reports as Not text. |

#### Conformance cases


##### IS_EMPTY

| Input | Result | Why |
| --- | --- | --- |
| (empty) | True | empty cell |
| "" | True | empty text |
| "   " | True | spaces only |
| "⍽ " | True | non-breaking space from copy-paste |
| "\n" | True | line break only (Alt+Enter) |
| 0 | False | zero is a value |
| False | False | FALSE is a value |
| #N/A | False | an error cell is not empty |
| " x " | False | text with spaces around it |

##### IS_NUMBER

| Input | Result | Why |
| --- | --- | --- |
| 25 | True | a number |
| 2.5 | True | a decimal |
| "25" | True | numeric text |
| " -3.5 " | True | numeric text with spaces and a sign |
| "007" | True | leading zeros are still a number (use TRIM for IDs) |
| (empty) | False | blank is not 0 |
| "" | False | empty text is not 0 |
| "1,250" | False | thousands separator: regional, so rejected |
| "1e3" | False | scientific notation text rejected |
| "25%" | False | percent text rejected |
| "(25)" | False | accounting negative rejected |
| True | False | TRUE is not a number |
| date 2026-10-15 | False | a date is a date, not a number |
| #N/A | False | an error is not a number |

##### IS_DATE

| Input | Result | Why |
| --- | --- | --- |
| date 2026-10-15 | True | a real date |
| "15/10/2026" | False | text that looks like a date |
| 46310 | False | a serial number is a number |
| (empty) | False | blank |
| #N/A | False | error |

##### IS_ERROR (VCL 0.2)

| Input | Result | Why |
| --- | --- | --- |
| #N/A | True | #N/A |
| #DIV/0! | True | #DIV/0! |
| "#N/A" | False | the text '#N/A' is just text |
| (empty) | False | an empty cell is not an error |
| 0 | False | zero is not an error |

##### TRIM

| Input | Result | Why |
| --- | --- | --- |
| "  W-01  " | "W-01" | spaces |
| "⍽W-01⍽" | "W-01" | non-breaking spaces (VBA Trim misses these) |
| "W-01\n" | "W-01" | trailing line break |
| "W  01" | "W  01" | inner spaces kept |
| 1001 | "1001" | number becomes its text |
| (empty) | "" | empty |

##### Text conversion (TRIM, UPPER, &, keys)

| Input | Result | Why |
| --- | --- | --- |
| 25.0 | "25" | whole number without .0 |
| 2.5 | "2.5" | decimal point is always '.' |
| date 2026-10-15 | "2026-10-15" | dates are yyyy-mm-dd on every PC |
| True | "TRUE" | boolean |
| #N/A | "#N/A" | error text, not VBA's 'Error 2042' |

##### UPPER

| Input | Result | Why |
| --- | --- | --- |
| "Open" | "OPEN" | case |

##### LEN

| Input | Result | Why |
| --- | --- | --- |
| "W-01" | 4 | length |
| (empty) | 0 | empty |

##### LEFT

| Input | Result | Why |
| --- | --- | --- |
| "Fit-up", 3 | "Fit" | left part |
| "Fit", 10 | "Fit" | longer than the text |
| "Fit", -1 | "" | negative counts as 0 |

##### RIGHT

| Input | Result | Why |
| --- | --- | --- |
| "W-01", 2 | "01" | right part |

##### = with text

| Input | Result | Why |
| --- | --- | --- |
| 25, "25" | True | number cell equals its text (plain VBA says False) |
| "Open", "Open" | True | same text |
| "open", "Open" | False | = is case-sensitive; use UPPER or SAME |
| "Open ", "Open" | False | = doesn't trim; use TRIM or SAME |
| (empty), "" | True | empty cell equals empty text |
| 2, 2.0 | True | numbers |

##### COL

| Input | Result | Why |
| --- | --- | --- |
| ["Part No", "Required Qty"], "required qty" | 2 | case ignored |
| ["Part No", "Required\nQty"], "Required Qty" | 2 | line break in header (Alt+Enter) |
| ["Part  No ", "Qty"], "Part No" | 1 | double and trailing spaces |
| ["Part No", "Part No"], "Part No" | 1 | duplicate header: first wins |
| ["Part No"], "Qty" | 0 | missing header gives 0 |

##### Dictionary keys

| Input | Result | Why |
| --- | --- | --- |
| 1001 | "1001" | number key |
| "1001" | "1001" | text key equals the number key |
| "W-01" | "w-01" | case ignored |

##### LASTROW

| Input | Result | Why |
| --- | --- | --- |
| ["Hdr", "a", (empty), "b", (empty)] | 4 | last value, gaps allowed |
| ["Hdr", "a", "  "] | 3 | a spaces-only cell still counts (like End(xlUp)) |
| [(empty), (empty)] | 1 | empty column gives the header row |

##### SAME (VCL 0.2)

| Input | Result | Why |
| --- | --- | --- |
| " ready ", "READY" | True | spaces at the ends and case are ignored |
| "⍽Ready\n", "ready" | True | non-breaking spaces and line breaks too, as TRIM |
| "Re ady", "Ready" | False | inner spaces count |
| 25, "25" | True | a number compares as its text |
| 25, "25.0" | False | not compared as numbers: use IS_NUMBER and = for that |
| 2.5, " 2.5 " | True | decimal point is always '.' |
| (empty), "" | True | an empty cell is the same as empty text |
| (empty), "   " | True | spaces only, after TRIM, is empty text |
| date 2026-10-15, "2026-10-15" | True | a date compares as yyyy-mm-dd |
| True, "true" | True | TRUE compares as its text, case ignored |
| #N/A, "#n/a" | True | an error compares as its text: use IS_ERROR to tell them apart |
| "straße", "STRASSE" | False | letter by letter, as LOWER: ß is not ss |

##### CLEAN_TEXT (VCL 0.2)

| Input | Result | Why |
| --- | --- | --- |
| "  W-01 ⍽\n  A " | "W-01 A" | runs of any whitespace become one space |
| "W-01\tA" | "W-01 A" | a tab inside becomes a space |
| "Fit\x01up" | "Fitup" | characters below code 32 are removed |
| "a \x07 b" | "a b" | removed first, so the spaces around it join |
| "Fit-Up" | "Fit-Up" | case is kept |
| 1001 | "1001" | a number becomes its text |
| (empty) | "" | empty |
| "   " | "" | spaces only |

##### PARSE_DATE (VCL 0.2)

| Input | Result | Why |
| --- | --- | --- |
| "15/10/2026", "dd/mm/yyyy" | date 2026-10-15 | day first |
| "10/15/2026", "mm/dd/yyyy" | date 2026-10-15 | month first |
| "2026-10-15", "yyyy-mm-dd" | date 2026-10-15 | ISO |
| "15-10-2026", "dd-mm-yyyy" | date 2026-10-15 | dashes |
| "15.10.2026", "dd.mm.yyyy" | date 2026-10-15 | dots |
| " 5/3/2026 ", "dd/mm/yyyy" | date 2026-03-05 | trimmed; 1-digit day and month |
| "03/04/2026", "dd/mm/yyyy" | date 2026-04-03 | the format decides, never the PC |
| "03/04/2026", "mm/dd/yyyy" | date 2026-03-04 | the same text, the other format |
| "31/02/2026", "dd/mm/yyyy" | (empty) | the date must exist |
| "29/02/2028", "dd/mm/yyyy" | date 2028-02-29 | leap day |
| "15/10/26", "dd/mm/yyyy" | (empty) | the year needs 4 digits |
| "15/10/0026", "dd/mm/yyyy" | (empty) | years before 1900 give empty (Excel can't hold them) |
| "015/10/2026", "dd/mm/yyyy" | (empty) | day takes at most 2 digits |
| "15-10-2026", "dd/mm/yyyy" | (empty) | other separators give empty |
| "15 / 10 / 2026", "dd/mm/yyyy" | (empty) | no spaces inside |
| "15/10/2026 08:30", "dd/mm/yyyy" | (empty) | no time of day in text |
| date 2026-10-15, "mm/dd/yyyy" | date 2026-10-15 | a real date passes through |
| date and time 2026-10-15 08:30, "dd/mm/yyyy" | date 2026-10-15 | time of day dropped |
| 46310, "dd/mm/yyyy" | (empty) | a number is not a date |
| (empty), "dd/mm/yyyy" | (empty) | empty gives empty |
| #N/A, "dd/mm/yyyy" | (empty) | an error gives empty |

##### IIF (VCL 0.2)

| Input | Result | Why |
| --- | --- | --- |
| True, "Short", "OK" | "Short" | TRUE gives the first value |
| False, "Short", "OK" | "OK" | FALSE gives the second |
| False, 1, (empty) | (empty) | a value can be empty |

##### MOD (VCL 0.2)

| Input | Result | Why |
| --- | --- | --- |
| 7, 3 | 1 | remainder |
| -7, 3 | 2 | sign of b, as Excel (VBA's Mod gives -1) |
| 7, -3 | -2 | sign of b |
| 7.5, 2 | 1.5 | decimals kept (VBA's Mod rounds them away) |
| 6, 3 | 0 | divides exactly |
| 0, 5 | 0 | zero |

##### Text written by SET_VALUE, WRITE_ROW and COPY_ROW (VCL 0.2)

| Input | Result | Why |
| --- | --- | --- |
| "Short" | False | plain words are stored as typed |
| "Days Open" | False | spaces are fine |
| "" | False | empty text empties the cell |
| "1001" | True | Excel would store the number 1001 |
| "1-2" | True | Excel would store a date |
| "W-01" | True | any digit counts, to be safe |
| "TRUE" | True | Excel would store the logical TRUE |
| " false " | True | after TRIM, in any case |
| "=A1" | True | Excel would store a formula |
| "-" | True | a leading minus could start a formula |
| "#N/A" | True | Excel would store the error value |
| "'quoted" | True | a leading ' would be hidden |
| "café" | True | a character above code 127 counts, to be safe |

##### SAFE_NAME (VCL 0.3)

| Input | Result | Why |
| --- | --- | --- |
| "J-101" | "J-101" | an ordinary job number |
| "  J-101 ⍽" | "J-101" | tidied like CLEAN_TEXT |
| "Report  2026.pdf" | "Report 2026.pdf" | inner spaces joined, extension kept |
| 1001 | "1001" | a number becomes its text |
| (empty) | "" | empty |
| ".." | "" | can't climb out of the folder |
| "." | "" | the folder itself |
| "..\..\Windows" | "" | a separator is never allowed |
| "A/B" | "" | neither is / |
| "J*1" | "" | wildcards are refused |
| "J?1" | "" | wildcards are refused |
| "C:x" | "" | a drive colon is refused |
| "J"1" | "" | quotes are refused |
| "J<1>" | "" | angle brackets are refused |
| "J|1" | "" | a pipe is refused |
| "J-1." | "" | Windows drops a trailing dot, so it is refused |
| "CON" | "" | a reserved device name |
| "con.txt" | "" | reserved with any extension and any case |
| "COM1" | "" | COM1 to COM9 are reserved |
| "LPT9.pdf" | "" | LPT1 to LPT9 are reserved |
| "CONSOLE" | "CONSOLE" | only the exact device names are reserved |
| "COM10" | "COM10" | COM10 is an ordinary name |
| "J\x01-1" | "J-1" | control characters are removed, as CLEAN_TEXT |
| "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx" | "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx" | 100 characters is the limit |
| "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx" | "" | over 100 characters is refused |

##### PATH_JOIN (VCL 0.3)

| Input | Result | Why |
| --- | --- | --- |
| "C:\Jobs", "J-101", "r1.pdf" | "C:\Jobs\J-101\r1.pdf" | base, then each name |
| "C:\Jobs\", "J-101" | "C:\Jobs\J-101" | a trailing separator on base is dropped |
| "\\server\share", "J-101" | "\\server\share\J-101" | a network (UNC) root works |
| "C:\Jobs", "" | "" | an empty part gives an empty path |
| "C:\Jobs", ".." | "" | an unsafe part gives an empty path |
| "C:\Jobs", "A\B" | "" | a part may not hide a separator |
| "", "J-101" | "" | no root, no path |

##### FILE_NAME (revision 3, VEP 0.4)

| Input | Result | Why |
| --- | --- | --- |
| "C:\Jobs\J-101\r1.pdf" | "r1.pdf" | the part after the last backslash |
| "r1.pdf" | "r1.pdf" | no backslash: all of it |
| "C:\Jobs\" | "" | a folder path ending in a backslash |
| "\\server\share\a b.pdf" | "a b.pdf" | a network path; spaces are kept |
| "C:/Jobs/r1.pdf" | "C:/Jobs/r1.pdf" | only a backslash separates |
| (empty) | "" | empty |

##### DEDUPE keys (VCL 0.4)

| Input | Result | Why |
| --- | --- | --- |
| ["S1", "W1"], ["S1", "W1"] | True | the same values in every key column |
| ["S1", "W1"], ["s1", "w1"] | True | case is ignored, as in dictionary keys (S10) |
| [1001], ["1001"] | True | a number and the same number as text |
| [2.5], ["2.50"] | False | text is compared as typed: 2.50 is not 2.5 |
| ["S1"], ["S1 "] | False | no TRIM: a trailing space makes another key |
| ["S1", "W1"], ["S1", "W2"] | False | one key column differs |
| ["ab", "c"], ["a", "bc"] | False | the columns stay apart |
| ["Straße"], ["STRASSE"] | False | letters are lowered one by one, as LCase$: ß stays ß (D48) |
| ["ＡB"], ["AB"] | False | a full-width letter is not its ASCII twin (D48) |
| [(empty)], [(empty)] | False | an empty key: never a duplicate (data is missing) |
| ["S1", "  "], ["S1", "  "] | False | a key cell of spaces only counts as empty |
| [#N/A], [#N/A] | False | an error in a key cell: never a duplicate |
| [date 2026-10-15], [date 2026-10-15] | True | the same date |

##### FIND (revision 3, VEP 0.4)

| Input | Result | Why |
| --- | --- | --- |
| ["Seam", "Welder"], "welder", "whole" | 2 | whole: SAME, so case is ignored |
| [" W1 ", "W1"], "W1", "whole" | 1 | whole: spaces at the ends are ignored |
| ["W12", "W1"], "W1", "whole" | 2 | whole: all of the text must match |
| [25, "25"], "25.0", "whole" | 0 | numbers compare as text: 25.0 is not 25 |
| [25], "25", "whole" | 1 | a number and the same number as text |
| ["a", (empty), "b"], "", "whole" | 2 | whole: an empty x finds the first empty cell |
| ["a", "b"], "c", "whole" | 0 | not there: NOTHING |
| ["W12", "W1"], "W1", "part" | 1 | part: the first cell that holds the text |
| ["Seam S-12 repair"], " s-12 ", "part" | 1 | part: TRIM(x), case ignored |
| [1250], 25, "part" | 1 | part: a number is its text |
| ["a", "b"], "  ", "part" | 0 | part: nothing to look for finds NOTHING |
| [date 2026-10-15], "2026-10-15", "whole" | 1 | a date is its yyyy-mm-dd text |

##### Sheet names for ADD_SHEET (VCL 0.4)

| Input | Result | Why |
| --- | --- | --- |
| "Summary" | True | a plain name |
| "NCR Log 2026-10" | True | spaces inside, digits and a dash |
| "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" | True | 31 characters, the most Excel allows |
| "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" | False | 32 characters is one too many |
| "보고서" | True | letters above U+7FFF are fine (Hangul, D48) |
| "" | False | empty |
| " Summary" | False | a space at the start |
| "Summary " | False | a space at the end |
| "Q1/Q2" | False | a slash |
| "Rev[2]" | False | square brackets |
| "Time 10:30" | False | a colon |
| "What?" | False | a question mark |
| "'Draft" | False | an apostrophe first |
| "Bob's" | True | an apostrophe inside is fine |
| "history" | False | History is kept by Excel, in any case |
| 2026 | True | a number is its text |

##### LIST_FILES (VCL 0.4)

| Input | Result | Why |
| --- | --- | --- |
| ["b.pdf", "A.PDF", "c.txt"], "*.pdf" | ["A.PDF", "b.pdf"] | case ignored, then name order |
| ["x.pdfa", "x.pdf"], "*.pdf" | ["x.pdf"] | the whole name: *.pdf is not x.pdfa |
| ["r1.pdf", "r10.pdf", "r2.pdf"], "r?.pdf" | ["r1.pdf", "r2.pdf"] | ? is exactly one character |
| ["r1.pdf", "r10.pdf", "r2.pdf"], "r*.pdf" | ["r1.pdf", "r10.pdf", "r2.pdf"] | * is any run; the order is by character, not by number |
| ["a.pdf", "A.pdf"], "*" | ["A.pdf", "a.pdf"] | the same name in two cases: as written breaks the tie |
| ["ab[1].txt", "ab1.txt"], "ab[1].txt" | ["ab[1].txt"] | [ ] are themselves, not a set |
| ["a.pdf"], "*.txt" | [] | nothing matches: an empty list |
| [], "*" | [] | an empty folder |

##### READ_TEXT (VCL 0.4)

| Input | Result | Why |
| --- | --- | --- |
| "J-101 Shell" | "J-101 Shell" | one line, as it is |
| "J-101\x0d\n" | "J-101" | one final CR LF is dropped |
| "J-101\n" | "J-101" | one final LF is dropped |
| "a\x0d\nb\x0d\n\x0d\n" | "a\x0d\nb\x0d\n" | only one: the inner and the other line breaks stay |
| "a\tb" | "a\tb" | a tab is text |
| "" | "" | an empty file |
| "café" | (empty) | an accented letter is not text: it reads differently with each code page |
| "x\x00y" | (empty) | a NUL byte is not text |

##### Text for READ_TEXT and WRITE_TEXT (VCL 0.4)

| Input | Result | Why |
| --- | --- | --- |
| "done: J-101" | True | printable ASCII |
| "a\x0d\nb" | True | line breaks |
| "€ 5" | False | the euro sign is not ASCII |

## Appendix F: ForgeLib, the VBA runtime (`vba/ForgeLib.bas`)

```vba
' ===== ForgeLib: VCL runtime for Excel 2016 VBA (generated code calls these) =====
' Every function here implements one decision in docs/vcl-semantics.md. Keep them in sync.

Private gForgeLog As Collection
Private gForgeFso As Object

' VCL 0.3 limits: the same numbers as vcl/semantics.py and compiler/vcl_interp.py (a unit test checks)
Private Const FORGE_MAX_NAME As Long = 100
Private Const FORGE_MAX_FILE_PATH As Long = 258
Private Const FORGE_MAX_FOLDER_PATH As Long = 247
Private Const FORGE_MAX_TEXT As Long = 32767   ' S25: the most characters a cell holds
Private Const FORGE_NAME_FORBIDDEN As String = "\/:*?""<>|"

Public Sub ForgeLog(ByVal procName As String, ByVal errNumber As Long, ByVal message As String, ByVal lineNo As Long)
    If gForgeLog Is Nothing Then Set gForgeLog = New Collection
    gForgeLog.Add procName & " | error " & errNumber & " | spec line " & (lineNo \ 10) & " | " & message
    Debug.Print "ForgeLog: " & procName & " | " & errNumber & " | " & message
End Sub

Public Sub ForgeInfo(ByVal procName As String, ByVal message As String)
    Debug.Print "ForgeInfo: " & procName & " | " & message
End Sub

Public Function ForgeErrorCount() As Long
    If gForgeLog Is Nothing Then ForgeErrorCount = 0 Else ForgeErrorCount = gForgeLog.Count
End Function

Public Sub ForgeClearLog()
    Set gForgeLog = New Collection
End Sub

Public Function ForgeLogText() As String
    Dim item As Variant
    Dim s As String
    If gForgeLog Is Nothing Then Exit Function
    For Each item In gForgeLog
        s = s & item & vbCrLf
    Next item
    ForgeLogText = s
End Function

Public Sub ForgeRestore(ByVal screenOn As Boolean, ByVal calcMode As Long, ByVal eventsOn As Boolean)
    On Error Resume Next
    Application.ScreenUpdating = screenOn
    Application.Calculation = calcMode
    Application.EnableEvents = eventsOn
    On Error GoTo 0
End Sub

' ---- S1 to S4: emptiness, numbers, dates ----

Private Function ForgeIsSpaceChar(ByVal code As Long) As Boolean
    ForgeIsSpaceChar = (code = 32 Or code = 160 Or code = 9 Or code = 10 Or code = 13)
End Function

Public Function ForgeIsEmpty(ByVal v As Variant) As Boolean
    Dim i As Long
    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then
        ForgeIsEmpty = True
        Exit Function
    End If
    If VarType(v) <> vbString Then Exit Function
    For i = 1 To Len(v)
        If Not ForgeIsSpaceChar(AscW(Mid$(v, i, 1))) Then Exit Function
    Next i
    ForgeIsEmpty = True
End Function

Private Function ForgeIsPlainNumberText(ByVal s As String) As Boolean
    Dim i As Long
    Dim ch As String
    Dim digits As Long
    Dim dots As Long
    Dim start As Long
    If Len(s) = 0 Then Exit Function
    start = 1
    If Left$(s, 1) = "+" Or Left$(s, 1) = "-" Then start = 2
    If start > Len(s) Then Exit Function
    For i = start To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits + 1
        ElseIf ch = "." Then
            dots = dots + 1
            If dots > 1 Then Exit Function
        Else
            Exit Function
        End If
    Next i
    If digits = 0 Then Exit Function
    If Right$(s, 1) = "." Then Exit Function
    ForgeIsPlainNumberText = True
End Function

Public Function ForgeIsNumber(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            ForgeIsNumber = True
        Case vbString
            ForgeIsNumber = ForgeIsPlainNumberText(ForgeTrim(v))
        Case Else
            ForgeIsNumber = False
    End Select
End Function

Public Function ForgeNum(ByVal v As Variant) As Double
    ' Only call after ForgeIsNumber is true. Val always uses "." as the decimal point.
    If VarType(v) = vbString Then
        ForgeNum = Val(ForgeTrim(v))
    Else
        ForgeNum = CDbl(v)
    End If
End Function

Public Function ForgeIsDate(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    ForgeIsDate = (VarType(v) = vbDate)
End Function

' ---- S6: canonical text ----

Private Function ForgeErrorText(ByVal v As Variant) As String
    If v = CVErr(xlErrNA) Then
        ForgeErrorText = "#N/A"
    ElseIf v = CVErr(xlErrDiv0) Then
        ForgeErrorText = "#DIV/0!"
    ElseIf v = CVErr(xlErrValue) Then
        ForgeErrorText = "#VALUE!"
    ElseIf v = CVErr(xlErrRef) Then
        ForgeErrorText = "#REF!"
    ElseIf v = CVErr(xlErrName) Then
        ForgeErrorText = "#NAME?"
    ElseIf v = CVErr(xlErrNum) Then
        ForgeErrorText = "#NUM!"
    ElseIf v = CVErr(xlErrNull) Then
        ForgeErrorText = "#NULL!"
    Else
        ForgeErrorText = "#ERROR"
    End If
End Function

Private Function ForgeTwo(ByVal n As Long) As String
    ForgeTwo = Right$("0" & CStr(n), 2)
End Function

Private Function ForgeNumberText(ByVal d As Double) As String
    Dim s As String
    If d = Int(d) And Abs(d) < 1E+15 Then
        ForgeNumberText = Format$(d, "0")
        Exit Function
    End If
    s = Trim$(Str$(d))
    If Left$(s, 1) = "." Then s = "0" & s
    If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
    ForgeNumberText = s
End Function

Public Function ForgeText(ByVal v As Variant) As String
    Dim d As Date
    If IsError(v) Then
        ForgeText = ForgeErrorText(v)
        Exit Function
    End If
    Select Case VarType(v)
        Case vbEmpty, vbNull
            ForgeText = ""
        Case vbString
            ForgeText = v
        Case vbBoolean
            If v Then ForgeText = "TRUE" Else ForgeText = "FALSE"
        Case vbDate
            d = v
            ForgeText = CStr(Year(d)) & "-" & ForgeTwo(Month(d)) & "-" & ForgeTwo(Day(d))
            If d <> Int(d) Then
                ForgeText = ForgeText & " " & ForgeTwo(Hour(d)) & ":" & ForgeTwo(Minute(d)) & ":" & ForgeTwo(Second(d))
            End If
        Case vbInteger, vbLong, vbByte
            ForgeText = CStr(v)
        Case vbSingle, vbDouble, vbCurrency, vbDecimal
            ForgeText = ForgeNumberText(CDbl(v))
        Case Else
            ForgeText = CStr(v)
    End Select
End Function

' ---- S2: TRIM (also removes non-breaking spaces, tabs and line breaks) ----

Public Function ForgeTrim(ByVal v As Variant) As String
    Dim s As String
    Dim a As Long
    Dim b As Long
    s = ForgeText(v)
    a = 1
    b = Len(s)
    Do While a <= b
        If Not ForgeIsSpaceChar(AscW(Mid$(s, a, 1))) Then Exit Do
        a = a + 1
    Loop
    Do While b >= a
        If Not ForgeIsSpaceChar(AscW(Mid$(s, b, 1))) Then Exit Do
        b = b - 1
    Loop
    If b >= a Then ForgeTrim = Mid$(s, a, b - a + 1) Else ForgeTrim = ""
End Function

Public Function ForgeLeft(ByVal v As Variant, ByVal n As Double) As String
    If n < 0 Then n = 0
    ForgeLeft = Left$(ForgeText(v), CLng(Int(n)))
End Function

Public Function ForgeRight(ByVal v As Variant, ByVal n As Double) As String
    If n < 0 Then n = 0
    ForgeRight = Right$(ForgeText(v), CLng(Int(n)))
End Function

' ---- S7: = between a value and text compares canonical text ----

Public Function ForgeEq(ByVal a As Variant, ByVal b As Variant) As Boolean
    If VarType(a) = vbString Or VarType(b) = vbString Or IsEmpty(a) Or IsEmpty(b) Then
        ForgeEq = (StrComp(ForgeText(a), ForgeText(b), vbBinaryCompare) = 0)
    ElseIf IsError(a) Or IsError(b) Then
        ForgeEq = (ForgeText(a) = ForgeText(b))
    Else
        ForgeEq = (a = b)
    End If
End Function

' ---- VCL 0.2 value functions: S5 and S12 to S16 ----

Public Function ForgeIsError(ByVal v As Variant) As Boolean
    ForgeIsError = IsError(v)
End Function

Public Function ForgeSame(ByVal a As Variant, ByVal b As Variant) As Boolean
    ' Binary compare of the lower-case texts: vbTextCompare would follow the PC's locale.
    ForgeSame = (StrComp(LCase$(ForgeTrim(a)), LCase$(ForgeTrim(b)), vbBinaryCompare) = 0)
End Function

Public Function ForgeCleanText(ByVal v As Variant) As String
    Dim s As String
    Dim ch As String
    Dim code As Long
    Dim i As Long
    Dim pending As Boolean
    Dim out As String
    s = ForgeText(v)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If ForgeIsSpaceChar(code) Then
            pending = True
        ElseIf code >= 32 Or code < 0 Then
            ' AscW is below 0 for characters above &H7FFF, which are kept
            If pending And Len(out) > 0 Then out = out & " "
            pending = False
            out = out & ch
        End If
    Next i
    ForgeCleanText = out
End Function

Private Function ForgeAllDigits(ByVal s As String, ByVal fewest As Long, ByVal most As Long) As Boolean
    Dim i As Long
    Dim ch As String
    If Len(s) < fewest Or Len(s) > most Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    ForgeAllDigits = True
End Function

Public Function ForgeParseDate(ByVal v As Variant, ByVal fmt As String) As Variant
    ' Never CDate or DateValue: those read the text by the PC's regional settings.
    Dim sep As String
    Dim parts As Variant
    Dim names As Variant
    Dim i As Long
    Dim dd As String
    Dim mm As String
    Dim yyyy As String
    Dim result As Date
    ForgeParseDate = Empty
    If IsError(v) Then Exit Function
    If VarType(v) = vbDate Then
        ForgeParseDate = DateSerial(Year(v), Month(v), Day(v))
        Exit Function
    End If
    If VarType(v) <> vbString Then Exit Function
    If Left$(fmt, 4) = "yyyy" Then sep = Mid$(fmt, 5, 1) Else sep = Mid$(fmt, 3, 1)
    parts = Split(ForgeTrim(v), sep)
    names = Split(fmt, sep)
    If UBound(parts) <> 2 Or UBound(names) <> 2 Then Exit Function
    For i = 0 To 2
        If names(i) = "dd" Then
            dd = parts(i)
        ElseIf names(i) = "mm" Then
            mm = parts(i)
        ElseIf names(i) = "yyyy" Then
            yyyy = parts(i)
        End If
    Next i
    If Not ForgeAllDigits(dd, 1, 2) Then Exit Function
    If Not ForgeAllDigits(mm, 1, 2) Then Exit Function
    If Not ForgeAllDigits(yyyy, 4, 4) Then Exit Function
    If CLng(yyyy) < 1900 Then Exit Function
    If CLng(mm) < 1 Or CLng(mm) > 12 Or CLng(dd) < 1 Then Exit Function
    ' DateSerial rolls 31/02 over into March, so the day must come back unchanged
    result = DateSerial(CLng(yyyy), CLng(mm), CLng(dd))
    If Day(result) <> CLng(dd) Then Exit Function
    ForgeParseDate = result
End Function

Public Function ForgeIif(ByVal cond As Boolean, ByVal a As Variant, ByVal b As Variant) As Variant
    ' Both values were worked out by the caller already, as with VBA's IIf.
    If cond Then ForgeIif = a Else ForgeIif = b
End Function

Public Function ForgeMod(ByVal a As Double, ByVal b As Double) As Double
    ' Excel's MOD: the sign of b, decimals kept. VBA's Mod rounds first. b = 0 raises error 11.
    ForgeMod = a - b * Int(a / b)
End Function

' ---- S9: COL matches headers ignoring case, extra spaces and line breaks ----

Public Function ForgeHeaderKey(ByVal v As Variant) As String
    Dim s As String
    s = ForgeText(v)
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    ForgeHeaderKey = LCase$(Trim$(s))
End Function

Public Function ForgeCol(ByVal ws As Worksheet, ByVal header As Variant) As Long
    Dim c As Long
    Dim lastCol As Long
    Dim want As String
    If ws Is Nothing Then Exit Function
    want = ForgeHeaderKey(header)
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If ForgeHeaderKey(ws.Cells(1, c).Value) = want Then
            ForgeCol = c
            Exit Function
        End If
    Next c
End Function

' ---- S11: last row / last column like End(xlUp) ----

Public Function ForgeLastRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    ' End(xlUp) skips rows a FILTER hid (D42), so the rows below it, up to the end of the used range,
    ' are checked too: the answer is the same whether rows are hidden or not.
    Dim last As Long
    Dim r As Long
    last = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    For r = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1 To last + 1 Step -1
        If Not IsEmpty(ws.Cells(r, col).Value) Then
            ForgeLastRow = r
            Exit Function
        End If
    Next r
    ForgeLastRow = last
End Function

Public Function ForgeLastCol(ByVal ws As Worksheet, ByVal rowNo As Long) As Long
    ForgeLastCol = ws.Cells(rowNo, ws.Columns.Count).End(xlToLeft).Column
End Function

' ---- ranges ----

Public Function ForgeUsedRange(ByVal ws As Worksheet) As Range
    Set ForgeUsedRange = ws.Range(ws.Cells(1, 1), ws.Cells(ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1, _
        ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1))
End Function

Public Function ForgeRowRange(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal c1 As Long, ByVal c2 As Long) As Range
    Set ForgeRowRange = ws.Range(ws.Cells(rowNo, c1), ws.Cells(rowNo, c2))
End Function

' ---- S17: SET_VALUE, WRITE_ROW and COPY_ROW store text as exactly that text ----

Public Function ForgeNeedsTextGuard(ByVal s As String) As Boolean
    ' True when Excel could turn the text into a number, date, TRUE/FALSE, error or formula.
    Dim i As Long
    Dim code As Long
    Dim t As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        code = AscW(Mid$(s, i, 1))
        If (code >= 48 And code <= 57) Or code > 127 Or code < 0 Then
            ForgeNeedsTextGuard = True
            Exit Function
        End If
    Next i
    If InStr(1, "=+-@'#", Left$(s, 1), vbBinaryCompare) > 0 Then
        ForgeNeedsTextGuard = True
        Exit Function
    End If
    t = UCase$(ForgeTrim(s))
    ForgeNeedsTextGuard = (t = "TRUE" Or t = "FALSE")
End Function

Public Sub ForgeSetValue(ByVal target As Range, ByVal v As Variant)
    ' The ' prefix is hidden by Excel: the cell holds the text unchanged.
    If VarType(v) = vbString Then
        If ForgeNeedsTextGuard(v) Then
            target.Value = "'" & v
            Exit Sub
        End If
    End If
    target.Value = v
End Sub

Public Sub ForgeWriteRow(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal firstCol As Long, ByVal values As Variant)
    Dim i As Long
    For i = LBound(values) To UBound(values)
        ForgeSetValue ws.Cells(rowNo, firstCol + i - LBound(values)), values(i)
    Next i
End Sub

Public Sub ForgeCopyRow(ByVal src As Worksheet, ByVal srcRow As Long, ByVal dst As Worksheet, ByVal dstRow As Long, ByVal nCols As Long)
    Dim c As Long
    For c = 1 To nCols
        ForgeSetValue dst.Cells(dstRow, c), src.Cells(srcRow, c).Value
    Next c
End Sub

Public Sub ForgeNoFill(ByVal rng As Range)
    rng.Interior.Pattern = xlNone
End Sub

Public Sub ForgeSort(ByVal rng As Range, ByVal keyCol As Long, ByVal ascending As Boolean, ByVal hasHeader As Boolean)
    Dim minRows As Long
    If hasHeader Then minRows = 3 Else minRows = 2
    If rng.Rows.Count < minRows Then Exit Sub
    ' Every option is given (D51): Range.Sort reuses the sheet's last Orientation, OrderCustom and
    ' MatchCase when they are left out, so a sheet once sorted left to right would sort columns.
    Dim ord As Long
    If ascending Then ord = xlAscending Else ord = xlDescending
    rng.Sort Key1:=rng.Worksheet.Cells(rng.Row, keyCol), Order1:=ord, Header:=IIf(hasHeader, xlYes, xlNo), _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
End Sub

Public Sub ForgeWriteArray(ByVal arr As Variant, ByVal topLeft As Range)
    ' S17 also for WRITE_ARRAY (D36): text Excel would convert gets the hidden ' prefix, in a copy.
    Dim out As Variant
    Dim r As Long
    Dim c As Long
    out = arr
    For r = LBound(out, 1) To UBound(out, 1)
        For c = LBound(out, 2) To UBound(out, 2)
            If VarType(out(r, c)) = vbString Then
                If ForgeNeedsTextGuard(out(r, c)) Then out(r, c) = "'" & out(r, c)
            End If
        Next c
    Next r
    topLeft.Resize(UBound(out, 1) - LBound(out, 1) + 1, UBound(out, 2) - LBound(out, 2) + 1).Value = out
End Sub

Public Function ForgeReadArray(ByVal rng As Range) As Variant
    ' READ_ARRAY (D36): always a 2-D array from 1, even for one cell, whose .Value is not an array.
    Dim one(1 To 1, 1 To 1) As Variant
    If rng.Cells.CountLarge = 1 Then
        one(1, 1) = rng.Value
        ForgeReadArray = one
    Else
        ForgeReadArray = rng.Value
    End If
End Function

Public Sub ForgeSetFormula(ByVal rng As Range, ByVal f As String, ByVal r1c1 As Boolean)
    ' SET_FORMULA (D43): Excel's own English form, .Formula or .FormulaR1C1 and never .FormulaLocal,
    ' so the same text means the same formula on every PC. The validator checked f against its list.
    ' A cell formatted as Text (@) would keep the formula as text: refused before anything is written.
    Dim c As Range
    For Each c In rng.Cells
        If c.NumberFormat = "@" Then Err.Raise 1004, "ForgeSetFormula", _
            "SET_FORMULA: " & c.Address(False, False) & " is formatted as Text, so Excel would keep the formula as text"
    Next c
    If r1c1 Then
        rng.FormulaR1C1 = f
    Else
        rng.Formula = f
    End If
End Sub

Public Sub ForgeFilter(ByVal rng As Range, ByVal col As Long, ByVal crit As Variant)
    ' FILTER (D42): below the range's first row (its header), a row stays visible when its cell in sheet
    ' column col is SAME as crit, and is hidden otherwise. Never AutoFilter, whose criteria read numbers
    ' and dates through the PC's settings and treat * and ? as wildcards; it never deletes a row.
    Dim ws As Worksheet
    Dim r As Long
    Set ws = rng.Worksheet
    For r = rng.Row + 1 To rng.Row + rng.Rows.Count - 1
        ws.Rows(r).Hidden = Not ForgeSame(ws.Cells(r, col).Value, crit)
    Next r
End Sub

Public Sub ForgeInsertRow(ByVal ws As Worksheet, ByVal rowNo As Long)
    ' INSERT_ROW (D36): a blank row with no fill and no number format, whatever the row above holds.
    ws.Rows(rowNo).Insert
    ws.Rows(rowNo).ClearFormats
    ws.Rows(rowNo).Hidden = False   ' never hidden like a filtered row next to it (D42)
End Sub

Public Function ForgeFind(ByVal rng As Range, ByVal x As Variant, ByVal whole As Boolean) As Range
    ' FIND (S22): the first cell, row by row, that is SAME as x (whole) or whose text holds TRIM(x),
    ' case ignored (part); Nothing when none does. It loops itself: Range.Find keeps LookIn, LookAt
    ' and MatchCase from the last search, even one made by hand in Excel.
    Dim cell As Range
    Dim want As String
    want = LCase$(ForgeTrim(x))
    If Not whole And Len(want) = 0 Then Exit Function
    For Each cell In rng.Cells
        If whole Then
            If ForgeSame(cell.Value, x) Then
                Set ForgeFind = cell
                Exit Function
            End If
        ElseIf InStr(1, LCase$(ForgeText(cell.Value)), want, vbBinaryCompare) > 0 Then
            Set ForgeFind = cell
            Exit Function
        End If
    Next cell
End Function

Public Function ForgeFileName(ByVal path As Variant) As String
    ' FILE_NAME (S20): the text after the last backslash; all of it when there is none.
    Dim s As String
    s = ForgeText(path)
    ForgeFileName = Mid$(s, InStrRev(s, "\") + 1)
End Function

' ---- S10: dictionaries with canonical-text keys, case ignored ----

Public Function ForgeNewDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set ForgeNewDict = d
End Function

Public Sub ForgeDictAdd(ByVal d As Object, ByVal key As Variant, ByVal value As Variant, ByVal mode As String)
    Dim k As String
    k = ForgeText(key)
    Select Case mode
        Case "overwrite"
            d(k) = value
        Case "keep"
            If Not d.Exists(k) Then d.Add k, value
        Case "sum"
            If d.Exists(k) Then
                d(k) = d(k) + ForgeNum(value)
            Else
                d.Add k, ForgeNum(value)
            End If
        Case "count"
            If d.Exists(k) Then
                d(k) = d(k) + 1
            Else
                d.Add k, 1
            End If
    End Select
End Sub

Public Function ForgeDictGet(ByVal d As Object, ByVal key As Variant) As Variant
    Dim k As String
    k = ForgeText(key)
    If d.Exists(k) Then ForgeDictGet = d(k) Else ForgeDictGet = Empty
End Function

Public Function ForgeDictExists(ByVal d As Object, ByVal key As Variant) As Boolean
    ForgeDictExists = d.Exists(ForgeText(key))
End Function

' ---- S21: DEDUPE keys (VEP 0.4 Phase 2, D38) ----

Public Function ForgeRowKey(ByVal values As Variant) As Variant
    ' The key of a row: each key value as a dictionary key (S10), its length first so the columns
    ' stay apart. Null when a key cell is empty or an error: such a row is never a duplicate.
    Dim i As Long
    Dim k As String
    Dim out As String
    For i = LBound(values) To UBound(values)
        If ForgeIsEmpty(values(i)) Or ForgeIsError(values(i)) Then
            ForgeRowKey = Null
            Exit Function
        End If
        k = LCase$(ForgeText(values(i)))   ' compared exactly below: vbTextCompare follows the PC's locale (D48)
        out = out & Len(k) & ":" & k
    Next i
    ForgeRowKey = out
End Function

Public Function ForgeSameRowKey(ByVal a As Variant, ByVal b As Variant) As Boolean
    Dim ka As Variant
    Dim kb As Variant
    ka = ForgeRowKey(a)
    kb = ForgeRowKey(b)
    If IsNull(ka) Or IsNull(kb) Then Exit Function
    ForgeSameRowKey = (StrComp(ka, kb, vbBinaryCompare) = 0)
End Function

Public Function ForgeDedupe(ByVal rng As Range, ByVal cols As Variant, ByVal hasHeader As Boolean) As Long
    ' DEDUPE: keeps the first row of each key; each later row with that key is deleted as a whole
    ' sheet row, bottom up, so the rows below move up. cols are sheet column numbers, as SORT key= is.
    ' It gives the number of rows it deleted (DEDUPE into=, VCL 0.5, D55).
    Dim seen As Object
    Dim doomed() As Boolean
    Dim parts() As Variant
    Dim key As Variant
    Dim first As Long
    Dim r As Long
    Dim i As Long
    Set seen = ForgeNewDict()
    seen.CompareMode = vbBinaryCompare   ' the keys are already lower case (ForgeRowKey, D48)
    If hasHeader Then first = 2 Else first = 1
    If rng.Rows.Count < first Then Exit Function
    ReDim doomed(first To rng.Rows.Count)
    ReDim parts(LBound(cols) To UBound(cols))
    For r = first To rng.Rows.Count
        For i = LBound(cols) To UBound(cols)
            parts(i) = rng.Worksheet.Cells(rng.Row + r - 1, cols(i)).Value
        Next i
        key = ForgeRowKey(parts)
        If Not IsNull(key) Then
            If seen.Exists(key) Then doomed(r) = True Else seen.Add key, True
        End If
    Next r
    For r = rng.Rows.Count To first Step -1
        If doomed(r) Then
            rng.Worksheet.Rows(rng.Row + r - 1).Delete
            ForgeDedupe = ForgeDedupe + 1
        End If
    Next r
End Function

' ---- VCL 0.4: the sheets of a workbook (S23, VEP 0.4 Phase 4, D40) ----

Public Function ForgeSheetNameOk(ByVal v As Variant) As Boolean
    ' S23: 1 to 31 characters, no spaces at the ends, none of \ / ? * [ ] :, no apostrophe first or
    ' last, no character below 32, and not History (Excel keeps that name for itself).
    Dim s As String
    Dim i As Long
    s = ForgeText(v)
    If Len(s) < 1 Or Len(s) > 31 Then Exit Function
    If ForgeTrim(s) <> s Then Exit Function
    If Left$(s, 1) = "'" Or Right$(s, 1) = "'" Then Exit Function
    If LCase$(s) = "history" Then Exit Function
    For i = 1 To Len(s)
        If InStr(1, "\/?*[]:", Mid$(s, i, 1), vbBinaryCompare) > 0 Then Exit Function
        If (AscW(Mid$(s, i, 1)) And &HFFFF&) < 32 Then Exit Function   ' AscW is below 0 above &H7FFF
    Next i
    ForgeSheetNameOk = True
End Function

Public Function ForgeSheetExists(ByVal book As Workbook, ByVal name As Variant) As Boolean
    ' SHEET_EXISTS: a sheet of that name, case ignored as Excel does. A loop, so no error handling.
    Dim sh As Object
    Dim s As String
    s = LCase$(ForgeText(name))
    For Each sh In book.Sheets
        If LCase$(sh.Name) = s Then
            ForgeSheetExists = True
            Exit Function
        End If
    Next sh
End Function

Public Function ForgeAddSheet(ByVal book As Workbook, ByVal name As Variant, ByVal after As Object) As Worksheet
    ' ADD_SHEET: a new empty sheet with this name, after the given sheet or else after the last one.
    ' A bad name, or one already used, stops the run before anything is added: never a replacement.
    Dim s As String
    Dim ws As Worksheet
    s = ForgeText(name)
    If Not ForgeSheetNameOk(s) Then Err.Raise 1004, "ForgeAddSheet", "ADD_SHEET: '" & s & "' is not a valid sheet name (S23)"
    If ForgeSheetExists(book, s) Then Err.Raise 1004, "ForgeAddSheet", "ADD_SHEET: a sheet named '" & s & "' already exists"
    If after Is Nothing Then
        Set after = book.Sheets(book.Sheets.Count)
    ElseIf Not after.Parent Is book Then
        Err.Raise 1004, "ForgeAddSheet", "ADD_SHEET: after= must be a sheet of the same workbook"
    End If
    Set ws = book.Worksheets.Add(After:=after)
    ws.Name = s
    Set ForgeAddSheet = ws
End Function

Public Function ForgeCopySheet(ByVal src As Worksheet, ByVal name As Variant, ByVal after As Object) As Worksheet
    ' COPY_SHEET: a copy of src in its own workbook, with this name, after the given sheet or else after
    ' the last one. A bad name, or one already used, stops the run before anything is copied.
    Dim book As Workbook
    Dim s As String
    Set book = src.Parent
    s = ForgeText(name)
    If Not ForgeSheetNameOk(s) Then Err.Raise 1004, "ForgeCopySheet", "COPY_SHEET: '" & s & "' is not a valid sheet name (S23)"
    If ForgeSheetExists(book, s) Then Err.Raise 1004, "ForgeCopySheet", "COPY_SHEET: a sheet named '" & s & "' already exists"
    If after Is Nothing Then
        Set after = book.Sheets(book.Sheets.Count)
    ElseIf Not after.Parent Is book Then
        ' Copy After:= a sheet of another workbook puts the copy there, and the rename below would then
        ' rename a sheet that was already here: refused before anything is copied (D48)
        Err.Raise 1004, "ForgeCopySheet", "COPY_SHEET: after= must be a sheet of the copied sheet's own workbook"
    End If
    src.Copy After:=after
    Set ForgeCopySheet = book.Sheets(after.Index + 1)
    ForgeCopySheet.Name = s
End Function

' ---- VCL 0.3: S18 SAFE_NAME, S19 PATH_JOIN, FOLDER_EXISTS and the file ops (E11) ----
' One attempt, then a status: the file ops never raise, overwrite, delete or retry.

Private Function ForgeFso() As Object
    ' Late-bound, so no References are needed. Made once per Excel session.
    If gForgeFso Is Nothing Then Set gForgeFso = CreateObject("Scripting.FileSystemObject")
    Set ForgeFso = gForgeFso
End Function

Private Function ForgeNoTrailingSlash(ByVal s As String) As String
    Do While Right$(s, 1) = "\"
        s = Left$(s, Len(s) - 1)
    Loop
    ForgeNoTrailingSlash = s
End Function

Private Function ForgeHasWildcard(ByVal s As String) As Boolean
    ' FSO reads * and ? as patterns, so a path with one is never used.
    ForgeHasWildcard = (InStr(1, s, "*", vbBinaryCompare) > 0 Or InStr(1, s, "?", vbBinaryCompare) > 0)
End Function

Private Function ForgeSamePath(ByVal a As String, ByVal b As String) As Boolean
    ' Paths compare ignoring case, as on Windows.
    ForgeSamePath = (StrComp(LCase$(a), LCase$(b), vbBinaryCompare) = 0)
End Function

Public Function ForgeSafeName(ByVal v As Variant) As String
    Dim s As String
    Dim stem As String
    Dim i As Long
    s = ForgeCleanText(v)
    If Len(s) = 0 Or Len(s) > FORGE_MAX_NAME Then Exit Function
    For i = 1 To Len(s)
        If InStr(1, FORGE_NAME_FORBIDDEN, Mid$(s, i, 1), vbBinaryCompare) > 0 Then Exit Function
    Next i
    ' A trailing dot also covers "." and "..".
    If Right$(s, 1) = "." Then Exit Function
    ' Device names are reserved with any extension and any case: con.txt is CON.
    stem = s
    i = InStr(1, stem, ".", vbBinaryCompare)
    If i > 0 Then stem = Left$(stem, i - 1)
    stem = UCase$(RTrim$(stem))
    If stem = "CON" Or stem = "PRN" Or stem = "AUX" Or stem = "NUL" Then Exit Function
    If Len(stem) = 4 And (Left$(stem, 3) = "COM" Or Left$(stem, 3) = "LPT") Then
        If AscW(Right$(stem, 1)) >= 49 And AscW(Right$(stem, 1)) <= 57 Then Exit Function
    End If
    ForgeSafeName = s
End Function

Public Function ForgePathJoin(ByVal root As Variant, ByVal names As Variant) As String
    ' names is Array(part, ...). An empty or unsafe part gives "", so it never becomes a path.
    Dim s As String
    Dim part As String
    Dim i As Long
    s = ForgeNoTrailingSlash(ForgeText(root))
    If Len(s) = 0 Or UBound(names) < LBound(names) Then Exit Function
    For i = LBound(names) To UBound(names)
        part = ForgeText(names(i))
        If Len(part) = 0 Then Exit Function
        If StrComp(ForgeSafeName(part), part, vbBinaryCompare) <> 0 Then Exit Function
        s = s & "\" & part
    Next i
    ForgePathJoin = s
End Function

Public Function ForgeFolderExists(ByVal path As String) As Boolean
    ' TRUE only for a folder written in full, as in the interpreter's mock. A relative path, or one
    ' Windows would rewrite (a trailing \, a /), depends on the PC, so it gives FALSE.
    On Error GoTo FailedOp
    If Len(path) = 0 Then Exit Function
    If Not ForgeSamePath(ForgeFso().GetAbsolutePathName(path), path) Then Exit Function
    ForgeFolderExists = ForgeFso().FolderExists(path)
    Exit Function
FailedOp:
    ForgeLog "ForgeFolderExists", Err.Number, Err.Description & ": " & path, 0
End Function

Private Function ForgeBelowRoot(ByVal root As String, ByVal path As String) As Boolean
    ' The run-time root check (docs/vep-0.3-file-ops.md, item 3): root is written in full, and
    ' path starts with root and \, as written and as Windows reads it. PATH_JOIN's safe names
    ' keep every generated path below its root; this catches a path built any other way.
    Dim r As String
    r = ForgeNoTrailingSlash(root)
    If Len(r) = 0 Then Exit Function
    If Not ForgeSamePath(ForgeNoTrailingSlash(ForgeFso().GetAbsolutePathName(root)), r) Then Exit Function
    If Not ForgeSamePath(Left$(path, Len(r) + 1), r & "\") Then Exit Function
    ForgeBelowRoot = ForgeSamePath(Left$(ForgeFso().GetAbsolutePathName(path), Len(r) + 1), r & "\")
End Function

Public Function ForgeMakeFolder(ByVal root As String, ByVal path As String) As String
    ' MAKE_FOLDER: path and every missing folder between it and root, never root itself. Every
    ' level is checked before any folder is made, so a blocked path changes nothing.
    Dim fso As Object
    Dim r As String
    Dim parts As Variant
    Dim cur As String
    Dim blocked As Boolean
    Dim i As Long
    On Error GoTo FailedOp
    Set fso = ForgeFso()
    r = ForgeNoTrailingSlash(root)
    If Len(path) = 0 Or Len(r) = 0 Or ForgeHasWildcard(path) Then
        ForgeMakeFolder = "Bad name"
    ElseIf Len(path) > FORGE_MAX_FOLDER_PATH Then
        ForgeMakeFolder = "Too long"
    ElseIf fso.FolderExists(path) Then
        ForgeMakeFolder = "Exists"
    ElseIf fso.FileExists(path) Then
        ForgeMakeFolder = "Blocked by file"
    ElseIf Not ForgeBelowRoot(root, path) Then
        ForgeMakeFolder = "Bad name"
    ElseIf Not fso.FolderExists(r) Then
        ForgeMakeFolder = "No folder"
    Else
        parts = Split(Mid$(path, Len(r) + 2), "\")
        cur = r
        For i = LBound(parts) To UBound(parts)
            cur = cur & "\" & parts(i)
            If fso.FileExists(cur) Then blocked = True
        Next i
        If blocked Then
            ForgeMakeFolder = "Blocked by file"
        Else
            cur = r
            For i = LBound(parts) To UBound(parts)
                cur = cur & "\" & parts(i)
                If Not fso.FolderExists(cur) Then fso.CreateFolder cur
            Next i
            ForgeMakeFolder = "Created"
        End If
    End If
    Exit Function
FailedOp:
    ForgeLog "ForgeMakeFolder", Err.Number, Err.Description & ": " & path, 0
    ForgeMakeFolder = "Failed"
End Function

Public Function ForgeNameMatches(ByVal name As String, ByVal pattern As String) As Boolean
    ' S24: * is any run of characters (also none), ? is exactly one, every other character is itself,
    ' case ignored, against the whole name. Not VBA's Like, which reads [ ] and # as well, and not Dir,
    ' which also matches Windows short names.
    Dim n As String
    Dim p As String
    Dim i As Long
    Dim j As Long
    Dim star As Long
    Dim mark As Long
    n = LCase$(name)
    p = LCase$(pattern)
    i = 1
    j = 1
    Do While i <= Len(n)
        If Mid$(p, j, 1) = "*" Then
            star = j
            mark = i
            j = j + 1
        ElseIf j <= Len(p) And (Mid$(p, j, 1) = "?" Or Mid$(p, j, 1) = Mid$(n, i, 1)) Then
            i = i + 1
            j = j + 1
        ElseIf star > 0 Then
            j = star + 1
            mark = mark + 1
            i = mark
        Else
            Exit Function
        End If
    Loop
    Do While Mid$(p, j, 1) = "*"
        j = j + 1
    Loop
    ForgeNameMatches = (j > Len(p))
End Function

Private Sub ForgeAddInOrder(ByVal names As Collection, ByVal name As String)
    ' S24 order: by LCase, then as written, both compared by character code, so it is the same on every PC.
    Dim k As Long
    Dim c As Long
    For k = 1 To names.Count
        c = StrComp(LCase$(name), LCase$(names(k)), vbBinaryCompare)
        If c = 0 Then c = StrComp(name, names(k), vbBinaryCompare)
        If c < 0 Then
            names.Add name, Before:=k
            Exit Sub
        End If
    Next k
    names.Add name
End Sub

Public Function ForgeMatchNames(ByVal names As Variant, ByVal pattern As String) As Collection
    ' S24 on a list of names: the ones that match, in S24 order.
    Dim i As Long
    Set ForgeMatchNames = New Collection
    For i = LBound(names) To UBound(names)
        If ForgeNameMatches(names(i), pattern) Then ForgeAddInOrder ForgeMatchNames, names(i)
    Next i
End Function

Public Function ForgeListFiles(ByVal root As String, ByVal folder As String, ByVal pattern As String) As Collection
    ' LIST_FILES (D44): the names of the files directly in folder that match pattern (S24), in S24 order.
    ' A bad, missing or too long folder gives an empty list: listing never raises and changes nothing.
    Dim f As Object
    Set ForgeListFiles = New Collection
    On Error GoTo FailedOp
    If Len(folder) = 0 Or Len(folder) > FORGE_MAX_FOLDER_PATH Or ForgeHasWildcard(folder) Then Exit Function
    If Not ForgeSamePath(folder, ForgeNoTrailingSlash(root)) Then
        If Not ForgeBelowRoot(root, folder) Then Exit Function
    End If
    If Not ForgeFolderExists(folder) Then Exit Function
    For Each f In ForgeFso().GetFolder(folder).Files
        ' Hidden and system files are left out (D51): Excel's ~$ owner file of an open workbook, Thumbs.db.
        If (f.Attributes And 6) = 0 Then
            If ForgeNameMatches(f.Name, pattern) Then ForgeAddInOrder ForgeListFiles, f.Name
        End If
    Next f
    Exit Function
FailedOp:
    ForgeLog "ForgeListFiles", Err.Number, Err.Description & ": " & folder, 0
    Set ForgeListFiles = New Collection
End Function

Public Function ForgeIsTextOk(ByVal v As Variant) As Boolean
    ' S25: printable ASCII (32 to 126), tabs and line breaks, at most 32,767 characters. Any other
    ' character reads or writes differently with each PC's code page.
    Dim s As String
    Dim i As Long
    Dim code As Long
    s = ForgeText(v)
    If Len(s) > FORGE_MAX_TEXT Then Exit Function
    For i = 1 To Len(s)
        code = AscW(Mid$(s, i, 1))
        If (code < 32 Or code > 126) And code <> 9 And code <> 10 And code <> 13 Then Exit Function
    Next i
    ForgeIsTextOk = True
End Function

Public Function ForgeFileText(ByVal raw As Variant) As Variant
    ' S25: what READ_TEXT gives for a file's contents: one final line break (CR LF, or LF) dropped;
    ' Null when it isn't text.
    Dim s As String
    s = ForgeText(raw)
    If Not ForgeIsTextOk(s) Then
        ForgeFileText = Null
    ElseIf Right$(s, 2) = vbCrLf Then
        ForgeFileText = Left$(s, Len(s) - 2)
    ElseIf Right$(s, 1) = vbLf Then
        ForgeFileText = Left$(s, Len(s) - 1)
    Else
        ForgeFileText = s
    End If
End Function

Private Function ForgeIsAsciiPath(ByVal path As String) As Boolean
    ' READ_TEXT and WRITE_TEXT (D48): VBA's Open reads a path through the PC's code page, so a path
    ' with any character outside printable ASCII is a Bad name, the same on every PC.
    Dim i As Long
    Dim code As Long
    For i = 1 To Len(path)
        code = AscW(Mid$(path, i, 1)) And &HFFFF&
        If code < 32 Or code > 126 Then Exit Function
    Next i
    ForgeIsAsciiPath = True
End Function

Private Function ForgeReadRaw(ByVal path As String) As String
    ' A file's bytes, read only and shared with any other reader. An error (70 when another program
    ' holds the file) goes on to the caller, which turns it into a status; the file is closed first.
    Dim f As Long
    Dim raw As String
    Dim errNo As Long
    Dim errText As String
    On Error GoTo FailedRead
    f = FreeFile
    Open path For Binary Access Read Shared As #f
    raw = Space$(LOF(f))
    Get #f, , raw
    Close #f
    ForgeReadRaw = raw
    Exit Function
FailedRead:
    errNo = Err.Number
    errText = Err.Description
    Close #f
    Err.Raise errNo, "ForgeReadRaw", errText
End Function

Public Function ForgeReadText(ByVal root As String, ByVal path As String, ByRef text As String) As String
    ' READ_TEXT (D45): one read, never a change. text gets the file's text (S25), or "" unless the
    ' status is Read: a missing, locked or binary file is a status, never an error.
    Dim fso As Object
    Dim got As Variant
    Dim reading As Boolean
    text = ""
    On Error GoTo FailedOp
    Set fso = ForgeFso()
    If Len(path) = 0 Or ForgeHasWildcard(path) Or Not ForgeIsAsciiPath(path) Then
        ForgeReadText = "Bad name"
    ElseIf Len(path) > FORGE_MAX_FILE_PATH Then
        ForgeReadText = "Too long"
    ElseIf Not ForgeBelowRoot(root, path) Then
        ForgeReadText = "Bad name"
    ElseIf Not fso.FileExists(path) Then
        ForgeReadText = "Missing"
    ElseIf fso.GetFile(path).Size > FORGE_MAX_TEXT Then
        ForgeReadText = "Too big"
    Else
        reading = True
        got = ForgeFileText(ForgeReadRaw(path))
        reading = False
        If IsNull(got) Then
            ForgeReadText = "Not text"
        Else
            text = got
            ForgeReadText = "Read"
        End If
    End If
    Exit Function
FailedOp:
    ' Error 70 (permission denied) while reading: another program holds the file.
    If reading And Err.Number = 70 Then
        ForgeReadText = "Locked"
    Else
        ForgeLog "ForgeReadText", Err.Number, Err.Description & ": " & path, 0
        ForgeReadText = "Failed"
    End If
End Function

Public Function ForgeWriteText(ByVal root As String, ByVal path As String, ByVal text As String) As String
    ' WRITE_TEXT (D45): a new file holding exactly text, with no line break added; never over anything.
    Dim fso As Object
    Dim ts As Object
    Dim slash As Long
    Dim stage As Long   ' 1 while reading the file that is there, 2 while writing the new one
    On Error GoTo FailedOp
    Set fso = ForgeFso()
    slash = InStrRev(path, "\")
    If Len(path) = 0 Or ForgeHasWildcard(path) Or Not ForgeIsAsciiPath(path) Then
        ForgeWriteText = "Bad name"
    ElseIf Len(path) > FORGE_MAX_FILE_PATH Then
        ForgeWriteText = "Too long"
    ElseIf Not ForgeBelowRoot(root, path) Then
        ForgeWriteText = "Bad name"
    ElseIf Not ForgeIsTextOk(text) Then
        ForgeWriteText = "Not text"
    ElseIf slash = 0 Then
        ForgeWriteText = "No folder"
    ElseIf Not fso.FolderExists(Left$(path, slash - 1)) Then
        ForgeWriteText = "No folder"
    ElseIf fso.FolderExists(path) Then
        ForgeWriteText = "Conflict"
    ElseIf fso.FileExists(path) Then
        If fso.GetFile(path).Size <> Len(text) Then
            ForgeWriteText = "Conflict"
        Else
            stage = 1
            If ForgeReadRaw(path) = text Then ForgeWriteText = "Already there" Else ForgeWriteText = "Conflict"
            stage = 0
        End If
    Else
        ' overwrite is False: a file that appears at path meanwhile raises an error, never replaced
        stage = 2
        Set ts = fso.CreateTextFile(path, False, False)
        ts.Write text
        ts.Close
        stage = 0
        If fso.GetFile(path).Size = Len(text) Then ForgeWriteText = "Written" Else ForgeWriteText = "Check write"
    End If
    Exit Function
FailedOp:
    ' Error 70 (permission denied) while reading what is there: another program holds that file.
    If stage = 1 And Err.Number = 70 Then
        ForgeWriteText = "Locked"
    Else
        ForgeLog "ForgeWriteText", Err.Number, Err.Description & ": " & path, 0
        ForgeWriteText = "Failed"
    End If
End Function

Public Function ForgeCopyFile(ByVal srcRoot As String, ByVal src As String, ByVal dstRoot As String, ByVal dst As String) As String
    ' COPY_FILE: one copy, never over anything at dst; src always stays, and a short copy is kept.
    Dim fso As Object
    Dim slash As Long
    Dim copying As Boolean
    On Error GoTo FailedOp
    Set fso = ForgeFso()
    slash = InStrRev(dst, "\")
    If Len(src) = 0 Or Len(dst) = 0 Or ForgeHasWildcard(src & dst) Then
        ForgeCopyFile = "Bad name"
    ElseIf Len(src) > FORGE_MAX_FILE_PATH Or Len(dst) > FORGE_MAX_FILE_PATH Then
        ForgeCopyFile = "Too long"
    ElseIf Not ForgeBelowRoot(srcRoot, src) Or Not ForgeBelowRoot(dstRoot, dst) Then
        ForgeCopyFile = "Bad name"
    ElseIf Not fso.FileExists(src) Then
        ForgeCopyFile = "Missing"
    ElseIf slash = 0 Then
        ForgeCopyFile = "No folder"
    ElseIf Not fso.FolderExists(Left$(dst, slash - 1)) Then
        ForgeCopyFile = "No folder"
    ElseIf fso.FileExists(dst) Then
        If fso.GetFile(dst).Size = fso.GetFile(src).Size Then ForgeCopyFile = "Already there" Else ForgeCopyFile = "Conflict"
    ElseIf fso.FolderExists(dst) Then
        ForgeCopyFile = "Conflict"
    Else
        ' overwrite is False: a file that appears at dst meanwhile raises an error, never replaced
        copying = True
        fso.CopyFile src, dst, False
        copying = False
        If fso.GetFile(dst).Size = fso.GetFile(src).Size Then ForgeCopyFile = "Copied" Else ForgeCopyFile = "Check copy"
    End If
    Exit Function
FailedOp:
    ' Error 70 (permission denied) during the copy: another program has src open.
    If copying And Err.Number = 70 Then
        ForgeCopyFile = "Locked"
    Else
        ForgeLog "ForgeCopyFile", Err.Number, Err.Description & ": " & src & " -> " & dst, 0
        ForgeCopyFile = "Failed"
    End If
End Function
```

## Appendix G: the brief, the plans and the language designs

CLAUDE.md (the brief Claude Code reads first), the plans, then every VEP, newest first.

### `CLAUDE.md`

#### VBA Forge: project brief for Claude Code

Read this first in every session. It is the standing brief; `docs/VBA_Forge_Project_Roadmap.docx`
has the full plan and `docs/vcl-mvp-plan.md` the current build sequence.

##### What this is

VBA Forge is an offline, VBA-only code generator for Excel automation at KGC Engineering Projects
(pressure-vessel fabrication paper trail: BOM, weld log, NDE, NCR, calibration and similar
registers). Claude writes task specs in VCL (VBA Command Language); VBA Forge turns each spec into
VBA and proves the macro works with tests. From March 2027 our own model, trained from zero on our
data (D25), will write VCL specs from plain English; the compiler stays the only thing that writes
VBA (D31, `docs/ai-plan.md`). Until then, the goal is a verified dataset of tasks and their proven
specs, each with its VBA.

##### Working policy

- **I don't test, debug or verify anything.** You verify everything automatically (unit tests,
  the `vclc` pipeline, CI) and keep going without waiting for me.
- **One PR per item.** When a PR's CI is green, merge it yourself with a merge commit, then start
  the next item.
- If something needs my decision, pick the more conservative option, record it in
  `DECISIONS.md`, and carry on.
- Never weaken a test, a check or a threshold to get to green.
- Running the generated VBA in Excel is optional, happens only when I choose, and never blocks
  work.

##### Working with me

- Call me Boss. I'm learning, so explain in plain words and keep it brief.
- Do one thing at a time and keep changes small.
- When I must do something by hand (for example update the Orchestrator Project on claude.ai),
  give exact clicks and what I should see afterwards.
- Ask only when truly necessary. Otherwise state your assumption and continue.
- When I ask for "baby steps": one action per reply, written as "🍼 Next action" plus
  "Expected:", then wait for my result.

##### Hard constraints until March 2027

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

##### Sources of truth

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

##### Repo layout

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
                     generator), vba_lint.py, vcl_mutate.py, gen_semantics_doc.py, gen_context.py,
                     vcl_formula.py (the SET_FORMULA check)
vba/                 ForgeLib.bas, the VBA runtime every generated procedure calls
dist/                vclc build output; only forge_module.bas, forge_selftest.bas and
                     forge_pilot.bas are committed (builds/<NAME>.bas: one module per BUILD)
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

##### Rules for code in this repo

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

##### How a spec is verified (no Excel needed)

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

##### Workflow for every change

1. Make one small change.
2. Verify it: the checks above, for what you touched.
3. Read your own diff.
4. Commit with a clear message that says what changed and why.
5. Push and open a pull request. When CI is green, merge it yourself with a merge commit, then
   start the next item.

##### Status and next steps

Phase 0, Foundation (Oct 14 to 31, 2026). Gate 0: all specs validate, type-check, pass their
TESTs in the reference interpreter and compile to lint-clean VBA, in CI.

Done: 12 specs (11 in VCL 0.2, task_12 in 0.3), VCL 0.4 in spec.md and the Orchestrator prompt (D46), the parser and validator, and the
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

Next, in order (`docs/vcl-mvp-plan.md` calls them items 4, 5, 5b and 6-7):
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
   and mutation scores a BUILD. Phase 4a is done (D55): ROW_OF(<range>) and DEDUPE `into=`; the
   reference teaches them at the switch to 0.5, and until then the drift guard wants them in the
   VEP. Phase 4b is done (D56): a header the spec writes to row 1 counts as checked for rule 14.
   Phase 4c is done (D57): an assertion names a sheet with spaces in quotes, `'Open NCRs'!B1`.
   **Next: Phase 5** (a proven BUILD and the switch to 0.5; Boss updates the Orchestrator once).
4. After the MVP: the AI plan, `docs/ai-plan.md` (D31, approved by Boss on 19 Oct 2026): the
   model writes VCL, never VBA. Its Stage A (a `GOAL` line, a formatter, error codes, the corpus,
   the data generator and the benchmark) starts when Boss says go for it, as every earlier
   item did (D51); no training before March 2027.

Update this section at the end of every session, and keep `docs/kgc-kronos-context.md` (the
context file Boss gives to other Claude chats) in step with it: edit its sections 1 to 12 by hand,
then run `python compiler/gen_context.py`, which rewrites its appendices from the repo (D54).

### `docs/vcl-mvp-plan.md`

#### VBA Forge MVP plan (revised 20 Oct 2026)

**Policy:** Boss has no testing, debugging or verification steps. Every check is automated: the
grammar, the type checker, the reference interpreter, mutation scoring, the VBA linter and CI.
Running the generated VBA in Excel is optional and never blocks work.

##### MVP: what it is and where it stands

`python compiler/vclc.py all specs/*.vcl -o dist` turns VCL specs into:

- `dist/forge_module.bas`: one paste-ready module, the ForgeLib runtime plus one procedure per spec
- `dist/procedures/*.bas`: each procedure on its own
- `dist/expected/*.json`: for every TEST, the expected RESULT, ERROR_COUNT and final workbook state,
  so any future Excel run can be compared automatically

**Status: working** for all 12 specs (task_06 uses the provable rewrite; task_12 is the router) and
Examples 1 and 3. Numbers as of 20 Oct 2026 (`docs/kgc-kronos-context.md` section 8 is kept
current):

| Automated check | Result |
| --- | --- |
| Grammar (LALR, proven unambiguous) | 12 of 12 specs parse |
| Type checker (every calculation and date comparison guarded) | 12 of 12 pass |
| Interpreter: the specs' own TESTs | 46 of 46 pass |
| Mutation score, TEST assertions only | 89.0% of 438 deliberate bugs caught (was 68% at the start) |
| Mutation score, whole-workbook comparison | 95.4% (418 of 438); 19 of the 20 survivors are equivalent mutants (D14, D24), 1 is a small task_11 fixture gap (D24) |
| VBA lint (the compile errors Excel would raise) | 0 problems in 1,850 lines |
| Unit tests | 403 pass |

##### Honest limits

- **The generated VBA is never executed.** Lint catches compile-time errors, not runtime
  behaviour. The main unproven risk is ForgeLib's behaviour in real Excel, for example comparing
  error values or Range.Sort edge cases.
- **One fixture generator** (item 1): `compiler/vcl_fixtures.py` builds the pilot's data
  bit for bit (`DECISIONS.md` D1).
- **The compiler and the interpreter cover every op and function of VCL 0.4** (item 5 below
  made them provable), except OPEN_BOOK, SAVE_AS and CLOSE_BOOK, which open and save real
  workbooks and are refused until March 2027 (D36), and CALL, which compiles but whose
  procedures the interpreter can't run. BUILD specs wait for their own item. `NOTE` steps go to
  Claude or Boss until VCL covers them (D31).
- **EXIT_IF was undefined in revision 3.** The MVP treats it as "leave the procedure". VCL 0.2
  adds EXIT_LOOP_IF for loops.

##### What the mutation scores teach (and the fix)

Every surviving mutant is missing test data, not missing assertions:

| Survivor | Missing test data | Fix (VEP) |
| --- | --- | --- |
| Deleting "clear the output sheet" goes unnoticed | output sheets start empty | fixtures pre-fill output sheets with stale rows |
| `mode=keep` vs `overwrite` goes unnoticed | no duplicate keys | `+dupes(n)` modifier |
| `<` vs `<=` goes unnoticed | no rows exactly on the boundary | `+boundary` rows (equal values) |
| task_10's 15-day window goes unnoticed | fixed dates don't move with TODAY | relative dates, `reldates(-30, 30)` |

Target after these: at least 95% with whole-workbook comparison. **Done in item 2 (VCL 0.1
revision 5):** `stale=<n>`, `+dupes(n)`, `+boundary(...)` and `reldates(lo, hi)` are in the
language, rule 19 warns when a spec doesn't use them, and the whole-workbook score reached
96.4% for the 10 specs of that time (95.4% for today's 12, in the table above).

##### Build sequence (Claude Code works through it without waiting; merges itself when CI is green)

CLAUDE.md, and Boss, number the open work differently: its item 1 (the router) is item 4 here,
its item 2 (the rest of revision 3) is item 5, and its item 3 (the AI plan's Stage A) is items 6
and 7.

1. **Integrate the MVP** into the repo. Unify the two fixture generators, apply the task_06
   rewrite, record the decisions, and make CI run the whole `vclc all` pipeline.
2. **Upgrade fixtures from the mutation evidence** (the table above), until the snapshot score
   is at least 95%.
3. **VCL 0.2 core, end to end:** grammar, types, semantics, interpreter, compiler, ForgeLib and
   conformance cases for SAME, SKIP_IF, EXIT_LOOP_IF, WRITE_ROW, MODIFIES with automatic
   UNCHANGED, PARSE_DATE, IS_ERROR and CLEAN_TEXT (plus IIF and MOD). Then migrate the 10 specs
   with a migration tool. Design: `docs/vep-0.2-core.md` (four PRs, A to D).
   **PR A (the six value functions), PR B (SKIP_IF, EXIT_LOOP_IF, WRITE_ROW, ForgeSetValue),
   PR C (MODIFIES with UNCHANGED), PR D1 (the 10 specs migrated) and PR D2 (spec.md and the
   Orchestrator prompt) are done: VCL 0.2 core is complete.**
4. **The Job-Pack Router MVP (VCL 0.3 file ops), approved 17 Oct 2026:** `docs/vep-0.3-file-ops.md`,
   Phases 0 to 5. Phases 0 to 3 (the compiler, ForgeLib and the sandbox builder: D29, D30) are
   done, and Phases 4 and 5 too (task_12 RouteJobPacks, the switch to VCL 0.3, the Excel pilot
   kit: D32 to D35). **The MVP is complete.** It came before the rest of item 5.
5. **The rest of rev 3** in the compiler and interpreter (DEDUPE with `cols=[...]`, FILTER,
   SET_FORMULA, sheet and file ops), plus list syntax. Designed in `docs/vep-0.4-rev3-rest.md`
   (D36), seven phases; OPEN_BOOK, SAVE_AS and CLOSE_BOOK wait until March. Phase 1 (the
   interpreter catches up, D37), Phase 2 (list syntax and DEDUPE, D38), Phase 3 (FIND, D39)
   Phase 4 (ADD_SHEET and COPY_SHEET, D40, D41), Phase 5 (FILTER, D42), Phase 6 (SET_FORMULA,
   D43) and Phase 7 (LIST_FILES, READ_TEXT, WRITE_TEXT; D44, D45) are done, and so is the switch:
   spec.md and the prompt teach VCL 0.4 (E12, D46). **Item 5 is complete.** CALL and BUILD specs
   are their own item.
5b. **CALL and BUILD specs** (VEP 0.5, `docs/vep-0.5-build.md`, D49): a BUILD becomes provable as
   every op of item 5 did: CALL with `into=`, MODIFIES that follows the arguments, BUILD TESTs in
   the interpreter, one module per BUILD from the compiler, and the review gaps of D47. FORM and
   xlam wait. Five phases, then one switch to VCL 0.5.
6. **Corpus:** 150 tasks, Orchestrator-written specs, and an expressibility report that drives
   the next VCL packs.
7. **Dataset generator:** tasks in plain English with their proven specs (and the compiler's
   VBA), plus repair pairs. The mutation engine doubles as a contrast-pair generator: every
   valid mutant is a near-identical spec with different behaviour. That is exactly the "read
   carefully" training data the Phase 2 plan wanted. See `docs/ai-plan.md`, Stages A and B.
8. **March:** train our own model from zero on the in-house dataset (D25: no pre-trained weights,
   no outside data). It writes VCL, never VBA; the compiler writes all the VBA (D31,
   `docs/ai-plan.md`). Excel conformance runs whenever Boss chooses.

### `docs/ai-plan.md`

#### VBA Forge AI plan: the model writes VCL (D31)

Status: **approved by Boss on 19 Oct 2026**. This is the general direction for the model after
the Job-Pack Router MVP. It is not a build plan yet: each stage gets its own design and PRs
when it starts. Nothing here starts before the MVP is done, and no training starts before the
new PC arrives in March 2027.

##### The decision

**Our own model writes VCL, never VBA. The compiler stays the only thing that writes VBA.**

The model is still ours (D25): designed and trained from zero, on in-house data only.

Why:
- The compiler already turns VCL into VBA without mistakes. A model doing the same job could
  only be a worse copy of it.
- Every VCL spec the model writes can be proven without Excel: the validator, the type checker,
  the TESTs in the reference interpreter and the MODIFIES safety check. VBA written by a model
  can't be proven until it runs in Excel.
- So nothing unproven ever reaches Excel, and every VBA line that does is lint-clean compiler
  output that calls ForgeLib.

##### The model's three jobs

1. **Plain English to VCL:** an offline Orchestrator for everyday tasks.
2. **Repair:** a spec that fails a check, plus the error messages, in; a fixed spec out.
3. **NOTE steps** (what VCL can't say yet): the model never writes raw VBA for them. Instead
   VCL grows until they are rare. Until VCL covers a NOTE, the task goes to Claude or Boss.

##### The flow

```text
Boss types a task in plain English
  -> the model writes VCL (several tries)
  -> validator, type checker and TESTs keep only the tries that pass
  -> the compiler writes the VBA (ForgeLib calls, lint-clean)
  -> Boss pastes it into Excel
No try passes  -> repair mode, with the error messages
Still failing  -> the task goes to Claude (the Orchestrator) or Boss
```

##### How VCL's structure helps

| VCL feature | What it gives the model |
| --- | --- |
| A small vocabulary (a few dozen keywords) | A small model, trainable from zero on one PC |
| One step per line | Easy to learn, and every error points at one line |
| A formal grammar (`vcl/grammar.lark`) | Output can be forced to be valid VCL (grammar-constrained decoding) |
| The validator and type checker | Wrong outputs are rejected at once; every error is a free "error -> fix" training pair |
| The interpreter, synthetic fixtures and TESTs | Outputs are judged by "do the tests pass", not by how they look |
| The mutation engine | Many near-miss specs: training for careful reading and for repair |
| MODIFIES and UNCHANGED | An automatic safety check: the model can't quietly touch other sheets or files |
| The migration tool (`vcl_migrate.py`) | Old training data moves to each new VCL version, so the model learns one version |

##### VCL changes planned for this goal

Each is a language change, so each comes as its own VEP and PRs, with spec.md, the prompt and
every check updated together. None starts before the MVP is done.

1. **A `GOAL` line** in every TASK: one plain-English sentence saying what the task does. Every
   spec then carries its own English -> VCL training pair.
2. **One canonical way to write each spec**, with a formatter (`vclc fmt`): the same meaning
   always gives the same text, so there is less for the model to learn.
3. **Error codes** in validator and type checker messages (for example `V107 line 12: ...`).
   Stable codes make repairs learnable.
4. **Shrink NOTE:** count every NOTE step in the corpus (the expressibility report), and turn
   the common ones into VCL ops.
5. **Keep VCL small:** a few ops that combine well, and one way to do each thing.
6. **Strong tests required:** a model-written spec counts only when its own TESTs catch enough
   mutants (a minimum snapshot mutation score), so it can't pass by writing weak tests.

##### Where the training data comes from (all in-house, D25)

- **Orchestrator specs:** synthetic KGC-style tasks (BOM, weld log, NDE, NCR, calibration and
  similar registers), each with its English description and its proven VCL.
- **Paraphrases:** the same task written several ways, by Claude through our pipeline, so the
  model learns how Boss words things.
- **Repairs:** made automatically from mutants and the errors they cause, in large numbers.
- **A locked benchmark:** held-out tasks never used for training; the only scoreboard.

Every pair passes the whole toolchain before it goes into `data/verified/`. Synthetic data only
until March 2027 (CLAUDE.md).

Scale: the planned 150-task corpus is a seed, not a dataset. A useful model needs thousands of
proven specs; paraphrases and mutants help multiply them.

##### Stages after the MVP

| Stage | When | What |
| --- | --- | --- |
| A | Before March (cloud) | The `GOAL` line, the formatter and error codes; the 150-task corpus; the data generator; the benchmark |
| B | Before March (cloud) | Grow VCL from the NOTE count; grow the data to thousands of proven specs |
| C | March 2027 (new PC) | A tokenizer and a small model trained from zero: English -> VCL, with grammar-constrained output |
| D | After C | Add repair. The model is used only when it passes enough benchmark tasks; Claude stays the Orchestrator until then |
| E | Later | Daily use: Boss types a task and gets proven VBA |

Stages A and B fit today's rules (no training before March), and they make the data valuable
even if the model takes longer than planned.

##### What doesn't change

- D25: the model and everything it learns from are ours; no pre-trained weights, outside data
  or retrieval.
- The compiler and ForgeLib stay the core, and the reference for what VBA a spec means.
- Running anything in Excel stays optional and never blocks work.

### `docs/excel-pilot.md`

#### Excel pilot (optional): ForgeLib and the router in Excel 2016

Run this only when you choose; nothing waits for it. It takes about 15 minutes.

**What it proves:**
1. ForgeLib, the VBA runtime every generated macro calls, behaves in Excel 2016 the way our
   Python reference says (229 cases).
2. The router spec (task_12, RouteJobPacks) does in Excel what its TESTs say.

**Why it's safe:**
- It works in a new, empty workbook, with made-up data only.
- It makes its own folders below `%TEMP%`, in a new folder named `KGC_KRONOS_SANDBOX_<date>_<n>`,
  and never touches any other file or folder.
- It never deletes anything.

##### What you need

- Excel 2016 with macros allowed (you confirmed they run on your PC). No Trust Center changes.
- A browser with GitHub open on `4ptech1/KGC-KRONOS`, branch `main`.
- Three files from the repo:
  - `dist/forge_module.bas`: ForgeLib plus every generated macro
  - `dist/forge_selftest.bas`: the ForgeLib self-test
  - `dist/forge_pilot.bas`: the router pilot

##### Steps

**1. Make a new workbook.**
In Excel, click **File → New → Blank workbook**. Then **File → Save As → Browse**, set
**Save as type** to **Excel Macro-Enabled Workbook (\*.xlsm)**, name it `VBAForge_Pilot` and
click **Save**.
*Expected:* the title bar shows `VBAForge_Pilot.xlsm`.

**2. Open the VBA editor.** Press **Alt+F11**.
*Expected:* the "Microsoft Visual Basic for Applications" window opens, and the left panel shows
`VBAProject (VBAForge_Pilot.xlsm)`.
If nothing opens, your PC blocks the editor: tell Claude and stop here.

**3. Add the three modules, one at a time.** For each file, in this order: `forge_module.bas`,
then `forge_selftest.bas`, then `forge_pilot.bas`:
1. On GitHub, open the file (for example `dist/forge_module.bas`) and click **Copy raw file**
   (the two-squares icon at the top right of the file).
2. In the VBA editor, click `VBAProject (VBAForge_Pilot.xlsm)`, then **Insert → Module**. A
   white code window opens.
3. Click inside the code window, press **Ctrl+A** and **Delete** (this removes an
   `Option Explicit` line Excel may have added), then press **Ctrl+V**.

*Expected:* the left panel shows Module1, Module2 and Module3 under Modules.

**4. Compile.** In the VBA editor, click **Debug → Compile VBAProject**.
*Expected:* nothing happens, no message.
If a message appears, copy its text, note the highlighted line, send both to Claude and stop
here.

**5. Run the self-test.** Go back to Excel (click its window), press **Alt+F8**, click
**ForgeSelfTest**, then **Run**.
*Expected:*
- A message says `ForgeLib self-test: 229 of 229 cases pass.`
- A new sheet, ForgeSelfTest, has one row per case.

Click cell **H4** on that sheet, press **Ctrl+C**, and paste it to Claude. It looks like
`SELFTEST 229/229 FAILED: none`.

**6. Run the router pilot.** Press **Alt+F8**, click **ForgePilotRun**, then **Run**.
*Expected:*
- After a few seconds, a message says `Pilot done: <n> lines on the sheet ForgePilot.`
- A new sheet, ForgePilot, has the lines in column A, from `PILOT 1 spec(s)` down to `END`.

Click the **A** column header on that sheet, press **Ctrl+C**, and paste everything to Claude.

**7. Claude checks the results.** Claude saves your paste to a file and runs:
```
python compiler/vclc.py pilot-check <results file> specs/task_12.vcl
```
It compares every line with the reference interpreter, on the date Excel wrote, and tells you
whether they agree. Any difference is a bug that Claude fixes.

##### Afterwards

- Keep or delete the workbook; it holds only made-up data.
- If Excel asks about macros when you reopen it, click **Enable Content**.
- The sandbox folders stay in `%TEMP%`, because VBA Forge never deletes anything. To remove them
  yourself:
  1. Press **Windows key + R**, type `%TEMP%` and press **Enter**.
  2. Delete the folders whose names start with `KGC_KRONOS_SANDBOX_`.

### `docs/vep-0.5-build.md`

#### VEP 0.5: BUILD specs and CALL (the VEP 0.4 group D)

Status: **design, 21 Oct 2026**. Boss said "continue building the next part according to your
recommendations" after item 2; the recommendation was CALL and BUILD specs before the AI plan's
Stage A, because the Orchestrator's rule 9 already asks for a BUILD when a task needs several
procedures or more than about 60 lines, and today no BUILD can be proven. Choices that needed
Boss's decision were made the conservative way and recorded in `DECISIONS.md` D49; Boss can
overrule any of them. Phases 1 to 4 are done (E13, E14, D50; D52; D53; D55, D56, D57).

##### The problem

A BUILD is several TASK procedures that make one tool: `modMain.RunHeadless` runs the whole tool
and is what TESTs call, `modMain.Run` is what the user starts, and CALL joins the procedures. The
validator has checked BUILDs since revision 3 (keys, SIGs, one TASK per SIG, CALL targets and
argument counts, the TESTs), but nothing runs or compiles them: `vclc test` and `vclc build` stop
with "only TASK specs". So a valid BUILD is not a provable BUILD, which is the gap item 2 closed
for the single ops. It matters for the corpus too: the larger register tasks (task_04 has 48 OP
lines) need several procedures.

What BUILD and CALL lack today:

| Part | Today |
| --- | --- |
| CALL | compiles to `Call Proc(...)`; the interpreter can't run it; a Function's result is lost |
| BUILD TESTs | validated, never run |
| MODIFIES in a BUILD | not required in its TASKs, so no UNCHANGED check is possible |
| BUILD output | none: `vclc build` refuses BUILDs |
| FORM, form events | in the language, but no op can show a form or read a control |

##### Principles (unchanged)

- The interpreter is the reference meaning; every meaning goes through ForgeLib or plain
  generated VBA that the compiler proves; a TEST proves it on synthetic data.
- MODIFIES covers every change, and UNCHANGED checks it after every TEST.
- One language version per spec; the new parts need `VCL 0.5`; spec.md and the prompt move to
  0.5 in one step after the last phase, so Boss updates the Orchestrator once.

##### Design

**A BUILD is one tool in one workbook.** The compiler writes one module per BUILD,
`dist/builds/<NAME>.bas`: ForgeLib plus every procedure of the BUILD, so Boss pastes one module,
as today. MODULE lines stay as the spec's own grouping; procedure names must be unique within a
BUILD (VBA can't call two same-named procedures from different modules without the module name,
and the compiler calls them by name). Two BUILDs never share a module, so every BUILD can have its
own Run and RunHeadless.

**CALL runs another procedure of the same BUILD.**
- `CALL(<module>.<Proc>, <args>)` runs that TASK; `CALL(..., into=<var>)` (VCL 0.5) keeps a
  Function's result, so RunHeadless can add up what its steps did and return it (rule 17).
- Each procedure keeps its own ERR standard and PRE, as the compiled VBA does: a failed PRE in the
  callee logs once and gives -1, a run-time error in it logs once and gives -2, and the caller goes
  on. A caller checks `into=` (IF n < 0 ...) when a step must not be skipped silently.
- A Worksheet, Workbook or Range parameter is passed one of the caller's own `$arguments` of that
  type, written in place, so MODIFIES can follow the sheet from the TEST down to the step that
  changes it. Values (Long, Double, String, Boolean, Date) are any expression of their type.
- No recursion: the CALL graph must have no cycle (every loop has a bound, CLAUDE.md).
- CALL is only for procedures of the same BUILD; in a VCL 0.5 TASK spec it is an error.
- FLOW lists the calls in order for the reader; every `a -> b` in FLOW must be a CALL from a to
  b, and every CALL edge must be in a FLOW line.

**MODIFIES follows the arguments.** Every TASK in a 0.5 BUILD declares MODIFIES for its own
writes, as a TASK spec does. A CALL that passes `$x` to a parameter the callee modifies needs a
MODIFIES line on `$x` in the caller that allows the same change (the whole sheet allows any
change; `["H"]` lines must name the same headers; `fills`, `files` and `sheets` must match). So
RunHeadless's MODIFIES lines cover the whole tool, and every BUILD TEST gets the same UNCHANGED
check a TASK TEST gets.

**modMain.Run starts the tool.** With no form, Run is a Sub with no ARGS that calls RunHeadless
with the sheets of the workbook the tool lives in: `CALL(modMain.RunHeadless, SHEET(THIS_BOOK,
"Orders"), ..., into=n)` and `MSG(...)`. It may only CALL RunHeadless, LET, IF and MSG, so it has
nothing left to test; TESTs call RunHeadless, as today. Run is the one caller that passes sheets
by their real names, with `SHEET(THIS_BOOK, "<name>")` written in place, so it needs no MODIFIES
lines; the names are the real workbook's, as its ASSUME lines describe it.

**Deferred (D49):**
- **FORM and form events:** a UserForm can't be added by pasting text (its controls live in a
  binary .frx file, and building one from code needs the Trust Center change CLAUDE.md rules out),
  and no op shows a form or reads a control. FORM lines stay valid, the compiler refuses them with
  a clear message, and the reference text stops teaching them until a design for after March 2027.
- **TARGET xlam:** an add-in's THIS_BOOK is the add-in, not the data, and installing one may need
  settings Boss's PC doesn't allow. `xlsm` only for now; `xlam` is refused with a clear message.
- **BUILDs before 0.5** stay check-only: their TASKs don't declare MODIFIES, so no TEST could check
  that nothing else changed.

##### Also in 0.5: the gaps the item 2 review found (D47)

Small language gaps the Orchestrator dry runs hit; each gets a decision in its phase:
- No row number of a found cell, so FIND can't lead to another column of the same row (D39).
  **Done (Phase 4a, D55):** `ROW_OF(<range>)` is the sheet row of the range's first cell, so
  `CELL($ws, ROW_OF(hit), cStatus)` reaches another column of the row FIND found. ROW_OF of
  NOTHING is a run-time error (VBA error 91), so the spec tests `IS_NOTHING(hit)` first.
- No count of the rows DEDUPE removed (rule 17 wants a count). **Done (Phase 4a, D55):**
  `DEDUPE(..., into=<var>)` sets the VAR to the number of rows it deleted (0 when none).
- Rule 14 has no form for a header the macro itself adds to an existing sheet. **Done (Phase 4b,
  D56):** in VCL 0.5 a header the spec writes to row 1 of a sheet argument with quoted text
  (`SET_VALUE(CELL($ws, 1, c), "Checked")` or `WRITE_ROW($ws, 1, ...)`) counts as checked, so
  `COL($ws, "Checked")` needs no PRE line (PRE can't check a header that the first run adds).
  Its column can't be named in `MODIFIES $ws ["Checked"]`, because it isn't there before the run:
  the spec declares the whole sheet.
- An assertion can't name a sheet whose name has spaces. **Done (Phase 4c, D57):** in VCL 0.5 an
  assertion writes such a sheet in single quotes, as Excel does: `VALUE('Open NCRs'!B1)`,
  `COUNTA('Open NCRs'![NCR No])`, `ROWS_USED('Open NCRs')`, with `''` for a quote in the name.
  Quotes are only for a name that isn't one word (`Ncr!B2` stays as it is), and the name must be
  one ADD_SHEET could give (S23).

##### Phases (one or two PRs each, in this order)

1. **The language surface and the validator (done: E13, E14, D50):** the `VCL 0.5` line, CALL `into=`, unique procedure
   names, the CALL graph without cycles, CALL only in a BUILD, FLOW as the CALL edges, MODIFIES
   in every TASK of a BUILD and its propagation, the rules for Run, FORM and xlam.
2. **The interpreter (done: D52):** BUILD TESTs run RunHeadless with a procedure table; CALL with per-procedure
   PRE and ERR; UNCHANGED from RunHeadless's MODIFIES; `vclc test` on BUILDs.
3. **The compiler (done: D53):** `dist/builds/<NAME>.bas`, CALL `into=` as a Function call, the BUILD's
   fixture builders and expected results, lint, and the mutation scorer over every TASK of a BUILD.
4. **The review gaps** above, as the evidence decides. 4a (done: D55): ROW_OF and DEDUPE `into=`.
   4b (done: D56): a header the spec writes counts as checked for rule 14. 4c (done: D57): sheet
   names with spaces in assertions, in quotes.
5. **A proven BUILD and the switch:** a sample BUILD in the tests that catches every mutant, a new
   Example 4 in the prompt, rule 9 and a CALL rule in the hard rules, spec.md and the prompt move
   to 0.5, and Boss updates the Orchestrator once.

##### Verification (every phase)

The four standard checks and `vclc mutate --snapshot` (the score stays 95.4% or better), plus for
each new rule a validator test, for each new meaning an interpreter test, and for the compiler a
lint-clean module. Every 0.4 spec stays valid under 0.5 with the same results (as E8 and E12).

### `docs/vep-0.4-rev3-rest.md`

#### VEP 0.4: the rest of revision 3 (MVP plan, item 2)

Status: **done, 20 Oct 2026** (Boss said "go item 2"; designed and built the same day). Choices
that needed Boss's decision were made the conservative way and recorded in `DECISIONS.md` D36,
as CLAUDE.md asks; Boss can overrule any of them. Phase 1 changes no language; later phases
change VCL and follow the usual rule (spec.md and the prompt together, every check, one
Orchestrator update per switch). Phases 1 to 7 are done (D37 to D45), and so is the switch (E12,
D46): spec.md and the Orchestrator prompt teach VCL 0.4, so Boss updates the Orchestrator once.
**Item 2 is complete.**

##### The problem

Revision 3 (VCL 0.1) defined more ops and functions than the toolchain runs. The validator
accepts them, so a spec that uses them passes `check` and then fails at `test` or `build`.
That breaks the promise that a valid spec is a provable spec, and it limits the corpus (item 6)
and the dataset (item 7).

| Group | What | Today |
| --- | --- | --- |
| A | INSERT_ROW, READ_ARRAY, WRITE_ARRAY, ARR, ROWS, COLS, SELECT_CASE, LOG, MSG, RANGE, SHEET, FILE_NAME, FOR_EACH over SHEETS and CELLS | compiled, but the interpreter can't run them, so no TEST can prove them |
| B | DEDUPE, FIND and IS_NOTHING, ADD_SHEET, COPY_SHEET, SHEET_EXISTS, FILTER, SET_FORMULA, LIST_FILES, READ_TEXT, WRITE_TEXT, ITEMS | neither the interpreter nor the compiler |
| C | OPEN_BOOK, SAVE_AS, CLOSE_BOOK | they open and save real workbooks |
| D | CALL, BUILD specs, NOTE | need several procedures, or our own model (D31) |

##### Principles (unchanged)

- The interpreter is the reference meaning, every meaning goes through a ForgeLib helper, and a
  TEST proves it on synthetic data.
- Where Excel's own behaviour depends on the PC (regional settings, locale, criteria parsing),
  ForgeLib does the work itself in plain VBA, as it does for SORT keys and dictionary keys.
- Nothing touches a real workbook or a real folder until March 2027 (CLAUDE.md).
- MODIFIES covers every change, and UNCHANGED checks it after every TEST.

##### Phases (one or two PRs each, in this order)

###### Phase 1: the interpreter catches up (group A, no language change; done: D37)

- **Arrays:** READ_ARRAY gives a 2-D array, 1-based, as `Range.Value` does; ROWS, COLS and ARR
  read it; SET_VALUE(ARR(...), x) writes it; WRITE_ARRAY writes it from a top-left cell. S17
  extends to WRITE_ARRAY: text stays exactly that text, so ForgeWriteArray guards each text
  with the ' prefix, as ForgeSetValue does (Excel would otherwise turn "1001" into a number).
- **INSERT_ROW:** a new empty row, with no fill and no number format. Excel copies the format of
  the row above by default; ForgeInsertRow clears it, so the result never depends on the row
  above.
- **SELECT_CASE:** a CASE matches when `=` would be TRUE (S7: text compares as text), the
  first matching CASE runs, and CASE_ELSE runs when none does. The compiler stops emitting
  VBA's `Select Case`, which compares a number with "25.0" as numbers. It evaluates the
  selector once into a variable, then emits `Select Case True` with one ForgeEq-style test per
  value.
- **LOG and MSG:** LOG writes to ForgeInfo and never counts as an error. MSG shows a message in
  Excel; the interpreter records it, and a TEST can't see it.
- **RANGE(ws, "A1 ref"):** the interpreter reads a cell ("B2") or a block ("B2:D9"). Any other
  ref ("A:A", "1:1", a name) stops the TEST with a clear message; the validator learns the same
  rule with the next version line, because that changes the language.
- **SHEET(book, "name") and FOR_EACH over SHEETS(book):** THIS_BOOK and ACTIVE_BOOK are the
  TEST's fixture workbook, sheets in FIXTURE order; a missing sheet is a run-time error. The
  type checker gives the loop variable the Sheet type (today it gives a cell value).
- **FOR_EACH over CELLS(range):** row by row, left to right, as VBA does; the loop variable is
  the cell's value.
- **FILE_NAME(path):** the part after the last backslash.

###### Phase 2: list syntax and DEDUPE (done: S21, D38)

- `cols=[1, 2]` (G3 of `docs/vcl-m2-findings.md`); a single number stays valid.
- DEDUPE keeps the first row of each key and deletes the later ones, moving rows up. The key is
  the listed columns' values compared as dictionary keys (S10: as text, ignoring case), not
  Excel's RemoveDuplicates, whose comparison isn't documented. ForgeDedupe marks the rows top
  down and deletes them bottom up. It needs the whole sheet in MODIFIES (as today).

###### Phase 3: FIND (done: S22, D39)

- FIND(range, x, match=whole) is the first cell, row by row, whose value is SAME as x;
  `match=part` is the first whose text contains TRIM(x), ignoring case. FIND gives a Range or
  NOTHING, and IS_NOTHING tests it. ForgeFind loops over the cells instead of calling
  Range.Find, whose settings (LookIn, LookAt, MatchCase, the last search) carry over between
  calls.

###### Phase 4: sheets of the workbook (done: 4a ADD_SHEET and SHEET_EXISTS, S23, D40; 4b COPY_SHEET, D41)

- ADD_SHEET(book, name=, after=, into=) and COPY_SHEET(sheet, after=, into=), and the
  SHEET_EXISTS assertion. They change the workbook's sheet list, so they need a new MODIFIES
  form, `MODIFIES $book sheets`: new sheets may be added, and every existing sheet stays as it
  was. A name that is already used is an error, never a replacement.

###### Phase 5: FILTER (decision D36; done: D42)

- FILTER hides the rows of a range whose column doesn't match; it never deletes. The match is
  SAME (exact text, case and spaces at the ends ignored), done by ForgeFilter setting
  `Rows(r).Hidden` itself, not by AutoFilter criteria, which read "1,250" or dates through the
  PC's settings and treat `*` and `?` as wildcards. The interpreter keeps hidden rows; the
  snapshot compares them; a new assertion VISIBLE_COUNT(<Sheet>!<range>) counts the visible
  rows.

###### Phase 6: SET_FORMULA (decision D36; done: D43)

- The interpreter can't calculate Excel formulas, so a TEST can only check what was written:
  a new assertion FORMULA(<Sheet>!<cell>) = "<formula>" compares the formula text. The
  validator refuses Microsoft 365 functions (XLOOKUP, FILTER, LET, LAMBDA, UNIQUE, SORT,
  SEQUENCE, TEXTJOIN, IFS and the like), as CLAUDE.md requires. A formula's value is proven
  only by an Excel run (the pilot of router Phase 5).

###### Phase 7: text files in the sandbox (decision D36; done: 7a S24, D44; 7b S25, D45)

Two PRs (D44): 7a lists files, 7b reads and writes text.

- **7a (done): LIST_FILES and ITEMS.** `LIST_FILES(<folder>, "<pattern>", into=<Collection VAR>)`
  (VCL 0.4) gives the names of the files directly in the folder that match the pattern, in S24
  order; `FOR_EACH var=<name> in=ITEMS(<list>)` walks them as text, and COUNT counts them. The
  folder is the String root argument itself or PATH_JOIN of it (T4), and PRE checks the root with
  FOLDER_EXISTS (rule 14). The pattern is quoted: `*` and `?` only, against the whole name, case
  ignored (S24); ForgeNameMatches does it itself, not VBA's Like or Dir, which also matches
  Windows short names (`*.pdf` finds `x.pdfa`). A bad, missing or too long folder gives an empty
  list: listing never raises and changes nothing. The names are names, not paths, so a spec that
  copies a listed file builds its path with PATH_JOIN and SAFE_NAME, as T4 asks.
- **7b (done): READ_TEXT and WRITE_TEXT**, each with a status into `into=`, as MAKE_FOLDER
  and COPY_FILE have (E11), so a missing or locked file never stops a run halfway. No new option
  word is added: `status` would become a reserved word (G2), and three specs have a VAR of that
  name.
  - `READ_TEXT(<path>, <String VAR>, into=<status VAR>)`: the file's text goes into the String
    VAR ("" unless the status is `Read`). Statuses in check order: `Bad name`, `Too long`,
    `Missing`, `Too big` (more than 32,767 characters, what a cell holds), `Locked` (before Not
    text: a locked file can't be read to tell), `Not text` (anything but printable ASCII, tabs and
    line breaks: other bytes read differently with each PC's code page, S25), `Read`; in Excel also
    `Failed`. One final line break is dropped.
  - `WRITE_TEXT(<path>, <text>, into=<status VAR>)` writes a **new** file below a
    `MODIFIES $root files` root, never over one: `Bad name`, `Too long`, `Not text`, `No folder`,
    `Conflict` (a folder or a different file is there), `Locked` (a file of the same size that
    can't be read to compare), `Already there` (the same text is there), `Written`; in Excel also
    `Check write` and `Failed`. `mode=` is gone: append changes a file that is there.
  - Folder FIXTUREs gain file contents, `files=notes\a.txt:"text"`: one line of printable ASCII.
    The assertion FILE_TEXT("<path>") gives a file's exact text.

###### Deferred (decision D36)

- **OPEN_BOOK, SAVE_AS, CLOSE_BOOK** (group C): they open and save real workbooks, which CLAUDE.md
  rules out until March 2027. They stay valid in the language, the compiler refuses them with a
  clear message, and a sandbox design for them comes after March, for Boss to approve.
- **CALL and BUILD specs** (group D): their own item, after this one.
- **NOTE**: by design, VCL grows until NOTE is rare (D31).

##### Verification (every phase)

The four standard checks, `vclc mutate --snapshot`, and for each new meaning: interpreter unit
tests, a conformance case where it is a value function, a ForgeLib self-test case (D34), and
lint. A phase that changes VCL bumps the version line once all its phases are in (as 0.2 and
0.3 did), so Boss updates the Orchestrator once.

### `docs/vep-0.3-file-ops.md`

#### VEP 0.3: file ops and the Job-Pack Router MVP

Status: **design, approved by Boss on 17 Oct 2026**. Phases 0 (D22), 1 (D26, D27), 2 (D28),
3 (D29, D30), 4 (D32, D33) and 5 (D34, D35) are done: the MVP is complete. Nothing below is in the language
until its phase lands.

##### Context

Boss sent a brief (written with another AI) with three parts:
- **Phase 0 is unblocked:** add blank key rows to the FIXTURE lines, to close the 22 SKIP_IF mutation survivors (D20).
- **The Orchestrator is unblocked:** Boss is pasting the VCL 0.2 prompt into the claude.ai project himself.
- **A pivot:** build a single-task MVP, the "Job-Pack Router". It reads a tracker, makes job folders, files PDF test reports from a dump folder, and writes a status back. It needs new file ops, and it would prove ForgeLib in Boss's Excel 2016.

Boss's answers to my questions:
1. **Sandbox only.** It never runs on real folders before March 2027, so the hard constraint stays.
2. **Copy first.** COPY_FILE comes now; MOVE_FILE comes later.
3. **Macros can run** on his work PC.

**Acknowledged:** both backlog items are unblocked. Note that the Orchestrator will need a second update at the end of Phase 4, because the file ops add language.

##### Critique of the File Ops VEP, and the revisions

| # | Issue in the brief | Revision |
|---|---|---|
| 1 | `MOVE_FILE` deletes the source. On a crash, a locked file or a network drop mid-run, files can end up half-filed. | Ship **`COPY_FILE`** first: the source is never removed. MOVE_FILE is a later VEP, after the Excel proof, and will need `ALLOW file_delete` (rule 8). |
| 2 | `CLEAN_TEXT` is not path-safe. A Job ID like `..\..\Windows`, `A/B`, `CON`, `J-1.` or `J*1` would escape the folder or crash FSO. | New value function **`SAFE_NAME(x)`**. It applies CLEAN_TEXT, then returns `""` when the result has any of `\ / : * ? " < > \|` or a character below code 32. It also returns `""` for `.` or `..`, for a trailing dot or space, for a reserved name (CON, PRN, AUX, NUL, COM1-9, LPT1-9, ignoring case and extension), or for anything over 100 characters. |
| 3 | Nothing stops a path from reaching outside the intended folders. | New value function **`PATH_JOIN(base, name, ...)`**. The type checker gives paths a **Path** type (this is the "path text typing" item): only a String `$arg` or a PATH_JOIN of SAFE_NAME results is a Path, and the file ops accept only Paths. At run time, ForgeLib re-checks that every resolved path (`GetAbsolutePathName`) starts with its `$root` followed by `\`. |
| 4 | FSO raises errors for many ordinary cases. Under the standard handler, any of them aborts the whole run halfway. The errors are 70 (file open or locked), 53 (not found), 58 (already exists), 76 (path not found) and network drops. | Each op returns a **status text** into a VAR instead of raising. For COPY_FILE: `Copied`, `Already there` (destination exists with the same size), `Conflict` (destination exists with a different size, so never overwrite), `Missing`, `Locked`, `Too long` (259 characters or more), `Bad name`, `Failed` (with the error number in ForgeInfo). For MAKE_FOLDER: `Created`, `Exists`, `Blocked by file`, `Bad name`, `Failed`. ERROR_COUNT stays reserved for real crashes. |
| 5 | Overwriting a destination. | `CopyFile` is always called with `overwrite:=False`, after a `FileExists` check (race-safe). Overwrite is never offered, so no `ALLOW file_overwrite` is needed. |
| 6 | `MAKE_FOLDER` when the folder exists or its parent is missing. | It is idempotent (`Exists` is fine). It creates missing parents one level at a time, because FSO `CreateFolder` isn't recursive, but only below the root. The root itself must exist: a new value function **`FOLDER_EXISTS(path)`** is checked in PRE. A file with the same name gives `Blocked by file`. |
| 7 | Network-drive latency. | No sleep loops or retries (bounded loops; Ctrl+Break is the only watchdog). One attempt, then a status. Re-runs are safe (rule 16): a second run reports `Already there`. UNC paths (`\\server\share`) work; there are no drive-letter assumptions. |
| 8 | FSO treats `*` and `?` in the source of a CopyFile as wildcards. | SAFE_NAME rejects them, and ForgeLib refuses any source path that contains them. |
| 9 | A copy that silently comes out short. | After the copy, the sizes are compared. On a mismatch the status is `Check copy`. The partial file is never deleted (no Kill or DeleteFile). |
| 10 | How a PDF is matched to a job. | Driven by the tracker: each row names its `Job ID` and `Report File`. That is deterministic, auditable and testable. A pattern search (LIST_FILES) is a later option. |
| 11 | The brief's file list is off. `catalog.yaml` doesn't exist, and a new YAML dependency breaks "standard library first". The file-system mock doesn't belong in `vcl/semantics.py`. | The op catalogue is the existing tables (`validator.py` OPS and VALUE_FUNCS, `vcl_types.py` FUNCS). `semantics.py` gets only the pure functions (`safe_name`, `path_join`) and their conformance cases. The mock file system lives in `vcl_fixtures.py` and `vcl_interp.py`. |
| 12 | Nothing proves that "nothing else changed" on disk. | **`MODIFIES $jobs files`**: files and folders may be added below `$jobs`, and nothing else on disk may change. This extends UNCHANGED to the mock file system: existing files keep their bytes, and nothing is ever removed. |
| 13 | Version and sync. | The new features need **`VCL 0.3`** on the first line, as 0.2 did. The shared reference text and the prompt switch once, in Phase 4. |
| 14 | The spec's file name. | `specs/task_12.vcl` (PROC `RouteJobPacks`), not `job_router.vcl`, to follow the repo's naming. (Was task_11; Boss's weld-log spec took that number first, D24.) |
| 15 | "Prove ForgeLib in Excel" has no deliverable. | Add **Phase 5: the Excel pilot kit**. A sandbox builder, a ForgeLib self-test that runs the conformance cases in VBA, and a 5-step guide for Boss. It is optional and runs when Boss chooses. |

Excel and FSO notes: always `CreateObject("Scripting.FileSystemObject")`, late-bound with no References, which works in Excel 2016. If it is blocked, the ops return `Failed` and log it. The lint bans `DeleteFile`, `DeleteFolder`, `MoveFile` and `Kill`.

##### Phased roadmap

Each phase is one or two PRs: CI green, merge, then the next phase. Every PR runs the four standard checks plus `vclc mutate --snapshot`.

###### Phase 0: baseline cleanup (blank key rows)
- **Change:** the FIXTURE and TEST lines only, in task_01, 02, 04, 05, 06, 08 and 10. Add `+blank(n)` (or a `""` boundary value) to each column that a SKIP_IF tests, then recompute the exact TEST expectations with the interpreter (as D13 did).
- **Frozen copies:** apply the same FIXTURE and TEST lines to the copies in `tests/data/specs_v01/`, so `tests/test_migrate.py` still proves migrate(v01) equals the specs. The tool never touches those lines; re-run `vcl_migrate.py --prove`.
- **Rule 19 warning:** a spec with `SKIP_IF IS_EMPTY(x)` whose fixture column for x has no blank cell gets a warning.
- **Done when:** the SKIP_IF probes are killed and the snapshot score is at least 96% again. D22 is recorded in `DECISIONS.md`, and CLAUDE.md, the README and the plan are updated.

###### Phase 1: language surface (VCL 0.3, no run-time code yet)
Two PRs (D26). **1a (done):** the version line, the three value functions, the Name and Path types, and the meaning of SAFE_NAME and PATH_JOIN (moved here from Phase 2, because the type checker uses `safe_name` for quoted parts). **1b (done):** the ops, MODIFIES `files`, folder FIXTUREs and the validator rules (E9, E10, F6, D27).
- **Version line:** `VCL 0.3` in `parser.py`, `vcl_parse.py` and the validator gate, reusing the VCL_02 pattern.
- **Value functions:** SAFE_NAME(x), PATH_JOIN(base, name, ...) and FOLDER_EXISTS(path).
- **Ops:** `MAKE_FOLDER(<path>, into=<var>)` and `COPY_FILE(<from>, <to>, into=<var>)`. The status VAR must be a String.
- **Types (`vcl_types.py`):** Name and Path. A String ARG used as a base is a Path. PATH_JOIN takes a Path and then Names (SAFE_NAME results or safe literals) and returns a Path. The file ops take only Paths.
- **Grammar:** the MODIFIES `files` form, and a FIXTURE folder form: `FIXTURE <name> folder=<Name> files=<a.pdf>,... [locked=<a.pdf>,...]`. TEST calls pass folder names bare, as they do sheet names.
- **Validator:** every file op path is rooted in an ARG declared by MODIFIES `files` (the copy source may be any rooted Path); the status VAR is declared; a fixture with missing, conflict, locked or unsafe cases gets the rule 19 extension.
- **Docs:** spec.md gets a "VCL 0.3 additions (being built)" section outside the shared text; `DECISIONS.md` gets E8 and above.

###### Phase 2: semantics and mock file system (done: E11, F7, D28)
- **`vcl/semantics.py`:** `safe_name` and `path_join` are done in Phase 1a (D26), with their conformance cases (path traversal, reserved names, trailing dots, wildcards, UNC roots).
- **`vcl_fixtures.py`:** builds folder fixtures as an in-memory tree of path to file record (size, locked), with the same seed rules.
- **`vcl_interp.py`:** runs FOLDER_EXISTS, MAKE_FOLDER and COPY_FILE on the mock with the status rules above. Adds the FILE_EXISTS assertion (already in the validator) and a new `FILE_COUNT("<folder>")`. The UNCHANGED check covers the mock file system.
- **Tests:** a unit test for each status, including locked, conflict, too long and bad name.

###### Phase 3: VBA emitter and runtime
Two PRs (D29). **3a (done):** the compiler, ForgeLib's file helpers and the lint bans. As built,
the helpers are Functions that return the status and also take each path's root, and each has
its own `On Error GoTo` handler instead of `On Error Resume Next` (D29). **3b (done):** the
sandbox builder, `ForgeSandbox_<file>()`, which also closes D28's gap: a TEST fails when the
longer Excel root could change a Too long status (D30).
- **`vcl_codegen.py`:** emits `ForgeSafeName`, `ForgePathJoin`, `ForgeFolderExists`, `ForgeMakeFolder path, status` and `ForgeCopyFile src, dst, status`.
- **`ForgeLib.bas`:** late-bound FSO wrappers, each with a local `On Error Resume Next` that is cleared and restored. They re-check the root prefix, use overwrite False and compare sizes after a copy. No Kill, DeleteFile or MoveFile.
- **`vba_lint.py`:** bans DeleteFile, DeleteFolder, MoveFile, Kill and `CopyFile` outside ForgeLib.
- **Fixture builder (`vcl_fixtures.vba_for_file`):** emits a sandbox builder that creates `Environ$("TEMP")\KGC_KRONOS_SANDBOX_<yyyymmdd_hhnnss>` and writes dummy `.pdf` text files with `Open ... For Output`. It never deletes anything, and a new timestamped folder means it never collides.

###### Phase 4: the MVP spec
Two PRs (D32). **4a (done):** `specs/task_12.vcl`, with its fixtures and TESTs; 29 of 29 mutants
caught. **4b (done):** the language switch, with hard rule 21 and Example 3 (D33).
- **`specs/task_12.vcl`** (`RouteJobPacks`, Function returning the number copied). For each tracker row it takes `SAFE_NAME` of the Job ID and of the Report File, skips bad names (status `Bad name`), runs MAKE_FOLDER on `PATH_JOIN($jobs, id)` and COPY_FILE from `PATH_JOIN($dump, file)`, and writes the status. It declares `MODIFIES $tracker["Status"]` and `MODIFIES $jobs files`.
- **Fixtures:** the cases Copied, Missing, Conflict, Locked, an unsafe ID, and a re-run with `runs=2` (Already there).
- **Run:** `vclc all` and rebuild `dist/forge_module.bas`.
- **Language switch:** the shared reference text and the prompt move to VCL 0.3, then Boss updates the Orchestrator a second time (exact clicks given).

###### Phase 5: Excel pilot kit (optional; Boss runs it when he chooses)
Two PRs (D34). **5a (done):** `dist/forge_selftest.bas`, committed, from `compiler/vcl_selftest.py`.
**5b (done):** `dist/forge_pilot.bas`, `vclc pilot-check` (which replaces the expected JSON files)
and `docs/excel-pilot.md` (D35).
- **`dist/forge_selftest.bas`:** runs every conformance case through ForgeLib in Excel and writes PASS or FAIL to a new sheet. This is the actual "ForgeLib runs in Excel 2016" proof.
- **A guide for Boss** (`docs/excel-pilot.md`): open a blank .xlsm, import 3 modules, run `ForgeSandboxBuild`, run `RouteJobPacks`, then paste the results back.
- **`dist/expected/task_12*.json`:** what the run should show. Claude compares the pasted results automatically.

##### Verification (every phase)

```
python -m unittest discover tests
python compiler/validator.py specs/*.vcl
python compiler/vclc.py all specs/*.vcl -o dist
python compiler/gen_semantics_doc.py --check
python compiler/vclc.py mutate specs/*.vcl --snapshot
```

For Phase 0, also run `python compiler/vcl_migrate.py --prove tests/data/specs_v01/*.vcl`. Excel runs are never required for any phase to merge.

### `docs/vep-0.2-core.md`

#### VEP 0.2 core: the ten additions of VCL 0.2

Status: **design, accepted by Claude under the working policy** (pick the conservative option,
record it, carry on). Boss can overturn any choice; each one has a `DECISIONS.md` id (D16).
Nothing here is in the language until its implementation PR lands. Each lands with spec.md,
the grammar, parser and validator, type checker, interpreter, compiler, ForgeLib, conformance
cases, `docs/vcl-semantics.md` and tests.

Every addition below removes a pattern the 10 specs already spell out the long way, or a gap
the M2 research found. The aim is shorter specs that say what they mean, so the spec-to-VBA
dataset teaches intent, not workarounds.

##### How it lands

| PR | Contents |
| --- | --- |
| A | Value functions: SAME, CLEAN_TEXT, IS_ERROR, PARSE_DATE, IIF, MOD |
| B | Loop and write ops: SKIP_IF, EXIT_LOOP_IF, WRITE_ROW |
| C | MODIFIES and the automatic UNCHANGED check |
| D | `vcl_migrate.py`, the 10 specs moved to `VCL 0.2`, the Orchestrator prompt switched |

Until PR D, specs and the prompt stay `VCL 0.1`. The new features work only under a
`VCL 0.2` version line, and spec.md documents them in a "VCL 0.2 additions" section outside the
text it shares with the prompt. PR D then moves that section into the shared reference text, so
the Orchestrator learns 0.2 in one step. Every `VCL 0.1` spec stays valid under 0.2.

##### The value functions (PR A)

###### SAME(a, b): equal as a person reads it

```text
SAME(<expr>, <expr>)   TRUE when TRIM(a) and TRIM(b) are the same text, ignoring case
```

- Both sides become canonical text (S6), then TRIM (S2), then compare ignoring case.
- So `SAME(" ready ", "READY")` is TRUE, `SAME(25, "25")` is TRUE, and `SAME(25, "25.0")` is
  FALSE. Numbers are not compared as numbers; use IS_NUMBER and `=` for that.
- `SAME(<empty>, "")` is TRUE.
- Replaces the `UPPER(TRIM(x)) = "READY"` pattern (6 times in 5 of the 10 specs) with
  `SAME(x, "Ready")`.
- ForgeLib: `ForgeSame(a, b)`.

###### CLEAN_TEXT(x): tidy typed text

```text
CLEAN_TEXT(<expr>)   TRIM, then every run of spaces, non-breaking spaces, tabs and line
                     breaks inside becomes one space; characters below code 32 are removed
```

- Case is kept. Same whitespace rule as COL's header match (S9), now available to values.
- `CLEAN_TEXT("  W-01 " & Chr(160) & Chr(10) & "  A ")` is `"W-01 A"`.
- ForgeLib: `ForgeCleanText(v)`.

###### IS_ERROR(x)

```text
IS_ERROR(<expr>)   TRUE for an Excel error value (#N/A, #DIV/0!, #VALUE!, #REF!, #NAME?,
                   #NUM!, #NULL!); FALSE for everything else, including the text "#N/A"
```

- Decided as S5 in revision 4; this implements it.
- ForgeLib: `ForgeIsError(v)` (VBA `IsError`).

###### PARSE_DATE(x, "<format>"): text dates, without guessing

```text
PARSE_DATE(<expr>, "<format>")   a real date, or empty when x is not a date in that format
formats: "dd/mm/yyyy"  "dd-mm-yyyy"  "dd.mm.yyyy"  "mm/dd/yyyy"  "yyyy-mm-dd"
```

- A real date value passes through unchanged (time of day dropped).
- Text is TRIMmed, then must match the format exactly. Day and month take 1 or 2 digits, and
  the year takes 4, from 1900 (D17). The date must exist, so 31/02/2026 gives empty.
- The format is required. "03/04/2026" means 3 April or 4 March depending on the PC, so VCL
  never guesses (S4).
- The result is a Variant, so the type checker requires the usual guard:
  `LET(d, PARSE_DATE(CELL($ws, r, c), "dd/mm/yyyy"))` then `IF IS_DATE(d)`.
- ForgeLib: `ForgeParseDate(v, fmt)`, which uses `DateSerial` on the parsed parts and never
  `CDate` or `DateValue`, because those follow regional settings.

###### IIF(cond, a, b)

```text
IIF(<condition>, <expr>, <expr>)   a when the condition is TRUE, else b
```

- **Both a and b are always evaluated**, as in VBA's `IIf` and as with AND/OR in VCL. So an
  IIF branch is never a guard: the type checker treats both branches as unguarded.
- Values only: IIF can't return a Range, Worksheet, Workbook or Dictionary.
- ForgeLib: `ForgeIif(cond, a, b)`.

###### MOD(a, b)

```text
MOD(<expr>, <expr>)   a - b * INT(a / b), as Excel's MOD: the result has the sign of b
```

- `MOD(7, 3)` is 1, `MOD(-7, 3)` is 2, and `MOD(7.5, 2)` is 1.5.
- VBA's `Mod` operator is not used, because it rounds both sides to whole numbers first
  (banker's rounding) and gives the sign of a.
- `MOD(a, 0)` is a run-time error, like `/` by 0. The standard handler logs it.
- It is a calculation, so its arguments need IS_NUMBER guards (rule 15).
- ForgeLib: `ForgeMod(a, b)`.

##### The loop and write ops (PR B)

###### SKIP_IF and EXIT_LOOP_IF

```text
SKIP_IF <condition>         go on with the next pass of the innermost loop
EXIT_LOOP_IF <condition>    leave the innermost loop
```

- Both are structural ops (they start an OP line), and they are allowed only inside LOOP_ROWS,
  LOOP_COLS or FOR_EACH. The validator rejects them elsewhere.
- EXIT_IF stays as it is and leaves the procedure (E1).
- They replace the "whole loop body inside `IF NOT IS_EMPTY(...)`" pattern, which nests every
  step one level deeper:

  ```text
  LOOP_ROWS dir=down from=2 to=LASTROW($weldLog, cSeam) var=r
    SKIP_IF IS_EMPTY(CELL($weldLog, r, cSeam))
    ...
  ```

- VBA: SKIP_IF becomes `If <cond> Then GoTo Next_<n>` with a `Next_<n>:` label just before
  the loop's `Next`, as EXIT_IF already does with `GoTo CleanExit`. EXIT_LOOP_IF becomes
  `Exit For`. VBA has no `Continue`, and every VCL loop compiles to `For` or `For Each`, so
  these are the only forms needed.
- A dir=up loop that deletes rows works unchanged, because SKIP_IF only skips forward.

###### WRITE_ROW(ws, row, v1, v2, ...)

```text
WRITE_ROW(<ws>, <row>, <expr>, <expr>, ...)   writes the values to columns 1, 2, 3, ... of row
WRITE_ROW(<ws>, <row>, <expr>, ..., col=<n>)  starts at column n instead of 1
```

- Each value is written as SET_VALUE writes it.
- It needs at least one value.
- It replaces runs of `SET_VALUE(CELL($out, 1, 1), "...")`, `SET_VALUE(CELL($out, 1, 2), "...")`
  and so on. For example, task_04's header goes from 5 lines to 1.
- `col` is already a reserved option word.
- VBA: each value goes through `ForgeSetValue`, exactly as SET_VALUE compiles. It is not a
  single array write, so each cell gets exactly what SET_VALUE would give it. (Built as one
  `ForgeWriteRow` line that loops over the values; D18.)

##### MODIFIES and the automatic UNCHANGED check (PR C)

###### MODIFIES

```text
MODIFIES  $<arg>                         anything on that sheet: values, fills, formats, rows
MODIFIES  $<arg>["<header>", ...]        values in those columns (row 2 down) only
MODIFIES  $<arg> fills                   fills only
```

- It is a new key, written after PRE. One line per sheet, or per kind of change.
- An argument may have one bracketed line and one `fills` line.
- The validator checks it two ways:
  - **Every write op names a declared sheet.** The ops are SET_VALUE, WRITE_ROW, FORMAT, CLEAR,
    COPY_ROW (target), INSERT_ROW, DELETE_ROW, SORT and WRITE_ARRAY. A write to an undeclared
    sheet is an error.
  - **Nothing is declared but never written.** A declared sheet no op writes to is a warning.
- Columns can't be checked statically, because COL() finds them at run time. The UNCHANGED
  check below covers them.
- Built with three additions (D19): SET_FORMULA, DEDUPE and FILTER are write ops too; the kind
  of change must match the kind of line (fills need `fills`, formats and rows need the whole
  sheet); and a write that doesn't go through a `$sheet` argument is an error.

###### UNCHANGED (automatic, no syntax)

For every TEST of a spec with MODIFIES, the interpreter compares the workbook before the first
run with the workbook after the last run.

- **Every sheet:** cells outside what MODIFIES allows must be identical, as values, fills and
  number formats.
- **Column-declared sheets:** the header row and every other column are compared too. Row
  inserts and deletes count as changes unless the whole sheet is declared.
- **A TEST that stops at PRE:** the whole workbook must be unchanged.

A difference fails the TEST with a line such as
`UNCHANGED: Inventory!A5 was "B-105", now "B-150"`.

This is stronger than today's hand-written collateral TESTs (`COUNTA(...) = PRE_COUNTA(...)`),
which only count cells. Those TESTs stay in the specs, because tests are never removed to make
room. The Excel batch runner, when it returns, gets the same check from `dist/expected/`.

A spec without MODIFIES gets no automatic check. From PR D the validator warns about that, and
every migrated spec has MODIFIES. (D20: a spec that writes nothing gets the check anyway.)

##### Migration (PR D)

`python compiler/vcl_migrate.py specs/*.vcl` rewrites a `VCL 0.1` spec into `VCL 0.2`. It
makes only proven-equivalent rewrites:

| 0.1 pattern | 0.2 |
| --- | --- |
| `UPPER(TRIM(x)) = "LIT"` with LIT already upper case and trimmed | `SAME(x, "LIT")` |
| `UPPER(TRIM(x)) <> "LIT"` (same condition) | `NOT SAME(x, "LIT")` |
| a loop whose whole body is one `IF NOT IS_EMPTY(x)` block | `SKIP_IF IS_EMPTY(x)` then the body, one level up |
| 2 or more consecutive `SET_VALUE(CELL($o, r, k), v)` with k = 1, 2, 3, ... | `WRITE_ROW($o, r, v1, v2, ...)` (or `col=` when k starts above 1) |
| (none) | MODIFIES lines from the sheets the write ops name: a whole sheet when it is CLEARed or rows are added or removed, `fills` when FORMAT is its only write, otherwise the written columns' headers |

The migration counts as done only when, for every spec:

1. the validator reports 0 errors and no new warnings
2. every TEST passes, and every TEST's final workbook and RESULT are identical before and after
3. the TESTs miss no more mutants than before (built as D20: the percentage can't be compared,
   because WRITE_ROW merges lines; mutants of the new SKIP_IF lines are reported separately)
4. the new UNCHANGED check passes on every TEST

The literal in a SAME rewrite keeps its original upper-case text (`SAME(x, "READY")`). Changing
the case would be a cosmetic edit beyond what was proven.

##### Found while writing this (fixed in PR B: S17, `ForgeSetValue`)

SET_VALUE compiles to a plain `.Value = <expr>`. When the value is text that Excel would read as
something else ("1-2", "25", "TRUE", "=A1"), Excel converts it. The interpreter keeps it as
text, so the two can disagree. The literal texts in the 10 specs are plain labels ("Short",
"Days Open"), but some computed values are TRIMmed keys: task_01 writes part number "1001" and
task_04 writes seam "101" as text, which Excel stores as numbers. Every current TEST gives the
same result either way, because COUNTIF and the snapshot compare canonical text (S6), so nothing
fails today. It is still a real difference between the interpreter and Excel. The fixture builder already guards
this case (`needs_text_guard`). Proposed fix, with its own conformance case: SET_VALUE and
WRITE_ROW go through a ForgeLib `ForgeSetValue` that applies the same guard. It is tracked as
the first step of PR B.

##### Out of scope for 0.2 core

- List syntax (`cols=[...]`), FILTER, DEDUPE with headers, SET_FORMULA, sheet and file ops:
  MVP plan item 4.
- Classes, events and form keys: later VCL 0.2 work (roadmap).
- NOTE stays the escape hatch.

### `docs/vcl-m2-findings.md`

#### VCL M2/M3 research: grammar, types and semantics (findings)

Prototype work for milestones M2 and M3 of `docs/vcl-language-plan.md`, done outside the repo
and tested against the 10 specs (revision 3), Example 1 of the Orchestrator prompt and a BUILD
example. Nothing here changes the language yet. Every recommendation is a decision or a VEP
for you to approve.

##### What was built (all tested: 31 unit tests, 66 conformance cases)

| File | What it is |
| --- | --- |
| `vcl/grammar.lark` | The formal grammar of VCL 0.1 rev 3, one start rule per key line |
| `compiler/vcl_parse.py` | Parser built on the grammar: file structure, key order, OP nesting. Never crashes |
| `compiler/vcl_types.py` | Type checker with flow-sensitive narrowing |
| `vcl/semantics.py` | Reference meaning of the core value functions |
| `tests/conformance_cases.py` | 66 conformance cases, shared by tests, docs and the future Excel batch |
| `docs/vcl-semantics.md` | Generated reference: decisions S1 to S11 plus the conformance tables |

##### 1. Grammar findings

**G1. The language is unambiguous.** The grammar builds as LALR(1), which Lark refuses to do
for an ambiguous grammar. All 10 specs and Example 1 parse. The parser also survived 400
randomly damaged specs without crashing.

**G2. Options vs comparisons inside calls.** `DICT_ADD(d, k, v, mode=sum)` and
`SET_VALUE(c, got = 0)` look the same to a parser. The grammar resolves this by making option
keys reserved words: dir, from, to, var, in, mode, into, key, order, header, what, fill,
numfmt, bold, width, style, cols, col, crit, readonly, format, save, after, name, match.
*Recommendation:* make the reserved list official. The validator rejects VAR names from the
list, and a comparison on such a name needs parentheses.

**G3. `DEDUPE(..., cols=1,2, ...)` is ambiguous** (is `2` a column or the next argument?).
*Recommendation (VEP):* list syntax `cols=[1, 2]`. The grammar already supports lists, and the
VCL 0.2 DO step (`reads=[...]`) needs them too.

**G4. BUILD folder fixtures are undefined.** The BUILD example's `fixture=folder_01` doesn't
parse, because FIXTURE only defines `.xlsx` workbooks. *Recommendation (VEP):* a FOLDER
fixture (`FIXTURE folder_01/ files=3 pattern=*.xlsx from=<workbook fixture>`) or a rule that
BUILD tests use workbook fixtures only.

**G5. A FIXTURE header can't contain a comma** ("Qty, kg"), because commas separate columns.
*Recommendation (VEP):* allow quoted headers, `cols=A:"Qty, kg",B:Part No`.

**G6. Strings use doubled quotes as the escape**, like VBA: `"Say ""hi"""`. Revision 3 never
said so. *Recommendation:* add this one line to spec.md.

**G7. Operator precedence now matches VBA:** arithmetic, then `&`, then comparison, then NOT,
AND, OR. `a < b < c` is rejected; VBA would silently compare a Boolean with c.
*Recommendation:* add the precedence table to spec.md.

##### 2. Type system (proposed for the validator)

**Types:** Num, Text, Date, Bool, Cell (a raw cell value: empty, number, text, date or error),
Sheet, Book, Range, Dict, Coll, Arr. VAR types map from VBA: Long/Double to Num, String to
Text, Variant to a flow-typed value.

**The core rule:** a raw Cell can't be used in arithmetic, in `<` `>` `<=` `>=`, or compared
with `=` to a number or date, and can't be assigned to a Long, Double or Date variable, until
an enclosing `IF IS_NUMBER(...)` or `IF IS_DATE(...)` narrows it. One-line IFs, ELSE branches
and EXIT_IF narrow too.

**T1. It works on real specs.** 9 of the 10 specs and Example 1 type-check as written.

**T2. task_06 can't be proven safe.** It guards a date comparison through a flag
(`IF NOT bad`) instead of the check itself. It is safe at runtime, but no checker can follow a
flag. The fix is a one-line change, proved equivalent over all 20 input combinations:

```
OP           IF NOT bad                                   (before)
OP           IF IS_DATE(expDate)                          (after)
OP             IF expDate < TODAY THEN LET(bad, TRUE)     (unchanged)
```

*Recommendation:* apply this edit, and add the style rule "guard with the check itself, never
with a flag".

**T3. Short-circuit trap caught.** `IF IS_NUMBER(x) AND x > 5` crashes in VBA when x is text,
because VBA evaluates both sides of AND. The checker rejects it. The spec already warned about
this in prose; now it is enforced.

**T4. Writes cancel guards precisely.** A write to the same cell, or a whole-sheet change
(CLEAR of a range, COPY_ROW, SORT, DELETE_ROW), cancels the guards it could affect. Writes to
other cells keep them. This is sound because PRE guarantees every COL() lookup found its
header, and different header names can't match the same header cell.

**T5. Dictionary values are tracked.** `DICT_GET(..., into=n)` into a Long is only allowed
when the dictionary's values are numbers (from `mode=sum` or `mode=count`, or guarded values).

**T6. 11 mistakes caught, 8 valid patterns accepted** (tests/test_types.py). Caught: the AND
trap, raw cell into a Long, raw arithmetic, a guard cancelled by a write or a SORT, an
unguarded mode=sum, a missing `$`, unguarded date comparison, cell = 0, a flag guard, and a
guard wrongly carried into a loop.

##### 3. Semantics decisions (S1 to S11, full table in docs/vcl-semantics.md)

The most important ones, because plain VBA gets them silently wrong:

| Id | Decision | The VBA trap it avoids |
| --- | --- | --- |
| S1 | IS_EMPTY includes spaces, non-breaking spaces, tabs and line breaks | Copy-pasted "blank" cells holding character 160 |
| S2 | TRIM also removes non-breaking spaces and line breaks | VBA Trim only removes ordinary spaces |
| S3 | IS_NUMBER accepts "25" but rejects "1,250", "25%", "1e3" | Regional settings decide what "1,250" means |
| S4 | IS_DATE is true only for real dates | "03/04/2026" could be March or April |
| S6 | Values become text the same way on every PC: dates as yyyy-mm-dd, "." as the decimal point, errors as "#N/A" | CStr on a date depends on regional settings; Trim on a #N/A cell crashes with a type mismatch |
| S7 | A cell holding the number 25 equals the text "25" | In VBA, `cell = "25"` is False when the cell holds a number |
| S10 | Dictionary keys compare as text, ignoring case | VBA's Dictionary treats 1001 and "1001" as different keys |

**Consequence for ForgeLib:** every text conversion, comparison with text and dictionary key
must go through a ForgeLib helper. Inline VBA would hit these traps. In particular, any `#N/A`
in a register crashes a macro that calls VBA's Trim on it.

**S5, IS_ERROR (VCL 0.2 candidate).** Revision 3 has no way to detect error cells. Under S1 to
S4 an error is not empty, not a number and not a date, so most specs already treat it as
"invalid", which is the right outcome. An explicit IS_ERROR is still needed for reports.

##### 4. Recommendations for the repo (in order)

1. **Decisions (DECISIONS.md):** reserved option words (G2), string escaping (G6), precedence
   (G7), semantics S1 to S11, and the style rule from T2. These clarify revision 3 rather than
   change it.
2. **Spec edit:** the task_06 rewrite (T2).
3. **M2 integration:** adopt `vcl/grammar.lark` as the normative syntax, with Lark pinned as a
   dependency (`lark==1.3.1`, pure Python). Either rebuild `compiler/parser.py` on it, or keep
   the hand-written parser and test it against the grammar on every spec. Add the type checker
   as a validator stage.
4. **M3 seed:** `vcl/semantics.py` becomes the core of the reference interpreter. The
   conformance cases move to `vcl/conformance/` and later become the Excel conformance batch.
5. **VEPs to write:** list syntax (G3), folder fixtures (G4), quoted headers (G5), IS_ERROR (S5).
