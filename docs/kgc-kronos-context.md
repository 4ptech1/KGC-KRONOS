# KGC KRONOS: AI context file

**Snapshot: 17 Oct 2026, `main` at 542e706.** Hand this file to any Claude chat or claude.ai
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
3. Every spec plus its proven VBA is a **verified spec-to-VBA pair**. From March 2027 those
   pairs train a fine-tuned local model that takes over code generation, with the compiler
   kept as the fallback and the reference.

The goal until March 2027: a large, verified dataset of spec-to-VBA pairs, and a toolchain
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
- **No local-model work** (llama.cpp, fine-tuning, GGUF) until the new PC arrives in March 2027.
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
DECISIONS.md              every language decision (G, S, T, E, F ids) and Claude's own (D1-D22)
README.md                 overview, how to run the checks, progress log, dataset count
requirements.txt          one pinned dependency: lark==1.3.1 (with hash)
vcl/spec.md               THE language reference (VCL 0.2)
vcl/orchestrator_prompt.txt  the Orchestrator's instructions; shares the reference text with spec.md
vcl/grammar.lark          formal LALR grammar
vcl/semantics.py          reference meaning of the value functions (S1-S17)
vcl/examples/             example1_rev3.vcl (= the prompt's Example 1), build_example.vcl (a sketch, invalid)
specs/task_01..10.vcl     the 10 KGC task specs (VCL 0.2)
compiler/vclc.py          the toolchain CLI: check, test, build, all, mutate
compiler/parser.py        hand-written parser   compiler/validator.py  rules, plus grammar and type stages
compiler/vcl_parse.py     grammar parser        compiler/vcl_types.py  type checker (guards, narrowing)
compiler/vcl_interp.py    reference interpreter, TEST assertions, the UNCHANGED check
compiler/vcl_codegen.py   VCL to VBA compiler   compiler/vcl_fixtures.py  the one fixture generator
compiler/vba_lint.py      static VBA checks     compiler/vcl_mutate.py  mutation scorer
compiler/vcl_migrate.py   VCL 0.1 to 0.2 rewriter, with --prove
compiler/gen_semantics_doc.py  writes docs/vcl-semantics.md
vba/ForgeLib.bas          the VBA runtime every generated procedure calls
dist/forge_module.bas     committed build output: ForgeLib plus one procedure per spec (paste-ready)
dist/forge_fixtures.bas, dist/procedures/, dist/expected/   rebuilt by vclc, git-ignored
tests/                    211 unit tests; conformance_cases.py is shared with the docs
tests/data/specs_v01/     the 10 specs and Example 1 frozen as VCL 0.1 (migration proof)
tests/data/task_01_pilot.vcl  task_01 as the Excel pilot ran it (pinned seed tests)
docs/vcl-mvp-plan.md      build sequence        docs/vep-0.2-core.md  VCL 0.2 design (done)
docs/vep-0.3-file-ops.md  file ops and Job-Pack Router design and roadmap (approved, Phase 0 done)
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

## 6. VCL 0.2 cheat sheet (summary only; `vcl/spec.md` is authoritative)

**TASK keys, in this order:** `VCL 0.2`, `TASK`, `# ASSUME:` lines, PROC, KIND (Sub or
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
- sheet and file ops: ADD_SHEET … WRITE_TEXT
- CALL, LOG, MSG, RETURN
- NOTE (the escape hatch)

**Value functions:**
- worksheet: COL, CELL, LASTROW, LASTCOL, ROW_RANGE, USED_RANGE, RANGE, SHEET
- arrays and dictionaries: ARR, ROWS, COLS, COUNT, EXISTS
- text: TRIM, UPPER, LOWER, LEN, LEFT, RIGHT
- checks: IS_EMPTY, IS_NUMBER, IS_DATE, IS_ERROR, IS_NOTHING
- constants and files: TODAY, NOW, FILE_NAME
- new in 0.2: SAME, CLEAN_TEXT, PARSE_DATE (with a required format), IIF, MOD

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

**MODIFIES** (a key after PRE):
- the three forms:
  - `MODIFIES $x`: the whole sheet
  - `MODIFIES $x["Header", ...]`: values in those columns, row 2 down
  - `MODIFIES $x fills`: fills only
- **Before the run**, the validator checks every write against these lines.
- **After every TEST**, the interpreter checks that nothing else changed (UNCHANGED).
- **A spec that writes nothing** needs no MODIFIES lines, and every TEST checks it changed
  nothing.

**The 20 hard rules (short form):**
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
20. Declare every sheet the macro changes with MODIFIES.

**FIXTURE syntax:**
`FIXTURE <file>.xlsx sheet=<Name> cols=A:Header,... rows=<n> [stale=<n>] data=<col>:<rule>;...`
- Rules: values("a"x3,...), seq, int, num, dates, reldates, text.
- Modifiers, in this order: `+boundary(...)`, `+blank(n)`, `+dupes(n)`.
- The data is deterministic: each column has its own seed, made from file, sheet and column.

**TEST assertions:**
- measures: COUNTIF, COUNTIFS, COUNTA, SUM, SUMIFS, VALUE, ROWS_USED, FILL_COUNT
- run results: RESULT, ERROR_COUNT
- PRE_<measure>, taken before the run
- ranges by header: `Sheet![Header]` and `Sheet![Header]<row>`

## 7. What the toolchain can run today (important when reviewing specs)

The validator accepts the whole language, but the interpreter and compiler implement only part
of it. A spec outside this subset passes the validator but fails at test or build.

| Fully supported (validate, test, compile) | Compile only (no interpreter yet) | Not built yet |
| --- | --- | --- |
| LET, SET_VALUE, WRITE_ROW, CLEAR, FORMAT, COPY_ROW, DELETE_ROW, DICT_ADD, DICT_GET, SORT, RETURN; LOOP_ROWS, LOOP_COLS, FOR_EACH over KEYS(...), IF/ELSE, EXIT_IF, SKIP_IF, EXIT_LOOP_IF; CELL, COL, LASTROW, LASTCOL, ROW_RANGE, USED_RANGE, TRIM, UPPER, LOWER, LEN, LEFT, RIGHT, IS_EMPTY, IS_NUMBER, IS_DATE, IS_ERROR, SAME, CLEAN_TEXT, PARSE_DATE, IIF, MOD, COUNT, EXISTS, KEYS; MODIFIES | INSERT_ROW, READ_ARRAY, WRITE_ARRAY, CALL, LOG, MSG, SELECT_CASE, ARR, ROWS, COLS, SHEET, RANGE, IS_NOTHING, FILE_NAME; FOR_EACH over SHEETS, CELLS or ITEMS | FILTER, DEDUPE, SET_FORMULA, ADD_SHEET, COPY_SHEET, OPEN_BOOK, SAVE_AS, CLOSE_BOOK, LIST_FILES, READ_TEXT, WRITE_TEXT, FIND, NOTE; assertions SHEET_EXISTS, FILE_EXISTS; BUILD specs |

**Known bug (waiting for Boss's "fix colours"):** `FORMAT(..., fill=red)` and the other colour
names pass the validator. But the interpreter doesn't turn them into a colour, so `FILL_COUNT`
doesn't count them, and the compiler refuses them. Until it is fixed, write fills as `#RRGGBB`
(red is `#FF0000`) or `fill=none`.

## 8. Status (17 Oct 2026)

**Roadmap Phase 0, Foundation (14 to 31 Oct 2026).** Gate 0: all specs validate, type-check,
pass their TESTs in the interpreter, and compile to lint-clean VBA, in CI. **This is met.**

| Measure | Now |
| --- | --- |
| Specs | 11 (task_01 to task_11), all VCL 0.2 |
| Unit tests | 214 pass |
| Validator | 0 errors; 1 known warning (task_04 has 48 OP lines, over the ~40 guideline; accepted until March) |
| Mutation score (snapshot) | 95.1% (389 of 409); 19 of the 20 survivors are equivalent mutants (D14, D24) that no test can ever catch; 1 is a small fixture gap in task_11 (D24) |
| VBA lint | 0 problems |
| Verified pairs | 0 (none has been run in Excel yet) |
| Orchestrator Project | updated to the VCL 0.2 prompt; its replies start with `VCL 0.2` and use MODIFIES and SAME |

The 11 specs:

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

**Still open on GitHub:**
- #4, the ForgeTest Excel pilot. Deferred, and nothing depends on it.
- #5, #6 and #7: older proposal documents (Phase 1 plan, Phase 2 dataset plan, VCL 0.2
  proposal). They are parked, and #7 is largely superseded by the VCL 0.2 that was actually
  built.

## 9. The plan

**Current: the Job-Pack Router MVP** (`docs/vep-0.3-file-ops.md`, approved by Boss).

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
| 1 | Language surface: `VCL 0.3`; SAFE_NAME, PATH_JOIN, FOLDER_EXISTS; MAKE_FOLDER and COPY_FILE returning a status into a VAR; Path and Name types; `MODIFIES $jobs files`; folder FIXTUREs | **Waiting for Boss: "go Phase 1"** |
| 2 | `safe_name` and `path_join` semantics with conformance cases; a mock file system in fixtures and the interpreter; FILE_EXISTS and FILE_COUNT assertions; UNCHANGED for files | Planned |
| 3 | Compiler and ForgeLib: late-bound FSO wrappers that never overwrite, delete or move; lint bans DeleteFile, MoveFile and Kill; a sandbox builder in `%TEMP%` | Planned |
| 4 | `specs/task_12.vcl` RouteJobPacks; the prompt switches to 0.3, and Boss updates the Orchestrator again | Planned |
| 5 | Excel pilot kit: a ForgeLib self-test of the conformance cases, a sandbox, and a 5-step guide for Boss | Planned, optional |

The key safety design for files:
- **Every path is built** with `PATH_JOIN` from a `$root` argument and `SAFE_NAME` parts.
  SAFE_NAME refuses `..`, `\ / : * ? " < > |`, reserved names like CON, and trailing dots.
- **Every file op returns a status** (`Copied`, `Already there`, `Conflict`, `Missing`,
  `Locked`, `Too long`, `Bad name`, `Failed`). An FSO error never aborts the whole run.
- **No overwrite, no delete, no retry loops.**

**After the MVP** (`docs/vcl-mvp-plan.md`):
1. The rest of revision 3 in the interpreter and compiler: DEDUPE with `cols=[...]`, FILTER,
   SET_FORMULA, sheet and file ops, list syntax.
2. A corpus of 150 tasks, Orchestrator-written specs, and an expressibility report.
3. A dataset generator: spec-to-VBA pairs from the compiler. Mutants double as "read
   carefully" contrast pairs.
4. **March 2027:** the new PC; train the local model; Excel conformance runs.

**Pending Boss decisions right now:**
1. "go Phase 1" for the router.
2. "fix colours", for the fill colour-name bug above.
3. Whether the Orchestrator's weld-highlight spec (HighlightRejectedWelds) becomes a repo spec.
   It needs `stale=10` instead of `stale=30` and `#FF0000` fills first.

## 10. Decisions worth knowing (full list in `DECISIONS.md`)

- **D1:** one fixture generator, a SplitMix64 PRNG seeded from SHA-256 of
  "file|sheet|column". The data never depends on the Python version.
- **D3:** only `dist/forge_module.bas` is committed. The expected results depend on TODAY.
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
| VCL | VBA Command Language: the spec language (`vcl/spec.md`). Current version 0.2; 0.3 is being designed. |
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
| Verified pair | A spec plus its VBA, proven by the toolchain; the training data for March 2027. |
| Gate 0 | The Phase 0 exit test: everything validates, passes, compiles and lints, in CI. |
