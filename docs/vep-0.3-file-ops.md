# VEP 0.3: file ops and the Job-Pack Router MVP

Status: **design, approved by Boss on 17 Oct 2026**. Phases 0 (D22), 1 (D26, D27), 2 (D28),
3 (D29, D30), 4 (D32, D33) and 5 (D34, D35) are done: the MVP is complete. Nothing below is in the language
until its phase lands.

## Context

Boss sent a brief (written with another AI) with three parts:
- **Phase 0 is unblocked:** add blank key rows to the FIXTURE lines, to close the 22 SKIP_IF mutation survivors (D20).
- **The Orchestrator is unblocked:** Boss is pasting the VCL 0.2 prompt into the claude.ai project himself.
- **A pivot:** build a single-task MVP, the "Job-Pack Router". It reads a tracker, makes job folders, files PDF test reports from a dump folder, and writes a status back. It needs new file ops, and it would prove ForgeLib in Boss's Excel 2016.

Boss's answers to my questions:
1. **Sandbox only.** It never runs on real folders before March 2027, so the hard constraint stays.
2. **Copy first.** COPY_FILE comes now; MOVE_FILE comes later.
3. **Macros can run** on his work PC.

**Acknowledged:** both backlog items are unblocked. Note that the Orchestrator will need a second update at the end of Phase 4, because the file ops add language.

## Critique of the File Ops VEP, and the revisions

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

## Phased roadmap

Each phase is one or two PRs: CI green, merge, then the next phase. Every PR runs the four standard checks plus `vclc mutate --snapshot`.

### Phase 0: baseline cleanup (blank key rows)
- **Change:** the FIXTURE and TEST lines only, in task_01, 02, 04, 05, 06, 08 and 10. Add `+blank(n)` (or a `""` boundary value) to each column that a SKIP_IF tests, then recompute the exact TEST expectations with the interpreter (as D13 did).
- **Frozen copies:** apply the same FIXTURE and TEST lines to the copies in `tests/data/specs_v01/`, so `tests/test_migrate.py` still proves migrate(v01) equals the specs. The tool never touches those lines; re-run `vcl_migrate.py --prove`.
- **Rule 19 warning:** a spec with `SKIP_IF IS_EMPTY(x)` whose fixture column for x has no blank cell gets a warning.
- **Done when:** the SKIP_IF probes are killed and the snapshot score is at least 96% again. D22 is recorded in `DECISIONS.md`, and CLAUDE.md, the README and the plan are updated.

### Phase 1: language surface (VCL 0.3, no run-time code yet)
Two PRs (D26). **1a (done):** the version line, the three value functions, the Name and Path types, and the meaning of SAFE_NAME and PATH_JOIN (moved here from Phase 2, because the type checker uses `safe_name` for quoted parts). **1b (done):** the ops, MODIFIES `files`, folder FIXTUREs and the validator rules (E9, E10, F6, D27).
- **Version line:** `VCL 0.3` in `parser.py`, `vcl_parse.py` and the validator gate, reusing the VCL_02 pattern.
- **Value functions:** SAFE_NAME(x), PATH_JOIN(base, name, ...) and FOLDER_EXISTS(path).
- **Ops:** `MAKE_FOLDER(<path>, into=<var>)` and `COPY_FILE(<from>, <to>, into=<var>)`. The status VAR must be a String.
- **Types (`vcl_types.py`):** Name and Path. A String ARG used as a base is a Path. PATH_JOIN takes a Path and then Names (SAFE_NAME results or safe literals) and returns a Path. The file ops take only Paths.
- **Grammar:** the MODIFIES `files` form, and a FIXTURE folder form: `FIXTURE <name> folder=<Name> files=<a.pdf>,... [locked=<a.pdf>,...]`. TEST calls pass folder names bare, as they do sheet names.
- **Validator:** every file op path is rooted in an ARG declared by MODIFIES `files` (the copy source may be any rooted Path); the status VAR is declared; a fixture with missing, conflict, locked or unsafe cases gets the rule 19 extension.
- **Docs:** spec.md gets a "VCL 0.3 additions (being built)" section outside the shared text; `DECISIONS.md` gets E8 and above.

### Phase 2: semantics and mock file system (done: E11, F7, D28)
- **`vcl/semantics.py`:** `safe_name` and `path_join` are done in Phase 1a (D26), with their conformance cases (path traversal, reserved names, trailing dots, wildcards, UNC roots).
- **`vcl_fixtures.py`:** builds folder fixtures as an in-memory tree of path to file record (size, locked), with the same seed rules.
- **`vcl_interp.py`:** runs FOLDER_EXISTS, MAKE_FOLDER and COPY_FILE on the mock with the status rules above. Adds the FILE_EXISTS assertion (already in the validator) and a new `FILE_COUNT("<folder>")`. The UNCHANGED check covers the mock file system.
- **Tests:** a unit test for each status, including locked, conflict, too long and bad name.

### Phase 3: VBA emitter and runtime
Two PRs (D29). **3a (done):** the compiler, ForgeLib's file helpers and the lint bans. As built,
the helpers are Functions that return the status and also take each path's root, and each has
its own `On Error GoTo` handler instead of `On Error Resume Next` (D29). **3b (done):** the
sandbox builder, `ForgeSandbox_<file>()`, which also closes D28's gap: a TEST fails when the
longer Excel root could change a Too long status (D30).
- **`vcl_codegen.py`:** emits `ForgeSafeName`, `ForgePathJoin`, `ForgeFolderExists`, `ForgeMakeFolder path, status` and `ForgeCopyFile src, dst, status`.
- **`ForgeLib.bas`:** late-bound FSO wrappers, each with a local `On Error Resume Next` that is cleared and restored. They re-check the root prefix, use overwrite False and compare sizes after a copy. No Kill, DeleteFile or MoveFile.
- **`vba_lint.py`:** bans DeleteFile, DeleteFolder, MoveFile, Kill and `CopyFile` outside ForgeLib.
- **Fixture builder (`vcl_fixtures.vba_for_file`):** emits a sandbox builder that creates `Environ$("TEMP")\KGC_KRONOS_SANDBOX_<yyyymmdd_hhnnss>` and writes dummy `.pdf` text files with `Open ... For Output`. It never deletes anything, and a new timestamped folder means it never collides.

### Phase 4: the MVP spec
Two PRs (D32). **4a (done):** `specs/task_12.vcl`, with its fixtures and TESTs; 29 of 29 mutants
caught. **4b (done):** the language switch, with hard rule 21 and Example 3 (D33).
- **`specs/task_12.vcl`** (`RouteJobPacks`, Function returning the number copied). For each tracker row it takes `SAFE_NAME` of the Job ID and of the Report File, skips bad names (status `Bad name`), runs MAKE_FOLDER on `PATH_JOIN($jobs, id)` and COPY_FILE from `PATH_JOIN($dump, file)`, and writes the status. It declares `MODIFIES $tracker["Status"]` and `MODIFIES $jobs files`.
- **Fixtures:** the cases Copied, Missing, Conflict, Locked, an unsafe ID, and a re-run with `runs=2` (Already there).
- **Run:** `vclc all` and rebuild `dist/forge_module.bas`.
- **Language switch:** the shared reference text and the prompt move to VCL 0.3, then Boss updates the Orchestrator a second time (exact clicks given).

### Phase 5: Excel pilot kit (optional; Boss runs it when he chooses)
Two PRs (D34). **5a (done):** `dist/forge_selftest.bas`, committed, from `compiler/vcl_selftest.py`.
**5b (done):** `dist/forge_pilot.bas`, `vclc pilot-check` (which replaces the expected JSON files)
and `docs/excel-pilot.md` (D35).
- **`dist/forge_selftest.bas`:** runs every conformance case through ForgeLib in Excel and writes PASS or FAIL to a new sheet. This is the actual "ForgeLib runs in Excel 2016" proof.
- **A guide for Boss** (`docs/excel-pilot.md`): open a blank .xlsm, import 3 modules, run `ForgeSandboxBuild`, run `RouteJobPacks`, then paste the results back.
- **`dist/expected/task_12*.json`:** what the run should show. Claude compares the pasted results automatically.

## Verification (every phase)

```
python -m unittest discover tests
python compiler/validator.py specs/*.vcl
python compiler/vclc.py all specs/*.vcl -o dist
python compiler/gen_semantics_doc.py --check
python compiler/vclc.py mutate specs/*.vcl --snapshot
```

For Phase 0, also run `python compiler/vcl_migrate.py --prove tests/data/specs_v01/*.vcl`. Excel runs are never required for any phase to merge.
