# The corpus: 150 requests with proven specs (VEP 0.6, D64)

The plan for `data/verified/` and `data/eval/`. Each pair is a request in Boss's words and the VCL spec the
toolchain proved. This file says what to cover, how a pair is made and what stops one from being stored.

## What to cover

13 registers, 9 families, 150 pairs. A pair names the registers its sheets are and the family of the task.

| Family | Planned | What it is |
| --- | --- | --- |
| flag | 30 | highlight or mark rows by a rule (fills, a Status column) |
| summarise | 25 | counts, totals, rates per key into a new or rebuilt sheet |
| reconcile | 20 | two or more registers against each other |
| clean | 15 | tidy, dedupe, validate or fix values in place |
| rebuild | 10 | list or copy rows into an output sheet that is rebuilt every run |
| files | 12 | folders and files: MAKE_FOLDER, COPY_FILE, LIST_FILES, READ_TEXT, WRITE_TEXT (sandbox only) |
| sheets_formulas | 10 | ADD_SHEET, COPY_SHEET, SET_FORMULA |
| find_filter | 8 | FIND with ROW_OF, FILTER, DEDUPE into=, text lookups |
| build | 20 | a BUILD of several TASKs joined by CALL, for a request one TASK cannot do |

Registers: `bom`, `grn`, `weld_log`, `nde`, `ncr`, `itp`, `welder_quals`, `calibration`, `consumables`,
`material_certs`, `wps_pqr`, `dispatch`, `job_tracker`. About 11 pairs each; a pair may name two.
`python compiler/vclc.py corpus --report` shows what is there and what is missing.

## How a pair is made

1. **The request.** Claude writes it as a person types it: short, practical, sometimes missing a detail
   the spec must assume. Invented sheet names, headers and figures only. Mixed difficulty: a simple flag,
   a two-register reconcile, a request with an edge (blank cells, text numbers, repeats, a re-run).
2. **The spec.** A helper that sees only `vcl/orchestrator_prompt.txt` and the request writes the spec, as the
   claude.ai Orchestrator would. It may run the toolchain and fix its spec up to two rounds.
3. **The gates** (`corpus-add` stops at the first one that fails):
   - the spec is valid with no warnings (meta.json may accept a named warning, such as V053), VCL 0.6, with
     its GOAL lines, in the canonical layout (`vclc fmt`);
   - every TEST passes in the reference interpreter;
   - it compiles and the VBA and the fixture builders are lint-clean;
   - its TESTs catch at least 90% of the mutants; a mutant that cannot change a result is counted as
     `equivalent` and the note says why.
4. **Failures are data.** Each failed attempt with its coded messages, and the spec that fixed it, is kept
   for `data/repair/` (Phase 7). A request the helper cannot express in VCL, or only by a workaround, is a
   **gap**: it is logged with the op family it needed. The gap count is what Stage B reads to grow VCL.
5. **Held out.** 15 pairs, spread over the registers and families, are stored in `data/eval/` instead and
   never used for training, paraphrasing or repair.

## Batches

Phase 6 makes the corpus in batches of about 20 requests, one PR each, with the gap report growing. The 12
committed specs (`specs/task_01.vcl` to `task_12.vcl`) are `c001` to `c012`; new pairs continue from `c013`.
Each batch keeps its requests, with registers, family and set, in `data/batches/batch_NN.json`; the failed
attempts of stored pairs go to `data/repair/attempts/<id>/` (never those of a held-out pair); what VCL could not
say goes to `data/gaps.md`.

| Batch | Pairs | State |
| --- | --- | --- |
| 1 | c013 to c032 | stored (D65): 18 verified, 2 held out |
