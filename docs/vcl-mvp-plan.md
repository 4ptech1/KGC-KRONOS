# VBA Forge MVP plan (revised 15 Oct 2026)

**Policy:** Boss has no testing, debugging or verification steps. Every check is automated: the
grammar, the type checker, the reference interpreter, mutation scoring, the VBA linter and CI.
Running the generated VBA in Excel is optional and never blocks work.

## MVP: what it is and where it stands

`python compiler/vclc.py all specs/*.vcl -o dist` turns VCL specs into:

- `dist/forge_module.bas`: one paste-ready module, the ForgeLib runtime plus one procedure per spec
- `dist/procedures/*.bas`: each procedure on its own
- `dist/expected/*.json`: for every TEST, the expected RESULT, ERROR_COUNT and final workbook state,
  so any future Excel run can be compared automatically

**Status: working** for all 10 KGC specs (task_06 uses the provable rewrite) and Example 1.

| Automated check | Result |
| --- | --- |
| Grammar (LALR, proven unambiguous) | 11 of 11 specs parse |
| Type checker (every calculation and date comparison guarded) | 11 of 11 pass |
| Interpreter: the specs' own TESTs | 39 of 39 pass |
| Mutation score, TEST assertions only | 87% of 417 deliberate bugs caught (was 68%) |
| Mutation score, whole-workbook comparison | 96.4% (was 84%); the 15 survivors are equivalent mutants (`DECISIONS.md` D14) |
| VBA lint (the compile errors Excel would raise) | 0 problems in 974 lines; 7 of 7 planted errors caught |
| Unit tests | 197 pass |

## Honest limits

- **The generated VBA is never executed.** Lint catches compile-time errors, not runtime
  behaviour. The main unproven risk is ForgeLib's behaviour in real Excel, for example comparing
  error values or Range.Sort edge cases.
- **One fixture generator** (item 1): `compiler/vcl_fixtures.py` builds the pilot's data
  bit for bit (`DECISIONS.md` D1).
- **The MVP compiler covers** every op the 10 specs use, plus loops, SELECT_CASE, arrays, CALL,
  LOG and MSG. Not yet covered: file ops, ADD_SHEET, COPY_SHEET, FIND, DEDUPE, FILTER,
  SET_FORMULA, OPEN_BOOK, SAVE_AS. `NOTE` steps need the local model.
- **EXIT_IF was undefined in revision 3.** The MVP treats it as "leave the procedure". VCL 0.2
  adds EXIT_LOOP_IF for loops.

## What the mutation scores teach (and the fix)

Every surviving mutant is missing test data, not missing assertions:

| Survivor | Missing test data | Fix (VEP) |
| --- | --- | --- |
| Deleting "clear the output sheet" goes unnoticed | output sheets start empty | fixtures pre-fill output sheets with stale rows |
| `mode=keep` vs `overwrite` goes unnoticed | no duplicate keys | `+dupes(n)` modifier |
| `<` vs `<=` goes unnoticed | no rows exactly on the boundary | `+boundary` rows (equal values) |
| task_10's 15-day window goes unnoticed | fixed dates don't move with TODAY | relative dates, `reldates(-30, 30)` |

Target after these: at least 95% with whole-workbook comparison. **Done in item 2 (VCL 0.1
revision 5):** `stale=<n>`, `+dupes(n)`, `+boundary(...)` and `reldates(lo, hi)` are in the
language, rule 19 warns when a spec doesn't use them, and the whole-workbook score is 96.4%.

## Build sequence (Claude Code works through it without waiting; merges itself when CI is green)

1. **Integrate the MVP** into the repo. Unify the two fixture generators, apply the task_06
   rewrite, record the decisions, and make CI run the whole `vclc all` pipeline.
2. **Upgrade fixtures from the mutation evidence** (the table above), until the snapshot score
   is at least 95%.
3. **VCL 0.2 core, end to end:** grammar, types, semantics, interpreter, compiler, ForgeLib and
   conformance cases for SAME, SKIP_IF, EXIT_LOOP_IF, WRITE_ROW, MODIFIES with automatic
   UNCHANGED, PARSE_DATE, IS_ERROR and CLEAN_TEXT (plus IIF and MOD). Then migrate the 10 specs
   with a migration tool. Design: `docs/vep-0.2-core.md` (four PRs, A to D).
   **PR A (the six value functions), PR B (SKIP_IF, EXIT_LOOP_IF, WRITE_ROW, ForgeSetValue) and
   PR C (MODIFIES with UNCHANGED) are done.**
4. **The rest of rev 3** in the compiler and interpreter (DEDUPE with `cols=[...]`, FILTER,
   SET_FORMULA, sheet and file ops), plus list syntax.
5. **Corpus:** 150 tasks, Orchestrator-written specs, and an expressibility report that drives
   VCL 0.2 packs.
6. **Dataset generator:** spec-to-VBA pairs straight from the compiler. The mutation engine
   doubles as a contrast-pair generator: every valid mutant is a near-identical spec with
   different behaviour, compiled to its own correct VBA. That is exactly the "read carefully"
   training data the Phase 2 plan wanted.
7. **March:** train the local model on the dataset; the compiler stays as the fallback and the
   reference. Excel conformance runs whenever Boss chooses.
