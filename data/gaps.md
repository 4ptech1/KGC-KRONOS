# Expressibility gaps

What the corpus helpers could not say in VCL, or only by a workaround (`docs/corpus-plan.md`). The count per
kind is what Stage B reads to decide which ops and functions to add. A gap is a fact about the language, not
a mistake of the helper. The "wording" table lists prompt sentences that misled a helper; they
are fixed together at the end of Phase 6 so the Orchestrator is updated once.

## Gaps in the language

| Pair | Could not do | What the spec does instead |
| --- | --- | --- |
| c013 | month arithmetic ("12 months before today") | `TODAY - 365`, on an ASSUME line |
| c015 | read the fill of a cell (so a macro cannot keep a highlight a person added) | clears every fill it owns on each run, on an ASSUME line |
| c019 | a `reldates` fixture column cannot hold text such as "TBC" | "no valid date" is tested with blank cells only |
| c023 | DEDUPE cannot trim its keys | a two-pass count dictionary with a `dir=up` loop |
| c028 | a SET_FORMULA formula cannot name a sheet whose name has spaces ("Calibration Tracker") | names `CalibrationTracker` and tells the user to rename the tab if it differs |
| c030 | FILTER keeps the rows equal to one value; it cannot hide only the Closed rows | keeps only Open rows, so blank or other statuses are hidden too |
| c030 | no TEST can prove that a row that was reopened shows again (a TEST cannot change a cell between runs, a fixture cannot start with hidden rows) | `runs=2` proves only that a re-run is safe |
| c034, c039, c051 | one fixture column that holds real numbers and non-numbers together (`+boundary` on `int()` or `num()` takes numbers and "" only; `values()` takes quoted text only and stores it as text) | two FIXTURE files, one with numbers and blanks, one with text values such as "n/a" |
| c036, c037, c038 | write rows in key order, or sort a dictionary's keys | writes the rows in first-seen order, then `SORT(USED_RANGE($out))`; every row-counter mutant is then equivalent, because the SORT moves the gap rows down |
| c039, c040, c051, c052 | a set: a dictionary needs a value and a mode on every `DICT_ADD` | `1, mode=keep`; the mode and value mutants are equivalent because only `EXISTS` reads it |
| c043 | the last used row of a sheet, MAX, and "is this row empty" | `LOOP_COLS` over the header columns with `LASTROW` and a flag; the one `>` mutant is equivalent |
| c044 | DEDUPE cannot trim its keys (case is folded already) | an ASSUME says " NCR-003" and "NCR-003" stay different |
| c049 | FIND limited to one column (no column range, and RANGE refuses whole columns) | `FIND(USED_RANGE(...))` with an ASSUME that the NCR number appears only in the NCR No column |
| c050 | the latest value per key (DICT_ADD has no max or latest mode) | `DICT_GET` + `IS_DATE` + compare + overwrite; two equivalent mutants |
| c050 | ROUND, or a percentage | a fraction shown with `FORMAT numfmt="0.0%"`; the fixture counts make every rate exact in binary |

## Wording of the prompt to fix

| Pair | Sentence | Problem |
| --- | --- | --- |
| c021 | rule 19: "repeat dictionary keys with +dupes(<n>) or values() counts of 2 or more" | different spellings of one key (" h-100 " and "H-100") do not count; the first data row must change the result |
| c025 | rule 19: "at least one blank cell" | V059 wants one in every FIXTURE sheet of every file the SKIP_IF reads |
| c026 | FIXTURE: "+boundary(v1,v2,...) the first data rows hold exactly these values" | every boundary value must also be in `values(...)`, listed no more often than its count (V264) |
| c028 | FIXTURE: "They may be short forms of the real sheet names" | a formula needs the sheet's exact name, so a short form does not work in SET_FORMULA |
| c031, c032 | rule 9 ("One macro or one tool is no reason for a BUILD") against Example 4 ("Make it one tool I can start with one macro") | Example 4's request invites a BUILD that one TASK would do |
| c032 | rule 5: "swap two of the sheets RunHeadless takes" | swap two whose headings differ, or the PRE still passes |
| c019 | rule 15: "guard ... in a nested IF" | SEMANTICS says SKIP_IF NOT IS_DATE(x) guards too |
| c033, c039 | rule 19: "(a "" boundary value in an early row, or +blank(<n>); LASTROW ignores a blank in the last row)" | a `LASTCOL($ws, 1)` mutant survives unless a data row is narrower than the header (a trailing column blank in row 2); and a blank key in the last row means the loop never visits that row |
| c036, c037, c051 | rule 19: "stale=<n> rows" (and Example 4's `stale=2`) | stale must be larger than the number of rows the macro writes, or a deleted CLEAR survives (raised to 6 and 8) |
| c034, c039, c050, c051 | FIXTURE: "+boundary(...) the first data rows hold exactly these values" and "for values(...) the values move up, so the counts stay the same" | a boundary value must fit the column rule (V262: `int()` and `num()` take numbers only) and the list must be the same multiset as the `values(...)` counts (V264); text is possible only with `values()`, which stores text (V033 for bare numbers); the form of a date in `+boundary` for `dates(...)` is not given |
| c042, c043 | rule 15: "Guard with the check itself ..., never with a flag variable set earlier" | read as banning every flag; it only concerns a guard, and a Boolean that is not a guard (`hasOther`) avoids counter mutants such as `+ 1` to `+ 2` |
| c042, c043, c044 | rule 5: "When the test workbook has another sheet, add a wrong-sheet TEST" | unclear whether a spec should add a Notes sheet only to have one |
| c050 | rule 17: "prefer Functions returning a count" | in a BUILD a callee's count that the caller discards cannot be seen by any TEST; the callee should be a Sub, or its result used |
| c050, c052 | rule 23: "its own GOAL, PRE, MODIFIES and ERR" and "a fixture where the callee's PRE fails and one where it gives 0 when Run checks n < 0" | the TASK format marks PRE and MODIFIES optional; the second sentence is hard to parse |
| c051 | rule 19: a blank cell in every column a `SKIP_IF IS_EMPTY` tests | with dictionary lookups this invites a guard pair (one when building, one at the lookup), and one of the two is always an equivalent mutant |
| c036, c038, c041 | DICT_GET and `SET_VALUE(cell, "")` | the prompt never says what DICT_GET gives for a missing key (a Long variable gets 0), or whether `SET_VALUE(cell, "")` leaves the cell empty (c034 used it; c035, c041, c042, c043 used `CLEAR(..., what=contents)`) |
| c037 | SORT | the prompt does not say whether SORT is stable or how it orders text; the spec has an ASSUME that ties come in no fixed order |
| c044 | TEST ASSERTIONS | `VALUE(cell) = "text"` compares with case, `COUNTIF` does not; this is said only under "= and <>" |

## Tooling notes (not the language)

| Seen in | Note |
| --- | --- |
| all six helpers | `vclc mutate` lists 4 survivors and then "... and N more"; there is no way to list them all (D58 knew this) |
| c036, c050, c051 | plain `mutate` samples, so the same spec scores differently from run to run, while `--snapshot` is exact and is what the corpus gate uses; the prompt never mentions `--snapshot`. The FORMAT mutants of c050 survive in plain mode |
