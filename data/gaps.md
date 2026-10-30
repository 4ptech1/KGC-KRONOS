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
| c019, c054, c059 | a `reldates` or `dates` fixture column cannot hold text such as "TBC", or a cell of spaces | "no valid date" is tested with blank cells only |
| c023 | DEDUPE cannot trim its keys | a two-pass count dictionary with a `dir=up` loop |
| c028 | a SET_FORMULA formula cannot name a sheet whose name has spaces ("Calibration Tracker") | names `CalibrationTracker` and tells the user to rename the tab if it differs |
| c030 | FILTER keeps the rows equal to one value; it cannot hide only the Closed rows | keeps only Open rows, so blank or other statuses are hidden too |
| c030, c069 | no TEST can prove that a row that was reopened shows again (a TEST cannot change a cell between runs, a fixture cannot start with hidden rows) | `runs=2` proves only that a re-run is safe |
| c069 | show the rows a FILTER hid, or read whether a row is hidden, so "show the other rows again if I run it twice" (a toggle) cannot be said | every run applies the filter again, on an ASSUME line |
| c034, c039, c051 | one fixture column that holds real numbers and non-numbers together (`+boundary` on `int()` or `num()` takes numbers and "" only; `values()` takes quoted text only and stores it as text) | two FIXTURE files, one with numbers and blanks, one with text values such as "n/a" |
| c036, c037, c038 | write rows in key order, or sort a dictionary's keys | writes the rows in first-seen order, then `SORT(USED_RANGE($out))`; every row-counter mutant is then equivalent, because the SORT moves the gap rows down |
| c039, c040, c051, c052, c057, c060, c070 | a set: a dictionary needs a value and a mode on every `DICT_ADD` | `1, mode=keep`; the mode and value mutants are equivalent because only `EXISTS` reads it (c060's first design left four such mutants and fell under the floor, so it writes the rows and runs `DEDUPE` instead; c061 stores "Closed" or "Open" and reads it) |
| c043 | the last used row of a sheet, MAX, and "is this row empty" | `LOOP_COLS` over the header columns with `LASTROW` and a flag; the one `>` mutant is equivalent |
| c044 | DEDUPE cannot trim its keys (case is folded already) | an ASSUME says " NCR-003" and "NCR-003" stay different |
| c049, c068 | FIND limited to one column (no column range, and RANGE refuses whole columns) | `FIND(USED_RANGE(...))` with an ASSUME that the NCR number appears only in the NCR No column (c049); a `LOOP_ROWS` with `SAME` down the one column (c068) |
| c050 | the latest value per key (DICT_ADD has no max or latest mode) | `DICT_GET` + `IS_DATE` + compare + overwrite; two equivalent mutants |
| c050 | ROUND, or a percentage | a fraction shown with `FORMAT numfmt="0.0%"`; the fixture counts make every rate exact in binary |
| c057, c058 | make an output sheet when it is missing and rebuild it when it is there (rule 22: only the run that adds a sheet can write to it) | `ADD_SHEET` after `SHEET_EXISTS`, and a re-run leaves the summary as it is, on an ASSUME line; c059 and c064 take the output sheet as an argument instead and rebuild it on every run |
| c065 | ask whether a file is there (`FILE_EXISTS` works only in a TEST), and COPY_FILE has no "Exists" status (an existing file gives Already there or Conflict) | `LIST_FILES(folder, "<exact name>")` and `IF COUNT(found) = 0` |
| c067 | set a number format on a column the spec writes while the collateral check still covers the rest of the sheet (`MODIFIES $ws["Next Due"]` covers values only; FORMAT needs the whole sheet) | leaves the format as it is, on an ASSUME line |
| c064 | a TEST that reads a cell's number format | the `FORMAT numfmt=` loop is proven only by the snapshot mutation score, which compares number formats |
| c068 | a TEST that names TODAY (to check a date the macro writes) | a `reldates` boundary of 0 in another column of the same row, and `VALUE(<Actual Date>) = VALUE(<Planned Date>)` |
| c070, c072 | in a BUILD, a later step that reads a column an earlier step adds (RunHeadless cannot PRE a header that step 1 adds, and rule 23 then asks for a TEST where the later step's PRE fails, which cannot happen) | both columns are ASSUMEd to exist and checked in PRE |

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
| c034, c039, c050, c051, c054, c059, c067 | FIXTURE: "+boundary(...) the first data rows hold exactly these values" and "for values(...) the values move up, so the counts stay the same" | a boundary value must fit the column rule (V262: `int()` and `num()` take numbers only) and the list must be the same multiset as the `values(...)` counts (V264); text is possible only with `values()`, which stores text (V033 for bare numbers); the form of a date in `+boundary` for `dates(...)` is not given |
| c042, c043 | rule 15: "Guard with the check itself ..., never with a flag variable set earlier" | read as banning every flag; it only concerns a guard, and a Boolean that is not a guard (`hasOther`) avoids counter mutants such as `+ 1` to `+ 2` |
| c042, c043, c044 | rule 5: "When the test workbook has another sheet, add a wrong-sheet TEST" | unclear whether a spec should add a Notes sheet only to have one |
| c050, c071, c072 | rule 17: "prefer Functions returning a count" | in a BUILD a callee's count that the caller discards cannot be seen by any TEST; the callee should be a Sub, or its result used |
| c050, c052 | rule 23: "its own GOAL, PRE, MODIFIES and ERR" and "a fixture where the callee's PRE fails and one where it gives 0 when Run checks n < 0" | the TASK format marks PRE and MODIFIES optional; the second sentence is hard to parse |
| c051 | rule 19: a blank cell in every column a `SKIP_IF IS_EMPTY` tests | with dictionary lookups this invites a guard pair (one when building, one at the lookup), and one of the two is always an equivalent mutant |
| c036, c038, c041, c057, c071 | DICT_GET and `SET_VALUE(cell, "")` | the prompt never says what DICT_GET gives for a missing key (a Long variable gets 0), or whether `SET_VALUE(cell, "")` leaves the cell empty (c034 used it; c035, c041, c042, c043 used `CLEAR(..., what=contents)`) |
| c037, c058 | SORT | the prompt does not say whether SORT is stable or how it orders text; the spec has an ASSUME that ties come in no fixed order |
| c044 | TEST ASSERTIONS | `VALUE(cell) = "text"` compares with case, `COUNTIF` does not; this is said only under "= and <>" |
| c053, c069 | rule 17 and Example 4's GOAL ("tell the user how many") | unclear whether "tell me how many" in one TASK means RESULT, a MSG or both: c053, c055 and c066 do both, c069 returns RESULT only |
| c064 | SEMANTICS, Dates: "TODAY + n and TODAY - n (a date plus or minus whole days)" | does not say that `TODAY - <date>` gives the number of days between them, which c064 needed and which works |
| c064 | FIXTURE: `reldates(lo, hi)` with `+boundary`, and `CLEAR(..., what=all)` | does not say a boundary value may lie outside lo..hi (it may), or whether CLEAR what=all removes number formats (the spec sets the date format again) |
| c067 | SET_FORMULA and FORMULA | no example has a quote inside a formula text; doubling it (`""""`) works in both |
| c059, c068 | TEST ASSERTIONS | no example checks a date; `COUNTIF(range, "2026-03-01")` and `VALUE(a) = VALUE(b)` work |
| c071, c072 | TEST ASSERTIONS: `<Sheet>![Header]` | does not say whether the header is found before or after the run, so a column the macro adds or rebuilds was checked by letters (`NDE!E1`, `Reorder!A1`) |
| c072 | FIXTURE: stale=<n> "with a red fill across the whole row" | spec.md says across the header's columns; and the prompt does not say stale may equal rows or sit on a sheet the macro only reads (both work) |

## Tooling notes (not the language)

| Seen in | Note |
| --- | --- |
| every helper of batches 2 and 3 | `vclc mutate` lists 4 survivors and then "... and N more"; there is no way to list them all (D58 knew this; batch 3's helpers used a scratch script that calls `mutate_report`) |
| c065 | V058 (rule 19's `+boundary` warning) fires on `COUNT(found) > 0`, a comparison of a count that no FIXTURE column can sit on; writing `= 0` avoids it |
| c036, c050, c051 | plain `mutate` samples, so the same spec scores differently from run to run, while `--snapshot` is exact and is what the corpus gate uses; the prompt never mentions `--snapshot`. The FORMAT mutants of c050 survive in plain mode |
