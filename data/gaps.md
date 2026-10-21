# Expressibility gaps

What the corpus helpers could not say in VCL, or only by a workaround (`docs/corpus-plan.md`). The count per
kind is what Stage B reads to decide which ops and functions to add. A gap is a fact about the language, not
a mistake of the helper. The "wording" table at the end lists prompt sentences that misled a helper; they
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
