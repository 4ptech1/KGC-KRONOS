# VEP 0.4: the rest of revision 3 (MVP plan, item 2)

Status: **design, 20 Oct 2026** (Boss said "go item 2"). Choices that needed Boss's decision were
made the conservative way and recorded in `DECISIONS.md` D36, as CLAUDE.md asks; Boss can
overrule any of them. Phase 1 changes no language; later phases change VCL and follow the usual
rule (spec.md and the prompt together, every check, one Orchestrator update per switch).

## The problem

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

## Principles (unchanged)

- The interpreter is the reference meaning, every meaning goes through a ForgeLib helper, and a
  TEST proves it on synthetic data.
- Where Excel's own behaviour depends on the PC (regional settings, locale, criteria parsing),
  ForgeLib does the work itself in plain VBA, as it does for SORT keys and dictionary keys.
- Nothing touches a real workbook or a real folder until March 2027 (CLAUDE.md).
- MODIFIES covers every change, and UNCHANGED checks it after every TEST.

## Phases (one or two PRs each, in this order)

### Phase 1: the interpreter catches up (group A, no language change)

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

### Phase 2: list syntax and DEDUPE

- `cols=[1, 2]` (G3 of `docs/vcl-m2-findings.md`); a single number stays valid.
- DEDUPE keeps the first row of each key and deletes the later ones, moving rows up. The key is
  the listed columns' values compared as dictionary keys (S10: as text, ignoring case), not
  Excel's RemoveDuplicates, whose comparison isn't documented. ForgeDedupe marks the rows top
  down and deletes them bottom up. It needs the whole sheet in MODIFIES (as today).

### Phase 3: FIND

- FIND(range, x, match=whole) is the first cell, row by row, whose value is SAME as x;
  `match=part` is the first whose text contains TRIM(x), ignoring case. FIND gives a Range or
  NOTHING, and IS_NOTHING tests it. ForgeFind loops over the cells instead of calling
  Range.Find, whose settings (LookIn, LookAt, MatchCase, the last search) carry over between
  calls.

### Phase 4: sheets of the workbook

- ADD_SHEET(book, name=, after=, into=) and COPY_SHEET(sheet, after=, into=), and the
  SHEET_EXISTS assertion. They change the workbook's sheet list, so they need a new MODIFIES
  form, `MODIFIES $book sheets`: new sheets may be added, and every existing sheet stays as it
  was. A name that is already used is an error, never a replacement.

### Phase 5: FILTER (decision D36)

- FILTER hides the rows of a range whose column doesn't match; it never deletes. The match is
  SAME (exact text, case and spaces at the ends ignored), done by ForgeFilter setting
  `Rows(r).Hidden` itself, not by AutoFilter criteria, which read "1,250" or dates through the
  PC's settings and treat `*` and `?` as wildcards. The interpreter keeps hidden rows; the
  snapshot compares them; a new assertion VISIBLE_COUNT(<Sheet>!<range>) counts the visible
  rows.

### Phase 6: SET_FORMULA (decision D36)

- The interpreter can't calculate Excel formulas, so a TEST can only check what was written:
  a new assertion FORMULA(<Sheet>!<cell>) = "<formula>" compares the formula text. The
  validator refuses Microsoft 365 functions (XLOOKUP, FILTER, LET, LAMBDA, UNIQUE, SORT,
  SEQUENCE, TEXTJOIN, IFS and the like), as CLAUDE.md requires. A formula's value is proven
  only by an Excel run (the pilot of router Phase 5).

### Phase 7: text files in the sandbox (decision D36)

- LIST_FILES(folder, "pattern", into=list) and READ_TEXT(path, into=text) read the mock file
  system; folder FIXTUREs gain file contents (`files=a.txt:"text"`). Paths follow T4 (PATH_JOIN
  of a String argument).
- WRITE_TEXT writes a **new** file only, below a `MODIFIES $root files` root, and reports a
  status like COPY_FILE (`Written`, `Already there`, `Bad name`, ...): it never overwrites.
  `mode=append` is refused, because it changes a file that already exists.
- ITEMS(list) iterates what LIST_FILES gave, in name order.

### Deferred (decision D36)

- **OPEN_BOOK, SAVE_AS, CLOSE_BOOK** (group C): they open and save real workbooks, which CLAUDE.md
  rules out until March 2027. They stay valid in the language, the compiler refuses them with a
  clear message, and a sandbox design for them comes after March, for Boss to approve.
- **CALL and BUILD specs** (group D): their own item, after this one.
- **NOTE**: by design, VCL grows until NOTE is rare (D31).

## Verification (every phase)

The four standard checks, `vclc mutate --snapshot`, and for each new meaning: interpreter unit
tests, a conformance case where it is a value function, a ForgeLib self-test case (D34), and
lint. A phase that changes VCL bumps the version line once all its phases are in (as 0.2 and
0.3 did), so Boss updates the Orchestrator once.
