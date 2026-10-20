# VEP 0.4: the rest of revision 3 (MVP plan, item 2)

Status: **design, 20 Oct 2026** (Boss said "go item 2"). Choices that needed Boss's decision were
made the conservative way and recorded in `DECISIONS.md` D36, as CLAUDE.md asks; Boss can
overrule any of them. Phase 1 changes no language; later phases change VCL and follow the usual
rule (spec.md and the prompt together, every check, one Orchestrator update per switch).
Phases 1 to 7 are done (D37 to D45): every phase of item 2 is in the toolchain. The toolchain accepts the `VCL 0.4` line; spec.md and the
prompt move to 0.4 after the last language phase (D38).

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

### Phase 1: the interpreter catches up (group A, no language change; done: D37)

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

### Phase 2: list syntax and DEDUPE (done: S21, D38)

- `cols=[1, 2]` (G3 of `docs/vcl-m2-findings.md`); a single number stays valid.
- DEDUPE keeps the first row of each key and deletes the later ones, moving rows up. The key is
  the listed columns' values compared as dictionary keys (S10: as text, ignoring case), not
  Excel's RemoveDuplicates, whose comparison isn't documented. ForgeDedupe marks the rows top
  down and deletes them bottom up. It needs the whole sheet in MODIFIES (as today).

### Phase 3: FIND (done: S22, D39)

- FIND(range, x, match=whole) is the first cell, row by row, whose value is SAME as x;
  `match=part` is the first whose text contains TRIM(x), ignoring case. FIND gives a Range or
  NOTHING, and IS_NOTHING tests it. ForgeFind loops over the cells instead of calling
  Range.Find, whose settings (LookIn, LookAt, MatchCase, the last search) carry over between
  calls.

### Phase 4: sheets of the workbook (done: 4a ADD_SHEET and SHEET_EXISTS, S23, D40; 4b COPY_SHEET, D41)

- ADD_SHEET(book, name=, after=, into=) and COPY_SHEET(sheet, after=, into=), and the
  SHEET_EXISTS assertion. They change the workbook's sheet list, so they need a new MODIFIES
  form, `MODIFIES $book sheets`: new sheets may be added, and every existing sheet stays as it
  was. A name that is already used is an error, never a replacement.

### Phase 5: FILTER (decision D36; done: D42)

- FILTER hides the rows of a range whose column doesn't match; it never deletes. The match is
  SAME (exact text, case and spaces at the ends ignored), done by ForgeFilter setting
  `Rows(r).Hidden` itself, not by AutoFilter criteria, which read "1,250" or dates through the
  PC's settings and treat `*` and `?` as wildcards. The interpreter keeps hidden rows; the
  snapshot compares them; a new assertion VISIBLE_COUNT(<Sheet>!<range>) counts the visible
  rows.

### Phase 6: SET_FORMULA (decision D36; done: D43)

- The interpreter can't calculate Excel formulas, so a TEST can only check what was written:
  a new assertion FORMULA(<Sheet>!<cell>) = "<formula>" compares the formula text. The
  validator refuses Microsoft 365 functions (XLOOKUP, FILTER, LET, LAMBDA, UNIQUE, SORT,
  SEQUENCE, TEXTJOIN, IFS and the like), as CLAUDE.md requires. A formula's value is proven
  only by an Excel run (the pilot of router Phase 5).

### Phase 7: text files in the sandbox (decision D36; done: 7a S24, D44; 7b S25, D45)

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
