# VCL 0.1: VBA Command Language (revision 5, 17 Oct 2026)

VCL is the contract between Claude (which writes specs) and VBA Forge (which turns them into
tested VBA). This file is the reference that `validator.py` and the training data follow.
The Orchestrator prompt contains the same reference text, so update both together.
The formal syntax is `vcl/grammar.lark`; the reference meaning of the value functions is
`vcl/semantics.py` (tables in `docs/vcl-semantics.md`); the reasons behind each decision are
in `DECISIONS.md`.

Target: Excel 2016 desktop. Specs never rely on Microsoft 365 functions such as XLOOKUP or
dynamic arrays; every op compiles to plain VBA. (`FILTER` here means AutoFilter, which
Excel 2016 has.)

## Hard rules

```text
1.  Reply with exactly one code block tagged vcl. No text before or after it.
2.  Never write VBA. Express all logic with VCL keys, ops and expressions only.
3.  Use only the keys, ops, functions and syntax defined here. Never invent new ones.
4.  If a step truly cannot be expressed with ops, write it as NOTE "<terse technical text>".
    Use NOTE rarely.
5.  Every spec has at least two TEST lines: one proves the goal was met, one proves nothing
    else changed (for read-only work: proves the source data is unchanged). At least one
    TEST includes ERROR_COUNT = 0. When the test workbook has another sheet, add a
    wrong-sheet TEST that passes it and expects RESULT = -1 AND ERROR_COUNT = 1.
6.  Put every assumption at the top of the spec as a line starting with # ASSUME:
    Prefer a sensible assumption over a question.
7.  Ask only when a wrong guess could destroy or overwrite data, or the request cannot be
    specified at all. Then the code block contains only: # QUESTION: <one short question>
8.  Use ALLOW only for risky actions the user explicitly asked for: file_delete,
    file_overwrite, shell, http. Otherwise design around them.
9.  Choose TASK for one procedure of up to about 60 lines of VBA. Choose BUILD when the
    request needs several procedures, a UserForm, or more than about 60 lines.
10. Delete rows only inside LOOP_ROWS dir=up, and only rows the user explicitly wants
    removed. Never delete rows because data is missing or pending: copy them to an
    exceptions sheet instead.
11. For more than about 1,000 rows: READ_ARRAY, process the array, WRITE_ARRAY, and add
    PERF screen_off calc_manual.
12. Naming: procedures PascalCase and verb-first (DeleteRowsByValue); modules mod +
    PascalCase (modCleanup); variables camelCase; fixtures snake_case_01.xlsx.
13. If the user's message is not a new request (for example a question about your last
    spec), answer in at most two plain sentences, then reissue the spec if it changed.
14. Find input columns by header name with COL(), never by fixed position, and check
    every header you use with a PRE line.
15. Assume data is typed by hand: trim IDs before matching, and guard every calculation
    and date comparison with IS_NUMBER or IS_DATE in a nested IF. Guard with the check
    itself (IF IS_DATE(x)), never with a flag variable set earlier.
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
    uses TODAY. Prefer exact assertions to ranges.
```

## Reference

```text
TASK FORMAT  (keys in this order; (opt) = optional)
VCL 0.1
TASK
# ASSUME: <assumption>                        (opt, one line each)
PROC     <PascalCase name>
KIND     Sub | Function
MODULE   <modName>
ARGS     <name> As <Type>, ...                (opt)
RET      <Type>                               (Function only)
VAR      <name> As <Type>                     (opt, one line per local variable)
PRE      <condition>; <condition>             (opt, checked before any change)
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
automatically; never repeat them in VAR.
For form events: MODULE is the form, PROC is <control>_<Event> (btnRun_Click).

BUILD FORMAT
VCL 0.1
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
<one complete TASK spec per SIG, each starting with VCL 0.1, separated by lines
containing only --->

BUILD rules: modMain always has RunHeadless (runs the whole tool without the form, so the
harness can test it) and Run (shows the form, or calls RunHeadless if there is no form).
Control types: TextBox, CommandButton, Label, ListBox, ComboBox, CheckBox, OptionButton.

OP SYNTAX
Structural ops begin an OP line:
  LOOP_ROWS dir=up|down from=<expr> to=<expr> var=<name>
  LOOP_COLS dir=left|right from=<expr> to=<expr> var=<name>
  FOR_EACH var=<name> in=SHEETS(<book>)|CELLS(<range>)|KEYS(<dict>)|ITEMS(<collection>)
  IF <condition> THEN <action>       (one-line form)
  IF <condition>                     (block form; steps below are nested)
  ELSE
  SELECT_CASE <expr>                 (block form)
  CASE <value>, <value>
  CASE_ELSE
  EXIT_IF <condition>                (leaves the procedure when the condition is true)

Every other op is a call, NAME(arguments, key=value):
  Assign  LET(<var>, <expr>)
  Data    READ_ARRAY(<range>, into=<var>)
          WRITE_ARRAY(<var>, to=<range>)
          DICT_ADD(<dict>, <key>, <value>, mode=overwrite|keep|sum|count)
          DICT_GET(<dict>, <key>, into=<var>)
          SORT(<range>, key=<col>, order=asc|desc, header=yes|no)
          FILTER(<range>, col=<n>, crit=<expr>)
          DEDUPE(<range>, cols=<n>|<n,n>, header=yes|no)
  Cells   SET_VALUE(<target>, <expr>)
          SET_FORMULA(<target>, "<formula>", style=a1|r1c1)
          FORMAT(<target>, numfmt="<format>", bold=yes|no, fill=<name>|#RRGGBB, width=auto|<n>)
          CLEAR(<target>, what=all|contents|formats)
          COPY_ROW(<fromWs>, <fromRow>, <toWs>, <toRow>, cols=<n>)   (values of columns 1 to n)
          INSERT_ROW(<ws>, <row>)
          DELETE_ROW(<ws>, <row>)
  Sheets  ADD_SHEET(<book>, name=<expr>, after=<sheet>, into=<var>)
          COPY_SHEET(<sheet>, after=<sheet>, into=<var>)
          OPEN_BOOK(<path>, readonly=yes|no, into=<var>)
          SAVE_AS(<book>, <path>, format=xlsx|xlsm|csv)
          CLOSE_BOOK(<book>, save=yes|no)
  Files   LIST_FILES(<folder>, "<pattern>", into=<var>)
          READ_TEXT(<path>, into=<var>)
          WRITE_TEXT(<path>, <expr>, mode=new|append)
  Calls   CALL(<module>.<Proc>, <args>)
  Output  LOG(<expr>)   MSG(<expr>)   RETURN(<expr>)
  Escape  NOTE "<terse technical text>"
FORMAT fill names: none (removes the fill), black, white, red, green, blue, yellow, magenta,
cyan. Write any other colour as #RRGGBB.

EXPRESSIONS
Value functions:
  LASTROW(<ws>, <col>)  LASTCOL(<ws>, <row>)  USED_RANGE(<ws>)
  COL(<ws>, "<header>")   column number of a row-1 header, or 0 if it is missing
  CELL(<ws>, <row>, <col>)  ROW_RANGE(<ws>, <row>, <fromCol>, <toCol>)  RANGE(<ws>, "<A1 ref>")
  SHEET(<book>, "<name>")  THIS_BOOK  ACTIVE_BOOK  ACTIVE_SHEET
  FIND(<range>, <expr>, match=whole|part)   returns a Range or NOTHING
  ROWS(<array>)  COLS(<array>)  ARR(<array>, <row>, <col>)
  COUNT(<dict or collection>)  EXISTS(<dict>, <key>)
  LEN  TRIM  UPPER  LOWER  LEFT(<expr>, <n>)  RIGHT(<expr>, <n>)
  IS_EMPTY  IS_NUMBER  IS_DATE  IS_NOTHING  TODAY  NOW  FILE_NAME(<path>)
Arguments are written $name. Variables and loop variables are bare names.
Text uses "double quotes"; write a quote inside text as two quotes: "Say ""hi""".
Also numbers, TRUE, FALSE, NOTHING.
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
  LEFT RIGHT  a count below 0 gives ""; a count beyond the length gives the whole text
  COL         matches the row-1 header ignoring case, extra spaces and line breaks; the first
              match wins; 0 if not found
  LASTROW     the last row holding any value (spaces too), like End(xlUp); 1 if the column
              is empty. LASTCOL works the same along a row
  Dictionary  keys are compared as text ignoring case, so 1001 and "1001" are one key;
              Dictionary and Collection VARs start empty
  PRE         if a condition is false: one ForgeLog entry, nothing changes, and a Function
              returns -1
  EXIT_IF     leaves the procedure; a Function then returns 0, "" or FALSE (use IF ... THEN
              RETURN(<value>) to return something else)
  AND / OR    both sides are always evaluated, so a guard on the left does not protect the
              right: guard risky comparisons with a nested IF on the check itself

TEST ASSERTIONS  (checked on the test workbook after the run)
  COUNTIF(<Sheet>!<range>, <criteria>)  COUNTA(<Sheet>!<range>)  SUM(<Sheet>!<range>)
  COUNTIFS(<Sheet>!<range>, <criteria>, <Sheet>!<range>, <criteria>, ...)
                 rows where every range meets its criteria
  SUMIFS(<Sheet>!<sumRange>, <Sheet>!<range>, <criteria>, ...)
                 sum of sumRange over the rows where every range meets its criteria
  VALUE(<Sheet>!<cell>)  ROWS_USED(<Sheet>)  SHEET_EXISTS("<name>")  FILE_EXISTS("<path>")
  FILL_COUNT(<Sheet>!<range>, #RRGGBB)   cells in the range with that fill colour
  RESULT         the Function's return value (from the last run)
  ERROR_COUNT    errors written to ForgeLog during the run
  PRE_<measure>  the same measurement taken before the first run, e.g. PRE_COUNTA(Data!A:A)
                 or PRE_COUNTIFS(Data![Status], "Open", Data![Owner], "QA")
COUNTIF, COUNTIFS and SUMIFS work as in Excel 2016: criteria such as "Ready", "<>" (not
blank), "<>B-109" or ">=5"; text matching ignores case. COUNTIFS and SUMIFS ranges must be
the same size, so use [Header] ranges from the same sheet.
Ranges by header: <Sheet>![Header] is the whole column under that row-1 header;
<Sheet>![Header]<n> is that column's cell in row n. Prefer them to column letters.
runs=<n> calls the procedure n times on the same test workbook before checking
(default 1).
Combine assertions with AND. In TEST call arguments, a bare sheet name (Orders) means that
worksheet in the test workbook.
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
  Sheet names have no spaces, so TEST calls can use them bare. They may be short forms of
  the real sheet names (WelderPerf for "Welder Performance"); ASSUME lines describe the
  real workbook.
  Each column has its own fixed seed, made from file name + sheet + column. The builder
  uses it for everything random in that column: the shuffled order of values(...), the
  int, num, dates, reldates and text values, and which cells +blank(<n>) leaves blank;
  +dupes(<n>) and stale=<n> have seeds of their own. Columns are not correlated,
  and the same spec always builds the same file.
```

## Fixture generation (exact algorithm)

The FIXTURE SPEC seed rule, in full. `compiler/vcl_fixtures.py` implements it, and the
interpreter, the expected results and the fixture-builder VBA all use that one
implementation. It must never change, because the expected results and the dataset depend
on it.

```text
seed       first 8 bytes (big-endian) of SHA-256("<file>|<sheet>|<column letter>"),
           for example "bom_grn_01.xlsx|BOM|C"
PRNG       SplitMix64 on that seed; below(n) = uniform 0..n-1 by rejection sampling
values()   every value in the order written, then a Fisher-Yates shuffle
           (i from the last index down to 1, j = below(i + 1))
seq()      start + i * step, no randomness
int()      lo + below(hi - lo + 1), row by row
num()      lo + (hi - lo) * next() / 2^64, rounded to 2 decimals, row by row
dates()    first date + below(days + 1) days, row by row
reldates() TODAY + lo + below(hi - lo + 1) days, row by row; TODAY is the day the file is
           built (the fixture-builder VBA writes Date + n)
text(n)    n letters A-Z, each "ABC...Z"[below(26)], row by row

Then the modifiers, always in this order:
+boundary(v1..vk)  rows 2 to k+1 hold exactly v1 to vk. For values() each vi is moved up
           (swapped with the first later row that holds it), so the counts stay the same;
           for the other rules vi replaces the generated value. No PRNG draws.
+blank(n)  a partial Fisher-Yates over the rows below the boundary rows (all rows when
           there are none) picks the n rows to blank, drawing from the column's PRNG
+dupes(n)  its own PRNG, seeded from "<file>|<sheet>|<column letter>|dupes": a partial
           Fisher-Yates over the non-blank rows below the boundary rows picks n targets;
           each target, in pick order, takes the value of source[below(len(source))],
           where source is the other non-blank rows below the boundary rows
stale=n    a red (#FF0000) fill across the header's columns. rows=0: n new rows of the
           text "stale". With data: a partial Fisher-Yates with its own PRNG, seeded from
           "<file>|<sheet>|stale", picks the n data rows; their values stay
```

A values() entry of `""` is a blank cell, and so is a column without a data rule.

## Changes in revision 5

Fixtures that can tell a right macro from a wrong one, from the mutation evidence in
`docs/vcl-mvp-plan.md`. Valid revision 4 specs stay valid and build the same data, because
a FIXTURE line without the new parts is generated exactly as before. Details: `DECISIONS.md`
(F1 to F5).

- `stale=<n>` on a FIXTURE line: leftover rows of an earlier run (F1).
- `+boundary(v1,...)`: exact values in the first data rows (F2).
- `+dupes(<n>)`: repeated values, for dictionary keys (F3).
- `reldates(lo, hi)`: dates relative to TODAY (F4).
- Modifiers go in the order +boundary, +blank, +dupes, each at most once (F5).
- Rule 19: fixtures able to catch a wrong macro. The validator warns when a rebuilt sheet
  has no stale rows, a spec with TODAY has no reldates, a spec with a dictionary never
  repeats a key, or a spec with `<`, `<=`, `>` or `>=` has no +boundary.

## Changes in revision 4

Clarifications from the M2 research (`docs/vcl-m2-findings.md`); valid revision 3 specs stay
valid, except task_06, which was rewritten (T2). Details and reasons: `DECISIONS.md`.

- Option names are reserved words (G2): only they make `name=value` options, and they can't
  be names.
- A quote inside text is written as two quotes (G6).
- Operator precedence matches VBA, and comparisons don't chain (G7).
- SEMANTICS now states S1 to S11: emptiness, TRIM, numbers, dates, canonical text, `=` with
  text, LEFT/RIGHT, COL, dictionary keys and LASTROW behave the same on every PC.
- Rule 15: guard with the check itself, never with a flag (T2).
- `EXIT_IF` leaves the procedure; a Function then returns 0, "" or FALSE.
- New section above: the exact fixture generation algorithm.

## Changes in revision 3

- Synthetic fixtures only: `fixture=target` is removed and every TEST names a FIXTURE file
  (new hard rule 18, also in the Orchestrator's FINAL CHECK). Example 1 now builds
  `orders_01.xlsx`.
- `COUNTIFS(...)` and `SUMIFS(...)` assertions with Excel 2016 criteria; they work with
  `PRE_` and with `[Header]` ranges.
- Fixtures: each column has its own fixed seed (file name + sheet + column) for the order of
  `values(...)`, the other random values and the `+blank(<n>)` cells; `values(...)` counts
  add up to rows.
- Fixture sheet names may be short forms of the real names, with no spaces; ASSUME lines
  describe the real workbook.
- FORMAT fill names are listed: none, black, white, red, green, blue, yellow, magenta, cyan.
- This file now shows all hard rules, 1 to 18.

## Changes in revision 2

- `COL(<ws>, "<header>")`: find columns by header name, so specs survive column moves.
- `ROW_RANGE(<ws>, <row>, <fromCol>, <toCol>)`: one row's cells, for highlighting.
- `COPY_ROW(<fromWs>, <fromRow>, <toWs>, <toRow>, cols=<n>)`: copy a row's values.
- `FILL_COUNT(<Sheet>!<range>, #RRGGBB)`: assertion that counts highlighted cells.
- `<Sheet>![Header]` and `<Sheet>![Header]<n>`: assertion ranges by header name.
- `runs=<n>` on TEST: run the procedure n times before checking, to prove re-runs are safe.
- SEMANTICS section: exact meaning of IS_EMPTY, IS_NUMBER, IS_DATE, COL, dictionary keys,
  PRE failure (Function returns -1) and AND / OR evaluation.
- Fixtures: `<rule>+blank(<n>)` for deliberate blanks; `rows=0` for a header-only sheet.
- Rules 10 and 14 to 17: no deleting rows for missing data, header-based columns checked by
  PRE, dirty-data guards, safe re-runs, and Functions that return a count.
