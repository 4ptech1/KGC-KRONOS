# VCL 0.1: VBA Command Language (revision 3, 14 Oct 2026)

VCL is the contract between Claude (which writes specs) and VBA Forge (which turns them into
tested VBA). This file is the reference that `validator.py` and the training data follow.
The Orchestrator prompt contains the same reference text, so update both together.

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
    and date comparison with IS_NUMBER or IS_DATE in a nested IF.
16. Make re-runs safe: rebuild output sheets from scratch, and recalculate totals from
    source data instead of adding to or subtracting from stored values. The goal TEST
    uses runs=2 to prove it.
17. Prefer KIND Function returning a count (rows written, flagged or removed), so TESTs
    can check RESULT.
18. Always test with a named synthetic FIXTURE. Never use fixture=target or real workbooks.
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
  EXIT_IF <condition>

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
Text uses "double quotes". Also numbers, TRUE, FALSE, NOTHING.
Operators: = <> < > <= >= AND OR NOT + - * / &
Columns are numbers (1 = A) everywhere except inside "A1 ref" strings.
An array element can be a SET_VALUE target: SET_VALUE(ARR(data, r, 3), 0).

SEMANTICS
  IS_EMPTY    true for empty cells, empty text, and text made only of spaces
  IS_NUMBER   true for numbers and for text that is a plain number ("25"); false for blanks
  IS_DATE     true only for real Excel date values; text that looks like a date is false
  COL         matches the header ignoring case and extra spaces; 0 if not found
  Dictionary  keys are compared ignoring case; Dictionary and Collection VARs start empty
  PRE         if a condition is false: one ForgeLog entry, nothing changes, and a Function
              returns -1
  AND / OR    both sides may be evaluated, so guard risky comparisons with a nested IF

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
  FIXTURE <name>.xlsx sheet=<name> cols=<A:Header,B:Header,...> rows=<n> data=<col>:<rule>;<col>:<rule>
  Rules: values("a"x<n>,"b"x<n>)  seq(<start>,<step>)  int(<lo>,<hi>)  num(<lo>,<hi>)
         dates(<yyyy-mm-dd>,<yyyy-mm-dd>)  text(<length>)
  Add +blank(<n>) to any rule to leave n random cells of that column blank.
  values(...) counts add up to rows. rows=0 creates the sheet with its header row only.
  Repeat FIXTURE lines with the same file name to add more sheets.
  Sheet names have no spaces, so TEST calls can use them bare. They may be short forms of
  the real sheet names (WelderPerf for "Welder Performance"); ASSUME lines describe the
  real workbook.
  Each column has its own fixed seed, made from file name + sheet + column. The builder
  uses it for everything random in that column: the shuffled order of values(...), the
  int, num, dates and text values, and which cells +blank(<n>) leaves blank. Columns are
  not correlated, and the same spec always builds the same file.
```

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
