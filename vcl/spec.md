# VCL 0.4: VBA Command Language (20 Oct 2026)

VCL is the contract between Claude (which writes specs) and VBA Forge (which turns them into
tested VBA). This file is the reference that `validator.py` and the training data follow.
The Orchestrator prompt contains the same reference text, so update both together.
The formal syntax is `vcl/grammar.lark`; the reference meaning of the value functions is
`vcl/semantics.py` (tables in `docs/vcl-semantics.md`); the reasons behind each decision are
in `DECISIONS.md`.

Specs start with `VCL 0.4`. The validator still reads `VCL 0.3`, `VCL 0.2` and `VCL 0.1` specs:
every 0.3 spec is also valid as 0.4 and gives the same results once a DEDUPE `cols=1,2` is
written `cols=[1, 2]` (E12), every 0.2 spec is also valid as 0.3 (E8), every 0.1 spec is also
valid as 0.2, and `compiler/vcl_migrate.py` rewrites 0.1 specs into 0.2.

Target: Excel 2016 desktop. Specs never rely on Microsoft 365 functions such as XLOOKUP or
dynamic arrays; every op compiles to plain VBA. (The `FILTER` op hides rows itself, without
AutoFilter; it is not Microsoft 365's FILTER function, which SET_FORMULA refuses.)

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
10. Delete rows only inside LOOP_ROWS dir=up or with DEDUPE, and only rows the user
    explicitly wants removed. Never delete rows because data is missing or pending: copy
    them to an exceptions sheet instead (DEDUPE never deletes a row with a blank key).
11. For more than about 1,000 rows: READ_ARRAY, process the array, WRITE_ARRAY, and add
    PERF screen_off calc_manual.
12. Naming: procedures PascalCase and verb-first (DeleteRowsByValue); modules mod +
    PascalCase (modCleanup); variables camelCase; fixtures snake_case_01.xlsx.
13. If the user's message is not a new request (for example a question about your last
    spec), answer in at most two plain sentences, then reissue the spec if it changed.
14. Find input columns by header name with COL(), never by fixed position, and check
    every header you use with a PRE line.
15. Assume data is typed by hand: trim IDs before matching, compare typed text with SAME,
    and guard every calculation and date comparison with IS_NUMBER or IS_DATE in a nested
    IF. Guard with the check itself (IF IS_DATE(x)), never with a flag variable set earlier.
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
    uses TODAY. Give each column a SKIP_IF IS_EMPTY tests at least one blank cell
    (+blank(<n>) or a "" boundary value), and give a spec that copies or reads files a
    locked= file in a FIXTURE folder. Prefer exact assertions to ranges.
20. Declare every sheet the macro changes with MODIFIES: the whole sheet when it is
    rebuilt, cleared, sorted, filtered or has rows added or removed; ["<header>", ...]
    when only those columns' values change; fills when only fills change; files for a
    String argument whose folder the macro adds folders or files to; sheets for a
    Workbook argument the macro adds sheets to. A spec that changes nothing has no
    MODIFIES lines.
21. Work with files only through MAKE_FOLDER, COPY_FILE, LIST_FILES, READ_TEXT and
    WRITE_TEXT, on paths built in place as PATH_JOIN($<root>, SAFE_NAME(...), ...) from a
    String argument (LIST_FILES may list the root itself). Check every root with
    FOLDER_EXISTS in PRE, and write each file op's status where the user can see it. No op
    moves, renames, overwrites or deletes a file.
22. Add sheets only with ADD_SHEET or COPY_SHEET, each with name=, on a Workbook argument
    declared with MODIFIES $<book> sheets, and skip the add when SHEET_EXISTS says the
    sheet is there, so a re-run is safe. Write formulas only with SET_FORMULA, using
    Excel 2016 functions; a TEST checks their text with FORMULA, never their value. No op
    removes or replaces a sheet.
```

## Reference

```text
TASK FORMAT  (keys in this order; (opt) = optional)
VCL 0.4
TASK
# ASSUME: <assumption>                        (opt, one line each)
PROC     <PascalCase name>
KIND     Sub | Function
MODULE   <modName>
ARGS     <name> As <Type>, ...                (opt)
RET      <Type>                               (Function only)
VAR      <name> As <Type>                     (opt, one line per local variable)
PRE      <condition>; <condition>             (opt, checked before any change)
MODIFIES $<arg> | $<arg>["<header>", ...] | $<arg> fills | $<arg> files | $<arg> sheets   (opt, one line each)
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
automatically; never repeat them in VAR. READ_ARRAY fills a Variant VAR, LIST_FILES a
Collection, ADD_SHEET and COPY_SHEET a Worksheet; status VARs and READ_TEXT's text VAR are
String. VARs start at 0, "" or FALSE; Dictionary and Collection VARs start empty.
For form events: MODULE is the form, PROC is <control>_<Event> (btnRun_Click).

BUILD FORMAT
VCL 0.4
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
<one complete TASK spec per SIG, each starting with VCL 0.4, separated by lines
containing only --->

BUILD rules: modMain always has RunHeadless (runs the whole tool without the form, so the
harness can test it) and Run (shows the form, or calls RunHeadless if there is no form).
Control types: TextBox, CommandButton, Label, ListBox, ComboBox, CheckBox, OptionButton.

OP SYNTAX
Structural ops begin an OP line:
  LOOP_ROWS dir=up|down from=<expr> to=<expr> var=<name>
  LOOP_COLS dir=left|right from=<expr> to=<expr> var=<name>
  FOR_EACH var=<name> in=SHEETS(<book>)|CELLS(<range>)|KEYS(<dict>)|ITEMS(<list>)
  IF <condition> THEN <action>       (one-line form)
  IF <condition>                     (block form; steps below are nested)
  ELSE
  SELECT_CASE <expr>                 (block form)
  CASE <value>, <value>
  CASE_ELSE
  EXIT_IF <condition>                (leaves the procedure when the condition is true)
  SKIP_IF <condition>                (inside a loop: goes on with its next pass)
  EXIT_LOOP_IF <condition>           (inside a loop: leaves that loop)

Every other op is a call, NAME(arguments, key=value):
  Assign  LET(<var>, <expr>)
  Data    READ_ARRAY(<range>, into=<var>)
          WRITE_ARRAY(<var>, to=<range>)
          DICT_ADD(<dict>, <key>, <value>, mode=overwrite|keep|sum|count)
          DICT_GET(<dict>, <key>, into=<var>)
          SORT(<range>, key=<col>, order=asc|desc, header=yes|no)
          FILTER(<range>, col=<n>, crit=<expr>)
          DEDUPE(<range>, cols=<n>|[<n>, <n>, ...], header=yes|no)
  Cells   SET_VALUE(<target>, <expr>)
          WRITE_ROW(<ws>, <row>, <expr>, <expr>, ..., col=<n>)
                  (the values go to columns n, n+1, ...; without col= they start at column 1)
          SET_FORMULA(<target>, "<formula>", style=a1|r1c1)
          FORMAT(<target>, numfmt="<format>", bold=yes|no, fill=<name>|#RRGGBB, width=auto|<n>)
          CLEAR(<target>, what=all|contents|formats)
          COPY_ROW(<fromWs>, <fromRow>, <toWs>, <toRow>, cols=<n>)   (values of columns 1 to n)
          INSERT_ROW(<ws>, <row>)
          DELETE_ROW(<ws>, <row>)
  Sheets  ADD_SHEET(<book>, name=<expr>, after=<sheet>, into=<var>)    (a new empty sheet)
          COPY_SHEET(<sheet>, name=<expr>, after=<sheet>, into=<var>)  (a copy, in its workbook)
          OPEN_BOOK, SAVE_AS and CLOSE_BOOK open and save real workbooks: refused until March 2027
  Files   MAKE_FOLDER(<path>, into=<status>)       (the folder, and any missing ones below its root)
          COPY_FILE(<from>, <to>, into=<status>)   (one file; never overwrites, never deletes)
          LIST_FILES(<folder>, "<pattern>", into=<list>)   (file names in the folder, in name order)
          READ_TEXT(<path>, <text var>, into=<status>)     (a text file's text into <text var>)
          WRITE_TEXT(<path>, <expr>, into=<status>)        (a new text file; never overwrites)
  Calls   CALL(<module>.<Proc>, <args>)
  Output  LOG(<expr>)   MSG(<expr>)   RETURN(<expr>)
  Escape  NOTE "<terse technical text>"
FORMAT fill names: none (removes the fill), black, white, red, green, blue, yellow, magenta,
cyan. Write any other colour as #RRGGBB. FILL_COUNT takes #RRGGBB: black #000000, white
#FFFFFF, red #FF0000, green #00FF00, blue #0000FF, yellow #FFFF00, magenta #FF00FF, cyan
#00FFFF.

EXPRESSIONS
Value functions:
  LASTROW(<ws>, <col>)  LASTCOL(<ws>, <row>)  USED_RANGE(<ws>)
  COL(<ws>, "<header>")   column number of a row-1 header, or 0 if it is missing
  CELL(<ws>, <row>, <col>)  ROW_RANGE(<ws>, <row>, <fromCol>, <toCol>)  RANGE(<ws>, "<A1 ref>")
  SHEET(<book>, "<name>")  SHEET_EXISTS(<book>, <name>)  THIS_BOOK
  FIND(<range>, <expr>, match=whole|part)   the first matching cell, or NOTHING
  ROWS(<array>)  COLS(<array>)  ARR(<array>, <row>, <col>)
  COUNT(<dict or list>)  EXISTS(<dict>, <key>)
  LEN  TRIM  UPPER  LOWER  LEFT(<expr>, <n>)  RIGHT(<expr>, <n>)
  IS_EMPTY  IS_NUMBER  IS_DATE  IS_ERROR  IS_NOTHING  TODAY  NOW  FILE_NAME(<path>)
  SAME(<expr>, <expr>)   TRUE when TRIM(a) and TRIM(b) are the same text, ignoring case
  CLEAN_TEXT(<expr>)   TRIM, and every inner run of whitespace becomes one space
  PARSE_DATE(<expr>, "<format>")   a real date, or empty when it isn't one in that format;
                       formats: "dd/mm/yyyy" "dd-mm-yyyy" "dd.mm.yyyy" "mm/dd/yyyy" "yyyy-mm-dd"
  IIF(<condition>, <expr>, <expr>)  MOD(<expr>, <expr>)
  SAFE_NAME(<expr>)   CLEAN_TEXT of it, or "" when that is not safe as one file or folder name
  PATH_JOIN($<root>, <part>, ...)   $<root>, then each part, joined by backslashes
  FOLDER_EXISTS(<path>)   TRUE when that folder exists
Arguments are written $name. Variables and loop variables are bare names.
Text uses "double quotes"; write a quote inside text as two quotes: "Say ""hi""".
Also numbers, TRUE, FALSE, NOTHING. A list [<n>, <n>, ...] is only for DEDUPE cols=.
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
  SAME        compares TRIM(a) and TRIM(b) as text, ignoring case: SAME(" ready ", "READY")
              and SAME(25, "25") are TRUE, SAME(25, "25.0") is FALSE; an error is SAME as its
              text, so check IS_ERROR first when that matters
  CLEAN_TEXT  also removes characters below code 32; case is kept
  IS_ERROR    true for Excel error values (#N/A, #DIV/0!, ...); the text "#N/A" is not one
  PARSE_DATE  a real date passes through; text is TRIMmed and must match the format exactly
              (day and month 1 or 2 digits, year 4 digits from 1900, a date that exists);
              anything else gives empty, so guard the result with IS_DATE. VCL never guesses
              whether 03/04/2026 is March or April
  IIF         both values are always evaluated, so the condition never guards them; values
              only, never a Range, Worksheet, Workbook or Dictionary
  MOD         a - b * INT(a / b), as Excel: MOD(-7, 3) = 2, MOD(7.5, 2) = 1.5; MOD(a, 0) is a
              run-time error; numbers only, guarded like any calculation
  SKIP_IF     past a SKIP_IF or EXIT_LOOP_IF line its condition is FALSE, so
              SKIP_IF NOT IS_NUMBER(x) guards x for the rest of that pass
  Writes      SET_VALUE, WRITE_ROW and COPY_ROW store text as exactly that text, even text
              Excel would read as a number, date, TRUE/FALSE, error or formula ("1001", "1-2")
  LEFT RIGHT  a count below 0 gives ""; a count beyond the length gives the whole text
  COL         matches the row-1 header ignoring case, extra spaces and line breaks; the first
              match wins; 0 if not found
  LASTROW     the last row holding any value (spaces too), like End(xlUp); 1 if the column
              is empty. LASTCOL works the same along a row
  Dictionary  keys are compared as text ignoring case, not trimmed (TRIM them first), so
              1001 and "1001" are one key;
              Dictionary and Collection VARs start empty
  PRE         if a condition is false: one ForgeLog entry, nothing changes, and a Function
              returns -1
  EXIT_IF     leaves the procedure; a Function then returns 0, "" or FALSE (use IF ... THEN
              RETURN(<value>) to return something else)
  AND / OR    both sides are always evaluated, so a guard on the left does not protect the
              right: guard risky comparisons with a nested IF on the check itself
  SAFE_NAME   "" when the CLEAN_TEXT result is empty, over 100 characters, holds any of
              \ / : * ? " < > |, is "." or "..", ends with a dot, or is a device name (CON,
              PRN, AUX, NUL, COM1-COM9, LPT1-LPT9, any case, any extension: con.txt)
  PATH_JOIN   trailing backslashes on $<root> are dropped; "" when $<root> is empty or a
              part is empty or not exactly its own SAFE_NAME, and a file op given "" reports
              Bad name
  Paths       a file op's path is PATH_JOIN($<root>, ...) of a String argument, never a bare
              root, and every part after the root is SAFE_NAME(...) or a quoted literal that
              is its own safe name, written in place (never a variable or a cell), so no path
              can climb out of its root. A SAFE_NAME result or a path is text elsewhere
  FOLDER_EXISTS  TRUE only for a folder, FALSE for a file or ""; in Excel the path must be
              written in full (C:\Jobs, not Jobs or C:\Jobs\)
  MAKE_FOLDER writes one status into its String VAR, checked in this order: Bad name ("" or
              with * or ?), Too long (over 247 characters), Exists, Blocked by file (a file
              has the folder's name, or the name of a folder on the way; then nothing is
              made), No folder (the root is missing), Created (with every missing folder
              between it and the root, never the root itself)
  COPY_FILE   writes, in this order: Bad name, Too long (over 258 characters), Missing (no
              file at <from>), No folder (make the folder first), Already there (same size:
              a re-run is safe), Conflict (another size, or a folder: never overwritten),
              Locked (open in another program), Copied (<from> stays). In Excel also Check
              copy (the copy came out a different size; it is kept) and Failed (logged to
              ForgeLog). A file op never stops the run
  LIST_FILES  the names (not paths) of the files directly in the folder whose whole name
              matches the pattern (* any run of characters, ? one character, case ignored),
              sorted by lowercase name; a missing or bad folder gives an empty list. ITEMS
              walks them in that order
  Text        READ_TEXT and WRITE_TEXT handle printable ASCII, tabs and line breaks, at most
              32,767 characters; anything else is Not text, because it reads differently on
              each PC. Their paths are printable ASCII too, else Bad name
  READ_TEXT   writes, in this order: Bad name, Too long, Missing, Too big, Locked, Not text,
              Read; with Read, <text var> holds the file's text without one final line
              break, else "". In Excel also Failed
  WRITE_TEXT  writes, in this order: Bad name, Too long, Not text, No folder, Conflict (a
              folder or a different file is there: never overwritten), Locked, Already there
              (the same text: a re-run is safe), Written (exactly the text, no line break
              added). In Excel also Check write and Failed
  RANGE       a cell (B2) or a block (B2:D9); whole columns, whole rows and names stop a TEST
  THIS_BOOK   the TEST's workbook; pass sheets as arguments rather than using the active one
  READ_ARRAY  always a 2-D array from 1, even for one cell; WRITE_ARRAY writes it from the
              top-left cell of to= and stores text as exactly that text
  SELECT_CASE a CASE matches when = would be TRUE; the first match runs, else CASE_ELSE
  FOR_EACH    over CELLS gives each cell's value, row by row; over SHEETS each sheet
  INSERT_ROW  a new empty row, visible, with no fill and no number format
  LOG MSG     LOG writes a line to ForgeLog and is not an error; MSG shows a message
  FILE_NAME   the text after the last backslash
  DEDUPE      keeps the first row of each key and deletes each later row with that key
              (whole sheet rows); cols= are sheet column numbers, as for SORT; key cells
              compare as their text, case ignored, not trimmed (" W-1 " and "W-1" differ); a
              row with an empty or error key cell stays
  FIND        looks row by row, left to right: match=whole (the default) means SAME,
              match=part means the cell's text holds TRIM(<expr>), case ignored (an empty
              text finds NOTHING). Test it with IS_NOTHING before writing to it. FIND gives
              that cell only, and VCL has no row number of a cell: to change another column
              of the matching row, loop with LOOP_ROWS and SAME
  FILTER      below the range's first row, a row stays visible when its cell in sheet column
              col= is SAME as crit=, and is hidden otherwise; it never deletes and never uses
              AutoFilter. A spec that FILTERs may not SORT
  SET_FORMULA writes the formula's English text (=SUM(B2:B9), or R1C1 such as =RC[-2]*RC[-1]
              with style=r1c1). Only Excel 2016 functions: no XLOOKUP, FILTER, LET, TEXTJOIN,
              IFS, INDIRECT, TEXT, VALUE, no @, no A1#, nothing outside the workbook. style=a1
              writes one cell; style=r1c1 writes the same text to every cell of the range.
              Later steps and TESTs can't read a formula's value; a spec that SET_FORMULAs
              may not INSERT_ROW, DELETE_ROW, DEDUPE or SORT. The text is fixed, so its
              references are fixed columns or offsets: check them in PRE, such as
              COL($ws, "Qty") = 3 for =RC3*RC4. A cell formatted as Text (@) stops the run,
              because Excel would keep the formula as text
  Sheet names ADD_SHEET and COPY_SHEET take 1 to 31 characters, no spaces at the ends, none
              of \ / ? * [ ] :, no apostrophe first or last, not History; a bad name or one
              already used (case ignored) stops the run before anything is added. Without
              after= the new sheet goes after the last one

MODIFIES  (one line per sheet the macro changes, or per kind of change)
  MODIFIES $<arg>                     anything on that sheet: values, fills, formats, rows
  MODIFIES $<arg>["<header>", ...]    values in those columns, row 2 down
  MODIFIES $<arg> fills               fills anywhere on that sheet
  MODIFIES $<arg> files               folders and files may be added below that String
                                      argument; nothing else on disk may change
  MODIFIES $<arg> sheets              sheets may be added to that Workbook argument; every
                                      sheet that was there stays as it was
  An argument has one whole-sheet line, or at most one [headers] line and one fills line.
  Every write op must go through a $sheet argument whose lines allow that change: values
  need [headers] or the whole sheet, fill-only FORMATs need fills, and anything else
  (other FORMATs, CLEAR what=all or formats, SORT, INSERT_ROW, DELETE_ROW, DEDUPE, FILTER)
  needs the whole sheet. Write through CELL, ROW_RANGE, USED_RANGE, RANGE or FIND of a
  $sheet, or through the into= VAR of ADD_SHEET or COPY_SHEET, which writes to the new
  sheet and which nothing else may set.
  MAKE_FOLDER's folder, COPY_FILE's <to> and WRITE_TEXT's file must be below a root with a
  files line; COPY_FILE's <from>, READ_TEXT's file and LIST_FILES' folder may be below any
  String argument.
  Every TEST then checks that nothing else changed: each other cell keeps its value, fill
  and number format, each row stays visible or hidden, and the header row never changes;
  no sheet is removed, and new sheets appear only with a sheets line; on disk nothing is
  removed or changed, and new folders and files appear only below a files root. A spec
  that changes nothing has no MODIFIES lines and gets the same check.

TEST ASSERTIONS  (checked on the test workbook after the run)
  COUNTIF(<Sheet>!<range>, <criteria>)  COUNTA(<Sheet>!<range>)  SUM(<Sheet>!<range>)
  COUNTIFS(<Sheet>!<range>, <criteria>, <Sheet>!<range>, <criteria>, ...)
                 rows where every range meets its criteria
  SUMIFS(<Sheet>!<sumRange>, <Sheet>!<range>, <criteria>, ...)
                 sum of sumRange over the rows where every range meets its criteria
  VALUE(<Sheet>!<cell>)  SHEET_EXISTS("<name>")
  ROWS_USED(<Sheet>)               the last used row (row 1, the headers, counts)
  FILE_EXISTS("<Folder>\<path>")   that file exists below a FIXTURE folder of the TEST
  FILE_COUNT("<Folder>[\<path>]")  files in that folder and in every folder below it
  FILL_COUNT(<Sheet>!<range>, #RRGGBB)   cells in the range with that fill colour
  VISIBLE_COUNT(<Sheet>!<range>)   non-empty cells of the range in rows that are not hidden
  FORMULA(<Sheet>!<cell>)          the formula text as SET_FORMULA wrote it (A1 or R1C1), or ""
  FILE_TEXT("<Folder>\<path>")     the exact text of that file, or "" when there is none
  RESULT         the Function's return value (from the last run)
  ERROR_COUNT    errors written to ForgeLog during the run
  PRE_<measure>  the same measurement taken before the first run, e.g. PRE_COUNTA(Data!A:A)
                 or PRE_COUNTIFS(Data![Status], "Open", Data![Owner], "QA")
COUNTIF, COUNTIFS and SUMIFS work as in Excel 2016: criteria such as "Ready", "<>" (not
blank), "<>B-109" or ">=5"; text matching ignores case. COUNTIFS and SUMIFS ranges must be
the same size, so use [Header] ranges from the same sheet.
Ranges by header: <Sheet>![Header] is that header's whole column from row 1, the header
cell included (5 values under a header: COUNTA, COUNTIF "<>" and VISIBLE_COUNT give 6);
<Sheet>![Header]<n> is that column's cell in row n. Prefer them to column letters.
An assertion names a sheet as one word, so it can read only sheets without spaces in
their names; SHEET_EXISTS("<name>") takes any name.
runs=<n> calls the procedure n times on the same test workbook before checking
(default 1).
Combine assertions with AND. Assertions have no NOT, TRUE or FALSE: write "is not there" as
SHEET_EXISTS("<name>") = 0 or FILE_EXISTS("<path>") = 0. In TEST call arguments, a bare
sheet name (Orders) means that worksheet in the test workbook, a bare folder name (Dump)
means that FIXTURE folder, the TEST's fixture file name (orders_01.xlsx) means the test
workbook, for a Workbook argument, and text and numbers are passed as written ("P-104", 5).
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
  FIXTURE <name>.xlsx folder=<Name> [files=<path>[:<bytes>|:"<text>"],...] [locked=<path>,...]
  adds a synthetic folder to the same test file: each file below it (\ between folders,
  each part a safe name of letters, digits, _ . and -), with that size or one from its own
  seed, or holding exactly "<text>" (one line of printable ASCII). locked= files act as if
  another program has them open. A folder exists when a file is below it. <Name> is one
  word and not a sheet's name. Nothing real is read or written: the folders sit in a
  sandbox (C:\KRONOS_SANDBOX in the interpreter, a new folder below %TEMP% in Excel, whose
  root may be up to 120 characters), so a TEST fails when a file op path could be Too long
  in Excel but not in the interpreter: keep paths far over or well under the limits.
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

VCL 0.3 folder FIXTUREs: a file without `:<bytes>` gets 100 + below(900) bytes, from the
PRNG seeded with "<file>|<folder>|<path as written>|size", for example
"router_01.xlsx|Dump|r3.pdf|size". Such a file holds that many "x" characters. A VCL 0.4
file written `:"<text>"` holds exactly that text, with no PRNG draws.

## Changes in VCL 0.4 (20 Oct 2026)

Designed in `docs/vep-0.4-rev3-rest.md` (the rest of revision 3, MVP plan item 2) and built in
its Phases 1 to 7. Every VCL 0.3 spec stays valid and gives the same results, once a DEDUPE
`cols=1,2` is written `cols=[1, 2]` (E12). Details and reasons: `DECISIONS.md` (E12, S20 to
S25, D36 to D46).

- The version line is `VCL 0.4` (E12).
- Every revision 3 op and function that compiled now also runs in the interpreter, so a TEST
  can prove it: arrays, INSERT_ROW, SELECT_CASE, LOG, MSG, RANGE, SHEET, FILE_NAME (S20), and
  FOR_EACH over SHEETS and CELLS (D37).
- DEDUPE, with key columns as a list, `cols=[1, 2]`; a row with a blank key stays (S21, D38).
- FIND and IS_NOTHING, by VCL's own matching (S22, D39).
- ADD_SHEET and COPY_SHEET with `name=` (S23), SHEET_EXISTS, Workbook arguments passed as the
  TEST's fixture file, and `MODIFIES $<book> sheets` (D40, D41).
- FILTER hides rows by SAME, without AutoFilter, and the VISIBLE_COUNT assertion (D42).
- SET_FORMULA on an Excel 2016 list of functions, and the FORMULA assertion (D43).
- LIST_FILES and ITEMS (S24); READ_TEXT and WRITE_TEXT with statuses (S25); text files in
  folder FIXTUREs and the FILE_TEXT assertion (D44, D45).
- Hard rules 10, 19, 20 and 21 grow to cover these, and the new rule 22 covers sheets and
  formulas. OPEN_BOOK, SAVE_AS and CLOSE_BOOK stay refused until March 2027 (D36).
- The specs stay in their versions (11 in 0.2, task_12 in 0.3); Examples 1 and 3 moved to 0.4
  by their first line only (D46).
- After a review with Orchestrator dry runs, the wording was corrected where it differed from
  the tools, with no change of meaning: a [Header] range includes the header cell, DEDUPE and
  dictionary keys are not trimmed, assertions have no NOT, TRUE or FALSE, and the VAR types
  the ops fill, the fill colours, ROWS_USED, FIND and formula references are spelled out (D47).
- A review of ForgeLib made three meanings the same in Excel and the interpreter on every PC:
  DEDUPE keys ignore case letter by letter, a READ_TEXT or WRITE_TEXT path outside printable
  ASCII is Bad name, and SET_FORMULA refuses a cell formatted as Text (D48).

## Changes in VCL 0.3 (20 Oct 2026)

Designed in `docs/vep-0.3-file-ops.md` (the Job-Pack Router) and built in router Phases 1 to 4.
Every VCL 0.2 spec stays valid and gives the same results (E8). Details and reasons:
`DECISIONS.md` (E8 to E11, S18, S19, T4, F6, F7, D26 to D30, D32, D33).

- The version line is `VCL 0.3` (E8).
- Value functions SAFE_NAME, PATH_JOIN and FOLDER_EXISTS, and the path rules of the type
  checker: no path can climb out of its root (S18, S19, T4).
- MAKE_FOLDER and COPY_FILE, which write a status into a String VAR instead of stopping the
  run (E9, E11), and `MODIFIES $<root> files` (E10). Nothing moves, overwrites or deletes a
  file.
- Folder FIXTUREs, the FILE_EXISTS and FILE_COUNT assertions, and the UNCHANGED check for
  files (F6, F7), run on a mock file system by the interpreter; in Excel, ForgeLib's file
  helpers (D29) and a sandbox builder (D30). A TEST fails when Excel's longer sandbox root
  could change a Too long status (D30).
- Hard rule 21 (files); rule 20 names files; rule 19 asks for a blank cell in each column a
  SKIP_IF IS_EMPTY tests and a locked= file when a spec copies files (D22, D33).
- `specs/task_12.vcl` RouteJobPacks is the first 0.3 spec (D32); the 11 others stay 0.2.
  Example 1 moved to 0.3 (its first line only), and the prompt gained Example 3, which copies
  files (D33).

## Changes in VCL 0.2 (17 Oct 2026)

Designed in `docs/vep-0.2-core.md` and built in five PRs (A, B, C, D1, D2). Every VCL 0.1 spec
stays valid. Details and reasons: `DECISIONS.md` (E3 to E7, S5, S12 to S17, T3, D17 to D20).

- The version line is `VCL 0.2` (E3).
- Value functions SAME, CLEAN_TEXT, IS_ERROR, PARSE_DATE, IIF and MOD (S5, S12 to S16, T3).
- SKIP_IF and EXIT_LOOP_IF inside loops (E4), and WRITE_ROW (E5).
- SET_VALUE, WRITE_ROW and COPY_ROW store text as exactly that text (S17).
- MODIFIES, a key after PRE, and the automatic UNCHANGED check on every TEST (E6, E7, D20);
  hard rule 20.
- The 10 specs and Example 1 were moved to 0.2 by `compiler/vcl_migrate.py`, with proven
  rewrites only (D20).

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
