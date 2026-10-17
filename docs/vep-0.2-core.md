# VEP 0.2 core: the ten additions of VCL 0.2

Status: **design, accepted by Claude under the working policy** (pick the conservative option,
record it, carry on). Boss can overturn any choice; each one has a `DECISIONS.md` id (D16).
Nothing here is in the language until its implementation PR lands. Each lands with spec.md,
the grammar, parser and validator, type checker, interpreter, compiler, ForgeLib, conformance
cases, `docs/vcl-semantics.md` and tests.

Every addition below removes a pattern the 10 specs already spell out the long way, or a gap
the M2 research found. The aim is shorter specs that say what they mean, so the spec-to-VBA
dataset teaches intent, not workarounds.

## How it lands

| PR | Contents |
| --- | --- |
| A | Value functions: SAME, CLEAN_TEXT, IS_ERROR, PARSE_DATE, IIF, MOD |
| B | Loop and write ops: SKIP_IF, EXIT_LOOP_IF, WRITE_ROW |
| C | MODIFIES and the automatic UNCHANGED check |
| D | `vcl_migrate.py`, the 10 specs moved to `VCL 0.2`, the Orchestrator prompt switched |

Until PR D, specs and the prompt stay `VCL 0.1`. The new features work only under a
`VCL 0.2` version line, and spec.md documents them in a "VCL 0.2 additions" section outside the
text it shares with the prompt. PR D then moves that section into the shared reference text, so
the Orchestrator learns 0.2 in one step. Every `VCL 0.1` spec stays valid under 0.2.

## The value functions (PR A)

### SAME(a, b): equal as a person reads it

```text
SAME(<expr>, <expr>)   TRUE when TRIM(a) and TRIM(b) are the same text, ignoring case
```

- Both sides become canonical text (S6), then TRIM (S2), then compare ignoring case.
- So `SAME(" ready ", "READY")` is TRUE, `SAME(25, "25")` is TRUE, and `SAME(25, "25.0")` is
  FALSE. Numbers are not compared as numbers; use IS_NUMBER and `=` for that.
- `SAME(<empty>, "")` is TRUE.
- Replaces the `UPPER(TRIM(x)) = "READY"` pattern (6 times in 5 of the 10 specs) with
  `SAME(x, "Ready")`.
- ForgeLib: `ForgeSame(a, b)`.

### CLEAN_TEXT(x): tidy typed text

```text
CLEAN_TEXT(<expr>)   TRIM, then every run of spaces, non-breaking spaces, tabs and line
                     breaks inside becomes one space; characters below code 32 are removed
```

- Case is kept. Same whitespace rule as COL's header match (S9), now available to values.
- `CLEAN_TEXT("  W-01 " & Chr(160) & Chr(10) & "  A ")` is `"W-01 A"`.
- ForgeLib: `ForgeCleanText(v)`.

### IS_ERROR(x)

```text
IS_ERROR(<expr>)   TRUE for an Excel error value (#N/A, #DIV/0!, #VALUE!, #REF!, #NAME?,
                   #NUM!, #NULL!); FALSE for everything else, including the text "#N/A"
```

- Decided as S5 in revision 4; this implements it.
- ForgeLib: `ForgeIsError(v)` (VBA `IsError`).

### PARSE_DATE(x, "<format>"): text dates, without guessing

```text
PARSE_DATE(<expr>, "<format>")   a real date, or empty when x is not a date in that format
formats: "dd/mm/yyyy"  "dd-mm-yyyy"  "dd.mm.yyyy"  "mm/dd/yyyy"  "yyyy-mm-dd"
```

- A real date value passes through unchanged (time of day dropped).
- Text is TRIMmed, then must match the format exactly. Day and month take 1 or 2 digits, and
  the year takes 4, from 1900 (D17). The date must exist, so 31/02/2026 gives empty.
- The format is required. "03/04/2026" means 3 April or 4 March depending on the PC, so VCL
  never guesses (S4).
- The result is a Variant, so the type checker requires the usual guard:
  `LET(d, PARSE_DATE(CELL($ws, r, c), "dd/mm/yyyy"))` then `IF IS_DATE(d)`.
- ForgeLib: `ForgeParseDate(v, fmt)`, which uses `DateSerial` on the parsed parts and never
  `CDate` or `DateValue`, because those follow regional settings.

### IIF(cond, a, b)

```text
IIF(<condition>, <expr>, <expr>)   a when the condition is TRUE, else b
```

- **Both a and b are always evaluated**, as in VBA's `IIf` and as with AND/OR in VCL. So an
  IIF branch is never a guard: the type checker treats both branches as unguarded.
- Values only: IIF can't return a Range, Worksheet, Workbook or Dictionary.
- ForgeLib: `ForgeIif(cond, a, b)`.

### MOD(a, b)

```text
MOD(<expr>, <expr>)   a - b * INT(a / b), as Excel's MOD: the result has the sign of b
```

- `MOD(7, 3)` is 1, `MOD(-7, 3)` is 2, and `MOD(7.5, 2)` is 1.5.
- VBA's `Mod` operator is not used, because it rounds both sides to whole numbers first
  (banker's rounding) and gives the sign of a.
- `MOD(a, 0)` is a run-time error, like `/` by 0. The standard handler logs it.
- It is a calculation, so its arguments need IS_NUMBER guards (rule 15).
- ForgeLib: `ForgeMod(a, b)`.

## The loop and write ops (PR B)

### SKIP_IF and EXIT_LOOP_IF

```text
SKIP_IF <condition>         go on with the next pass of the innermost loop
EXIT_LOOP_IF <condition>    leave the innermost loop
```

- Both are structural ops (they start an OP line), and they are allowed only inside LOOP_ROWS,
  LOOP_COLS or FOR_EACH. The validator rejects them elsewhere.
- EXIT_IF stays as it is and leaves the procedure (E1).
- They replace the "whole loop body inside `IF NOT IS_EMPTY(...)`" pattern, which nests every
  step one level deeper:

  ```text
  LOOP_ROWS dir=down from=2 to=LASTROW($weldLog, cSeam) var=r
    SKIP_IF IS_EMPTY(CELL($weldLog, r, cSeam))
    ...
  ```

- VBA: SKIP_IF becomes `If <cond> Then GoTo Next_<n>` with a `Next_<n>:` label just before
  the loop's `Next`, as EXIT_IF already does with `GoTo CleanExit`. EXIT_LOOP_IF becomes
  `Exit For`. VBA has no `Continue`, and every VCL loop compiles to `For` or `For Each`, so
  these are the only forms needed.
- A dir=up loop that deletes rows works unchanged, because SKIP_IF only skips forward.

### WRITE_ROW(ws, row, v1, v2, ...)

```text
WRITE_ROW(<ws>, <row>, <expr>, <expr>, ...)   writes the values to columns 1, 2, 3, ... of row
WRITE_ROW(<ws>, <row>, <expr>, ..., col=<n>)  starts at column n instead of 1
```

- Each value is written as SET_VALUE writes it.
- It needs at least one value.
- It replaces runs of `SET_VALUE(CELL($out, 1, 1), "...")`, `SET_VALUE(CELL($out, 1, 2), "...")`
  and so on. For example, task_04's header goes from 5 lines to 1.
- `col` is already a reserved option word.
- VBA: one assignment per value, exactly as SET_VALUE compiles. It is not a single array
  write, so each cell gets exactly what SET_VALUE would give it.

## MODIFIES and the automatic UNCHANGED check (PR C)

### MODIFIES

```text
MODIFIES  $<arg>                         anything on that sheet: values, fills, formats, rows
MODIFIES  $<arg>["<header>", ...]        values in those columns (row 2 down) only
MODIFIES  $<arg> fills                   fills only
```

- It is a new key, written after PRE. One line per sheet, or per kind of change.
- An argument may have one bracketed line and one `fills` line.
- The validator checks it two ways:
  - **Every write op names a declared sheet.** The ops are SET_VALUE, WRITE_ROW, FORMAT, CLEAR,
    COPY_ROW (target), INSERT_ROW, DELETE_ROW, SORT and WRITE_ARRAY. A write to an undeclared
    sheet is an error.
  - **Nothing is declared but never written.** A declared sheet no op writes to is a warning.
- Columns can't be checked statically, because COL() finds them at run time. The UNCHANGED
  check below covers them.

### UNCHANGED (automatic, no syntax)

For every TEST of a spec with MODIFIES, the interpreter compares the workbook before the first
run with the workbook after the last run.

- **Every sheet:** cells outside what MODIFIES allows must be identical, as values, fills and
  number formats.
- **Column-declared sheets:** the header row and every other column are compared too. Row
  inserts and deletes count as changes unless the whole sheet is declared.
- **A TEST that stops at PRE:** the whole workbook must be unchanged.

A difference fails the TEST with a line such as
`UNCHANGED: Inventory!A5 was "B-105", now "B-150"`.

This is stronger than today's hand-written collateral TESTs (`COUNTA(...) = PRE_COUNTA(...)`),
which only count cells. Those TESTs stay in the specs, because tests are never removed to make
room. The Excel batch runner, when it returns, gets the same check from `dist/expected/`.

A spec without MODIFIES gets no automatic check. From PR D the validator warns about that, and
every migrated spec has MODIFIES.

## Migration (PR D)

`python compiler/vcl_migrate.py specs/*.vcl` rewrites a `VCL 0.1` spec into `VCL 0.2`. It
makes only proven-equivalent rewrites:

| 0.1 pattern | 0.2 |
| --- | --- |
| `UPPER(TRIM(x)) = "LIT"` with LIT already upper case and trimmed | `SAME(x, "LIT")` |
| `UPPER(TRIM(x)) <> "LIT"` (same condition) | `NOT SAME(x, "LIT")` |
| a loop whose whole body is one `IF NOT IS_EMPTY(x)` block | `SKIP_IF IS_EMPTY(x)` then the body, one level up |
| 2 or more consecutive `SET_VALUE(CELL($o, r, k), v)` with k = 1, 2, 3, ... | `WRITE_ROW($o, r, v1, v2, ...)` (or `col=` when k starts above 1) |
| (none) | MODIFIES lines from the sheets the write ops name: a whole sheet when it is CLEARed or rows are added or removed, `fills` when FORMAT is its only write, otherwise the written columns' headers |

The migration counts as done only when, for every spec:

1. the validator reports 0 errors and no new warnings
2. every TEST passes, and every TEST's final workbook and RESULT are identical before and after
3. the snapshot mutation score is not lower
4. the new UNCHANGED check passes on every TEST

The literal in a SAME rewrite keeps its original upper-case text (`SAME(x, "READY")`). Changing
the case would be a cosmetic edit beyond what was proven.

## Found while writing this

SET_VALUE compiles to a plain `.Value = <expr>`. When the value is text that Excel would read as
something else ("1-2", "25", "TRUE", "=A1"), Excel converts it. The interpreter keeps it as
text, so the two can disagree. The literal texts in the 10 specs are plain labels ("Short",
"Days Open"), but some computed values are TRIMmed keys: task_01 writes part number "1001" and
task_04 writes seam "101" as text, which Excel stores as numbers. Every current TEST gives the
same result either way, because COUNTIF and the snapshot compare canonical text (S6), so nothing
fails today. It is still a real difference between the interpreter and Excel. The fixture builder already guards
this case (`needs_text_guard`). Proposed fix, with its own conformance case: SET_VALUE and
WRITE_ROW go through a ForgeLib `ForgeSetValue` that applies the same guard. It is tracked as
the first step of PR B.

## Out of scope for 0.2 core

- List syntax (`cols=[...]`), FILTER, DEDUPE with headers, SET_FORMULA, sheet and file ops:
  MVP plan item 4.
- Classes, events and form keys: later VCL 0.2 work (roadmap).
- NOTE stays the escape hatch.
