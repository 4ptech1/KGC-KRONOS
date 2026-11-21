# VEP 0.7: KGC's own documents (Phase 3 of `docs/kgc-documents-plan.md`, D109)

VCL 0.7 lets a recipe fill KGC's own forms: a WPS, the estimation workbook, the enquiry register, and the Word estimation
sheet and technical proposal. It is built from the survey (`docs/kgc-documents.md`, D107) and tested only on the
look-alike templates (`data/templates/`, D108). Boss asked for Phases 1 to 3 in one session (D107).

Every 0.6 spec stays valid and gives the same results in 0.7 (E16). As for every earlier version, the toolchain learns
each part first (spec.md and the prompt keep teaching 0.6), and one last PR, the switch, teaches 0.7 to the
Orchestrator prompt, so Boss pastes it once.

## Parts, one PR each

| Part | What | Runs where |
| --- | --- | --- |
| 3a | `FIXTURE <file> template=<name>`: a practice workbook from a template, with formulas and their saved values, merged cells and number formats | everywhere |
| 3b | merged cells: writing inside a merged range, `MERGE`, the `MERGED` assertion | everywhere |
| 3c | `ROUND`, `ROUNDUP`, `MIN`, `MAX` as value functions; `TEXT` stays refused in SET_FORMULA | everywhere |
| 3d | two workbooks: `TEST ... read=<file>`, a read-only Workbook argument whose formulas give their saved values | KRONOS Offline only |
| 3e | `AI_TEXT("...")`: text the AI wrote, listed for a person to check | everywhere |
| 3f | `DOCUMENT`, the Word kind: fields, repeated rows, columns and sections, optional sections | KRONOS Offline only |
| 3g | the switch: spec.md and the prompt teach 0.7 | |

"Everywhere" means the interpreter, the compiler and ForgeLib (so the macro runs in Excel), and the VBA simulator.
"KRONOS Offline only" means the interpreter proves it and the compiler refuses it with a clear message (D105: macros
open no second workbook until March 2027, and VBA can't be proven to drive Word).

## 3a. Template FIXTUREs

```text
FIXTURE  wps_01.xlsx template=wps_form
```

- `<name>` is a file `data/templates/<name>.xlsx` (letters, digits, `_`). Every sheet of the template becomes a sheet of
  the practice workbook, in the template's order and with its names, which may hold spaces and other characters
  (`BOM-RM`, `Enq Summary`). A template FIXTURE is the only line for its file: it can't be mixed with `sheet=` lines
  (a `folder=` line may be added, as today).
- What the practice workbook holds: each cell's value; each formula's text and its saved value; number formats; solid
  fills; merged ranges. Fonts, borders, widths, pictures, print areas and comments are not part of the model (they don't
  change what a recipe reads or writes; KRONOS Offline keeps them, plan Phase 4a).
- The reader is the standard library (`compiler/vcl_template.py`: `zipfile` and `xml.etree`), so the toolchain and
  KRONOS Offline's page need nothing new. It refuses, with a message, what it doesn't read: shared formulas, array
  formulas, a link to another workbook, a cell holding an error value. A date is a number with a date format, read as a
  date.
- A formula cell is a `Formula` in the model, as SET_FORMULA's are, now with its saved value. Reading it in a TASK stops
  the TEST, as today (R009): a recipe that writes cells can't know what Excel would recalculate. (3d gives one
  exception.) A TEST checks it with `FORMULA(...)`; `VALUE(...)` of a formula cell stops the TEST as today, since the
  saved value may be stale once the recipe has changed the cells it reads.
- A TEST names a template sheet whose name isn't one word in single quotes, as for added sheets
  (`VALUE('BOM-RM'!Q10)`); in a TEST call, a bare sheet name must be one word, so a recipe on the estimation workbook
  takes the Workbook (`est_01.xlsx`) and reaches its sheets with `SHEET($wb, "BOM-RM")`.
- In Excel, `forge_fixtures.bas` builds the same workbook: values, formulas (written after the values), number formats,
  fills and merges. The simulator loads it from the same model.

## 3b. Merged cells

- The model keeps each sheet's merged ranges. `SET_VALUE`, `WRITE_ROW`, `COPY_ROW`, `WRITE_ARRAY` and `CLEAR` on a cell
  inside a merged range that is not its top-left cell stop the run (a new R code): Excel would either refuse or keep the
  value hidden. Writing the top-left cell writes the merged range's value. `INSERT_ROW`, `DELETE_ROW`, `SORT`, `DEDUPE`
  and `FILTER` on a sheet with merged ranges stop the run: they would cut a merged block in Excel.
- `MERGE(<range>)` merges a block (two or more cells), with the whole sheet in MODIFIES. It stops the run when a cell
  other than the top-left holds a value (Excel would drop it) or when the block overlaps a merged range that is not the
  same block; merging the same block again does nothing, so a re-run is safe.
- `BLOCK($ws, <row1>, <col1>, <row2>, <col2>)` is the block between two corners known only at run time (a new
  enquiry's item rows), for MERGE and anything else that takes a range (D113: RANGE takes only a quoted A1 text).
- Assertion `MERGED(<Sheet>!<range>)` is 1 when exactly that block is merged, else 0. The automatic unchanged-check also
  compares merged ranges.

## 3c. Rule functions

The rules of the survey (section 6) that VCL can't compute today:

| Function | Meaning (as Excel's worksheet function) |
| --- | --- |
| `ROUND(<x>, <n>)` | x rounded to n digits, a half away from zero, n from -15 to 15 written in place; `ROUND(1234.5, -3)` is 1000 |
| `ROUNDUP(<x>, <n>)` | x rounded away from zero to n digits; `ROUNDUP(288364, -3)` is 289000 (the estimate's selling price) |
| `MIN(<a>, <b>)`, `MAX(<a>, <b>)` | the smaller or larger of two numbers |

`PI()` was planned and left out (D111): the weight formulas that use it are Excel formulas already (SET_FORMULA
allows PI), and no job computes a weight in a recipe. Numbers only, guarded like any calculation (rule 15). Excel rounds the decimal it shows, not the binary double, so
`ROUND(2.675, 2)` is 2.68: `semantics.py` rounds the value's 15-digit decimal form, and ForgeLib calls Excel's own
`WorksheetFunction.Round` and `RoundUp` (the simulator gives them `semantics.py`'s meaning). The digits are a whole
number from -15 to 15 written in place. `TEXT` stays off the SET_FORMULA allowlist: the WPS's `TEXT(R33,"0.00")`
formulas are already in the template, and a recipe that needs a range as text builds it with `&` and `ROUND`.

## 3d. Two workbooks (KRONOS Offline only)

```text
ARGS     old As Workbook, ws As Worksheet, thk As Double
...
FIXTURE  wps_old_01.xlsx template=wps_form
FIXTURE  wps_new_01.xlsx template=wps_form
TEST     fixture=wps_new_01.xlsx read=wps_old_01.xlsx call=MakeWps(wps_old_01.xlsx, WPS, 16) assert=...
```

- `read=<file>` opens a second FIXTURE workbook beside the test workbook. In the call, its file name is that Workbook;
  bare sheet names still mean the test workbook's sheets.
- The read workbook is read-only: no MODIFIES line may name its argument, and the automatic check proves it unchanged.
  Writes already go only through `$sheet` arguments (rule 20), so nothing in it can be written.
- Because nothing in the read workbook changes, reading one of its formula cells gives the **saved value** (Excel's
  last calculation), the one exception to R009. KRONOS Offline will refuse a formula with no saved value (plan 4c).
- The validator marks a spec with a `read=` TEST "KRONOS Offline only" (a warning, once); the compiler refuses it
  with that message, so `vclc build` and KRONOS.exe never make a .bas that would need a second workbook open.

## 3e. AI_TEXT

```text
OP       SET_VALUE(CELL($ws, 81, 2), AI_TEXT("Clean the joint before each run."))
```

- `AI_TEXT("<text>")` is a text value: its quoted text, nothing more, in the interpreter and in VBA.
- It marks text the AI wrote (an assumption, a remark, a query) that no TEST can prove right. `vclc check` lists every
  AI_TEXT of a spec as a note, and KRONOS shows them in Review and Save as "written by the AI, not proven: check it"
  (plan 4d). The argument is a quoted string written in place, never an expression.

## 3f. DOCUMENT, the Word kind (KRONOS Offline only)

A third spec kind beside TASK and BUILD. It fills a Word template marked up as `data/templates/README.md` says.

```text
VCL 0.7
DOCUMENT
# ASSUME: The estimation workbook was saved in Excel after its last change.
GOAL     Fill the internal estimation sheet from the estimation workbook.
NAME     FillEstimateSheet
TEMPLATE estimate_template.docx
ARGS     wb As Workbook, estNo As String
FIELD    Est_No = $estNo
FIELD    Client = CELL(SHEET($wb, "Enq Summary"), 2, 5)
FIELD    Grand_Total = NUMBER_TEXT(CELL(SHEET($wb, "Summary"), 17, 4), "#,##,##0")
REPEAT   row table=2 from=RANGE(SHEET($wb, "Enq Summary"), "A2:AC3") Sl=1, Item=10, Qty=15
REPEAT   section Assumption from=[[AI_TEXT("Prices ex-works.")], [AI_TEXT("Rates are budgetary.")]] Text=1
FIXTURE  est_01.xlsx template=estimation
TEST     fixture=est_01.xlsx call=FillEstimateSheet(est_01.xlsx, "EST-9001") assert=FIELD_TEXT(Client) = "Example Refinery Ltd" AND TABLE_ROWS(2) = 4
```

- **Keys, in order:** GOAL, NAME, TEMPLATE, ARGS (opt), FIELD (one per field), REPEAT (opt), KEEP (opt), FIXTURE (opt),
  TEST (one or more). No OP, PRE, MODIFIES or ERR: a DOCUMENT changes no workbook.
- **TEMPLATE** names a `.docx` in `data/templates/` (in KRONOS Offline, the uploaded template). The validator reads it
  and checks the spec against it: every `{{Name}}` in the template is bound by exactly one FIELD or by a REPEAT that
  holds it, every name the spec binds is in the template, and every table number exists.
- **FIELD `<Name>` = `<expr>`**: the text that replaces `{{Name}}` wherever it stands outside a repeated part (body,
  header, footer). The expression is a VCL value expression over the ARGS (CELL, VALUE of a range, SHEET, the text
  functions, ROUND, AI_TEXT, `&`); its value becomes text as S6 says, so a number is written plainly unless the spec
  formats it with `NUMBER_TEXT(<x>, "<format>")`: formats `0`, `0.0`, `0.00`, `#,##0`, `#,##0.00`, and the Indian
  `#,##,##0` and `#,##,##0.00` (`2,89,000`); `DATE_TEXT(<date>, "dd Mmm yyyy" | "dd.mm.yyyy" | "dd/mm/yyyy")`.
- **REPEAT `row` | `column` | `section` `<target>` `from=<source>` `<Name>=<n>, ...`**: the target is `table=<n>` (the
  template's n-th table, counted in the template's order) for a row or column, or a section name. The source is a
  range (one repeat per row of it) or a list of rows written in place, `[[...], [...]]`, for values typed in the
  request (Boss's "no workbook" route, D107). Each `Name=<n>` binds a field to the source's n-th column. A repeated row
  is the table's one row whose cells hold the bound fields; a repeated column likewise; a section is everything between
  its `{{BEGIN}}` and `{{END}}` paragraphs. Zero source rows remove the row, column or section.
- **A repeat inside a section** adds `match=<n>=<Name>`: only the source rows whose n-th column is SAME as the section
  row's field `<Name>` (the proposal's nozzle rows of each item).
- **KEEP `<section>` IF `<condition>`**: an optional section, kept once (its fields come from the FIELDs) or removed;
  the proposal's `KEEP Nace IF $sour` with a Boolean argument `sour`.
- **Filling** keeps every run's formatting: a field is replaced inside the run that holds its `{{`, after the
  paragraph's runs are read together (Word splits typing over runs), and the header, footer, pictures, styles and
  page fields stay as they are. The interpreter writes the new .docx with `zipfile` and `xml.etree`; nothing changes the
  template.
- **Values from a workbook** are what it holds as Excel last saved it: a formula cell gives its saved value (3d's rule,
  since a DOCUMENT changes no workbook); a formula with no saved value stops the TEST.
- **TESTs** pass no `runs=` and no `read=` (a DOCUMENT is filled once and reads every workbook it is given); at
  least one checks `ERROR_COUNT = 0`; a DOCUMENT with no workbook (Boss's typed-values route) still names a small
  FIXTURE file for its TESTs.
- **TEST assertions** read the filled document: `FIELD_TEXT(<Name>)` (the text a FIELD wrote), `TABLE_ROWS(<t>)`,
  `TABLE_COLS(<t>)`, `TABLE_CELL(<t>, <row>, <col>)` (tables of the filled document, counted in its order),
  `SECTION_COUNT(<section>)`, `TEXT_HAS("<text>")` (1 when the document's text holds it), and `LEFT_OVER` (the `{{`
  left in the document: always checked to be 0, so a TEST also proves every field was filled). Each TEST also proves
  every workbook unchanged.
- **The compiler refuses DOCUMENT** with "a DOCUMENT is filled by KRONOS Offline, not by a macro" (no VBA).
- **Built (D115):** `compiler/vcl_docx.py` fills, `compiler/vcl_document.py` checks, `vcl_interp.run_document_tests`
  runs the TESTs; a FIELD is written `FIELD    <Name> = <expr>`; the order of filling is sections (with their
  matched repeats), then rows and columns, then KEEP; `vclc mutate` mutates a DOCUMENT's FIELD, REPEAT and KEEP
  lines and compares the filled documents.

## Error codes

New messages are written `coded("...")` and numbered by `compiler/gen_error_codes.py`; none is reused.

## How each part is proven

Each part's PR adds: grammar and both parsers, the validator and type checker rules with tests of every refusal, the
interpreter, the compiler and ForgeLib (or the compiler's refusal), the simulator (`vclc sim` proves the compiler's VBA
equals the interpreter on every spec), conformance cases for each value function (3c, 3e, NUMBER_TEXT, DATE_TEXT), the
semantics doc, and a practice spec in `tests/data/` on a look-alike template whose TESTs pass and whose mutants are
caught. Every check of CLAUDE.md runs green before each merge; no committed spec changes.
