# VBA Forge decisions

One line per decision: what was decided, why, and where it is enforced. Boss can overturn any
of them. Ids: **G** grammar, **S** semantics, **T** types (from `docs/vcl-m2-findings.md`),
**E** language rules, **F** fixtures, **D** decisions Claude made on its own under the rule "pick the more
conservative option, record it here, and carry on".

## Language (VCL 0.1 revision 4, 16 Oct 2026)

These clarify revision 3; they don't change any valid spec except task_06 (T2). spec.md and the
Orchestrator prompt state them in the reference text.

| Id | Decision | Why | Enforced by |
| --- | --- | --- | --- |
| G2 | Option names are reserved words: dir, from, to, var, in, mode, into, key, order, header, what, fill, numfmt, bold, width, style, cols, col, crit, readonly, format, save, after, name, match. Only these make `name=value` options (no space before `=`); any other `name=value` is a comparison. They can't be argument, variable or loop variable names. | `DICT_ADD(d, k, v, mode=sum)` and `SET_VALUE(c, got = 0)` look alike; reserved keys make the grammar unambiguous (it builds as LALR). | `vcl/grammar.lark` (OPTKEY), validator (reserved names, option splitting) |
| G6 | A quote inside text is written as two quotes, as in VBA: `"Say ""hi"""`. | Revision 3 never said how. | grammar (STRING), `compiler/parser.py` tokenizer and `values(...)` |
| G7 | Operator precedence matches VBA, tightest first: `-` (negation), `* /`, `+ -`, `&`, comparisons, NOT, AND, OR. Comparisons don't chain. | `a < b < c` would make VBA compare a Boolean with c, silently. | grammar |
| S1 | IS_EMPTY: empty cell, empty text, or text made only of spaces, non-breaking spaces (character 160), tabs or line breaks. An error cell is not empty. | Copy-pasted "blank" cells often hold character 160. | `vcl/semantics.py`, ForgeLib `ForgeIsEmpty` |
| S2 | TRIM removes spaces, non-breaking spaces, tabs and line breaks at both ends; inner spaces stay. | VBA's Trim removes only ordinary spaces. | semantics, `ForgeTrim` |
| S3 | IS_NUMBER: numbers, and text that is a plain number after TRIM (optional sign, digits, optional `.` and digits). False for blanks, TRUE/FALSE, dates, errors, "1,250", "1e3", "25%", "(25)". | Regional settings decide what "1,250" means. | semantics, `ForgeIsNumber` |
| S4 | IS_DATE is true only for real date values; text that looks like a date is false. | "03/04/2026" could be March or April. | semantics, `ForgeIsDate` |
| S5 | IS_ERROR (true for #N/A, #DIV/0!, ...) is decided, and arrives with VCL 0.2. In 0.1 an error is not empty, not a number and not a date, so most specs already treat it as invalid. | Reports need to detect error cells. | semantics (`is_error`), VCL 0.2 |
| S6 | Every value becomes text the same way on every PC: empty is "", whole numbers without ".0", "." as the decimal point, dates as yyyy-mm-dd, TRUE/FALSE, errors as their Excel text (#N/A). | VBA's CStr on a date follows regional settings, and Trim on a #N/A cell crashes. | semantics, `ForgeText` |
| S7 | `=` and `<>` with text on either side compare the canonical text exactly (case and spaces count), so a cell holding 25 equals "25". | In VBA `cell = "25"` is False when the cell holds the number 25. | semantics, `ForgeEq` |
| S8 | LEFT and RIGHT: a count below 0 counts as 0; a count beyond the length gives the whole text. | VBA raises an error for a negative count. | semantics, `ForgeLeft`, `ForgeRight` |
| S9 | COL matches a header by its canonical text with line breaks and other whitespace turned into single spaces, ends trimmed, case ignored; the first match wins; 0 if none. | Headers typed with Alt+Enter or double spaces must still match. | semantics, `ForgeCol` |
| S10 | Dictionary keys compare as canonical text, ignoring case: 1001 and "1001" are one key. | VBA's Dictionary treats 1001 and "1001" as different keys. | semantics, `ForgeDictAdd` and friends |
| S11 | LASTROW and LASTCOL work like End(xlUp) and End(xlToLeft): the last cell holding any value, spaces-only text included; an empty column gives 1. | Matches what the generated VBA does. | semantics, `ForgeLastRow`, `ForgeLastCol` |
| T2 | Guard a calculation or date comparison with the check itself (`IF IS_DATE(x)`), never with a flag set earlier (`IF NOT bad`). task_06 was rewritten this way; the rewrite was proved equivalent over all 20 input combinations. | No checker can follow a flag. | rule 15, FINAL CHECK, type checker |
| E1 | `EXIT_IF <condition>` leaves the procedure when the condition is true. A Function then returns its RET type's default (0, "" or FALSE), as VBA does; to return something else write `IF <condition> THEN RETURN(<value>)`. VCL 0.2 adds EXIT_LOOP_IF to leave only a loop. | Revision 3 never defined EXIT_IF. | interpreter, compiler (`GoTo CleanExit`) |

## Language (VCL 0.1 revision 5, 17 Oct 2026)

Fixture syntax from the mutation evidence (`docs/vcl-mvp-plan.md`). A FIXTURE line without the
new parts builds exactly the same data as before. spec.md and the Orchestrator prompt state
them in the reference text; spec.md has the exact algorithm.

| Id | Decision | Why | Enforced by |
| --- | --- | --- | --- |
| F1 | `stale=<n>` after `rows=<n>`: on a rows=0 sheet, n rows of the text "stale" with a red fill; on a sheet with data, a red fill on n data rows picked by their own PRNG ("<file>\|<sheet>\|stale"), values unchanged. stale may not exceed rows when rows > 0. | Without leftovers, a macro that forgets to clear its output sheet or old fills passes every TEST. | grammar (`stale_opt`), `vcl_fixtures.py`, validator |
| F2 | `+boundary(v1,...)` puts exactly these values in the first data rows. For `values()` the values are moved up, so the counts in the rule stay true; `""` is a blank cell; values must fit the rule (text, numbers, yyyy-mm-dd dates, whole days for reldates). No PRNG draws. | Random data almost never sits exactly on a `<` or `<=` edge, or puts the interesting row in row 2 where `from=2` vs `from=3` shows. | grammar, `vcl_fixtures.py`, validator (`check_modifiers`) |
| F3 | `+dupes(<n>)` gives n non-blank rows below the boundary rows the value of another such row, with its own PRNG ("<file>\|<sheet>\|<col>\|dupes"). | `mode=keep` and `mode=overwrite` only differ when a key repeats. | grammar, `vcl_fixtures.py`, validator |
| F4 | `reldates(lo, hi)`: TODAY + lo to TODAY + hi days. The generator keeps the offset; the interpreter resolves it with the run's TODAY and the fixture-builder VBA writes `Date + n`. | Fixed dates drift past TODAY, so a TEST that is exact today fails next month. With offsets, every result is the same on any day. | `vcl_fixtures.py` (`RelDate`), interpreter, `vba_literal` |
| F5 | Modifiers go in the order +boundary, +blank, +dupes, each at most once. +blank and +dupes only touch rows below the boundary rows. | One order means one algorithm, and the boundary rows stay exactly as written. | validator |
| E2 | Rule 19: a FIXTURE should be able to catch a wrong macro. The validator warns (not errors) when a sheet the macro rebuilds or whose fills it owns has no stale rows in a TEST that gets past PRE, a spec with TODAY has no reldates, a spec with DICT_ADD never repeats a key, or a spec with `<`, `<=`, `>` or `>=` has no +boundary. | Warnings, so every revision 4 spec stays valid; the 10 specs now have none. | validator (`check_fixtures_can_catch_mistakes`) |

## Language (VCL 0.2 value functions, PR A, 17 Oct 2026)

The first of the four VCL 0.2 PRs (`docs/vep-0.2-core.md`). spec.md lists them in its
"VCL 0.2 additions" section, outside the text it shares with the Orchestrator prompt, until the
migration PR.

| Id | Decision | Why | Enforced by |
| --- | --- | --- | --- |
| E3 | The first line may be `VCL 0.1` or `VCL 0.2`. The VCL 0.2 additions are errors under `VCL 0.1`. Every TASK inside a BUILD repeats the BUILD's own version line. | One spec, one language version; a 0.1 spec never silently uses 0.2. | `parser.py`, `vcl_parse.py`, validator |
| S12 | SAME(a, b): TRIM(a) and TRIM(b), as canonical text, are equal ignoring case, letter by letter (as LOWER). Never compared as numbers: SAME(25, "25.0") is FALSE. | Replaces `UPPER(TRIM(x)) = "READY"`, the most repeated pattern in the 10 specs. | semantics (`same`), `ForgeSame` |
| S13 | CLEAN_TEXT: characters below code 32 (other than tabs and line breaks) are removed, every run of spaces, non-breaking spaces, tabs and line breaks becomes one space, and the ends are trimmed. Case is kept. | Typed and pasted text needs the same tidy-up as COL's header match (S9). | semantics (`clean_text`), `ForgeCleanText` |
| S14 | PARSE_DATE(x, "<format>") with one of five formats: a real date passes through without its time; text is TRIMmed and must match exactly (day and month 1 or 2 digits, year 4 digits from 1900, the date must exist); anything else gives empty. | `CDate` and `DateValue` read "03/04/2026" by the PC's regional settings; a stated format never guesses. | semantics (`parse_date`), validator (format list), `ForgeParseDate` (DateSerial only) |
| S15 | IIF(cond, a, b) gives a or b; both are always evaluated, as VBA's IIf. | Matches the generated VBA exactly. | semantics (`iif`), interpreter, `ForgeIif` |
| S16 | MOD(a, b) = a - b * INT(a / b), as Excel's MOD: sign of b, decimals kept; b = 0 is a run-time error. | VBA's `Mod` rounds both sides (banker's rounding) and gives the sign of a. | semantics (`mod`), `ForgeMod` |
| T3 | IIF's condition never guards its values, so `IIF(IS_NUMBER(x), x * 2, 0)` is rejected. IIF of two values of one type has that type; of two different types it is a raw value that needs a guard. MOD takes numbers only: a raw cell needs IS_NUMBER, and a date is rejected. PARSE_DATE gives a raw value that needs IS_DATE. | Both IIF values run, as both sides of AND / OR do; MOD of a date has no clear meaning. | type checker |

## Decisions Claude made on its own

| Id | Date | Decision | Why it is the conservative choice |
| --- | --- | --- | --- |
| D1 | 16 Oct 2026 | The one fixture generator (`compiler/vcl_fixtures.py`) uses the pilot's SplitMix64 algorithm from `harness/fixtures.py` (PR #4), not the MVP's Python `random` module. It builds bit-for-bit the same data as the pilot. | The algorithm is written down in full (spec.md), so the data never depends on a Python version, and the pilot's expected results stay valid if PR #4 is revived. All specs still pass their TESTs on this data. |
| D2 | 16 Oct 2026 | The batch builder is `vclc build`: besides `forge_module.bas` it writes `forge_fixtures.bas`, one `ForgeFixture_<file>(wb)` builder per fixture file, from the same generator. | Boss asked for one generator used by the interpreter, the batch builder and the expected results; PR #4's runner is deferred, so nothing in main may depend on it. |
| D3 | 16 Oct 2026 | Only `dist/forge_module.bas` is committed (CI fails if it is out of date). The rest of `dist/` is rebuilt by `vclc build` and git-ignored. | The expected results depend on TODAY (task_06, task_09, task_10), so a committed copy would go stale every day. |
| D4 | 16 Oct 2026 | `compiler/validator.py` keeps its hand-written parser and rules and adds two stages: the grammar and the type checker. `vclc check` runs this validator. | Keeps every existing rule and message; a differential run (see the PR) showed where the checkers disagreed, and each disagreement was fixed or is now reported. |
| D5 | 16 Oct 2026 | Where the parsers disagreed, they were aligned on what spec.md already allows: a UTF-8 byte order mark and blank lines before `VCL 0.1` are fine; `IF ... THEN NOTE "..."` is fine. | Accepts only what the validator already accepted; no valid spec is rejected. |
| D6 | 16 Oct 2026 | Assertions allow `+` and `-` only, not `*` and `/`; `NOT` right after a comparison operator needs brackets. | That is what the grammar and the interpreter support; spec.md never listed more. Stricter is safer. |
| D7 | 16 Oct 2026 | ForgeLib stays in `vba/ForgeLib.bas` (the MVP's place), not in `library/`. The conformance cases stay in `tests/`. | `library/` is for approved code, and ForgeLib has never run in Excel. Fewer moves, fewer surprises. |
| D8 | 16 Oct 2026 | These clarifications are VCL 0.1 revision 4; the version line in specs stays `VCL 0.1`. | No valid spec's syntax changes. |
| D9 | 16 Oct 2026 | PRs #5 and #6 (Phase 1 and Phase 2 plans) stay open and untouched. No M0, M1 or M2-research PRs exist; the M2 research arrived inside the MVP zip. | They are proposals waiting for Boss, and nothing depends on them. |
| D10 | 16 Oct 2026 | A RETURN value is converted to the Function's RET type, as VBA does (a Long rounds half to even). | Matches the generated VBA; no spec's results change. |
| D11 | 16 Oct 2026 | Removed `pyflakes-4.0.3-py2.py3-none-any.whl` from the repo root. | It was committed by accident in 7c8e5fe; nothing uses it. |
| D12 | 17 Oct 2026 | task_01 as the pilot ran it is frozen in `tests/data/task_01_pilot.vcl`. The seed-stability test and the pilot-model test (26 shortages, sums 986 and 661) run on that copy instead of being rewritten for the new task_01. | Keeps both pinned tests exactly as strong as before; changing their numbers would have weakened them. |
| D13 | 17 Oct 2026 | Only the FIXTURE and TEST lines of the 10 specs changed for item 2 of the MVP plan, never an OP line. Where a column had to be told apart from its neighbour, a column was added (task_09 NCR gets `F:Remarks`) or given old values (task_08 Inventory `Balance` and `Remark`). Goal TESTs now assert exact results, worked out by the reference interpreter. | The specs' logic is untouched; only the test data and assertions got stronger. |
| D14 | 17 Oct 2026 | The 15 mutants that still survive are equivalent: they can't change any result. They are `LET(flagged, 0)` or `LET(unknown, 0)` on a Long that starts at 0 (task_02, task_06, task_08); the value or mode of a dictionary only used through KEYS or EXISTS (task_04 `seams` x8, task_08 `known`); the value of `mode=count`, which ignores it (task_03 x2); and task_10's `<` vs `<=` TODAY, which the next line overrides. They still count as survivors in the score. | Excluding them would make the score look better than the evidence; the score is 96.4% with them. |
| D15 | 17 Oct 2026 | The interpreter now moves fills and number formats with their rows on SORT and DELETE_ROW, and the expected results include number formats. | That is what Excel does; with stale fills on data sheets the old behaviour would have hidden real differences. |
| D16 | 17 Oct 2026 | The VCL 0.2 core is designed in `docs/vep-0.2-core.md` before any code, and lands in four PRs (value functions; loop and write ops; MODIFIES with UNCHANGED; migration). The key choices: SAME compares trimmed text ignoring case, never as numbers; PARSE_DATE needs an explicit format and never guesses; IIF evaluates both branches, so it is never a guard; MOD follows Excel, not VBA's `Mod`; MODIFIES is checked statically per sheet and at run time per cell; specs and the prompt stay `VCL 0.1` until the migration PR switches them all at once. | Each choice is the stricter or more predictable option, and nothing reaches the Orchestrator until all of 0.2 is proven on the 10 specs. |
| D17 | 17 Oct 2026 | Choices inside PR A that the VEP left open: SAME compares lower-case text with a binary compare, not VBA's `vbTextCompare` (which follows the PC's locale); PARSE_DATE rejects years before 1900, because Excel can't hold them and VBA's DateSerial reads a year below 100 as 19xx or 20xx; CLEAN_TEXT removes control characters before it joins spaces, so "a " & Chr(7) & " b" gives "a b"; a BUILD's TASK sections must use the BUILD's version line; spec.md's revision number and the Orchestrator prompt stay as they are, because nothing in the shared reference text changed. | Each is the stricter or more predictable option, and the Orchestrator Project needs no update until the migration PR. |
