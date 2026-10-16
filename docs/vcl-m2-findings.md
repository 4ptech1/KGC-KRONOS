# VCL M2/M3 research: grammar, types and semantics (findings)

Prototype work for milestones M2 and M3 of `docs/vcl-language-plan.md`, done outside the repo
and tested against the 10 specs (revision 3), Example 1 of the Orchestrator prompt and a BUILD
example. Nothing here changes the language yet. Every recommendation is a decision or a VEP
for you to approve.

## What was built (all tested: 31 unit tests, 66 conformance cases)

| File | What it is |
| --- | --- |
| `vcl/grammar.lark` | The formal grammar of VCL 0.1 rev 3, one start rule per key line |
| `compiler/vcl_parse.py` | Parser built on the grammar: file structure, key order, OP nesting. Never crashes |
| `compiler/vcl_types.py` | Type checker with flow-sensitive narrowing |
| `vcl/semantics.py` | Reference meaning of the core value functions |
| `tests/conformance_cases.py` | 66 conformance cases, shared by tests, docs and the future Excel batch |
| `docs/vcl-semantics.md` | Generated reference: decisions S1 to S11 plus the conformance tables |

## 1. Grammar findings

**G1. The language is unambiguous.** The grammar builds as LALR(1), which Lark refuses to do
for an ambiguous grammar. All 10 specs and Example 1 parse. The parser also survived 400
randomly damaged specs without crashing.

**G2. Options vs comparisons inside calls.** `DICT_ADD(d, k, v, mode=sum)` and
`SET_VALUE(c, got = 0)` look the same to a parser. The grammar resolves this by making option
keys reserved words: dir, from, to, var, in, mode, into, key, order, header, what, fill,
numfmt, bold, width, style, cols, col, crit, readonly, format, save, after, name, match.
*Recommendation:* make the reserved list official. The validator rejects VAR names from the
list, and a comparison on such a name needs parentheses.

**G3. `DEDUPE(..., cols=1,2, ...)` is ambiguous** (is `2` a column or the next argument?).
*Recommendation (VEP):* list syntax `cols=[1, 2]`. The grammar already supports lists, and the
VCL 0.2 DO step (`reads=[...]`) needs them too.

**G4. BUILD folder fixtures are undefined.** The BUILD example's `fixture=folder_01` doesn't
parse, because FIXTURE only defines `.xlsx` workbooks. *Recommendation (VEP):* a FOLDER
fixture (`FIXTURE folder_01/ files=3 pattern=*.xlsx from=<workbook fixture>`) or a rule that
BUILD tests use workbook fixtures only.

**G5. A FIXTURE header can't contain a comma** ("Qty, kg"), because commas separate columns.
*Recommendation (VEP):* allow quoted headers, `cols=A:"Qty, kg",B:Part No`.

**G6. Strings use doubled quotes as the escape**, like VBA: `"Say ""hi"""`. Revision 3 never
said so. *Recommendation:* add this one line to spec.md.

**G7. Operator precedence now matches VBA:** arithmetic, then `&`, then comparison, then NOT,
AND, OR. `a < b < c` is rejected; VBA would silently compare a Boolean with c.
*Recommendation:* add the precedence table to spec.md.

## 2. Type system (proposed for the validator)

**Types:** Num, Text, Date, Bool, Cell (a raw cell value: empty, number, text, date or error),
Sheet, Book, Range, Dict, Coll, Arr. VAR types map from VBA: Long/Double to Num, String to
Text, Variant to a flow-typed value.

**The core rule:** a raw Cell can't be used in arithmetic, in `<` `>` `<=` `>=`, or compared
with `=` to a number or date, and can't be assigned to a Long, Double or Date variable, until
an enclosing `IF IS_NUMBER(...)` or `IF IS_DATE(...)` narrows it. One-line IFs, ELSE branches
and EXIT_IF narrow too.

**T1. It works on real specs.** 9 of the 10 specs and Example 1 type-check as written.

**T2. task_06 can't be proven safe.** It guards a date comparison through a flag
(`IF NOT bad`) instead of the check itself. It is safe at runtime, but no checker can follow a
flag. The fix is a one-line change, proved equivalent over all 20 input combinations:

```
OP           IF NOT bad                                   (before)
OP           IF IS_DATE(expDate)                          (after)
OP             IF expDate < TODAY THEN LET(bad, TRUE)     (unchanged)
```

*Recommendation:* apply this edit, and add the style rule "guard with the check itself, never
with a flag".

**T3. Short-circuit trap caught.** `IF IS_NUMBER(x) AND x > 5` crashes in VBA when x is text,
because VBA evaluates both sides of AND. The checker rejects it. The spec already warned about
this in prose; now it is enforced.

**T4. Writes cancel guards precisely.** A write to the same cell, or a whole-sheet change
(CLEAR of a range, COPY_ROW, SORT, DELETE_ROW), cancels the guards it could affect. Writes to
other cells keep them. This is sound because PRE guarantees every COL() lookup found its
header, and different header names can't match the same header cell.

**T5. Dictionary values are tracked.** `DICT_GET(..., into=n)` into a Long is only allowed
when the dictionary's values are numbers (from `mode=sum` or `mode=count`, or guarded values).

**T6. 11 mistakes caught, 8 valid patterns accepted** (tests/test_types.py). Caught: the AND
trap, raw cell into a Long, raw arithmetic, a guard cancelled by a write or a SORT, an
unguarded mode=sum, a missing `$`, unguarded date comparison, cell = 0, a flag guard, and a
guard wrongly carried into a loop.

## 3. Semantics decisions (S1 to S11, full table in docs/vcl-semantics.md)

The most important ones, because plain VBA gets them silently wrong:

| Id | Decision | The VBA trap it avoids |
| --- | --- | --- |
| S1 | IS_EMPTY includes spaces, non-breaking spaces, tabs and line breaks | Copy-pasted "blank" cells holding character 160 |
| S2 | TRIM also removes non-breaking spaces and line breaks | VBA Trim only removes ordinary spaces |
| S3 | IS_NUMBER accepts "25" but rejects "1,250", "25%", "1e3" | Regional settings decide what "1,250" means |
| S4 | IS_DATE is true only for real dates | "03/04/2026" could be March or April |
| S6 | Values become text the same way on every PC: dates as yyyy-mm-dd, "." as the decimal point, errors as "#N/A" | CStr on a date depends on regional settings; Trim on a #N/A cell crashes with a type mismatch |
| S7 | A cell holding the number 25 equals the text "25" | In VBA, `cell = "25"` is False when the cell holds a number |
| S10 | Dictionary keys compare as text, ignoring case | VBA's Dictionary treats 1001 and "1001" as different keys |

**Consequence for ForgeLib:** every text conversion, comparison with text and dictionary key
must go through a ForgeLib helper. Inline VBA would hit these traps. In particular, any `#N/A`
in a register crashes a macro that calls VBA's Trim on it.

**S5, IS_ERROR (VCL 0.2 candidate).** Revision 3 has no way to detect error cells. Under S1 to
S4 an error is not empty, not a number and not a date, so most specs already treat it as
"invalid", which is the right outcome. An explicit IS_ERROR is still needed for reports.

## 4. Recommendations for the repo (in order)

1. **Decisions (DECISIONS.md):** reserved option words (G2), string escaping (G6), precedence
   (G7), semantics S1 to S11, and the style rule from T2. These clarify revision 3 rather than
   change it.
2. **Spec edit:** the task_06 rewrite (T2).
3. **M2 integration:** adopt `vcl/grammar.lark` as the normative syntax, with Lark pinned as a
   dependency (`lark==1.3.1`, pure Python). Either rebuild `compiler/parser.py` on it, or keep
   the hand-written parser and test it against the grammar on every spec. Add the type checker
   as a validator stage.
4. **M3 seed:** `vcl/semantics.py` becomes the core of the reference interpreter. The
   conformance cases move to `vcl/conformance/` and later become the Excel conformance batch.
5. **VEPs to write:** list syntax (G3), folder fixtures (G4), quoted headers (G5), IS_ERROR (S5).
