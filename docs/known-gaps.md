# Known gaps: KRONOS Offline version 6 on the real KGC samples (21 Nov 2026)

Found by a dry run of the published version 6 (main `95bf184`) on Boss's samples in `samples/kgc/` (D101, D106): a
read-only check in a scratch folder, nothing from the samples copied into the repo. Boss's word on 21 Nov 2026: "record
and save the warning and the gaps and work on them later" (D125). None of them is fixed yet.

## What passed (so nobody redoes it)

- All 4 real workbooks open, are described and are read as Excel saved them, each in under a second.
- A small recipe on each real workbook gives a copy that keeps everything (`kept_everything`): every part of the file
  is kept byte for byte except the sheet, the styles and `workbook.xml`; pictures, print areas and merged cells stay.
- Each of the 4 real Word proposals, with markers added by hand, is filled with the proven Example 5 recipe; every part
  other than the text parts is unchanged; LibreOffice opens the result, and every page after the filled part reads the
  same as the original.

## The warning (gap 1)

**A recipe that writes over a formula gives no warning.** On one real WPS form, an `OP SET_VALUE` at the address of a
formula cell replaced the formula with plain text; `kept_everything` was true, `warnings` was empty, and Review showed
only the old and new values. The copy is "right" by the recipe's own words, but the person loses a formula without being
told.
- Fix: when a write replaces a formula, the result says so in `warnings` (and Review shows it first, like the other
  warnings to tick before the download): "H36 held a formula, =..., and now holds text".
- Where: `compiler/kronos_book.py` (`_warnings`, `_summary`), `web/app.js` (Review and the download ticks).
- Test: an invented look-alike with a formula cell; a recipe that overwrites it; the warning is present; a recipe that
  overwrites a plain cell gives none.

## The gaps

**Gap 2: the look-alike WPS form doesn't match the real ones.** On the look-alike (`data/templates/wps_form.xlsx`) the
second page's revision and date lines are in rows 35 and 36. On the two real forms they sit lower, and the row is not
the same in each form; on one of them, rows 35 and 36 hold formulas, so a recipe that writes by those addresses hits
gap 1. The repo's WPS revision recipe (`tests/data/template_wps.vcl`) refused both real forms only because a label
differs (a " :" at the end); that was luck, not a rule.
- Fix: add a second look-alike WPS (`wps_form_c`, invented values) whose second page is lower and holds a formula where
  the first look-alike has text; teach the 0.7 prompt and the corpus recipes to find a line by its text (FIND) instead of
  a fixed address; a recipe's PRE should check what it relies on.
- Where: `data/templates/make_templates.py`, `tests/test_templates.py`, `vcl/spec.md` and the prompt together (a hint,
  not a language change), `tests/data/template_wps.vcl`.

**Gap 3: a Word marker inside a table cell is accepted by the check and refused by the fill.** Section markers
(`{{BEGIN Name}}` and `{{END Name}}`) must be paragraphs of the document body. The real proposals lay their text out in
tables, so a person will naturally type a marker in a cell. `vcl_docx.template_fields` accepts it (it walks every
paragraph in the file), but `fill` looks only at the body's own children and says "the template has no section Name",
which doesn't tell the person what to change.
- Fix: `template_fields` refuses a marker that isn't in the body, naming the table, with the way to fix it ("move the
  marker out of the table, onto its own line above it"), in the check, so the fill never meets it. A new error code
  (`python compiler/gen_error_codes.py`).
- Where: `compiler/vcl_docx.py`, the tests of the Word fill, the guide's section 8.

**Gap 4: a Word recipe counts tables by position.** A recipe says `table=2`, proven on a look-alike with 5 tables. The
real proposals have 12 to 25 tables, so a recipe made for the look-alike fits a real template only if its tables are in
the same order. It refuses safely, but the Word fill won't work on Boss's own files until this is designed.
- Fix: this is a design change to VCL, so it comes as a VEP, not a patch: a table named by a marker (a field in its
  first cell, or a caption) instead of a number, with a migration for the recipes that use numbers. Boss approves the VEP
  first.
- Where: a new `docs/vep-` file, then `vcl/spec.md` and the prompt, the grammar, the validator, `vcl_interp.py`,
  `vcl_docx.py`.

## Order of work, when Boss says "go"

1. Gap 1 (the warning): small, no language change.
2. Gap 3 (the table marker): small, one new message code.
3. Gap 2 (the second look-alike and the hint): small; its recipes are new corpus items.
4. Gap 4: a VEP for Boss to read before anything is built.

Each is its own PR with an invented look-alike test, never the real file. The corpus batch of about 20 document jobs
(Phase 5) should wait for gaps 3 and 4, since the recipes it adds depend on how tables are named.
