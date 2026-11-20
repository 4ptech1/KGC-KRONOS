# Look-alike KGC templates (invented, for tests)

Practice copies of KGC's own documents (Phase 2 of `docs/kgc-documents-plan.md`, D108). They have the layout the
survey found (`docs/kgc-documents.md`) and **only invented names, numbers and prices**. Every TEST, FIXTURE and corpus
pair about KGC documents uses these files, never the real ones (D105).

They are made by `make_templates.py` (run `python data/templates/make_templates.py` from the repo root);
`tests/test_templates.py` checks that each committed file is what the script makes today, and what each file holds.

| File | Look-alike of | What it has |
| --- | --- | --- |
| `wps_form.xlsx` | the WPS form (ISO 15609-1) | one sheet `WPS`: 97 merged ranges, a logo and a weld sketch, print area `A1:I50` on two pages, three run rows (29 to 31) with range formulas, and the heat-input check in `K24:AC31` outside the print area (k-factor `N25`, limit `N26`; run 3 is OVER LIMIT, run 2's PQR run is outside the WPS range) |
| `wps_form_b.xlsx` | a second WPS on the same form (KGC's samples are two WPS) | the same layout with other numbers, PQR values and IDs (DM-WPS-105, DM-PQR-105), every run inside its limits: the old WPS a new one is made from (D114) |
| `estimation.xlsx` | the cost estimation workbook | the same 7 sheets (Summary, Inputs, BOM-RM, Conversion, Hydrotest Options, Enq Summary, Assumptions & Clarif.) and the same kinds of formulas: ring and plate weights, SUMIF and SUMIFS by scope, mark-ups, `ROUNDUP(...,-3)` selling prices, the register's columns |
| `enquiry_register.xlsx` | the marketing enquiry register | three enquiries: two with columns A to I merged over their item rows, one with a single item; Shell OD and kg/cm2 formulas; two cell notes; dates as text |
| `estimate_template.docx` | the internal estimation sheet | fields, a repeated row, a repeated column, a repeated paragraph and a repeated section, a header with a logo and fields, a footer with page numbers |
| `proposal_template.docx` | the technical proposal (Part A) | fields, repeated rows, a section repeated per item with a repeated row inside it, an optional section, a header with a logo |

The workbooks hold each formula's value as Excel saves it (worked out by the script; LibreOffice, made to recalculate,
gives the same 168 values), so a reader of "the last values saved by Excel" finds them.

## How a Word template is marked up

Boss turns his own Word file into a template the same way (`docs/kgc-documents-plan.md` Phase 4).

- `{{Name}}`: a field, anywhere: a paragraph, a table cell, the header or the footer. Names are letters, digits and
  `_`. Word may split what a person typed over several runs (`estimate_template.docx` has `{{End_User}}` in two
  runs, on purpose), so a filler reads a paragraph's runs together.
- **A repeated row:** a table row whose cells hold fields; the recipe says which table's row repeats and where the rows
  come from. The rows above it (headings) and below it (totals) stay as they are.
- **A repeated column:** the same, for a column (the estimation sheet's weight table: one column per item).
- **A repeated section:** everything between a paragraph `{{BEGIN Name}}` and a paragraph `{{END Name}}`, paragraphs
  and tables, repeated once per source row; inside it, fields come from that row, and a table may have its own repeated
  row (the proposal's nozzle table). The two marker paragraphs are removed.
- **An optional section:** a repeated section with zero or one source rows (the proposal's `{{BEGIN Nace}}`): kept
  when the job needs it, removed otherwise.
