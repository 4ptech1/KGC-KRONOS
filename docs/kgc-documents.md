# The KGC documents survey (Phase 1 of `docs/kgc-documents-plan.md`, D107)

What KRONOS must fill in each of Boss's KGC samples (`samples/kgc/`, D106), where each value comes from, the jobs a
person will ask for, the engineering rules the samples' own formulas use, and what VCL 0.6 can't say yet. It was made
by reading copies of the samples outside the repo (`python -I`, openpyxl for the workbooks, `zipfile` and
`xml.etree` for the Word files); nothing in the repo reads the samples (`tests/test_samples.py`).

The survey names fields by their cell or table and their form label. It does not copy the samples' client names,
prices or people's names: the look-alikes of Phase 2 invent their own.

**Sources**, used in every table below:

| Source | Meaning |
| --- | --- |
| **sheet** | another cell of the same workbook (a formula today, or a value KRONOS copies) |
| **file** | the second uploaded file: another workbook (D105 answer 4), for example an older WPS or the estimation workbook |
| **typed** | the person's request ("for 16 mm plate", "client is ...") |
| **rule** | one of the engineering rules in section 6, each taken from a sample's own formula |
| **AI** | text the AI writes (assumptions, remarks, queries); KRONOS marks it `AI_TEXT`, "check this" |
| **fixed** | the form's own text: labels, format numbers, terms and conditions; KRONOS never changes it |

**Kinds:** text, number, date, formula (KRONOS writes the formula, Excel computes it), picture (KRONOS keeps it).

## 1. Welding Procedure Specification (WPS), Excel

Two files on one form, format QF/QA/31 to ISO 15609-1 (`WPS_KGC-P104-191-101_R1.xlsx`, `WPS_KGC-P105-101-101_R1.xlsx`).
One sheet each, 176 merged ranges, print area `A1:I87` (P104) and `A1:I91` (P105), portrait, one manual page break
(row 39 / 45, "Continued to Page No. 2"), 4 pictures: the KGC logo on each page (A1 and the page-2 header) and two weld
sketches (row 16 / 17, columns A and E). Columns A to I are the printed form; columns K to AC, outside the print area,
are the heat-input check.

**The two WPS do not share row numbers.** P104 has an extra "Application" row (row 8) and one run row; P105 has eight
run rows. Every field below therefore moves by a few rows between WPS: a recipe that reads one WPS and writes another
finds each field by **its label in column A**, not by a fixed address: VCL 0.6 can already do this with
`FIND(RANGE($ws, "A1:A100"), "Material Thickness", match=part)` and `ROW_OF`. (The look-alike template of Phase 2 has one
fixed layout, so a recipe for that template may use fixed cells.)

### Page 1

| Field (label in column A) | Cell (P105) | Kind | Source |
| --- | --- | --- | --- |
| Format number, ISO edition | H1, C3 | text | fixed |
| Date (`DATE:- dd/mm/yyyy`), on both pages | H2, H47 | text | typed (today by default) |
| WPS Revision (`R1`), on both pages | H3, H48 | text | typed |
| Company Name, Company Address | C4, C5 | text | fixed |
| Welding Procedure Specification No.; its revision note | C6; F6 | text | typed |
| WPQR No. (the PQR this WPS rests on) and its certificate | C7; F7 | text | file (the source WPS) or typed |
| Application (P104 only) | C8 | text | typed |
| Welding Process; Degree of mechanization; Mode of metal transfer | C8, C9, C10 | text | file or typed |
| Method of preparation and cleaning | C12 | text | file or typed |
| Welding Position: Actual / Qualified | C13 / F13 | text | file (Actual from the PQR) / rule-backed text |
| Joint design; joint and weld type; the sketches' notes | B14, D14, A15, D15, A21, D22 | text, picture | file (kept), typed |
| Weld Details: Actual / Qualified | C23 / F23 | text | file / text |
| Parent material type, designation, group, thickness (mm), pipe OD: Actual (part 1 # part 2) / Qualified | C25:E29 / F25:F29 | text, number | typed (Actual), rule (Qualified thickness: 0.5t to 2t, section 6) |
| Welding details, one row per run: run, process, filler size, current (A), voltage (V), type of current / polarity, travel speed (mm/min), heat input (kJ/mm) | A33:H42 | text; D, E, G, H are formulas | file (the run table), sheet (D, E, G, H from the heat-input block) |
| Notes between the runs (side A, back-gouge, side B) | A32, A37, A38 | text | file or typed |

### Page 2

| Field | Cell (P105) | Kind | Source |
| --- | --- | --- | --- |
| Column heads "Actual (PQR-nnn)" / "Qualified / WPS range" | C50 / F50 | text | typed (the PQR number) |
| Weld deposited thickness: Actual / Qualified | C51 / F51 | text | typed / rule (max 2t) |
| Filler designation, make and brand, size, baking: Actual / Qualified | C53:C56 / F53:F56 | text | file |
| Shielding and backing gas, flow rate: Actual / Qualified | C58:C61 / F58:F61 | text | file |
| Tungsten electrode, nozzle size, back gouging | C62:C64 / F62:F64 | text | file |
| Preheat, interpass, post-heating, pre-heat maintenance, PWHT | C65:C69 / F65:F69 | text | file, typed |
| Other information (oscillation, pulse, contact tube, plasma, single or multi run, electrodes, torch angle) | C71:C77 / F71:F77 | text | file |
| Heat input (kJ/mm): Actual / Qualified | C78 / F78 | text | sheet (the block's limits) |
| Impact test, hardness test (PQR) | C79 (P104: C73, C74) / F79 | text | file |
| Remarks 1 to 5 | B81:B85 | text | AI (marked) or file |
| Prepared by / Examiner: name, manufacturer, date | B89:B91, F89:F91 | text, date | typed (never AI: a person signs) |

### The heat-input check (outside the print area)

| Field | Cells (P105) | Kind | Source |
| --- | --- | --- | --- |
| k-factor of the process | N26 | number | rule (0.6 for process 141, as its label says) |
| WPS maximum heat input (kJ/mm); P104 also a minimum | N27 (P104: N27, N28) | number | file (the PQR's recorded maximum) |
| Per run: I min, I max (A), U min, U max (V), v min, v max (mm/min) | L:Q, rows 33 to 42 | number | typed or file |
| Per run: HI min, HI max | R, S | formula | rule (heat input) |
| Per run: "Max <= limit?" | T | formula | rule (`IF(S<=$N$27,"OK","OVER LIMIT")`; P104 also checks the minimum) |
| Per run: PQR I min, I max, U min, U max, v, HI recorded | V:AA | number | file (the PQR) |
| Per run: PQR HI calculated; "PQR run inside WPS range?" | AB, AC | formula | rule |
| Source of the PQR values (a note) | K44 | text | typed |

## 2. The enquiry register, Excel (`NEA_Enquiry-Summary-MKT.xlsx`)

One sheet, header in row 1, one row per **item**, 27 columns A to AA; no print area, no pictures.

- **An enquiry is a block of rows.** Columns A to I (Sr. No., Enquiry, Date, Result, End User, EPC/PMC, LSTK, Eng
  Contractor, Project) are written once per enquiry and **merged down** over its item rows (22 merged ranges, for
  example `A9:A12`, `I4:I8`). Some blocks merge only part of a column (`I2:I3` and `I4:I8` inside one enquiry), so a
  column may hold two sub-blocks.
- **Columns J to AA are per item:** Item, Type, U Stamp, Old Tag No, Revised Tag No, Qty, MOC, Shell OD, Shell ID
  (mm), Shell THK (mm), TL to TL (mm), Drawing unit weight (kg), Actual weight, Design pressure (kg/cm2), Design
  temperature, Sour/NACE, HIC/SSC testing, PWHT.
- **Formulas:** Shell OD `Q = R + 2*S` (rows 9 to 14); design pressure `W = ROUND(<barg>*1.01972,2)` with the bar value
  written inside the formula. Older rows hold typed values (OD as `12" NPS`, thickness as `SCH 100`, so numbers and text
  share a column).
- **Dates are text** (`13.11.2026`), not Excel dates; VCL already writes such text unchanged (S17's `'` prefix).
- **Cell notes** (comments) on 28 cells explain a value (where the shell thickness came from, the conversion used, why
  PWHT is or isn't needed). They are part of the record.

| Field | Kind | Source |
| --- | --- | --- |
| Sr. No. | number | sheet (the next number) |
| Enquiry, Date, Result, End User, EPC/PMC, LSTK, Eng Contractor, Project | text, date as text | typed, or file (the estimation workbook's Enq Summary) |
| Item, Type, U Stamp, tags, Qty, MOC | text, number | typed or file |
| Shell OD | formula | rule (Shell OD) |
| Shell ID, THK, TL to TL, unit weight | number | file or typed |
| Design pressure (kg/cm2) | formula | rule (bar to kg/cm2) from typed barg |
| Design temperature, Sour/NACE, HIC/SSC, PWHT | number, text | typed or file |
| Cell notes | text | AI (marked) |

## 3. The estimation workbook, Excel (`Estimation_Sheet_E-1604AB_Channel_FH.xlsx`)

Seven sheets, no merged cells, no pictures, about 270 formulas, all inside the workbook.

| Sheet | What a person fills | What the formulas do |
| --- | --- | --- |
| **Inputs** (`A4:E30`: Ref, Parameter, Value, Unit, Basis) | the rates and factors in column C: exchange rate C5, steel density C6, material rates C7:C13, deposition efficiency C14, labour, machining, PWHT, NDT, dish rates C15:C22, plate margin C23, scrap C24, forging allowance C25, overheads C26, contingency C27, margin C28, warranty C29, residual credit C30 | none (all typed; "yellow = confirm with vendor quotes") |
| **BOM-RM** (`A4:R25`, items in rows 5 to 21) | per item: Scope (CH or FH), BOM item no., description, material, form (Ring, Plate, Pipe, Flange, Bar), Qty/set, Sets, finished L or OD, W or ID, T or H, BOM finished weight, rate unit, remarks | H total qty, M raw weight each (ring or plate rule), N raw total, O rate from Inputs, Q amount; totals per scope in rows 23 to 25 (SUMIF on column A) |
| **Conversion** (`A4:J35`, activities in rows 5 to 30) | per activity: scope (CH, FH, ALL), activity, basis, qty/set, sets, unit, rate | G total qty, J amount, E for welding = weld metal / deposition efficiency; totals per scope J32:J35 |
| **Hydrotest Options** (`A4:L30`) | per adder: option (1 or 2), scope, item, basis, qty, unit, sizes, rate | J plate weight, L amount, the warranty provision from Summary, the residual credit; totals per option and scope (SUMIFS) |
| **Summary** (`A1:F32`) | the title A1, the end user line A2, notes F5:F11 | the cost build-up, overheads, contingency, both options, selling prices `ROUNDUP(.../2,-3)`, USD, price per kg |
| **Enq Summary** (`A1:AC5`) | the register's 27 columns for this enquiry's items | Shell OD, unit weight from BOM-RM, design pressure in kg/cm2, prices from Summary (AB, AC) |
| **Assumptions & Clarif.** (`A1:C17`) | numbered assumptions and clarifications | none |

Sources: Inputs are typed (rates) and stay "to be confirmed"; BOM-RM rows come from a file (a BOM) or typed; the
formulas are rules (section 6) and stay formulas, so Excel recomputes them; Summary notes and Assumptions are AI
(marked); Enq Summary feeds the register (job E5).

## 4. The internal estimation sheet, Word (`KGC_Estimation_Sheet_NEA-28982_INTERNAL.docx`, `...28983...`)

A filled document, no placeholders. Header: the estimation sheet number, the date and the line "INTERNAL - KGC
CONFIDENTIAL - NOT FOR SUBMISSION TO CLIENT"; footer: a fixed line and "Page n of m" (Word fields); one logo. The two
samples differ in item count (4 items against 2), which is how the fields below were found: what differs between them
is a field, what is the same is the form.

| Part | Shape | Fields | Source |
| --- | --- | --- | --- |
| Title (P1, P2) | two paragraphs | enquiry no., item line | typed |
| Table 1, enquiry facts | 6 rows of label and value | client, end user, enquiry, related proposal, price basis, rate basis | typed or file |
| Table 2, price summary | one row per item, then total, GST and total incl. GST | Sl., item, qty, weight per bottle, unit price, total, INR/kg | file (the estimation workbook) or typed; totals are rules |
| Table 3, optional items | one row per option, **one column per item**, then a total column | option name, price per item, total | file or typed |
| Section 3, basis and assumptions | a list of paragraphs | each assumption | AI (marked), typed |
| Table 4, preliminary design check (ASME VIII-1) | one row per parameter, **one column per item** | quantity, pressure (barg / kg/cm2), temperature, ID x TL-TL, t req, nominal thk, head thk, OD, MAWP, vacuum check, hydrotest pressure, flange class, under-rated classes, nozzle necks | typed or file; the conversion and OD are rules; the code calculations (UG-27, UG-32, UG-28, UG-45) have **no formula in any sample**, so they are typed values to check (section 6) |
| Table 5, weight build-up | one row per part, **one column per item** | weights, totals, painted area, RT films | file or typed |
| Sections 6.n, cost build-up | **one heading and one table per item** (35 or 39 rows: materials, conversion, services, subtotals, overheads, contingency, margin, selling price, rounded unit price) | each cost line: head, basis, amount | file or typed; subtotals and mark-ups are rules |
| Table 10 (or 8), rate schedule | one row per rate | item, rate | file (the workbook's Inputs) or typed |
| Last table, sign-off | Prepared, Checked, Approved | names, signatures, dates | typed by people, never AI |

Numbers are written in the Indian grouping (`2,89,000`, `24,40,000`): the field's text must be formatted that way.
Boss's answer (D107): this sheet is filled **either** from an estimation workbook of the E-1604AB kind **or** from typed
values and AI text when no workbook exists.

## 5. The technical proposal, Word (three samples: `KGC_Technical_Proposal_NEA-28982_...`, `..._NEA-28983_..._NACE_...`, `Technical_Proposal-P-XXXX-...`)

Part A of KGC's offer. Header: the proposal number and date; two footers (first pages QF-MKT-009, last pages
QF-MKT-010) with "Page n of m"; one logo. Annexures A to H in order.

| Part | Shape | Fields | Source |
| --- | --- | --- | --- |
| Cover (P2 to P6, table 1) | 7 rows of label, ":", value | client, item, end user / project, project or OEM ref., enquiry ref., proposal no., date | typed |
| Table 2, contents | one row per annexure with its page numbers | page numbers ("3-6") | **cannot be proven**: they change with the item count; typed, marked to check |
| Table 3, enquiry documents | 1 row | enquiry e-mail date and sender | typed |
| Table 4, items | one row per item, then total quantity | Sl., item, tag no., qty, document no. | file (register or estimation) or typed |
| Table 5, enquiry documents list | one row per document | document no., name, rev., remarks | typed; remarks AI |
| Table 6, codes and standards | one row per code | code, application | typed or AI |
| **Annexure B, one item section per item** | a heading ("Annexure - B (contd.)"), the delivery conditions table (TDC no., equipment, tags and qty, 15 vessel details), the nozzle table (one row per nozzle; absent in the P-XXXX sample), the materials table, a nozzle-notes paragraph | per item: the TDC fields, nozzle rows, material rows | file or typed; notes AI. **Repeated once per item** (4 sections in 28982, 2 in 28983 and P-XXXX; D107) |
| Annexure C, inspection and confirmations | numbered tables (A quality control, E confirmations against the enquiry's notes, F NACE requirements only for sour service) | type, extent, remark; KGC's confirmation per note | AI (marked) and typed; **table F appears only when the job is sour service** |
| Annexure D, scope and commercial basis | scope rows; commercial rows, "highlighted items to be confirmed by KGC" | scope lines, price basis, taxes, delivery, validity | typed; highlighted text is already the form's own "check this" |
| Annexure E, technical queries | one row per query | TQ no., enquiry ref., requirement, KGC proposal or query, type | AI (marked) |
| Annexure F, document schedule | one row per document | document, for approval, final data book, submission | typed or fixed |
| Annexure G, terms and conditions | 17 numbered clauses | none | fixed |
| Annexure H, vendor list | two vendor columns side by side | none per job | fixed |

## 6. The engineering rules in the samples' formulas

Only rules a sample writes as a formula (D25). A value a sample shows without a formula (the ASME calculations of
section 4, the qualified ranges written as text in the WPS) is not a rule KRONOS may compute: it is typed, or AI text
marked to check, until Boss gives the rule.

| Rule | Formula (as the sample writes it) | Sample cell |
| --- | --- | --- |
| Heat input, kJ/mm | `Q = k * U * I * 60 / (v * 1000)`: `=$N$26*N33*L33*60/(Q33*1000)` (min: min U and I over max v); max: `=$N$26*O33*M33*60/(P33*1000)` | WPS P105 R33:S42, P104 R34:S34 |
| Heat input of a PQR run | `=$N$26*Y33*W33*60/(Z33*1000)` | WPS P105 AB33:AB42 |
| Heat input within the WPS limit | `=IF(S33<=$N$27,"OK","OVER LIMIT")`; P104 `=IF(AND(S34<=$N$27,R34>=$N$28),"OK","OUT OF RANGE")` | WPS P105 T33, P104 T34 |
| PQR run inside the WPS range | `=IF(AND(V>=L,W<=M,X>=N,Y<=O,Z>=P,Z<=Q),"Yes","No")` | WPS AC33 |
| A range shown as text | `=L33&"–"&M33`; `=TEXT(R33,"0.00")&" – "&TEXT(S33,"0.00")` | WPS D33:H33 |
| Qualified thickness | 0.5t to 2t, written as text ("7 to 28 mm (0.5t to 2t ...)") with no formula | WPS P105 F28 (text only: KRONOS computes it only once Boss confirms the rule) |
| Shell OD | `=R9+2*S9` (OD = ID + 2 x thickness) | register Q9:Q14; Enq Summary Q2:Q5 |
| bar to kg/cm2 | `=ROUND(18*1.01972,2)` | register W9:W14; Enq Summary W2:W5 |
| Ring raw weight | `=PI()/4*((OD+2*fa)^2-(ID-2*fa)^2)*(H+2*fa)*density` | BOM-RM M5 |
| Plate raw weight | `=(L+2*pm)*(W+2*pm)*T*density*(1+scrap)` | BOM-RM M7 |
| Pipe weight | `=L/1000*kg_per_m` | BOM-RM M12 |
| Plate weight (no margin) | `=G9*H9*I9*Inputs!$C$6` | Hydrotest Options J9 |
| Amounts and totals | `=N5*O5`, `=F5*G5`, `SUMIF` by scope, `SUMIFS` by option and scope | BOM-RM, Conversion, Hydrotest Options |
| Weld consumable | `=35/Inputs!$C$14` (weld metal / deposition efficiency) | Conversion E7 |
| Mark-ups | overheads `=B8*Inputs!$C$26`, contingency `=B8*Inputs!$C$27`, option adder `*(1+overheads)` | Summary B9, B10, B14 |
| Selling price | `=ROUNDUP(B16*(1+Inputs!$C$28)/2,-3)` (per unit, rounded up to the next 1,000) | Summary B21:C22 |
| USD price | `=B21/Inputs!$C$5` | Summary B23 |
| Price per kg | `=D21/('BOM-RM'!$L$23+'BOM-RM'!$L$24)` | Summary B31 |

## 7. The jobs, in a person's words

| # | Request | Files | Route |
| --- | --- | --- | --- |
| W1 | "Make WPS KGC-P105-101-102 from WPS-105 for 16 mm plate, revision R0, dated today." | the old WPS (read), the new form (fill) | two workbooks, Offline |
| W2 | "Change the WPS date and revision to R2 on both pages." | one WPS | one workbook |
| W3 | "Add run 9 with 200 to 240 A, 13 to 16 V, 75 to 110 mm/min and check its heat input." | one WPS | one workbook |
| W4 | "Fill the heat-input block from the PQR runs in the old WPS and mark every run that is over the limit." | two WPS | two workbooks, Offline |
| W5 | "Set the qualified thickness range for 16 mm." (needs the 0.5t to 2t rule from Boss) | one WPS | one workbook |
| E1 | "Fill BOM-RM from this BOM: ..." (typed rows) | the estimation workbook | one workbook |
| E2 | "Change the plate rate to 95 and the margin to 12%." | the estimation workbook | one workbook |
| E3 | "Add a Conversion line: PWHT, FH, 1 lot, rate 40000." | the estimation workbook | one workbook |
| E4 | "Write the assumptions sheet for a sour-service job." | the estimation workbook | one workbook, AI text |
| E5 | "Fill the Enq Summary sheet for this enquiry's items." | the estimation workbook | one workbook |
| R1 | "Add enquiry 29001 with its three items to the register." | the register | one workbook, merged block |
| R2 | "Add the items of the estimation workbook's Enq Summary to the register." | estimation (read), register (fill) | two workbooks, Offline |
| R3 | "Set the Result of enquiry 28982 to Won." | the register | one workbook, merged cell |
| R4 | "Fill Shell OD and design pressure in kg/cm2 for every new row." | the register | one workbook, rules |
| S1 | "Fill the internal estimation sheet from the estimation workbook." | Word template, workbook | DOCUMENT, Offline |
| S2 | "Fill the internal estimation sheet for enquiry 29001 from these values: ..." (no workbook) | Word template | DOCUMENT, typed and AI |
| S3 | "Write the basis and assumptions for this estimate." | Word template | DOCUMENT, AI text |
| P1 | "Make the technical proposal for enquiry 28982 from the register and the estimation workbook." | Word template, workbooks | DOCUMENT, Offline |
| P2 | "Add a fourth item section to the proposal for the 2nd stage discharge bottle." | Word template | DOCUMENT, repeated section |
| P3 | "Include the NACE annexure; this job is sour service." | Word template | DOCUMENT, optional section |
| P4 | "Write the technical queries for these deviations: ..." | Word template | DOCUMENT, AI text |
| P5 | "Fill the proposal cover and header: client ..., proposal P-2750-0001, today's date." | Word template | DOCUMENT, typed |

## 8. What VCL 0.6 can't say (the rows of `data/gaps.md`, "KGC documents")

1. A FIXTURE built from a template workbook (merged cells, formulas, number formats, a picture, a print area).
2. Write into a merged range, and merge a new enquiry block (register R1).
3. ROUND, ROUNDUP to a multiple (`-3`), PI, MIN and MAX as value functions (SET_FORMULA's allowlist already has them).
4. `TEXT` in a SET_FORMULA formula (the WPS writes `TEXT(R33,"0.00")`): it is not on the allowlist today (`&` is allowed).
5. Read one workbook and write another (two Workbook ARGS with one in MODIFIES).
6. Mark a value as written by the AI.
7. Fill a Word document: fields, a repeated row per item, **a repeated column per item** (estimation sheet tables 3 to
   5), **a repeated section per item** (proposal Annexure B; estimation sections 6.n), **an optional section** (NACE
   table F), fields in the header and footer.
8. Write a number as text in the Indian grouping (`2,89,000`).
9. Write a cell note (comment) (the register's notes).
10. Page numbers in a contents table, which depend on Word's layout: KRONOS can't prove them.
