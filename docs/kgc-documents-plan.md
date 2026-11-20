# KRONOS fills KGC's own documents (Excel and Word): the plan (D105)

Status: **approved by Boss on 20 Nov 2026.** Written to be handed to another Claude Code session: read `CLAUDE.md`
first, then this file. **Phase 0 is done (D106):** all 11 samples are in `samples/kgc/`, so the next session needs no
uploads. **Phase 1 is done (D107):** the survey is `docs/kgc-documents.md`. **Phase 2 is done (D108):** the look-alike
templates are in `data/templates/`, with their Word mark-up in its README. Each phase is its own PR.

**Boss's answers of 20 Nov 2026, second round (D107):** a new WPS is made from an existing WPS workbook (its PQR
values; a two-workbook job, KRONOS Offline only); the Word estimation sheet is filled either from an estimation
workbook or from typed values and AI text with no workbook; and a technical proposal repeats a whole item section
(Annexure B's tables) once per item.

## Context

On 20 Nov 2026 Boss uploaded KGC's own documents (`Samples.rar` and five files from it) and said: "KRONOS will be
used to fill in such files." Until now KRONOS has done QA register jobs (weld log, NDE, NCR...). These samples are
different. They are estimation and welding documents:

| Sample | What it is | Shape |
| --- | --- | --- |
| `NEA_Enquiry-Summary-MKT.xlsx` | marketing's enquiry register | 1 sheet, 27 columns, one row per item, an enquiry's first cells merged over its item rows, formulas (Shell OD = ID + 2 x THK, bar to kg/cm2) |
| `Estimation_Sheet_E-1604AB_Channel_FH.xlsx` | a cost estimate | 7 linked sheets: Inputs (rates), BOM-RM, Conversion, Hydrotest Options, Summary, Enq Summary (the register's columns plus prices), Assumptions; about 270 cross-sheet formulas (weights from dimensions, ROUNDUP of prices, totals) |
| `KGC_Estimation_Sheet_NEA-28982_INTERNAL.docx` | the internal estimation sheet | Word, 11 tables: price summary, options, basis and assumptions, a preliminary ASME VIII-1 design check; header, footer, a logo |
| `KGC_Estimation_Sheet_NEA-28983_INTERNAL.docx` | a second internal estimation sheet | the same Word form for a second enquiry |
| `KGC_Technical_Proposal_NEA-28982_..._Rev0.docx`, `..._NEA-28983_..._NACE_Rev0.docx`, `Technical_Proposal-P-XXXX-XXXX-R0-Dt_20-Nov-2026.docx` | technical proposals (Part A) | Word: a cover table, annexures A to H (basis of offer, technical delivery conditions per item, inspection, scope and exclusions, technical queries, document schedule, terms, vendor list), header and footer with the proposal number and form number |
| `WPS_KGC-P104-191-101_R1.xlsx`, `WPS_KGC-P105-101-101_R1.xlsx` (and the same two as `.pdf`) | Welding Procedure Specifications (ISO 15609-1, format QF/QA/31) | a 2-page form: 176 merged cells, 4 pictures (joint sketches), a print area, an "Actual / Qualified" pair of columns, and a heat-input check with formulas outside the print area |

**Boss's answers (20 Nov 2026):**
1. KRONOS should fill all of these: every Excel file and the Word files.
2. The values come from all four places: another sheet or file, the person's request, engineering rules, and text
   the AI writes.
3. Word: KRONOS Offline makes the Word file in the browser (no macro).
4. Two files: KRONOS Offline takes two uploaded files (for example, the estimation workbook fills the enquiry
   register). Macros still open no second workbook until March 2027 (D36).
5. Save the real samples in the public repo. This reverses part of the standing rule (CLAUDE.md, D101) for these
   files only; Boss chose it knowing the repo is public. Everything else stays synthetic: specs, TESTs, FIXTUREs,
   the corpus and unit tests use invented look-alikes.

**The outcome:** a person uploads (or opens) a KGC workbook or Word template, types a request, and gets the filled
file back, proven by tests on look-alike practice files, with every AI-written text marked "check this".

## The approach in one picture

| Job | Route | Why |
| --- | --- | --- |
| Fill one workbook (a WPS, an estimation workbook, the enquiry register) | VCL recipe, run as a **macro in Excel** (KRONOS.exe or the .bas) or on a **copy in KRONOS Offline** | Excel keeps pictures, print areas and formulas itself; Offline must learn to keep them too (Phase 4a) |
| Fill one workbook from another | VCL recipe with two Workbook ARGS, **KRONOS Offline only** | Boss's answer 4 |
| Fill a Word file | a new VCL kind, **DOCUMENT**, run by **KRONOS Offline only** | Boss's answer 3; VBA can't be proven to drive Word |
| Engineering rules | the formulas the KGC workbooks already use, written by the recipe as values or SET_FORMULA | they are KGC's own rules, already in the samples; a rule not in a sample needs Boss's word |
| AI-written text (assumptions, remarks) | marked in the recipe with `AI_TEXT("...")` | it can't be proven, so KRONOS lists it for an engineer to check |

## Phases (one PR each unless it says otherwise; CI green, merge with a merge commit, then the next)

### Phase 0: publish the samples (D105) - done (D106)
- Read `Samples.rar` in full first (install a RAR reader, `unar` or `7z`, in the container only; extract into the
  scratchpad, never inside the repo; read every file with `python -I`). Publish only what has been read.
- Commit the real files to `samples/kgc/` with a `README.md`: what each is, that Boss published them on
  20 Nov 2026 with KGC's permission (D105), and that nothing in tests, specs, FIXTUREs or the corpus may use them.
- `.gitignore`: an exception for `samples/kgc/*.xlsx` (workbooks are ignored on purpose), as for `docs/demo/`.
- `CLAUDE.md` hard constraints: say exactly this exception; the rest of the synthetic-only rule stands.
- A test: `samples/kgc/` holds only the files the README lists, and no test, spec or corpus file reads from it.

### Phase 1: the document survey (`docs/kgc-documents.md`) - done (D107)
- For each sample, every field KRONOS may fill: its cell (or Word table cell), its kind (text, number, date, formula),
  and where its value comes from (another sheet or file, typed, a rule, AI text). Use `kronos_book.describe`
  (`compiler/kronos_book.py`) for the workbook facts; read the Word XML with the standard library.
- The jobs a person will ask for, about 20, in their words. For example: "make a new WPS from PQR-105 for 16 mm
  plate"; "fill the estimation workbook's BOM-RM from this BOM"; "add this enquiry's items to the register";
  "fill the internal estimation sheet from the workbook".
- The rules found in the formulas (heat input Q = k x U x I x 60 / (v x 1000); qualified thickness 0.5t to 2t;
  Shell OD = ID + 2t; kg/cm2 = bar x 1.01972; ring and plate weights; ROUNDUP of prices to thousands), each with the
  sample cell it came from. Nothing from outside sources (D25).
- What VCL can't say yet, as rows in `data/gaps.md` (a new "KGC documents" section).

### Phase 2: look-alike templates (synthetic, for tests) - done (D108)
- `data/templates/`: invented copies with the same layout as each sample: the WPS form (merged cells, a sketch
  picture, print area, the heat-input block), the 7-sheet estimation workbook with its formulas, the enquiry
  register with merged enquiry blocks, the Word estimation template and a Word technical proposal template with `{{field}}` placeholders and one
  repeating row per table. Invented names, numbers and prices only.
- Made by scripts beside them (the `docs/demo/` pattern), with a test that the files equal what the scripts make
  today (as `tests/test_demo.py`).

### Phase 3: VCL 0.7 (a VEP first: `docs/vep-0.7-documents.md`, then one PR per part) - the VEP is done (D109)
Every language change follows CLAUDE.md "Any VCL change means": spec.md and the prompt together, grammar, parser,
validator, type checker, interpreter, compiler, ForgeLib, conformance cases, `docs/vcl-semantics.md`, tests,
DECISIONS.md, `dist/forge_module.bas`; and `vclc sim` must still prove every spec.
- **3a. Template FIXTUREs:** `FIXTURE wps_01.xlsx template=wps_form` builds a practice workbook from a committed
  template in `data/templates/` (cells, formulas, merged ranges, number formats; pictures are noted, not drawn), so a
  TEST can run on the real layout. `vcl_fixtures.py` stays the one fixture generator; `forge_fixtures.bas` rebuilds
  it in Excel; the VBA simulator loads it.
- **3b. Form writes and checks:** writing a fixed cell already works (`SET_VALUE(RANGE($ws, "C6"), ...)`). Add what
  the survey shows is missing, for example: writing into a merged range (its top-left cell, with a clear message
  otherwise), MERGE for a new enquiry block if the register needs it, and assertions on a template's cells.
- **3c. Rule functions:** only the value functions the samples' rules need and VCL lacks (likely ROUND, ROUNDUP to a
  multiple, PI, MIN and MAX of two), each in `vcl/semantics.py` with conformance cases and a ForgeLib helper. The
  SET_FORMULA allowlist (`compiler/vcl_formula.py`) already has ROUND and ROUNDUP; check PI, TEXT and AND.
- **3d. Two workbooks:** a TASK with two Workbook ARGS (D40 already allows Workbook ARGS) reads one and writes the
  other; `MODIFIES` names only the one it writes, and the type checker refuses a write to the other. The validator
  marks such a recipe "KRONOS Offline only" (no .bas until macros may open a second workbook, March 2027).
- **3e. `AI_TEXT("...")`:** a text value that the AI wrote. It means its text; KRONOS lists every AI_TEXT in Review
  and Save ("written by the AI, not proven: check it").
- **3f. DOCUMENT, the Word kind:** a spec kind beside TASK and BUILD:
  `TEMPLATE estimate_template.docx`, Workbook ARGS, `FIELD Client = VALUE($wb!Inputs!C5)`,
  `ROWS 1 FROM RANGE($wb, "Summary!A5:F9")` (fill table 1's repeating row once per source row), and TESTs that
  assert `FIELD_TEXT(Client)` and table cells, with a FIXTURE workbook and the look-alike template. The interpreter
  fills the template with the standard library (`zipfile` and `xml.etree`), keeping every run's formatting, the
  header, footer and pictures. The compiler refuses DOCUMENT with a clear message (it makes no VBA).
  The survey adds (D107): zero or more Workbook ARGS, so a DOCUMENT can be filled from typed ARGS and AI_TEXT alone; a
  column repeated per item (the estimation sheet's tables 3 to 5); a whole section repeated per item (the proposal's
  Annexure B, the estimation sheet's sections 6.n); an optional section (the NACE table); fields in the header and footer.
- **3g. The switch:** spec.md and the prompt teach VCL 0.7 once (hard rules for templates, two workbooks, AI_TEXT and
  DOCUMENT; one example of each); Boss pastes the prompt once.

### Phase 4: KRONOS Offline version 6 (`docs/kronos-offline.md` Phase 4; the old version 6 scope moves to 7)
- **4a. A copy that keeps everything:** write only the changed cells into the uploaded file's own XML (values,
  formulas; styles untouched), so pictures, print areas, merged cells, data validation and page setup survive.
  Fall back to today's openpyxl copy, with today's warnings, only when rows or sheets move. Test: the look-alike WPS
  with its picture comes back identical except the changed cells.
- **4b. Two files:** "Workbook to fill" and "Workbook to read from"; each Workbook ARG is bound to one upload.
- **4c. Word:** upload a template .docx (up to 2 MB) and a workbook, run a DOCUMENT recipe, download the filled
  .docx. Values come from the workbook's last values saved by Excel: a workbook KRONOS just changed must be opened and
  saved in Excel first, and a formula with no saved value is refused, never guessed.
- **4d. Review shows the AI's text** (3e) and which cells or fields each recipe fills.
- The guide `docs/KRONOS_Offline_Guide.docx` gets the new steps, including how Boss turns his own Word file into a
  template (type `{{Client}}` where the client's name goes).
- Released as `kronos-offline-v6` by `build-offline.yml`, checked as before (Chromium 109, Firefox 115, PowerShell).

### Phase 5: corpus batch 9, documents (about 20 pairs)
- Requests from Phase 1's job list, written as Boss would type them; specs by helpers who see only the prompt,
  gated by `corpus-add` as before (90% mutation floor), FIXTUREs from `data/templates/` only. A new family
  `documents` in `compiler/vcl_corpus.py` (QUOTAS) and registers `wps`, `estimation`, `enquiry`.
- `vclc datagen` (D103) and the paraphrase check (D104) cover them; 3 paraphrases each.

### Phase 6: the benchmark (Boss's step 7)
- As `docs/vep-0.6-stage-a.md` section 7 plans, with at least 5 of its held-out items being KGC document jobs.

### Later (not in this plan)
- KRONOS.exe learns DOCUMENT and two-workbook recipes; the old version 6 scope (History, practice-run check, corpus
  suggestion, Help) as Offline version 7; storage S1; the VCL memory; the day-one kit.

## Rules for the implementing session
- Read `CLAUDE.md` first. One PR per phase or part; never weaken a test or threshold; specs are inputs.
- Real KGC files only in `samples/kgc/`; every test, spec, FIXTURE and corpus pair uses `data/templates/` look-alikes.
- Engineering rules come from the samples' own formulas or from Boss; nothing from outside sources (D25).
- When a choice is Boss's, take the conservative option, record it in DECISIONS.md and carry on.
- Suggested model: Opus, high effort for Phases 3 and 4a; medium for the rest; Sonnet helpers for Phase 5 specs.

## Verification (every PR)
- The five checks of CLAUDE.md, `python compiler/vclc.py sim` on every spec, `python compiler/vclc.py corpus`, and
  `build-offline.yml` for Offline PRs.
- New tests per phase: Phase 0 (samples folder rule), 2 (templates equal their scripts), 3 (each language part:
  validator, interpreter, compiler or its refusal, simulator, conformance), 4 (keep-everything round trip, two-file
  binding, Word fill keeps formatting, stale formula values refused), 5 (corpus gates).
- End to end: in Offline, upload the look-alike WPS template, fill it from a typed request with a proven recipe,
  download, and check the picture, print area and merged cells are still there; then fill the look-alike Word
  template from the look-alike estimation workbook and check every field.
