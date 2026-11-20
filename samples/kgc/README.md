# KGC samples (real documents, published by Boss on 20 Nov 2026)

These are **real KGC Engineering Projects documents**, the ones KRONOS is to learn to fill
(`docs/kgc-documents-plan.md`). Boss shared them as safe samples he has permission for, and on 20 Nov 2026 asked that
they be saved in this public repo (D105, D106). They are the one exception to the repo's synthetic-only rule.

**Rules (D105):**
- Nothing in tests, specs, FIXTUREs, the corpus (`data/`) or KRONOS's code reads these files. Tests use invented
  look-alikes in `data/templates/` (Phase 2 of the plan). `tests/test_samples.py` checks both rules.
- No other real file joins this folder without a new decision from Boss.
- The files are kept exactly as Boss gave them (`Samples.rar`); the SHA-256 below proves it.

| File | What it is | SHA-256 |
| --- | --- | --- |
| `NEA_Enquiry-Summary-MKT.xlsx` | marketing's enquiry register: one row per item, an enquiry's first cells merged over its rows, formulas, cell comments | `c9bc98721fe96e938307e597ffe9c0ae1674b0291e1f9f5824b9b34d2f2fe844` |
| `Estimation_Sheet_E-1604AB_Channel_FH.xlsx` | a cost estimate in 7 linked sheets (Summary, Inputs, BOM-RM, Conversion, Hydrotest Options, Enq Summary, Assumptions) | `2ed149184f4e45032ec6e58a828e067bce8268690183173edb4dfcadbe4ef9ba` |
| `KGC_Estimation_Sheet_NEA-28982_INTERNAL.docx` | internal Word estimation sheet: price summary, options, assumptions, design check, weights, cost build-up per item, rates | `f3f384f08684319b5280950cf33471abbb0bc701ce078b450bf87f39702d995f` |
| `KGC_Estimation_Sheet_NEA-28983_INTERNAL.docx` | the same form for a second enquiry (sour service) | `ee180de3e356372ede8fe4e89a9c04ed4f41499a77be3ad1297da4ad84ba394d` |
| `KGC_Technical_Proposal_NEA-28982_Pulsation_Dampeners_Rev0.docx` | Word technical proposal (Part A): annexures A to H, equipment-wise delivery conditions, technical queries, terms | `716ae1d56843fbc2c07bec24f2c78ef35f16a3bf762e5e3885d5ca1146ca5a4f` |
| `KGC_Technical_Proposal_NEA-28983_Pulsation_Dampeners_NACE_Rev0.docx` | the same proposal form for the sour-service enquiry | `f7b88ef81b69b03dc00fdef528445e35fd939a2031bed3234655103aac398006` |
| `Technical_Proposal-P-XXXX-XXXX-R0-Dt_20-Nov-2026.docx` | a technical proposal for the channel and floating head job (the estimation workbook's job) | `84d684cefa4ae2026456b258bd26570e66c61b9c215517982112a392b10aa142` |
| `WPS_KGC-P104-191-101_R1.xlsx` | Welding Procedure Specification (ISO 15609-1, format QF/QA/31): a 2-page form with merged cells, weld sketches, a print area and a heat-input check | `1c286e69fc582421b8201f891ce5303542f1ff0a34f253c5cb550965bde3134d` |
| `WPS_KGC-P104-191-101_R1.pdf` | the same WPS, printed to PDF | `5d6c32f612eb6fce28ee5c89882bb3237ef2eb94cdde06a661e1b12f978c0ba3` |
| `WPS_KGC-P105-101-101_R1.xlsx` | a second WPS on the same form | `43627fd37ba79ee1fad2d66a3c5e3de14dfefbec8d13ceee66f36f300aaffc32` |
| `WPS_KGC-P105-101-101_R1.pdf` | the same WPS, printed to PDF | `5a15f935487054d26658e2373523ddaec58d2bef116a660eada88a7178745260` |
