# KRONOS gets mechanical engineering and welding knowledge: the plan (D133)

Status: **approved by Boss on 25 Nov 2026** (his answers below). One PR per phase; each is merged when CI is green.

## Context

KRONOS knows no engineering by itself (D132). It works out a formula it is given, and its TESTs prove the arithmetic a
recipe states, never the physics or the units. On 25 Nov 2026 Boss asked to give it mechanical engineering and welding
knowledge, and answered:

1. **What:** all four kinds: welding (WPS and PQR) rules, vessel design checks, estimation rules, and the words
   (a glossary).
2. **Where the formulas come from:** KGC's own samples (`samples/kgc/`, D106). I list every formula the samples already
   use, Boss confirms the list once, and each rule gets one worked example.
3. **Values from licensed code books** (ASME II-D allowable stresses, ASME B16.5 flange ratings, ISO tables): Boss
   keeps the values he uses in his own workbook; KRONOS reads it as "Workbook to read from" (D114, D120) and never
   stores it. They never go into the public repo.
4. **Order:** the glossary and the welding rules first, then the corpus batch of document jobs and the benchmark (they
   use these rules), then the vessel design checks.

"Knowledge" here means three layers, each proven in its own way:

| Layer | What it is | How it is kept honest |
| --- | --- | --- |
| Words | a glossary of welding, pressure-vessel and estimation terms and KGC's shop words | used only to understand a request (the recipe finder, the AI's prompt, the training data), never shown as a fact |
| Rules | engineering formulas, each from a sample's own formula or from Boss | each one confirmed by Boss, tested against a worked example, units in its name, the source cell named |
| Tables | values from licensed code books | only in Boss's own workbook, read when a job runs; the repo has invented look-alike tables for tests |

## The rules (from the samples' own formulas, `docs/kgc-documents.md` section 6)

| Group | Rule | The sample's formula |
| --- | --- | --- |
| Welding | Heat input per run, kJ/mm | Q = k x U x I x 60 / (v x 1000); min from min U, I and max v; max from max U, I and min v |
| Welding | Heat input of a PQR run | the same formula on the PQR's recorded values |
| Welding | Heat input within the WPS limit | max <= the WPS limit (P104 also: min >= the minimum) |
| Welding | PQR run inside the WPS range | every PQR value inside the WPS's min and max |
| Welding | Qualified thickness 0.5t to 2t | **text only in the sample: computed only once Boss confirms the rule** |
| Vessel | Shell OD | OD = ID + 2 x thickness |
| Vessel | bar to kg/cm2 | ROUND(bar x 1.01972, 2) |
| Estimation | Ring, plate and pipe weights | the BOM-RM formulas, with the Inputs sheet's allowances and density |
| Estimation | Weld consumable | weld metal / deposition efficiency |
| Estimation | Mark-ups, selling price, USD price, price per kg | the Summary formulas; price per unit rounded up to the next 1,000 |

The ASME VIII-1 design checks the estimation sheets show (UG-27 shell, UG-32 head, MAWP, hydrotest 1.3 x MAWP x LSR,
external pressure) are **values without a formula** in the samples. They are built only when Boss gives the formula, the
code edition and a worked example for each (D132); until then they stay typed values or AI text marked "check it".

## Rules for this work (my conservative choices)

- A formula comes only from a sample's own formula or from Boss, never from an AI chat or from Claude's own knowledge
  (D132). Each rule names its source (sample file and cell, or "Boss, <date>").
- Each rule is tested against a worked example: the sample's own inputs and result where the sample has them (Boss
  confirms them), or Boss's hand-worked numbers. The TESTs themselves use invented look-alike values (D105).
- Units are in the GOAL and the column names (kJ/mm, mm, barg, kg/cm2, kg, INR).
- Review says, for every recipe with an engineering rule: "Engineering calculation: KGC's own formula (<source>). The
  result is preliminary; a qualified engineer checks it before use." KRONOS never approves, signs or certifies anything,
  and the "Checked by" and "Approved by" lines stay for people.
- Values from code books never enter the repo, a test, a Claude prompt or the training data; the look-alike tables have
  invented values that are clearly not code values.
- Words in the glossary never become facts: KRONOS uses them to understand a request, not to answer an engineering
  question.

## Phases (one PR each)

- **E0, this plan** and D133.
- **E1, the glossary** (`data/domain/glossary.md`, D124's file): welding (WPS, PQR, WPQ, P-No and group, heat input,
  interpass, preheat, PWHT, back-gouge, run, positions PA to PG and 1G to 6G...), pressure vessels (shell, head, nozzle,
  repad, MAWP, MDMT, corrosion allowance, joint efficiency, hydrotest, NACE, HIC...), estimation (BOM, raw weight,
  finished weight, conversion, overheads, contingency, margin...), and KGC's shop words and abbreviations as the
  samples write them. Each term has a plain meaning and the ways people write it. The recipe finder's synonym table
  (`compiler/kronos_finder.py`) and the training data (`train/make_dataset.py`) read it; a test keeps them in step.
  Boss may read and correct it; nothing in it is used as a number.
- **E2, the rule sheet for Boss** (`docs/engineering-rules.md`): every rule above with its formula, units, inputs,
  source cell and its worked example taken from the sample. **Boss's one action:** reply "confirmed", or correct a row.
  The 0.5t to 2t rule is confirmed or left out here.
- **E3, the welding recipes:** proven recipes on the look-alike WPS (`data/templates/wps_form*.xlsx`): fill each run's
  heat-input range, check every run against the WPS limit, check that each PQR run is inside the WPS range, write the
  qualified thickness range (if confirmed in E2), and make a new WPS's numbers from an old WPS and a new thickness (two
  workbooks). They join the corpus (a new family `calc`) and the recipe finder's library, and Review shows the
  engineering line. If the recipes need value functions VCL lacks (PI, SQRT, POWER for exact results rather than
  formula text), that is a small VCL change with its own design note first (CLAUDE.md "Any VCL change").
- **E4, the estimation and vessel-data recipes:** Shell OD, kg/cm2, ring, plate and pipe weights, consumables,
  mark-ups, selling price and price per kg, on the look-alike estimation workbook and enquiry register.
- **E5, the code tables:** an invented look-alike table workbook (allowable stress by material and temperature, flange
  rating by class and temperature) and recipes that read a value from "Workbook to read from"; a short page in the guide
  on how Boss lays out his own table workbook.
- **E6, the vessel design checks** (UG-27, UG-32, MAWP, hydrotest, as Boss chooses): only after Boss gives each formula,
  its code edition and a worked example.

After E3, the plan goes on as before: the corpus batch of document jobs (`docs/kgc-documents-plan.md` Phase 5), which
can now use the welding and estimation rules, and the benchmark (Boss's step 7), with E4 to E6 between or after them
as Boss chooses.

## How the model learns it (March 2027)

The glossary and the proven calculation recipes, with their requests and paraphrases, go into the training data like
every other proven pair (`train/make_dataset.py`, D126). The model learns to write recipes that use KGC's formulas; it
never learns engineering facts as free text, and every number it produces still comes from a proven recipe (D31, D90).

## Verification (every PR)

The five checks of CLAUDE.md, `python compiler/vclc.py sim` and `python compiler/vclc.py corpus`; a test per rule
against its worked example; `build-offline.yml` when the page or the library changes; `tests/test_samples.py` keeps the
real samples out of tests and training data.
