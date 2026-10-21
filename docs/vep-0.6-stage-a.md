# VEP 0.6: AI plan Stage A, the language side and the data (CLAUDE.md item 4)

Status: **design, 21 Oct 2026**. Boss said "go item 4" after item 3 (CALL and BUILD specs).
This is Stage A of `docs/ai-plan.md` (D31): the language changes the plan lists (a `GOAL` line,
a formatter, error codes), then the 150-task corpus, the data generator and the benchmark.
Choices that needed Boss's decision were made the conservative way and recorded in
`DECISIONS.md` D59, as CLAUDE.md asks; Boss can overrule any of them. No model training and no
local-model work before March 2027 (CLAUDE.md), and everything stays in-house (D25).

## The problem

The toolchain proves a spec, but nothing yet turns it into training data. Today a spec has no
English next to it, the same meaning can be written many ways, a failing check says only
"line 12: ..." in words that can change, and there is no collection of tasks, no way to make
repair examples at scale, and no scoreboard. Without these, March 2027 starts with a model and
no data to teach it, and no way to tell whether it is any good.

## Principles (unchanged)

- The model writes VCL, never VBA (D31). The compiler writes all VBA, so every model output is
  proven before anything reaches Excel.
- Every pair in the dataset passed the whole toolchain: validator, type checker, TESTs in the
  reference interpreter, MODIFIES and UNCHANGED, the compiler and the linter, and a mutation
  floor (below). Nothing unproven enters `data/`.
- Synthetic data only: invented sheet names, headers and figures; no real KGC file, name or
  customer (CLAUDE.md). Claude may write tasks and specs "through our pipeline" (D25): its
  output is trusted only after the toolchain verifies it.
- Any VCL change updates spec.md and the prompt together, the grammar, parser, validator,
  type checker, interpreter and compiler as needed, the tests and `DECISIONS.md`; Boss updates
  the Orchestrator once per switch.

## Design

### 1. Error codes (no language change)

Every message the validator, the type checker, the grammar stage and the fixture stage can give
carries a stable code: `V` the validator's rules and warnings (including the BUILD rules), `G`
the grammar, `T` the type checker, `X` the fixture builder, `R` a failing TEST. A code is written
at the call site (`self.error("...", code="V107")`), so it is explicit and reviewable. Rules:
- A code, once published, never changes meaning and is never reused; a message may be reworded
  and keep its code.
- A new message must bring a new code. A test scans the source and fails when a call has no
  code, a code is used for two different messages, or a code is missing from the generated
  catalogue `docs/error-codes.md` (written by `compiler/gen_error_codes.py`, checked in CI).
- The code is a field of the problem (`Problem.code`); the message text stays as it is, so no
  test that compares messages changes. The command line shows `path:line: error: V107: ...`.
- A repair example (below) carries the codes, so "error -> fix" is learnable and stays valid
  when a message is reworded.

### 2. The `GOAL` line (VCL 0.6)

`GOAL <one plain-English sentence>` says what the spec does. It sits right before `PROC` in a
TASK (after the `# ASSUME:` lines) and right before `NAME` in a BUILD; every TASK inside a BUILD
has its own. From `VCL 0.6` it is required; before, it is an error ("GOAL is VCL 0.6"). One line,
10 to 200 characters, printable ASCII, not starting with `#`. The validator warns when it only
repeats the PROC name. Every older spec stays valid and gives the same results (as E8, E12, E13).
The 12 committed specs are inputs and stay as they are; their pairs live in `data/` (below).

### 3. One canonical text: `vclc fmt`

`python compiler/vclc.py fmt <files>` prints each spec in its canonical form and `--check` exits
1 when a file differs. The same meaning then always gives the same text, so the model learns
one way. The formatter works on the token stream and must be:
- **safe:** the formatted spec parses to the same tree, validates the same, and every TEST gives
  the same result (a test proves it on every committed spec and example);
- **idempotent:** formatting twice changes nothing;
- **small:** key names padded to column 9, OP nesting two spaces per level, one space after a
  comma and around operators, keywords and option names in the case spec.md writes them, one
  blank-line rule, `---` between a BUILD's specs. It never reorders keys (the validator already
  insists on the order) and never changes a quoted text or a FIXTURE rule.

### 4. The data layout

```text
data/verified/<id>/task.txt       the request, in the words Boss would type
data/verified/<id>/spec.vcl       the proven spec: VCL 0.6, canonical, with GOAL
data/verified/<id>/meta.json      id, kind (TASK or BUILD), registers, ops used, mutation score,
                                  source, toolchain commit
data/verified/<id>/paraphrase_N.txt   the same request worded differently (Phase 7)
data/eval/<id>/                   the held-out benchmark items (below); same files, never trained on
data/repair/                      generated "failing spec + codes -> fixed spec" pairs (Phase 7)
```
The VBA is not stored: it is derived and goes stale. `vclc export` writes the dataset as JSON
Lines (request, GOAL, spec, VBA freshly compiled, repairs) into `dist/`, which is git-ignored,
and a test proves every stored pair still passes the toolchain. The `data/*/README.md` files,
which describe the old Excel-based plan, are rewritten to match.

### 5. The corpus: 150 tasks

- **Coverage is planned, not accidental.** A task list crosses the registers (BOM, GRN, weld
  log, NDE, NCR, ITP, welder qualifications, calibration, consumables stock, material
  certificates, WPS/PQR, dispatch, job tracker) with task families (flag or highlight, summarise
  or count, reconcile two registers, clean or dedupe, rebuild an output sheet, route or read
  files, add sheets or formulas, a multi-step BUILD). About 100 single TASKs, 30 using the
  item 2 ops (files, sheets, formulas, FIND, FILTER), and 20 BUILDs.
- **Process per task.** (1) Claude writes the request as a person types it, with invented
  headers and figures. (2) A helper that sees only the Orchestrator prompt and the request writes
  the spec, as the claude.ai Orchestrator would (the two dry runs of D58 did exactly this).
  (3) The toolchain gates it: valid with no warnings, TESTs pass, builds lint-clean, snapshot
  mutation score at least 90%, canonical, GOAL present. (4) A failing spec gets up to two repair
  rounds with the coded messages; every failed attempt and its fix is kept as a repair pair.
  (5) A task the helper cannot express becomes an **expressibility gap** (a NOTE, a QUESTION or
  a workaround), counted per op family; this is the NOTE count Stage B needs.
- **The 12 existing specs** become the first 12 pairs, with a request written for each and a
  GOAL added to a copy (the originals stay).
- **A report** (`vclc corpus`) lists coverage, op use, mutation scores and the gaps.

### 6. The data generator

All of it is automatic and verified, or marked as not verified:
- **Repairs from mutants.** The mutation engine already makes thousands of near-miss specs. A
  mutant that fails a check or a TEST becomes "broken spec + coded messages or failing TEST ->
  the original spec". About 30 per task, about 4,500 for the corpus.
- **Repairs from real failures.** The failed attempts of step (4) above.
- **Contrast pairs.** A mutant that passes its own TESTs is a near-identical spec with different
  behaviour: "read carefully" data (marked, not trained as a fix).
- **Paraphrases.** Three wordings per request, written by Claude through our pipeline. They are
  not provable, so a mechanical check (every number, header and quoted status of the original
  appears in the paraphrase) filters them and the file is marked `paraphrase`; the spec they
  point to is the proven part.

### 7. The benchmark (the only scoreboard)

About 10% of the corpus (15 tasks) is chosen when the corpus is made, spread over registers and
families, kept in `data/eval/`, and never used for training, paraphrasing or repair.
- Each item has the request and a **contract**: the PROC name, ARGS and RET (for a BUILD,
  RunHeadless's SIG), plus **hidden reference FIXTURE and TEST lines**.
- `vclc bench <items> <candidates>` scores a candidate spec in four levels: **V** valid, **T**
  its own TESTs pass, **M** its own TESTs catch at least 90% of mutants, **H** its procedure
  passes the hidden reference TESTs (its own FIXTURE and TEST lines replaced by the reference
  ones, so a weak test can't help it). The headline is the share of items at H.
- A baseline comes first: Claude as the Orchestrator (helpers that see only the prompt) over the
  15 items. The March 2027 model is judged against it; Claude stays the Orchestrator until the
  model matches it (D31).

## Out of scope (Stage B and later)

Growing VCL from the gap counts, shrinking NOTE, going from 150 to thousands of specs, the
tokenizer, any training. Stage A only prepares the data and the scoreboard.

## Phases (one or two PRs each, in this order)

0. **This design** and D59.
1. **Error codes:** the call-site codes, `Problem.code`, `gen_error_codes.py`, the catalogue and
   its test, the CLI output.
2. **The `GOAL` line:** the `VCL 0.6` line, the key in the parser, grammar and validator, the
   migration helper (`vcl_migrate.py` adds a GOAL and moves a spec to 0.6), compatibility tests.
3. **`vclc fmt`:** the formatter, `--check`, the safety and idempotence tests on every committed
   spec and example.
4. **The switch:** spec.md and the prompt teach VCL 0.6 (GOAL, the canonical form, Examples 1, 3
   and 4 with a GOAL line), and Boss updates the Orchestrator once. The corpus helpers read the
   same prompt, so this comes before the corpus.
5. **Corpus tooling and the first pairs:** the layout, `vclc corpus`, `vclc export`, the 12
   existing specs as pairs, the task list for the 150, the first batch.
6. **Corpus to 150:** batches of about 20, each its own PR; the gap report grows.
7. **The data generator:** repairs, contrast pairs, paraphrases, the JSON Lines export.
8. **The benchmark:** the 15 held-out items, contracts, `vclc bench`, the baseline.

## Verification (every phase)

The four standard checks and `vclc mutate --snapshot` (the score stays 95.4% or better), plus a
test for each new rule. Every older spec stays valid with the same results. The data phases add
a check that every stored pair still passes the toolchain, so nothing in `data/` can go stale
without CI failing.
