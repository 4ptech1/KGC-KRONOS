# VBA Forge AI plan: the model writes VCL (D31)

Status: **approved by Boss on 19 Oct 2026**. This is the general direction for the model after
the Job-Pack Router MVP. It is not a build plan yet: each stage gets its own design and PRs
when it starts. Nothing here starts before the MVP is done, and no training starts before the
new PC arrives in March 2027.

## The decision

**Our own model writes VCL, never VBA. The compiler stays the only thing that writes VBA.**

> **Changed by D90 (Boss, 17 Nov 2026):** the model may also write the VBA from a proven VCL recipe, with its context and memory in VCL. Its VBA is used only when it passes the recipe's own TESTs in a VBA simulator; otherwise KRONOS gives the compiler's VBA. See `docs/ai-vba-plan.md`. The reasons below still explain why the recipe and the compiler stay the proof and the fallback.

> **Changed by D102 (Boss, 19 Nov 2026):** D25's ban on pre-trained weights is lifted. In March 2027 two candidates are
> built and scored on the same benchmark, and the better one is used: **A**, our own model trained from zero (at most
> about 350 million parameters), and **B**, a small open-weight coder model (Apache 2.0 or MIT licence, at most about
> 7 billion parameters) tuned with LoRA on our data and run in 4-bit, offline, on the new PC. No giant model is trimmed.

What the model learns from is still ours (D25): in-house data only. Since D70,
in-house also covers a spec that ChatGPT Codex wrote through the KGC KRONOS app and the toolchain
proved; Codex writes data, never part of the model.

Why:
- The compiler already turns VCL into VBA without mistakes. A model doing the same job could
  only be a worse copy of it.
- Every VCL spec the model writes can be proven without Excel: the validator, the type checker,
  the TESTs in the reference interpreter and the MODIFIES safety check. VBA written by a model
  can't be proven until it runs in Excel.
- So nothing unproven ever reaches Excel, and every VBA line that does is lint-clean compiler
  output that calls ForgeLib.

## The model's three jobs

1. **Plain English to VCL:** an offline Orchestrator for everyday tasks.
2. **Repair:** a spec that fails a check, plus the error messages, in; a fixed spec out.
3. **NOTE steps** (what VCL can't say yet): the model never writes raw VBA for them. Instead
   VCL grows until they are rare. Until VCL covers a NOTE, the task goes to Claude or Boss.

## The flow

The KGC KRONOS desktop app (`docs/kronos-app.md`, D68) is where this flow lives. Until our model
is ready, ChatGPT Codex writes the VCL behind the same window; from March 2027 our model takes its
place.

```text
Boss types a task in plain English (in the KRONOS window)
  -> the model writes VCL (several tries)
  -> validator, type checker and TESTs keep only the tries that pass
  -> the compiler writes the VBA (ForgeLib calls, lint-clean)
  -> Boss pastes it into Excel
No try passes  -> repair mode, with the error messages
Still failing  -> the task goes to Claude (the Orchestrator), Codex or Boss
```

## How VCL's structure helps

| VCL feature | What it gives the model |
| --- | --- |
| A small vocabulary (a few dozen keywords) | A small model, trainable from zero on one PC, or a small open model tuned on one graphics card (D102) |
| One step per line | Easy to learn, and every error points at one line |
| A formal grammar (`vcl/grammar.lark`) | Output can be forced to be valid VCL (grammar-constrained decoding) |
| The validator and type checker | Wrong outputs are rejected at once; every error is a free "error -> fix" training pair |
| The interpreter, synthetic fixtures and TESTs | Outputs are judged by "do the tests pass", not by how they look |
| The mutation engine | Many near-miss specs: training for careful reading and for repair |
| MODIFIES and UNCHANGED | An automatic safety check: the model can't quietly touch other sheets or files |
| The migration tool (`vcl_migrate.py`) | Old training data moves to each new VCL version, so the model learns one version |

## VCL changes planned for this goal

Each is a language change, so each comes as its own VEP and PRs, with spec.md, the prompt and
every check updated together. None starts before the MVP is done.

1. **A `GOAL` line** in every TASK: one plain-English sentence saying what the task does. Every
   spec then carries its own English -> VCL training pair.
2. **One canonical way to write each spec**, with a formatter (`vclc fmt`): the same meaning
   always gives the same text, so there is less for the model to learn.
3. **Error codes** in validator and type checker messages (for example `V107 line 12: ...`).
   Stable codes make repairs learnable.
4. **Shrink NOTE:** count every NOTE step in the corpus (the expressibility report), and turn
   the common ones into VCL ops.
5. **Keep VCL small:** a few ops that combine well, and one way to do each thing.
6. **Strong tests required:** a model-written spec counts only when its own TESTs catch enough
   mutants (a minimum snapshot mutation score), so it can't pass by writing weak tests.

## Where the training data comes from (all in-house, D25)

- **Orchestrator specs:** synthetic KGC-style tasks (BOM, weld log, NDE, NCR, calibration and
  similar registers), each with its English description and its proven VCL.
- **Paraphrases:** the same task written several ways, by Claude through our pipeline, so the
  model learns how Boss words things.
- **Repairs:** made automatically from mutants and the errors they cause, in large numbers.
- **Codex through KRONOS (D70):** a spec ChatGPT Codex wrote for a request typed into the KRONOS
  app, once it passes every corpus gate, marked `"source": "codex via KRONOS"`. While the repo is
  public, only practice requests with invented data, added through a reviewed PR. From March 2027
  KRONOS lives on the new PC's 300 GB (D91): real requests and their proven recipes may become
  training data there, in `KRONOS\train\data`, never on GitHub, and a model trained on them never
  leaves that PC.
- **A locked benchmark:** held-out tasks never used for training; the only scoreboard.

Every pair passes the whole toolchain before it goes into `data/verified/`. Synthetic data only in
the repo, which is public (CLAUDE.md, D69).

Scale: the planned 150-task corpus is a seed, not a dataset. A useful model needs thousands of
proven specs; paraphrases and mutants help multiply them.

## Stages after the MVP

| Stage | When | What |
| --- | --- | --- |
| A | Before March (cloud) | The `GOAL` line, the formatter and error codes; the 150-task corpus (done, D96); the data generator (done, D103, D104); the benchmark (next) |
| B | Before March (cloud) | Grow VCL from the NOTE count; grow the data to thousands of proven specs |
| C | March 2027 (new PC) | Two candidates, English -> VCL with grammar-constrained output, scored on the benchmark (D102): A, a tokenizer and a model trained from zero (up to about 350 million parameters); B, a small open-weight model tuned with LoRA (up to about 7 billion) |
| D | After C | Add repair. The model is used only when it passes enough benchmark tasks; Claude stays the Orchestrator until then |
| E | From item 5 (now) | Daily use: the KGC KRONOS desktop app (`docs/kronos-app.md`, D68). Boss types a task and gets proven VBA; Codex writes the VCL until our model passes the benchmark |

Stages A and B fit today's rules (no training before March), and they make the data valuable
even if the model takes longer than planned.

**Where Stage C lives (D91, `docs/kronos-storage-plan.md`):** on the new PC's 300 GB, the datasets in
`KRONOS\train\data`, the training runs and their checkpoints in `KRONOS\train\runs`, and the released
models in `KRONOS\models`, each with a budget. The budgets assume candidate A of up to about 350 million
parameters and candidate B of up to about 7 billion (D102: its base models in `KRONOS\models\base`), and are resized in March to the PC's graphics card; a run never starts if it would eat
into the 30 GB kept free.

## Stage C in detail: candidate B and the meaning check (D124)

Boss, 21 Nov 2026: "trim, translate, etc. an open source AI model for KRONOS so that it is capable of translating user
request to VCL and asking Codex to recheck if VCL is correct." His answers: shrink the file, not a big model cut down;
tune it for mechanical engineering and welding work, without what KRONOS doesn't need; its job is English to VCL; Codex
checks the meaning, as an option; prepare now, train in March 2027.

**The base model.** Qwen2.5-Coder-7B-Instruct is the main candidate and Qwen2.5-Coder-1.5B-Instruct the fast one; the
benchmark picks. Both are Apache 2.0 today; not every size in that family is, so each file's licence, model card and
SHA-256 are checked and written down in March (D102). They are downloaded once onto the new PC with IT's approval, into
`KRONOS\models\base`, and run offline; never into the repo or a Claude Code session.

**"Trim" means a small file, not a cut model.** Knowledge can't be cut out of a model cleanly, and what KRONOS never
uses costs nothing when it runs beyond the file's size. The model is kept to KRONOS's job in three ways:
1. LoRA tuning (QLoRA on one graphics card) on our proven pairs, so its habits become KRONOS's;
2. grammar-locked output: `vcl/grammar.lark` is turned into the runner's grammar, so the model can only write text
   that parses as VCL, never chat, VBA or anything else;
3. a short fixed prompt.
The tuned model is stored in 4-bit (about 5 GB for 7B, about 1 GB for 1.5B) and run with an offline runner such as
llama.cpp (MIT licence), which IT approves in March.

**Mechanical engineering and welding.** The base model already knows the common words (WPS, PQR, NDE, RT, UT, ASME IX,
hydrotest). KGC's own shop words come from an in-house glossary, `data/domain/glossary.md`: Claude writes it from Boss's
words, with invented examples only. Requests written in shop language, each paired with a proven recipe, are checked
the way the paraphrases are (D104). No downloaded welding or engineering datasets (D25). The samples in `samples/kgc/`
are never trained on (D106). From March, Boss's real proven runs join the data on the new PC (D91).

**The meaning check.** The toolchain proves a recipe parses, type-checks, passes its TESTs and touches only what it
says. It can't prove the recipe does what the person meant. So, as an option, after the proof KRONOS asks Codex: here
is the request, here is the proven recipe and what it changes; do they match? Codex answers in a fixed format, MATCH
or MISMATCH with its reasons, and Review shows the answer. Rules:
- It runs only on a recipe the toolchain has already proven. A MATCH never lets an unproven recipe through, and a
  MISMATCH never changes a recipe by itself: the person decides, or sends it back for a new try.
- It sends what already goes to Codex, the request and a recipe, and never a file (D69).
- KRONOS works without it: no Codex, no check, and Review says so.
- In KRONOS.exe it is a call through the Codex CLI, as the writing step is. In KRONOS Offline it is a text to copy
  into the AI chat and a box to paste the answer (version 7).
- Today Codex checks recipes Codex wrote. From March it checks our model's.

**The checker is scored too.** The data generator's contrast examples (D103) are recipes that are valid but mean the
wrong thing. The benchmark counts how many of them Codex's check catches, and how often it wrongly says MISMATCH to a
right recipe. That tells us how far a MATCH can be trusted.

**Before March, in the cloud (one PR each; nothing is downloaded and no model is run):**
1. this design;
2. the benchmark, `vclc bench` (Stage A Phase 8, Boss's step 7), with today's Codex as the baseline, then the
   checker's score;
3. the meaning check in KRONOS.exe (`kronos_engine.py`, `kronos_ai.py`, Review in `app/`), tested with the fake AI;
4. the meaning check in KRONOS Offline, as part of version 7;
5. the glossary and the shop-language requests;
6. the training kit in `train/` (`train/README.md`), as code only. Done: the dataset export (D126), the grammar
   conversion (D127: `train/vcl.gbnf`, a GBNF grammar made from `vcl/grammar.lark`, accepted by every recipe in the repo
   and read back by an independent checker), the QLoRA script and the 4-bit export (D128: written and tested as far as
   possible without a card, never run). The first run in March is the script's `--smoke` option.

**In March, on the new PC:** download the base models (IT), tune them, score them beside candidate A on the benchmark,
and use the best only once it passes. Until then, and whenever it fails, Codex stays the writer.

## What doesn't change

- D25: everything the model learns from is ours; no outside data or retrieval. D70 only adds proven Codex specs to
  what counts as in-house data, and D102 only allows a small open model's weights as a starting point.
- The compiler and ForgeLib stay the core, and the reference for what VBA a spec means.
- Running anything in Excel stays optional and never blocks work.
