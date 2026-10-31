# VBA Forge AI plan: the model writes VCL (D31)

Status: **approved by Boss on 19 Oct 2026**. This is the general direction for the model after
the Job-Pack Router MVP. It is not a build plan yet: each stage gets its own design and PRs
when it starts. Nothing here starts before the MVP is done, and no training starts before the
new PC arrives in March 2027.

## The decision

**Our own model writes VCL, never VBA. The compiler stays the only thing that writes VBA.**

The model is still ours (D25): designed and trained from zero, on in-house data only. Since D70,
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
| A small vocabulary (a few dozen keywords) | A small model, trainable from zero on one PC |
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
  public, only practice requests with invented data, added through a reviewed PR. When KRONOS lives
  on the company drive, real requests may become training data there, never on GitHub.
- **A locked benchmark:** held-out tasks never used for training; the only scoreboard.

Every pair passes the whole toolchain before it goes into `data/verified/`. Synthetic data only in
the repo, which is public (CLAUDE.md, D69).

Scale: the planned 150-task corpus is a seed, not a dataset. A useful model needs thousands of
proven specs; paraphrases and mutants help multiply them.

## Stages after the MVP

| Stage | When | What |
| --- | --- | --- |
| A | Before March (cloud) | The `GOAL` line, the formatter and error codes; the 150-task corpus; the data generator; the benchmark |
| B | Before March (cloud) | Grow VCL from the NOTE count; grow the data to thousands of proven specs |
| C | March 2027 (new PC) | A tokenizer and a small model trained from zero: English -> VCL, with grammar-constrained output |
| D | After C | Add repair. The model is used only when it passes enough benchmark tasks; Claude stays the Orchestrator until then |
| E | From item 5 (now) | Daily use: the KGC KRONOS desktop app (`docs/kronos-app.md`, D68). Boss types a task and gets proven VBA; Codex writes the VCL until our model passes the benchmark |

Stages A and B fit today's rules (no training before March), and they make the data valuable
even if the model takes longer than planned.

## What doesn't change

- D25: the model and everything it learns from are ours; no pre-trained weights, outside data
  or retrieval. D70 only adds proven Codex specs to what counts as in-house data.
- The compiler and ForgeLib stay the core, and the reference for what VBA a spec means.
- Running anything in Excel stays optional and never blocks work.
