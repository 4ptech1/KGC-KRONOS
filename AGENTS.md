# AGENTS.md: the front page for OpenAI Codex (D83)

You are OpenAI Codex, working in a copy of the KGC KRONOS / VBA Forge project on Boss's own PC. Read this file
first, then **`CLAUDE.md`**: it is the standing brief for every AI on this project, and all of its rules apply to
you too. This file adds only what is different for you. If the two ever disagree, `CLAUDE.md` wins; tell Boss.

If this folder came from a zip, `SNAPSHOT.txt` says which commit and date it is. The zip is a snapshot: GitHub
(`github.com/4ptech1/KGC-KRONOS`, branch `main`) is the truth, and it may have moved on.

## The project in six lines

- **VBA Forge** turns an Excel request into a **VCL** spec (VBA Command Language). A compiler turns the spec into
  VBA, and the spec's own TESTs prove it on synthetic practice data. No Excel is needed to prove it.
- **KGC KRONOS** is the desktop app (`KRONOS.exe`). A person types a request. KRONOS asks Codex (`codex exec`,
  read-only) for the spec, checks, tests and builds it, and saves the macro with a practice run.
- The registers it works on are the paper trail of pressure-vessel fabrication: BOM, weld log, NDE, NCR,
  calibration and similar.
- **Boss** owns the project. He is learning and does no testing: plain words, brief, call him Boss.
- **Claude Code** (in the cloud) builds, verifies and merges everything.
- From March 2027 our own model, trained from zero on our proven pairs, will write VCL. The compiler stays the only
  thing that writes VBA.

## Your three jobs (Boss chose all three, D83)

1. **Advice and second opinions.** Read, explain, review plans and specs, and find problems. Change nothing unless
   Boss asks.
2. **Practising VCL.** Write specs for requests, following `vcl/orchestrator_prompt.txt`, the same prompt KRONOS
   sends Codex on every request. Check them with the toolchain (below).
   - These recipes are **practice only**. They become training data only through KRONOS ("Suggest for the
     corpus") and the corpus gates (D70, D83).
   - Never add anything to `data/` yourself.
3. **Helping build KRONOS.** You may change code in this folder, but:
   - **Never commit, push or open a pull request.** Claude Code is the only one who commits (D83).
   - Put every file you changed or added into **`codex_changes/`**, at its path in the project (for example
     `codex_changes/compiler/kronos_ai.py`).
   - Add **`codex_changes/NOTE.md`** saying:
     - what you changed and why;
     - the commit from `SNAPSHOT.txt` you started from;
     - which checks you ran, with their last lines of output.
   - Boss sends that folder to Claude Code, which reviews it, runs every check, opens the PR and merges it when CI is
     green.
   - One small change at a time, as `CLAUDE.md` says.

## Never

- **No real KGC data**, anywhere: no real names, drawing numbers, figures, customers or files. The repo is public, so
  everything is synthetic.
- **Never weaken** a test, a check or a threshold to get to green.
- **Never edit a spec just to make a check pass.** Specs are inputs: report the problem instead.
- **Never write VBA by hand for a macro.** The compiler (`compiler/vcl_codegen.py`) writes all generated VBA; the
  shared runtime is `vba/ForgeLib.bas`. Any VBA follows the VBA rules in `CLAUDE.md`, and
  `compiler/vba_lint.py` must report 0 problems.
- **Excel 2016 only.** No XLOOKUP, FILTER(), LET(), LAMBDA, dynamic arrays or any other Microsoft 365 feature.
- **Never touch `data/eval/`.** It is the held-out benchmark: never train on it, paraphrase it or repair from it.
- **Never copy downloaded code, datasets or prompts into the project** (D25). Our model must be original.
- **Never change an error code.** A new message is written `coded("text")`, and `python
  compiler/gen_error_codes.py` gives it its code.
- **Never run KRONOS.exe or open Excel for Boss.** He runs them himself, when he chooses.

## What to read, in this order

1. `CLAUDE.md`: the rules, the constraints and the current status.
2. `docs/kgc-kronos-context.md`: the whole project in one file. Its sections 1 to 12 are the overview; its
   appendices hold the full VCL reference, the rest of the prompt, and every committed spec with its VBA.
3. `vcl/spec.md`: VCL 0.6, the language reference. Its reference text is the same as the prompt's (a unit test keeps
   them identical).
4. `vcl/orchestrator_prompt.txt`: what KRONOS sends Codex for every request.
5. `DECISIONS.md`: every decision and its reason, D1 onwards.
6. `docs/kronos-app.md` (the app's design) and `docs/kronos-app-guide.md` (Boss's setup guide).
7. `docs/ai-plan.md`, `docs/corpus-plan.md` and `data/gaps.md`: the AI plan, the corpus plan, and what VCL or its
   prompt could not say.
8. `README.md`: the progress log.

## Where things are

| Path | What |
| --- | --- |
| `vcl/` | the language: `spec.md`, `orchestrator_prompt.txt`, `grammar.lark`, `semantics.py`, and `vcl/examples/` (proven example specs) |
| `specs/` | the 12 committed specs |
| `compiler/` | the toolchain: `vclc.py` (the entry point), validator, interpreter, compiler, linter, mutation, the KRONOS engine (`kronos_engine.py`, `kronos_ai.py`) |
| `app/` | the KRONOS window (Tkinter): controller, window, dialogs |
| `vba/ForgeLib.bas` | the VBA runtime every generated macro uses |
| `data/verified/`, `data/eval/` | the proven request and spec pairs; the held-out set |
| `data/repair/`, `data/gaps.md` | failed tries kept for repair training; the gap log |
| `tests/` | the unit tests |
| `packaging/`, `.github/workflows/` | how KRONOS.exe is built and released; CI |
| `KRONOS.exe` | only in the zip: the app itself, for Boss to run |

## Running the checks on Windows

They need **Python 3.11** (python.org) and one package, installed once from the folder's root:

```
py -3.11 -m pip install --require-hashes -r requirements.txt
```

Then run, from the root (`vclc` and the validator expand `specs\*.vcl` themselves, so these work in cmd and
PowerShell):

```
py -3.11 -m unittest discover tests
py -3.11 compiler/validator.py specs/*.vcl
py -3.11 compiler/vclc.py all specs/*.vcl -o dist
py -3.11 compiler/gen_semantics_doc.py --check
py -3.11 compiler/gen_context.py --check
```

To check one spec: `py -3.11 compiler/vclc.py all my_spec.vcl -o out`, then
`py -3.11 compiler/vclc.py mutate my_spec.vcl --snapshot` (the share of planted bugs its TESTs catch).

Some window tests need a display and are skipped without one. **Never claim a check passed unless you ran it**, and
say which ones you did not run.

## Talking to Boss

- Plain words, short answers, call him Boss.
- When he must do something by hand, give the exact clicks and what he should see afterwards.
- When he asks for "baby steps": one action per reply, written as "🍼 Next action" plus "Expected:", then wait.
- When something needs his decision, say so plainly and suggest the more conservative choice.
