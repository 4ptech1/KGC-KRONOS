# The KGC KRONOS desktop app (item 5): design

Status: **design, 30 Oct 2026**, from Boss's answers of that day (DECISIONS.md D68 to D71).
**Phase 1, the engine and `vclc ask`, is built (D72, 31 Oct 2026)**. The window, a four-step
wizard designed with Boss on 31 Oct 2026, is next and comes before everything else (D74). Each
phase below is one PR, verified by unit tests and CI.

## What it does

A person types an Excel automation request in a window. KRONOS asks the AI on the PC (ChatGPT
Codex) for the VCL spec, proves the spec with the toolchain, and hands back the VBA module.

```text
the person types a request in the KRONOS window
  -> KRONOS sends the Orchestrator prompt + the request to Codex (codex exec, read-only, on stdin)
  -> Codex replies with one vcl block, or a # QUESTION: line
       question -> KRONOS shows it; the person's answer continues the conversation
  -> KRONOS checks, tests and builds the spec (validator, TESTs in the interpreter, compiler, lint)
       problems -> sent back to Codex with their codes, up to 2 more tries
  -> KRONOS shows what the spec does, assumes and changes, and its TESTs, in plain words
  -> the person ticks the assumptions, gives the real sheet names and saves the .bas module,
     then imports it into Excel (exact steps on screen)
with Show the recipe on (Settings), the person may also edit the spec and press Re-check
```

Only the request-to-VCL step is online. The validator, interpreter, compiler, ForgeLib and the
VBA are the same offline toolchain as today. KRONOS never runs the VBA; the person does, in
Excel, when they choose.

## Decisions it rests on

| | Decision |
| --- | --- |
| Where | Boss's Windows work PC, with IT's approval to install KRONOS.exe and the Codex CLI (D69) |
| AI link | The Codex CLI (`codex exec`), signed in with the person's ChatGPT account. The ChatGPT desktop app, which includes Codex since July 2026, can't be driven by another program (D68) |
| Data | Real work requests may be typed (the company allows it); no file is ever sent. Personal ChatGPT Plus/Pro accounts, so "Improve the model for everyone" is turned off in Data Controls. Practice requests only at first (D69) |
| Repo | Public on GitHub, so it stays synthetic. KRONOS's runs live in its own data folder, never in the repo (D69) |
| Training data | A spec Codex wrote that passes every corpus gate may join the corpus, marked `codex via KRONOS`. While the repo is public: practice requests only, through a reviewed PR (D70) |
| Later | KRONOS moves off GitHub to the company's storage drive (100 GB), where real data may live (D69, D70) |
| Users | Boss first; then a team, each person with their own Codex CLI and login, sharing a folder on the company drive (D71) |
| Window | Tkinter (Python's built-in), shipped as one KRONOS.exe that GitHub Actions builds (D68) |

## The parts

### 1. The engine (no window): `compiler/kronos_engine.py`, `compiler/kronos_ai.py`
- **AI adapters**, one small class each with `ask(conversation) -> reply text`:
  - `fake`, for the tests: replies from a script;
  - `codex`: runs the Codex CLI;
  - later, `own` (our model, March 2027, D31) and others, such as the Claude Code CLI.
  Settings choose the adapter.
- **The Codex call:**
  - KRONOS keeps an empty working folder in its data folder. Its only file is a short
    `AGENTS.md`: run no commands, write no files.
  - It runs `codex exec` there, with the read-only sandbox. The Orchestrator prompt and the
    conversation go on stdin, as one message (D72): the prompt is about 48 KB, too long for a
    Windows command line and, as far as we know, longer than the 32 KiB of `AGENTS.md` that
    Codex reads. The final message comes back through `--output-last-message`.
  - The default command line: `codex exec --sandbox read-only --skip-git-repo-check
    --output-last-message {out} -` (the working folder is not a git repo; `-` reads stdin).
  - It has a time limit, and the Windows process opens no console window.
  - Each call is stateless: KRONOS resends the conversation so far (the request, the earlier
    replies, the answers or problems), so nothing depends on Codex's session store.
  - The exact flags are checked against `codex exec --help` when KRONOS first runs. A settings
    field can override the command line if Codex changes them.
- **Reading a reply** (nothing in the repo does this yet):
  - take the one ```` ```vcl ```` block;
  - otherwise, take the whole reply if it starts with `VCL` or `# QUESTION:`.
  - Text outside the block is kept for the log: rule 13 allows two sentences before a
    reissued spec.
  - A QUESTION reply is caught before the tests run (the interpreter refuses it, R046).
- **The pipeline**, in process (not through `vclc`, whose import needs `tests/`):
  - `validator.validate_text`; then `vcl_fmt.format_text` (shown, not forced);
  - `vcl_interp.run_tests`;
  - `vcl_codegen.compile_text` plus `assemble` (a TASK) or `compile_build` (a BUILD);
  - `vba_lint.lint`;
  - optionally `vcl_mutate.mutate_report` as "test strength".

  One result object holds every problem with its code, every TEST with its failed
  assertions, the module text and timings.
- **The repair loop:**
  - If the spec has errors or a failing TEST, KRONOS sends back the coded messages and failed
    assertions (as `vclc` prints them; this is the feedback the corpus helpers used), asking
    for the corrected spec in one vcl block.
  - At most 2 more tries. Then the best attempt is shown with its problems, and the person can
    edit it.
- **The starter macro (a compiler addition):**
  - A TASK compiles to a Function that takes sheets, and Alt+F8 can't run that.
  - KRONOS asks for the real sheet name behind each sheet argument (for `$ncr`, say,
    "NCR Register").
  - The compiler then adds a `Run<Proc>` Sub that passes `ThisWorkbook.Worksheets("<name>")`
    and shows the result. A BUILD already has `modMain.Run`.
  - It is compiler output, lint-clean and tested like the rest (D31).
- **The module file:** ForgeLib plus the procedures, as `assemble` writes it today. It is
  saved with Windows line endings and an `Attribute VB_Name` line, so File > Import works as
  well as paste. One KRONOS module per workbook: each module carries its own ForgeLib, and two
  in one workbook would clash.
- **The run folder:** each run gets
  `<data folder>\runs\<date>_<time>_<short name>\`, holding:
  - `request.txt` and `conversation.jsonl` (every message both ways);
  - `attempt_N.vcl`;
  - `report.txt` and `report.json`;
  - the `.bas`.

  The data folder is `Documents\KGC KRONOS` by default (a setting) and never inside the repo.
- **Corpus candidates (D70; built in Phase 2d with its button, D72, D74):**
  - "Suggest for the corpus" runs the corpus gates (`vcl_corpus`) on a proven run, then writes
    `<data folder>\corpus_candidates\<id>\` with `task.txt`, `spec.vcl` and `meta.json`
    (`"source": "codex via KRONOS"`).
  - While the repo is public, a candidate reaches `data/verified/` only from a practice
    request, through `corpus-add` in a reviewed PR.
- **`python compiler/vclc.py ask "<request>"`**: the same engine from the command line, for
  tests and for anyone without the window.

### 2. The window: a four-step wizard (`app/kronos_app.py`, Tkinter; D74)

Boss chose a step-by-step wizard (31 Oct 2026), the recipe hidden behind a setting, and
screenshots of the window in every window PR.

**Principles**
- Plain words for non-technical colleagues, one thing per screen. The recipe (the VCL) stays
  hidden unless Settings turns it on.
- The window never freezes, always says what it is doing, and Cancel always works.
- Nothing unproven is offered as a macro: Save exists only for a proven recipe.
- KRONOS never runs a macro and never opens a workbook.
- It fits a 1366x768 laptop (smallest window 900x640, resizable) and reads well at 100 to 150%
  display scaling.

**The frame (every screen)**
```text
+----------------------------------------------------------------------+
| KGC KRONOS 0.2                              History  Settings  Help   |
|  (1) Ask  >  (2) Questions  >  (3) Review  >  (4) Save               |
+----------------------------------------------------------------------+
|                         ... the current step ...                     |
+----------------------------------------------------------------------+
|                                   [ Back ]   [ Next / main action ]  |
+----------------------------------------------------------------------+
```
The step bar ticks the done steps and highlights the current one. Questions is greyed out ("no
questions") when Codex asked none.

**Step 1, Ask**
- "What should Excel do?": a large box, with an example as its placeholder ("In the Calibration
  sheet, list every instrument due in the next 15 days on a sheet called Due List").
- Three tips: name the sheets and the column headers; say what must change and what must stay;
  practice requests with invented data only, for now.
- [Ask Codex] (or Ctrl+Enter) works once the first-run checklist is done and the box holds at
  least 20 characters.

**The working panel** (whenever Codex or the checks run)
```text
  Codex is writing the recipe (try 1 of 3)
  [==========            ]   0:42
  v Asking Codex   v Checking   ... Testing   . Building
  Codex calls used: 1                                        [ Cancel ]
```
Each try is a line in plain words ("Try 1: 2 of 3 tests passed, so KRONOS sent the problems back
to Codex"). Cancel stops Codex and returns to the step it came from, with [Try again].

**Step 2, Questions** (only when Codex asks)
- "Codex has a question before it can write the macro", and the question in a box.
- A "Your answer" box and [Send answer]. [Show the conversation] opens the history so far.
- Back returns to Ask. Changing the request there starts a new request; a confirm box says the
  old one stays in History.

**Step 3, Review**
- A banner:
  - PROVEN (green): "Passed every check and 3 of 3 practice tests".
  - NOT PROVEN (amber): "Codex couldn't get it right in 3 tries", with [Try again], [Change my
    request] and, when the recipe setting is on, [Edit the recipe].
  - When Codex gives no reply, or Cancel is pressed, the wizard goes back to the step it came
    from, with the reason in plain words and [Try again] (D75).
- **What it does:** the GOAL lines.
- **What it assumes:** the ASSUME lines, under "Check these match your workbook", with one
  required tick ("I checked these assumptions") before Next.
- **What it changes:** the MODIFIES lines in words ("the Status column of the NCR sheet"; "the
  whole Due List sheet, rebuilt on every run").
- **What it never does:** fixed text (no deleting or moving files, no internet, no Excel
  security changes, nothing outside the named sheets).
- **How it was proven:** the checks and n of n tests; [Show the tests] (a table of what each test
  checks and its result); [Measure strength] (mutation) when Settings turns it on.
- **Ask for a change:** a small box and [Send change]. It continues the conversation (rule 13)
  and comes back to Review.
- [Next: Save the macro] works only when the recipe is PROVEN and the assumptions are ticked.

**Step 4, Save**
- "Your real sheet names": one field per sheet argument, labelled with the recipe's name for it.
  The placeholder is the practice sheet name its TEST used. Each field is checked as you type
  (`sheet_name_ok`); text and number arguments get their own fields. A BUILD shows "this tool
  already names its sheets" and no fields.
- [Save macro file...]: a Save As dialog. The default is `modKronos<Name>.bas` in
  `Documents\KGC KRONOS\macros`. Then "Saved: <path>" and [Open the folder].
- "How to use it in Excel": the 5 numbered steps (`import_steps`) and [Copy steps].
- [Start a new request] goes back to step 1.

**Dialogs**
- **First-run checklist** (on first start; Ask Codex stays off until it is done; built in 2b, D76):
  1. the Codex CLI is found (`CodexAI.program`), with [Check again] and a hint;
  2. its flags match `codex exec --help` (`CodexAI.check`);
  3. [Test Codex], optional: one tiny prompt, using one call, showing "signed in" or the
     `codex login` hint;
  4. a required tick: "I turned off Improve the model for everyone (ChatGPT > Settings > Data
     Controls)";
  5. the data folder, with [Change...];
  6. a required tick: "Practice requests with invented data only, for now".
- **Settings** (kept in the data folder as `settings.json`):
  - the data folder;
  - the Codex timeout (default 300 seconds);
  - the Codex command line (advanced);
  - **Show the recipe (advanced)**, off by default;
  - Measure test strength, off by default;
  - Run the checklist again.
- **Recipe window** (only when Show the recipe is on): the VCL, editable, with line numbers; the
  problems (click one to jump to its line); [Re-check], [Tidy layout] (`vclc fmt`) and [Ask
  Codex to fix].
- **History:** past runs (date, request, status). Selecting one shows its report. [Open the
  folder]. [Use this macro again] re-proves its best attempt in a new run and goes to Save.
- **Help/About:** the version, the data folder and the guide.

**Working with Codex**
- The `CodexAI` adapter as built in Phase 1: the prompt and the whole conversation on stdin, the
  read-only sandbox, `--output-last-message`, no console window, stateless calls.
- Codex runs through `subprocess.Popen` with a time limit, so Cancel can stop it. The time limit
  is a setting.
- **Turns and tries:** a request, an answer, a change, or "Ask Codex to fix" is one turn, and each
  turn gets 3 tries (the first spec and 2 repairs). A question is not a try; a 4th question ends
  the run.
- **Codex errors in plain words**, each with [Try again]:
  - not found: install it, and see the guide;
  - an exit that mentions login: run `codex login` in a Command Prompt;
  - too slow: try again, or make the request smaller;
  - Cancel: "Cancelled".
- The working panel counts the Codex calls (1 to 3 per turn), for ChatGPT's usage limits.

**How it is built**
- `app/kronos_app.py` holds only Tkinter: ttk, the "vista" theme on Windows and "clam"
  elsewhere, Segoe UI and Consolas with fallbacks. It has the frame, one frame per step, the
  working panel and the dialogs, and no logic.
- `app/kronos_controller.py` is the wizard's state machine between the window and the engine
  (ask, question, review and save, with the working panel over them; D75).
  Engine work runs on a worker thread, and its events come back on a queue that the window
  reads every 100 ms (Tk is not thread-safe). It is unit-tested without a display.
- The engine gains:
  - a progress callback;
  - `Run.cancel()`;
  - `Run.follow_up(text)` and `Run.fix(text)`, each a new turn;
  - `summary(proof)` (the GOAL, ASSUME and MODIFIES lines and the warnings in words);
  - `sheet_hints(spec)`;
  - the settings `show_recipe`, `measure_strength`, `data_controls_confirmed`,
    `practice_confirmed` and `macro_folder`.
- `python app/kronos_app.py [--data DIR] [--ai fake --replies FILE]` starts it. `--smoke` runs
  one fake request through all four steps, saves the module and closes. CI runs it under a
  virtual display, and later it is KRONOS.exe's `--self-test`.

### 3. KRONOS.exe
- **The build:**
  - A Windows job in GitHub Actions builds one KRONOS.exe with PyInstaller. Its pin and hash
    go in a separate `requirements-build.txt`: it is a build tool and never runs inside
    KRONOS.
  - It bundles `compiler/`, `vcl/` (the prompt, grammar and semantics) and `vba/ForgeLib.bas`.
- **The self-test:** `KRONOS.exe --self-test` runs one request end to end with the fake AI,
  and CI runs it on the built exe.
- **Delivery:**
  - Boss downloads the exe from a GitHub Release, with exact clicks in the guide.
  - The window title shows the version.
  - The exe is unsigned, so IT may need to allow it.
- **The guide**, `docs/kronos-app-guide.md`, with exact clicks: install the Codex CLI,
  `codex login`, Data Controls, download KRONOS.exe, the first run, a practice request,
  importing the .bas, running the starter macro.

### 4. The team version (later, D71)
- Each person: KRONOS.exe, their own Codex CLI and ChatGPT login.
- **Shared folder on the company drive:**
  - the team's history;
  - approved specs;
  - corpus candidates;
  - the settings the team shares.
- **Writing safely:** each run is written as a new folder (so two people never write the same
  file), and settings changes take a lock file.
- A server program waits until a shared folder is not enough.
- **The move off GitHub:** the data folder and the training store point at the company drive.
  Real data may then go into the training store there, never into GitHub (D70).

## Phases (one PR each)

| Phase | What | Done when |
| --- | --- | --- |
| 0 | This design and D68 to D71; the reference files say the same | merged (PR #58) |
| 1 | The engine, `vclc ask`, the fake and Codex adapters, reply reading, the repair loop, the run folder, the starter macro | done (D72): unit tests with the fake AI cover every path (spec, question, repair, give up), and the starter macro is lint-clean |
| 2.0 | The window's design (the wizard) and the new order of work (D74) | the docs agree |
| 2a | The engine hooks (progress, cancel, follow-up, fix, summary, sheet hints, settings) and the wizard's controller | done (D75): controller tests with the fake AI cover every path |
| 2b | The wizard window: the frame, the four steps, the working panel, Cancel, the first-run checklist and `--smoke` | done (D76): it starts and runs a fake request in CI under a virtual display; Boss got screenshots of each step |
| 2c | Settings (with Show the recipe and the Recipe window), History, Help, the one-call Codex test, Codex errors in plain words | controller tests, the smoke test and screenshots |
| 2d | Suggest for the corpus: the corpus gates on a proven run, and a candidate folder in the data folder (never `data/`) | tests; the candidate passes the gates |
| 3 | KRONOS.exe, `--self-test`, the release and the guide | the Windows job builds the exe and its self-test passes |
| 4 | Boss's pilot on his PC, after IT approves the installs | three practice requests through Codex, each proven; problems logged |
| 5 | The team version on the company drive | two copies share one folder in a test |

**Order (D74):** the window comes first, and nothing else starts until it is done. Then Phase 3, so
it reaches Boss's PC, then the pilot once IT approves. The rest of Stage A (corpus batch 3, the
data generator, the benchmark) resumes after Phase 3 and may run while the pilot waits for IT.
The team version follows the pilot.

## Risks

- **Codex CLI flags change.** KRONOS checks the version and the help text, and the command
  line is a setting.
- **Codex writes weaker specs than Claude.** The toolchain proves every spec, the repair loop
  helps, and the prompt is the same. The pilot measures it.
- **IT blocks an unsigned exe.** IT allows it, or KRONOS runs from a folder with Python as the
  fallback.
- **Plus/Pro usage limits.** A request costs one to three Codex calls.
- **Data.** Personal accounts, so Data Controls are off. The repo is public, so practice data
  only, and runs stay on the PC.

## Notes for Phase 1 (found while designing, 30 Oct 2026)

Phase 1 used these notes; D72 records what it decided where they left a choice.

**Reading a reply and running the toolchain**
- Nothing in the repo extracts a ```` ```vcl ```` block from an AI reply. Left in, the fence gives only V007
  and V008 on lines 1 and 2. Rule 13 allows up to two sentences before a reissued spec, and a BUILD
  holds `---` lines.
- A `# QUESTION:` reply passes the validator, but `run_tests` raises R046 and the compiler refuses
  it, so catch it first. A NOTE step validates, gives RESULT -2 in the interpreter, and the compiler
  refuses it.
- Call the toolchain in process, not through `vclc`. Importing `vclc` pulls in `vcl_selftest`, which
  imports `tests/conformance_cases.py`, and the exe won't have `tests/`. The functions:
  - `validator.validate_text` (Problem: line, severity, message with `.code`);
  - `vcl_interp.run_tests(text, today)` (TestResult: passed, result, errors, failed with codes, log);
  - for a TASK, `vcl_codegen.compile_text` plus `assemble(procs, module_title(1))`; for a BUILD,
    `compile_build`;
  - `vba_lint.lint`, `vcl_fmt.format_text`, and `vcl_mutate.mutate_report` for "test strength".
- `vclc build` alone skips the validator, and with the default `-o dist` it overwrites the committed
  `dist/forge_module.bas`. `assemble` reads `vba/ForgeLib.bas` on every call, and the compiler needs
  `compiler/` on `sys.path`.
- The corpus helpers' repair feedback was plain `vclc` output, and it works well:
  - `path:line: error: message [Vnnn]`;
  - `ASSERT: ... is false (left X, right Y) [R051]`.

**Windows**
- Files are written with CRLF unless `newline=` is given (good for a .bas).
- cmd.exe doesn't expand `specs/*.vcl`.
- Messages can echo non-ASCII characters, so print them with care.
- The prompt is about 48 KB, too long for a command line: send it on stdin or in `AGENTS.md`.
- Run Codex with `CREATE_NO_WINDOW`, from a worker thread so the window keeps responding.

**The Codex CLI** (web search, 30 Oct 2026; the cloud can't reach developers.openai.com, so check
`codex exec --help` on Boss's PC, and make the command line a setting)
- `codex exec` prints only the final message to stdout; progress goes to stderr.
- Flags: `-o`/`--output-last-message <file>`, `--json` (JSON lines), `--sandbox read-only` (reported
  as exec's default), `--output-schema <file>`, and `codex exec resume --last`.
- It reads `AGENTS.md` at the working folder's root.
- It signs in with ChatGPT (`codex login`) or an API key. A native Windows installer exists.
- A Codex SDK exists too (TypeScript; the Python `openai-codex` is still beta), but use the CLI. Keep
  every call stateless.

**Out of date when designing, now fixed (D73)**
- `data/gaps.md` said plain `mutate` samples; mutation is deterministic, and the note now says
  how the two modes differ.
- R028 and `vcl_codegen` said "NOTE steps need the local model", against D31; they now say a spec
  with a NOTE goes back to Claude or Boss.
- The roadmap .docx listed "Desktop Excel (Microsoft 365 or 2021+)"; it now says Excel 2016.

## Not in scope

- Running VBA, opening or editing workbooks.
- UserForms (FORM waits, D49).
- A server program (the team version starts with a shared folder).
- Training any model before March 2027.
