# KRONOS Offline: the plan for Phase 1, and Boss's brief

A handover file, written on 14 Nov 2026. Boss approved this plan in one Claude Code session and
asked for it as a file so that another session can build it. Everything that session needs is
here: the approved plan (section 2), what was already checked (section 3), notes that save time
(section 4), and Boss's whole brief, word for word, with BlueForce's four launcher and server files
(Appendix A). BlueForce's repository is private, so Appendix A is the only copy.

**To the session that builds this:** read `CLAUDE.md` first, then this file to the end, then start
at section 2, step 1. Work as `CLAUDE.md` says (call him Boss, plain words, small changes, one PR,
merge when CI is green). The plan is approved: don't ask Boss to approve it again. When Phase 1's
PR writes `docs/kronos-offline.md`, move what still matters from this file into it and delete this
file in the same PR.

## 1. What Boss decided and answered

Boss's four decisions (14-11-2026, in his BlueForce session) are in Appendix A, section 1. Record
them as **D86** in `DECISIONS.md`, quoting him, with his reasons.

In the planning session he also answered two questions:

| Question | Boss's answer |
| --- | --- |
| Which browser is on the locked-down work PC? | **Chrome.** On old Windows (7, 8.1, Server 2008 R2 or 2012) that is version 109, its last. |
| How is version 1 handed over? | **A GitHub Release link plus the zip as a file in the chat.** A workflow builds the zip from main and publishes it as its own release, `kronos-offline-v1`, with its SHA-256, the way KRONOS.exe is published (D79). |

## 2. The approved plan: Phase 1, the proof

### Context

Boss's work PC is locked down (D69) and KRONOS.exe still waits for IT. BlueForce's offline copy
already runs on that PC with nothing installed: a tiny local web server started by a double-click,
and the app in the browser. KRONOS gets the same: **KRONOS Offline**, in the browser, no install,
AI by copy-paste, no password page, in BlueForce's look. KRONOS.exe stays as it is.

Phase 1 proves the route end to end and hands Boss **version 1** to try on the work PC: a
one-screen page that loads Python in the browser offline and checks, tests and builds a pasted
recipe, showing PROVEN or the problems. The wizard (Phase 2) and the rest (Phase 3) come later,
each with its own plan and PR (Appendix A, section 6).

### Steps (one PR, branch `claude/kronos-offline-phase-1`)

**Step 1. Prove the route first.** Download `pyodide@0.24.1`, run it in Node, load the toolchain
files and call `prove()` on `vcl/examples/example1_rev3.vcl`; then the same in a page with
Chromium, in a classic Web Worker. The Node half is already done (section 3). If importing a
module fails in the browser, fix it in the shared module so KRONOS.exe and the browser keep one
code path. **If the route can't work, stop and tell Boss why** (Appendix A, section 7).

**Step 2. The decision and the recipe.**
- `DECISIONS.md` **D86**: Boss's four decisions quoted, with his reasons; Pyodide 0.24.1 and why
  (section 3); port 5181; the zip published as its own release; third-party files fetched at build
  time by pinned hash and never committed (licence texts are committed).
- `docs/kronos-offline.md`: KRONOS's own copy of the recipe (the work PC, the four error rules,
  the zip layout, versions, the handover checks and words, the look), plus the version list.
- `CLAUDE.md`: the standing rule ("When Boss wants KRONOS offline, in whatever words, build the
  offline copy as docs/kronos-offline.md says and give its download link first"), the repo layout
  lines and the status section. `docs/kgc-kronos-context.md` by hand, then
  `python compiler/gen_context.py`.

**Step 3. One shared bridge module.** `compiler/kronos_web.py`: `check(text, today=None)` calls
`kronos_engine.prove`, `feedback` and `bas_text`, and returns plain data (proven or not, kind,
name, findings, TEST lines, the vclc-style problem lines, the .bas text) as canonical JSON. The
page's worker and the CPython tests both call it, so "identical results" is one string comparison.
No second copy of any engine logic.

**Step 4. The launcher and the three servers.**
`packaging/offline/Start KRONOS (Windows).bat` and `packaging/offline/server/serve.{mjs,py,ps1}`:
BlueForce's four files (Appendix A, section 3.7) with only the name, the port (5181) and the
"I can't find the KRONOS app" line changed, plus the version and date in the start message. Every
PowerShell fix stays (`param([string]$AppFolder)`, `Open-Browser`, `"%CD%\app"`). The .bat is
committed with CRLF, and a `.gitattributes` line keeps git from converting it.

**Step 5. The page: `web/`.** Plain HTML, CSS and JavaScript, no framework, no build step, and
only features that Chrome 109 and Firefox 115 have (no CSS nesting, no `:has()`, no `color-mix()`;
a classic Web Worker, not a module worker).
- `index.html`, `styles.css` (BlueForce's values as CSS variables, Appendix A section 4),
  `app.js`, `worker.js`, `logo.svg` (a navy circle with a white K).
- Top bar (96 px, navy, sticky): logo, "KGC KRONOS", "Offline", the version line.
- One screen: a "Paste a recipe" box, **Load the practice example** (so Boss can try it with no
  AI), **Check the recipe**; a working panel with Cancel (stops the worker and starts a fresh
  one); then the PROVEN (green) or NOT PROVEN (amber) banner with the TEST count or the problems.
- A small line showing the browser and its version, so Boss can say what the work PC runs.
- `worker.js` loads Pyodide from `pyodide/`, writes the toolchain files into its file system from
  a manifest, unpacks the lark wheel locally, imports `kronos_web`. Nothing loads from the internet.

**Step 6. The builder: `packaging/build_offline.py`** (Python, standard library only). It
assembles `KRONOS-Offline/` and `KRONOS-Offline-v<N>-<DD-MM-YYYY>.zip` (India date):
- `web/` into `app/`; every `compiler/*.py`, `vcl/semantics.py` and the data files of
  `packaging/kronos.spec` into `app/toolchain/`, same relative paths, with a manifest;
- Pyodide's five core files, the lark wheel and Inter's woff2 files, each downloaded once into a
  git-ignored cache and checked against a SHA-256 in `packaging/offline_pins.json` (lark's must
  equal the one in `requirements.txt`);
- the launcher and servers (CRLF), and licence texts in `app/licences/` (Pyodide MPL-2.0, Inter
  OFL, lark MIT, Python PSF);
- the version and date written into the page and the servers' start message. The version number
  comes from the version list; the zip is written with sorted names and fixed times so the same
  inputs give the same bytes.

**Step 7. Tests and CI.**
- `tests/test_offline.py`. Always-on checks: CRLF launcher, the PowerShell fixes, port 5181, pins
  well formed, lark's pin equals `requirements.txt`, the manifest covers everything `kronos_web`
  imports and every `kronos.spec` data file, version list in order. Checks that need the
  downloads: build the zip and verify its file list, CRLF and version line; under Pyodide in Node,
  `kronos_web.check` must equal CPython's for the 12 specs and the examples, with a fixed date;
  with Playwright and every non-localhost request blocked, the page shows the version, proves the
  example, shows problems for a broken recipe, Cancel works, and there are no page errors. These
  skip without network locally and **fail instead of skipping in CI**
  (`KRONOS_OFFLINE_REQUIRED=1`, as `KRONOS_LIVE_CODEX` does for the Codex tests).
- `.github/workflows/build-offline.yml`:
  - **build** (Ubuntu): builds the zip, runs all of the above, runs `serve.mjs` and `serve.py`
    from a freshly unpacked zip, uploads the zip and its checksum;
  - **old browsers** (Ubuntu 22.04): the same page test on Chromium 109 (Playwright 1.29) and
    Firefox 115 (Playwright 1.37). If those old builds can't be installed in CI, say so plainly
    and Boss's own PC is the check;
  - **powershell** (Windows): starts `serve.ps1` exactly as the .bat does and fetches the page and
    the .wasm file. It runs Windows PowerShell 5.1; version 2 can't be run there, so say that;
  - **release** (manual run from main with "release" ticked): publishes `kronos-offline-v<N>` with
    the zip and its SHA-256, as `build-exe.yml` does for KRONOS.exe.
- Test tools (Playwright) are pinned with hashes in their own requirements file and installed
  only by CI.

**Step 8. Hand over v1.** Merge when CI is green, run the release, then tell Boss: the link first,
"This is version 1, built DD-MM-2026.", the five steps (Appendix A, section 3.6), that Internet
Explorer can't show KRONOS, and what he can do now (load the example, check it, see PROVEN). Send
the zip and screenshots in the chat. Add the release's commit to the version list in a docs-only
follow-up.

### Not in Phase 1

The four-step wizard, the copy-paste AI step, the .bas download, history, the practice-run check,
the corpus suggestion, Help. No change to KRONOS.exe, the Codex adapter or its release. No real
KGC data.

### Verification

- The five project checks (`unittest discover`, validator, `vclc all`, both `--check` scripts).
- `python packaging/build_offline.py`, then the brief's handover checks on the built zip
  (Appendix A, section 3.5): the name, the `unzip -Z -1` list, the PowerShell fixes, CRLF
  (`file`), `node KRONOS-Offline/server/serve.mjs` and `python3 KRONOS-Offline/server/serve.py`
  from a fresh folder, the page through Playwright with the internet blocked, and the .bas equal
  byte for byte to CPython's.
- CI green on all four jobs of `build-offline.yml` and on `check-specs.yml`; `build-exe.yml` still
  green (KRONOS.exe untouched).
- After the release: the link opens without signing in, and the file's SHA-256 matches.

## 3. What was checked while planning (14 Nov 2026)

**The Pyodide version: 0.24.1.** Pyodide's own docs (`docs/usage/index.md` at each release tag)
give these minimum browsers:

| Pyodide | Python | Chrome | Edge | Firefox |
| --- | --- | --- | --- | --- |
| 0.21.3 to 0.24.1 | 3.10.2, then 3.11.2 and 3.11.3 | 71 | 80 | 70 |
| 0.25.1, 0.26.4, 0.27.8, 0.28.3, 0.29.5 | 3.11.3 to 3.13.2 | 112 | not listed | 112 |

So every release from 0.25 on documents a minimum above Chrome/Edge 109. **0.24.1** is the last
that covers Chrome 109, and it also covers Firefox 115. It runs Python 3.11.3, the Python this
repo's CI uses, which helps the byte-for-byte comparison. The work PC's browser is Chrome.

**Where the files come from, with their hashes** (checked by downloading them):

| File | Where | SHA-256 |
| --- | --- | --- |
| `pyodide-0.24.1.tgz` (the npm package `pyodide@0.24.1`) | `https://registry.npmjs.org/pyodide/-/pyodide-0.24.1.tgz` | `028563a8ecf77e48a520e54fff5b7d6c955f2f8a69387f42ca38969cdbf67eae` |
| `lark-1.3.1-py3-none-any.whl` | PyPI (take the URL from `https://pypi.org/pypi/lark/1.3.1/json`) | `c629b661023a014c37da873b4ff58a817398d12635d3bbb2c5a03be7fe5d1e12`, the hash already in `requirements.txt` |

The npm package holds the five core files (`pyodide.js`, `pyodide.asm.js`, `pyodide.asm.wasm`,
`python_stdlib.zip`, `pyodide-lock.json`) and `pyodide.mjs` for Node. It holds no licence file, so
the licence texts are committed in the repo. Pin each core file's own SHA-256 in
`packaging/offline_pins.json` when the builder is written. Inter is on npm as `@fontsource/inter`
(5.3.0 on 14 Nov 2026, OFL-1.1); pin the version and hash that are used.

**The route works in Node, with no change to the toolchain.** Pyodide 0.24.1 under Node 22, the
lark wheel unpacked locally, and the repo's files written unchanged into its file system as
`/kronos/compiler/*.py`, `/kronos/vcl/{semantics.py,grammar.lark,orchestrator_prompt.txt}`,
`/kronos/vcl/examples/example1_rev3.vcl` and `/kronos/vba/ForgeLib.bas`:

| | Normal Python 3.11 | Pyodide 0.24.1 (Python 3.11.3) |
| --- | --- | --- |
| `kronos_engine.prove(example1_rev3.vcl, date(2026, 11, 14))` | proven, 3 of 3 TESTs, 0 findings | proven, 3 of 3 TESTs, 0 findings |
| SHA-256 of the built module | `6cb9dc49…9acaaf53` | `6cb9dc49…9acaaf53`, the same |
| Time | | Pyodide loads in about 1.5 s; import and prove take about 1.1 s |

`import kronos_engine` works there as it is, although it imports `kronos_ai` (which imports
`subprocess` and `threading`): importing them is fine, only using them would fail.

**Not checked yet:** the same in a real page (a classic Web Worker in Chromium), anything on
Chromium 109 or Firefox 115, and `serve.ps1`. Nothing is built: the branch
`claude/kronos-offline-phase-1` holds only this file.

**Other facts from the repo:**
- `kronos_engine.prove(text, today, mutation, step)` is the whole Phase 1 job: validator, TESTs,
  compiler and lint, in order, stopping at the first stage that fails. `feedback(proof)` gives the
  problems as vclc prints them; `bas_text(name, module)` writes the .bas with Windows line endings.
- The toolchain finds its data files through `compiler/resources.py` (`ROOT`), so the layout
  `compiler/`, `vcl/`, `vba/` under one root is all the browser needs.
- `packaging/kronos.spec` bundles `vcl/grammar.lark`, `vcl/orchestrator_prompt.txt`,
  `vcl/examples/example1_rev3.vcl` and `vba/ForgeLib.bas`.
- The window's texts to reuse later are in `app/kronos_app.py`: `EXAMPLE`, `TIPS`, `NEVER`,
  `STEP_NAMES`, `COLOURS`.
- The last decision is D85, so this work is **D86**.

## 4. Notes that save time

- **Node and Pyodide 0.24.1:** pass `new Uint8Array(buffer)` to `pyodide.unpackArchive` and
  `FS.writeFile`. A Node `Buffer` fails with "Unknown typed array type 'Buffer'". In Node import
  `pyodide.mjs`; in the page's classic worker use `importScripts('pyodide/pyodide.js')`.
- **The servers' file types:** BlueForce's `serve.ps1` already sends `.wasm` as
  `application/wasm`. Pyodide also fetches `python_stdlib.zip`, and the page fetches `.py`,
  `.lark`, `.txt`, `.bas` and `.whl` files; the default `application/octet-stream` is fine for
  those, but check each of the three servers serves them.
- **Old browsers in CI:** Playwright 1.29 ships Chromium 109 and Playwright 1.37 ships Firefox
  115. Those versions list Ubuntu 22.04, not 24.04, so use `runs-on: ubuntu-22.04` for that job.
- **GitHub from a Claude Code cloud session** (as met on 14 Nov 2026): `gh pr create` fails
  because GraphQL is not available; open and merge pull requests through the REST API
  (`gh api -X POST repos/4ptech1/kgc-kronos/pulls ...`, then
  `gh api -X PUT repos/4ptech1/kgc-kronos/pulls/<n>/merge -f merge_method=merge`).
  `gh workflow run` works. Deleting a branch and other direct writes are refused, so a release is
  published by a workflow job with `contents: write`, as `build-exe.yml` does.
- **A fresh clone tracks only `main`:** after the first push of a new branch, add
  `+refs/heads/claude/*:refs/remotes/origin/claude/*` to `remote.origin.fetch`, or the session's
  stop check reports unpushed commits that are in fact pushed.
- **Python here:** the project wants Python 3.11 (`python3.11 -m venv`, then
  `pip install --require-hashes -r requirements.txt`). The Tk window tests need Python 3.12 with
  `python3-tk` under `xvfb-run`; they skip without it.
- **This file** is on the public repo: keep real KGC data and the work PC's real folder path out
  of everything, as Appendix A says.

## Appendix A. Boss's brief, word for word (14-11-2026)

Boss wrote this in his BlueForce session and gave it to the planning session. It is copied here
unchanged, apart from Markdown code fences around BlueForce's four files.

You are Claude Code, working in the KGC-KRONOS repository (github.com/4ptech1/KGC-KRONOS). Boss
asks you to build KRONOS OFFLINE: a copy of KGC KRONOS he can use on his locked-down work PC with
nothing installed, the same way he already uses the offline copy of BlueForce, another project of
his. It must look like BlueForce.

You can't see the BlueForce repository (it is private and not in your session). So everything you
need from it is in this message: Boss's decisions, BlueForce's offline-copy recipe, its four
launcher and server files, and its look. Read the whole message before you start.

### 0. HOW TO WORK

- CLAUDE.md and AGENTS.md rule everything here: call him Boss, plain words, keep it brief, one
  thing at a time, small changes. The repo is public, so synthetic data only. Never weaken a test
  or a check. D25 still holds (general tools such as Python are allowed; nothing downloaded goes
  into the model's training data).
- Start in plan mode. First read CLAUDE.md, docs/kronos-app.md, compiler/kronos_engine.py,
  compiler/kronos_ai.py, compiler/resources.py, packaging/kronos.spec and
  app/kronos_controller.py. Then write a plan for Phase 1 only (section 6) and ask Boss to
  approve it.
- Ask Boss only what you truly can't decide from this message or the repo. Otherwise state your
  assumption and carry on.
- Record Boss's decisions (section 1) as the next D-number in DECISIONS.md, quoting him, with the
  reasons given here.

### 1. WHAT BOSS DECIDED (14-11-2026, in his BlueForce session)

1. KRONOS Offline runs in the browser, with no install. Like BlueForce's offline copy: he unzips
   a folder, double-clicks "Start KRONOS (Windows)", and KRONOS opens in Chrome, Edge or Firefox.
   KRONOS's checks, tests and macro writer (the Python toolchain) run inside the browser through
   Pyodide (Python compiled to WebAssembly), bundled in the zip. No exe, no Python and no
   installs, so no IT approval is needed. KRONOS.exe stays as it is, for when IT approves it.
2. The AI step is copy-paste with an AI chat. Offline there is no Codex, and a web page can't
   start a program. KRONOS gives a "Copy for the AI" button. Boss pastes it into his claude.ai
   Orchestrator project or into ChatGPT, then pastes the AI's reply back. KRONOS checks, tests and
   builds the macro in the browser. When something fails it offers "Copy the problems for the AI",
   for up to 2 more tries, as the engine does now.
3. No password page. BlueForce's offline copy has one; KRONOS doesn't need it, because its code
   is public and holds no secrets.
4. It looks like BlueForce (section 4).

Why: the work PC is locked down (D69) and KRONOS.exe is still waiting for IT, while BlueForce's
offline copy already runs on that same PC with nothing installed, through Windows' own PowerShell
and the browser. A browser page starts no new program, so it is the copy most likely to work
there.

### 2. WHAT CAN AND CAN'T RUN IN THE BROWSER (checked on 14-11-2026)

- compiler/*.py imports only Python's standard library and lark (pure Python; lark==1.3.1 in
  requirements.txt). The parser, validator, type checker, interpreter, fixtures, compiler
  (vcl_codegen), vba_lint, vcl_fmt and vcl_pilot can all run under Pyodide.
- Not in the browser: tkinter (the window), and subprocess and threading for Codex
  (kronos_ai.CodexAI and the controller's process handling). The browser copy needs its own small
  "AI" adapter that waits for Boss's pasted reply instead of calling Codex. Reuse the engine's own
  message builders and retry logic; don't write a second copy of the loop.
- One source of truth: the browser runs the very same compiler/*.py files and data files that
  resources.py and kronos.spec bundle. Never fork or edit the toolchain just for the browser. If
  a module needs a change to run there (a thread, a path), make it in the shared module, so
  KRONOS.exe and the browser copy keep passing the same tests.

### 3. BLUEFORCE'S OFFLINE-COPY RECIPE (copy it, adapted to KRONOS)

#### 3.1 The work PC, as BlueForce met it

- Windows, an older version (Explorer looks like Windows 7 or Server 2008 R2, perhaps over Remote
  Desktop). The black window's title is set by the launcher.
- The folder sits on a mapped network drive (a drive letter such as Y:), not on the PC's own
  disk. Never write the real folder path into the repo: it is public.
- No Node.js or Python, so the launcher falls back to PowerShell, probably version 2.
- No default web browser: starting a URL with Start-Process fails with "No application is
  associated with the specified file".
- A browser that runs the app: Chrome or Edge 109+, or Firefox 115+, the last versions for
  Windows 7 and 8.1. Internet Explorer can't.

#### 3.2 The errors already met, and the rule for each (keep every rule)

1. HTTP 500 on every page (PowerShell): serve.ps1 found its folder through $PSScriptRoot, which
   was empty on that PC.
   Rule: the .bat passes the app folder in:
   `powershell -NoProfile -ExecutionPolicy Bypass -File "server\serve.ps1" "%CD%\app"`.
   serve.ps1 takes it as `param([string]$AppFolder)`. Never rely on $PSScriptRoot alone.
2. A red "Start-Process: No application is associated" error: the PC has no default browser. The
   app was running anyway; only opening the browser failed.
   Rule: never call a bare `Start-Process "http://..."`. Open-Browser uses the default browser
   only when one is set and isn't Internet Explorer. Otherwise it looks for Chrome, Edge or
   Firefox by name in %LOCALAPPDATA%, Program Files (x86) and Program Files. If none is found, it
   prints one plain line telling the person to open Chrome, Edge or Firefox and type the address.
   Nothing in a server script may end in a red PowerShell error.
3. A text file shown as one jumbled line: Notepad before Windows 10 (1809) ignores Unix line
   endings.
   Rule: the .bat, and any text file in the zip, keeps Windows line endings (CRLF). Check with
   `file`: "with CRLF line terminators".
4. A blank page in Internet Explorer.
   Rule: the message to Boss says to use Chrome, Edge or Firefox.

Also keep:
- The servers listen on localhost / 127.0.0.1 only, never on the network. "Port ... is busy"
  means the app is already running.
- No Node.js advice for old Windows: current Node.js needs Windows 10 or later. On older Windows
  the PowerShell server is the way.
- Boss never runs commands in a terminal: everything is a double-click.

#### 3.3 The zip

- One file, KRONOS-Offline-v<N>-<DD-MM-YYYY>.zip (its version number and build date, India time:
  `TZ=Asia/Kolkata date +%d-%m-%Y`), for Windows only. The folder inside is always
  KRONOS-Offline/, whatever the version, so Boss's folder keeps its name.

```
KRONOS-Offline/
  Start KRONOS (Windows).bat   node -> py -3 -> python -> PowerShell, in that order (CRLF)
  server/serve.mjs             Node.js: http://localhost:5181/, 127.0.0.1 only
  server/serve.py              Python 3: the same
  server/serve.ps1             PowerShell only: the same (3.2 rules 1 and 2)
  app/                         the page, its scripts, fonts, Pyodide and the toolchain files
```

- Port 5181, not BlueForce's 5180: Boss may run both at once, and a different port also keeps
  each app's browser storage apart.
- No guide inside the zip and no Mac or Linux launcher: what Boss does is said in the chat (3.6).
  The Node and Python servers stay, because the .bat falls back to them.
- Versions: N starts at 1 and goes up by one with every rebuild. The version and date show in the
  page ("Version 1 · 14-11-2026", in its top bar or footer) and in the server window's start
  message. Keep a version list: one row per copy, with the version, the date, the main commit it
  was built from, and what changed in a few words.

#### 3.4 Where it lives (adapt to KRONOS's own way)

- BlueForce keeps its zip on a downloads branch that is never merged. KRONOS already publishes
  KRONOS.exe as a GitHub Release from a manual workflow run (D79), with its checksum, and the
  repo is public, so a release link works without signing in.
- Recommended: publish the offline zip the same way, as its own release (a tag such as
  kronos-offline-v1) built from main by a workflow, with its SHA-256. If you see a better fit
  with KRONOS's rules, tell Boss.
- Whatever you choose, every handover gives Boss the download link first, with the version and
  date, and also sends the zip as a file in the chat if your tools can.

#### 3.5 Checks before handing over (all must pass)

- Name: KRONOS-Offline-v<N>-<DD-MM-YYYY>.zip, with N one more than the last.
- File list (`unzip -Z -1`): the launcher, server/serve.mjs, server/serve.ps1, server/serve.py
  and app/. Windows only, nothing stray.
- PowerShell fixes still there: serve.ps1 has `param([string]$AppFolder)` and Open-Browser, the
  .bat passes `"%CD%\app"`, and the .bat is CRLF.
- It runs: unpack the new zip into a fresh folder, run `node KRONOS-Offline/server/serve.mjs`
  (and once `python3 KRONOS-Offline/server/serve.py`), and open http://localhost:5181/ with
  Playwright, with every request outside localhost blocked (to prove nothing loads from the
  internet):
  - the page shows the version line;
  - a practice request goes through all four steps, using a scripted "AI reply" (a proven spec
    pasted in);
  - the downloaded .bas is byte for byte what CPython's engine writes for the same reply;
  - no page errors.
  Then stop the server.
- PowerShell can't run in the cloud container. Better: a Windows job in GitHub Actions
  (windows-latest has PowerShell) starts serve.ps1 with the app folder, as the .bat does, and
  fetches the page. If no job runs it, say plainly that serve.ps1 was checked by reading only.

#### 3.6 What to tell Boss at each handover (short, step by step)

Start with the download link and "This is version N, built DD-MM-YYYY.", then:
1. Close the black window if it's open.
2. Delete the old KRONOS-Offline folder.
3. Right-click the zip, choose Extract All..., then Extract.
4. Double-click Start KRONOS (Windows).
5. If the browser doesn't open by itself, open Chrome, Edge or Firefox and type localhost:5181.

Mention anything new in one line, and say that Internet Explorer can't show KRONOS.

#### 3.7 BlueForce's four files (from BlueForce-Offline v25, 14-11-2026), to copy and adapt

Copy them, then change only: "BlueForce" to "KRONOS" in titles, comments and messages; the port
5180 to 5181; and the app's name in "I can't find the ... app here". Keep everything else, above
all the PowerShell fixes. Save the .bat with CRLF line endings.

**Start BlueForce (Windows).bat** (save as CRLF)

```bat
@echo off
title BlueForce
pushd "%~dp0"
node -v >nul 2>nul
if not errorlevel 1 goto usenode
py -3 -c "pass" >nul 2>nul
if not errorlevel 1 goto usepy
python -c "pass" >nul 2>nul
if not errorlevel 1 goto usepython
goto usepowershell

:usenode
node "server\serve.mjs"
goto done

:usepy
py -3 "server\serve.py"
goto done

:usepython
python "server\serve.py"
goto done

:usepowershell
powershell -NoProfile -ExecutionPolicy Bypass -File "server\serve.ps1" "%CD%\app"
goto done

:done
popd
echo.
echo BlueForce has stopped. You can close this window.
pause
```

**server/serve.ps1**

```powershell
# BlueForce offline: the same tiny server, for Windows computers with neither Node.js nor Python.
# Uses only PowerShell, which comes with Windows. The launcher passes the app folder in; this
# script does not depend on knowing where it is itself.
param([string]$AppFolder)

$port = 5180

if (-not $AppFolder) {
  $here = $PSScriptRoot
  if (-not $here) { $here = Split-Path -Parent $MyInvocation.MyCommand.Path }
  if (-not $here) { $here = (Get-Location).Path }
  $AppFolder = [System.IO.Path]::Combine($here, '..', 'app')
}
$root = [System.IO.Path]::GetFullPath($AppFolder).TrimEnd('\')
$index = [System.IO.Path]::Combine($root, 'index.html')
if (-not [System.IO.File]::Exists($index)) {
  Write-Host ""
  Write-Host "I can't find the BlueForce app here: $root"
  Write-Host "Make sure you unzipped the whole folder, and start BlueForce from inside it."
  exit 1
}

$types = @{
  '.html'='text/html; charset=utf-8'; '.js'='text/javascript; charset=utf-8'; '.mjs'='text/javascript; charset=utf-8';
  '.css'='text/css; charset=utf-8'; '.svg'='image/svg+xml'; '.png'='image/png'; '.jpg'='image/jpeg'; '.ico'='image/x-icon';
  '.json'='application/json'; '.wasm'='application/wasm'; '.woff'='font/woff'; '.woff2'='font/woff2'
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
try {
  $listener.Start()
} catch {
  Write-Host ""
  Write-Host "Port $port is busy. BlueForce may already be running: open http://localhost:$port/ in your browser."
  exit 1
}

Write-Host ""
Write-Host "BlueForce is running at http://localhost:$port/"
Write-Host "App folder: $root"
Write-Host "Keep this window open while you use it. Close it (or press Ctrl+C) to stop."
Write-Host ""

# Internet Explorer can't run BlueForce, and some computers have no default browser at all
# (Start-Process then fails with "No application is associated"). So unless another default
# browser is set, look for Chrome, Edge or Firefox by name.
function Open-Browser([string]$url) {
  $default = $null
  try {
    $default = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice' -ErrorAction Stop).ProgId
  } catch { }
  if ($default -and $default -notlike 'IE.*') {
    try { Start-Process $url -ErrorAction Stop; return $true } catch { }
  }
  $places = @($env:LOCALAPPDATA, ${env:ProgramFiles(x86)}, $env:ProgramFiles) | Where-Object { $_ }
  foreach ($name in @('Google\Chrome\Application\chrome.exe', 'Microsoft\Edge\Application\msedge.exe', 'Mozilla Firefox\firefox.exe')) {
    foreach ($place in $places) {
      $exe = [System.IO.Path]::Combine($place, $name)
      if ([System.IO.File]::Exists($exe)) {
        try { Start-Process -FilePath $exe -ArgumentList $url -ErrorAction Stop; return $true } catch { }
      }
    }
  }
  return $false
}

if (-not (Open-Browser "http://localhost:$port/")) {
  Write-Host "Your browser did not open by itself. Open Chrome, Edge or Firefox and type:  localhost:$port"
  Write-Host "(Internet Explorer can't show BlueForce.)"
  Write-Host ""
}

while ($listener.IsListening) {
  $context = $listener.GetContext()
  $response = $context.Response
  try {
    $relative = [System.Uri]::UnescapeDataString($context.Request.Url.AbsolutePath).TrimStart('/')
    if ($relative -eq '') { $relative = 'index.html' }
    $file = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($root, $relative.Replace('/', '\')))
    $inside = $file.StartsWith($root + '\', [System.StringComparison]::OrdinalIgnoreCase)
    if (-not $inside -or -not [System.IO.File]::Exists($file)) {
      $response.StatusCode = 404
    } else {
      $bytes = [System.IO.File]::ReadAllBytes($file)
      $type = $types[[System.IO.Path]::GetExtension($file).ToLower()]
      if (-not $type) { $type = 'application/octet-stream' }
      $response.StatusCode = 200
      $response.ContentType = $type
      $response.AddHeader('Cache-Control', 'no-cache')
      $response.ContentLength64 = $bytes.Length
      $response.OutputStream.Write($bytes, 0, $bytes.Length)
    }
  } catch {
    Write-Host ("Problem with " + $context.Request.Url.AbsolutePath + ": " + $_.Exception.Message)
    try { $response.StatusCode = 500 } catch { }
  } finally {
    try { $response.Close() } catch { }
  }
}
```

**server/serve.mjs**

```js
// BlueForce offline: a tiny web server for the app folder, using only what Node.js already has.
// Nothing leaves this computer: it listens on localhost only.
import { createServer } from 'node:http'
import { readFile, stat } from 'node:fs/promises'
import { extname, join, normalize, resolve, sep } from 'node:path'
import { fileURLToPath } from 'node:url'
import { exec } from 'node:child_process'

const ROOT = resolve(fileURLToPath(new URL('../app/', import.meta.url)))
const PORT = 5180
const TYPES = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.svg': 'image/svg+xml', '.png': 'image/png', '.jpg': 'image/jpeg', '.ico': 'image/x-icon',
  '.json': 'application/json', '.wasm': 'application/wasm', '.woff': 'font/woff', '.woff2': 'font/woff2', '.map': 'application/json',
}

const server = createServer(async (req, res) => {
  try {
    let path = normalize(decodeURIComponent(new URL(req.url, 'http://x').pathname))
    let file = resolve(join(ROOT, path))
    if (file !== ROOT && !file.startsWith(ROOT + sep)) throw new Error('outside')
    if ((await stat(file).catch(() => null))?.isDirectory()) file = join(file, 'index.html')
    const body = await readFile(file)
    res.writeHead(200, { 'Content-Type': TYPES[extname(file).toLowerCase()] ?? 'application/octet-stream', 'Cache-Control': 'no-cache' })
    res.end(body)
  } catch {
    res.writeHead(404, { 'Content-Type': 'text/plain' })
    res.end('Not found')
  }
})

server.on('error', (error) => {
  if (error.code === 'EADDRINUSE') console.log(`\nPort ${PORT} is busy. BlueForce may already be running: open http://localhost:${PORT}/ in your browser.\n`)
  else console.log(error.message)
  process.exit(1)
})

server.listen(PORT, '127.0.0.1', () => {
  const url = `http://localhost:${PORT}/`
  console.log(`\nBlueForce is running at ${url}\nKeep this window open while you use it. Close it (or press Ctrl+C) to stop.\n`)
  const open = process.platform === 'win32' ? `start "" "${url}"` : process.platform === 'darwin' ? `open "${url}"` : `xdg-open "${url}"`
  exec(open, () => {})
})
```

**server/serve.py**

```python
# BlueForce offline: the same tiny server, for computers with Python instead of Node.js.
import functools, http.server, os, socketserver, sys, threading, webbrowser

PORT = 5180
ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'app')

class Handler(http.server.SimpleHTTPRequestHandler):
    extensions_map = {**http.server.SimpleHTTPRequestHandler.extensions_map,
                      '.js': 'text/javascript', '.mjs': 'text/javascript', '.wasm': 'application/wasm',
                      '.woff2': 'font/woff2', '.woff': 'font/woff', '.svg': 'image/svg+xml'}
    def log_message(self, *args):
        pass
    def end_headers(self):
        self.send_header('Cache-Control', 'no-cache')
        super().end_headers()

# Mac and Linux keep a just-closed port reserved for a while, so allow re-using it. Not on Windows,
# where that would let two copies share the port without noticing.
socketserver.TCPServer.allow_reuse_address = os.name != 'nt'
try:
    server = socketserver.TCPServer(('127.0.0.1', PORT), functools.partial(Handler, directory=ROOT))
except OSError:
    print(f'\nPort {PORT} is busy. BlueForce may already be running: open http://localhost:{PORT}/ in your browser.\n')
    sys.exit(1)
url = f'http://localhost:{PORT}/'
print(f'\nBlueForce is running at {url}\nKeep this window open while you use it. Close it (or press Ctrl+C) to stop.\n')
threading.Timer(0.5, lambda: webbrowser.open(url)).start()
try:
    server.serve_forever()
except KeyboardInterrupt:
    pass
```

### 4. BLUEFORCE'S LOOK (use these exact values)

BlueForce is built with React and Ant Design. KRONOS Offline is plain HTML, CSS and JavaScript,
so recreate the look with CSS variables, not those libraries.

- Font: Inter. Bundle its woff2 files and its OFL licence in app/fonts (no Google Fonts: the page
  must work offline). Fallback "Segoe UI", system-ui, sans-serif. Text black #000000; secondary
  text #595959.
- Top bar: 96 px high, navy #0B2A6F (the BlueForce logo's navy), white title (small lines in
  #D6DEEF). It stays at the top while the page scrolls.
- Logo: a navy #0B2A6F circle with a bold white mark, like BlueForce's "BF". For KRONOS a white
  "K". Make it an SVG in app/.
- Left sidebar (if the layout has one): #F7F7F7, black text at 15 px, items 40 px high with
  radius 8, hover #EFE6D4 (sand beige), the current item underlined.
- Page: white #FFFFFF.
- Buttons, all of them: light orange #FFD8A8, bold (700) black text, no border, radius 8. On
  hover or press: navy #0B2A6F with white text. Main and secondary buttons look the same.
- Boxes (text boxes, text areas, dropdowns, date fields): sand-beige fill #EFE6D4, no border;
  hover #E5D9C2; pressed #DBCDB2. While typing in it: white, with a mid-grey edge #595959. An
  open dropdown list is white, its chosen row #EFE6D4, its hovered row #F7F2E8.
- Tables: 12 px text; 12 px semibold headers on a light band #F7F2E8; a thin black line
  (#000000) under every row and under the header, none between columns; cells 8 px top and bottom
  and 4 px each side; every other row #EFE6D4; hover #E5D9C2.
- Cards and panels: light grey #F5F5F5, no border, no line under the title, radius 8.
- Status colours, for status only: green #1E8E5A (the PROVEN banner may use #00870E with white
  text), amber #B26A00, red #C8372D.
- KRONOS's current COLOURS map onto these: accent -> navy #0B2A6F; ok -> green; warn -> amber;
  note -> red; bg -> white; panel -> #F5F5F5; line -> #000000 for table lines, #BFBFBF for soft
  edges.
- Pop-ups with fields: centred, the label on the left and its box on the right, 12 px between
  fields, one main button, no Cancel (close with the X). The one exception is KRONOS's "working"
  panel, which keeps its Cancel button.
- Corners: radius 8 everywhere. Keyboard focus: a visible navy outline.
- Words: plain and short, as the KRONOS window writes them. Keep its EXAMPLE, TIPS and NEVER
  texts.

### 5. WHAT TO BUILD

#### 5.1 In the repo

- web/ (or a name you prefer): the page's source: index.html, styles.css, app.js, a Web Worker
  (worker.js) that loads Pyodide and the toolchain, the logo and the fonts. Plain HTML, CSS and
  JavaScript: no framework and no build step (no React, no Ant Design, no npm build) unless Boss
  decides otherwise.
- packaging/build_offline.py (Python, standard library only), which assembles KRONOS-Offline/
  and the zip:
  - copies web/ into app/;
  - copies the toolchain's .py and data files (the same list as resources.py and kronos.spec)
    into app/;
  - adds Pyodide: one pinned version and its core files (pyodide.js, pyodide.asm.js,
    pyodide.asm.wasm, python_stdlib.zip, pyodide-lock.json), from the npm registry (package
    "pyodide"; package registries are usually reachable from Claude Code on the web) or from
    Pyodide's releases, each file checked against a SHA-256 kept in the repo. Choose a version
    whose browser support covers Chrome/Edge 109 and Firefox 115, and name it in the decision;
  - adds the lark==1.3.1 wheel (py3-none-any) from PyPI, checked against the hash in
    requirements.txt. The page loads it from the local file (micropip with a local URL, or
    unpacked into Pyodide's file system), never from the internet;
  - adds the three servers and the launcher (CRLF), and the licence files (Pyodide MPL-2.0,
    Inter OFL, lark MIT, Python PSF) in app/licences/;
  - writes the version and date into the page and into the servers' start message;
  - zips KRONOS-Offline/ as KRONOS-Offline-v<N>-<date>.zip.
- docs/kronos-offline.md: KRONOS's own copy of this recipe (sections 3 to 5 as they apply to
  KRONOS), so the next session can rebuild it without this message. Add a standing rule to
  CLAUDE.md: "When Boss wants KRONOS offline, in whatever words, build the offline copy as
  docs/kronos-offline.md says and give its download link first."
- Tests, in tests/ and CI:
  1. build_offline.py makes the expected file list, a CRLF launcher, the PowerShell fixes and the
     pinned hashes;
  2. under Pyodide (it also runs in Node.js, so CI can use it), the browser engine gives
     byte-identical results to CPython for the repo's specs (validator messages, TEST results,
     the .bas) and for a scripted copy-paste replay;
  3. the page itself, with Playwright, offline (every non-localhost request blocked).

#### 5.2 The page: KRONOS.exe's four-step wizard, in BlueForce's look

- Top bar: the logo, "KGC KRONOS", the steps (1 Ask, 2 Questions, 3 Review, 4 Save) with the
  current one marked, History and Help on the right, and the version line.
- Step 1, Ask: the request box, with KRONOS's EXAMPLE and TIPS. Then:
  - "Copy request for the Orchestrator": Boss's claude.ai Orchestrator project has KRONOS's
    instructions as its project instructions, so this copies the request (and the conversation
    so far) only. If vcl/orchestrator_prompt.txt is newer than what Boss last put in that
    project, tell him in one line, with the exact clicks to update it.
  - "Copy request with instructions (ChatGPT or a new chat)": the Orchestrator prompt and the
    request as one message, exactly what KRONOS sends Codex today.
  - One short line of steps: paste it into the AI chat, copy the AI's whole reply, paste it
    below.
  - A large "Paste the AI's reply" box, and "Check the reply".
  The message with instructions is about 48 KB: call navigator.clipboard.writeText inside the
  click handler, and fall back to selecting the text in a box for Ctrl+C.
- Step 2, Questions: when the reply is a question (the engine already tells them apart), show
  it. Boss's answer goes into the next "Copy..." (the conversation continues), and he pastes the
  next reply.
- While the toolchain works: a panel covers the step with what is happening and a Cancel button
  (Cancel stops the worker and starts a fresh one). The page never freezes.
- Step 3, Review: the PROVEN (green) or NOT PROVEN (amber) banner, and the same parts as the
  KRONOS window: what the macro does, its assumptions with the tick, what it changes, what it
  never does, how it was tested, and "Show the recipe (advanced)". When the checks fail: the
  problems in plain words, and "Copy the problems for the AI" (the engine's own repair message),
  up to 2 more tries; after that, the closest attempt with its problems, as KRONOS does.
- Step 4, Save: the real sheet names (as KRONOS asks today), then "Download the macro (.bas)"
  (the browser saves it to Downloads), with the numbered import steps and the starter macro's
  name. The practice run: download its module, and a box to paste Excel's result back for
  vcl_pilot.check, with the verdict.
- History: past runs kept in this browser on this PC (IndexedDB, every read and write in
  try/catch), newest first, with "Download the run record" (like the runs folder KRONOS keeps).
  Say plainly that clearing the browser's data removes them.
- "Suggest for the corpus" (D70): downloads the candidate file for Boss to send to Claude Code.
  Nothing goes into data/ from the page.
- Never: run VBA, open Excel, touch other files, or call the internet. The page loads only from
  localhost.
- Works from 1280 px down to about 900 px wide. Phone width is not needed, but nothing may
  scroll sideways.

#### 5.3 Not in scope

- No change to KRONOS.exe, the Codex adapter or KRONOS.exe's release.
- No password page.
- No real KGC data anywhere: the repo is public.

### 6. PHASES: one at a time, each a pull request with green CI and screenshots for Boss

Phase 1, the proof: the decision entry; docs/kronos-offline.md (first version); build_offline.py;
the four server files; a one-screen page in BlueForce's look that loads Pyodide offline and runs
the validator and the TESTs on a pasted spec, showing PROVEN or the problems; the CI tests (file
list, CRLF, hashes, offline page, identical results). Hand Boss v1, with the link and the steps
from 3.6, so he can check that it starts on his work PC.

Phase 2, the wizard: the four steps with the copy-paste AI step, the review and the .bas
download. Hand over v2.

Phase 3, the rest: history, the practice-run check, the corpus suggestion, Help, and the release
workflow if it isn't done yet. Hand over v3.

After each phase, tell Boss in plain words what he can do now, with the link first.

### 7. IF SOMETHING HERE IS WRONG

This message was written from outside the KRONOS repo, on 14-11-2026. If the repo says otherwise
(a newer decision, a file that moved), the repo wins: tell Boss in one line and carry on. If the
Pyodide route can't work on Chrome/Edge 109 or Firefox 115, stop and tell Boss why before trying
another way.
