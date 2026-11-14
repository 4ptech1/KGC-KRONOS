# KRONOS Offline

KRONOS in the browser, with nothing installed (D86). Boss unzips a folder on his locked-down work PC, double-clicks
**Start KRONOS (Windows)**, and KRONOS opens in Chrome, Edge or Firefox. A tiny local server (Node.js, Python, or
Windows' own PowerShell) shows the page, and the page runs this repo's toolchain through **Pyodide** (Python compiled
to WebAssembly), bundled in the zip. No exe, no Python, no installs, so no IT approval. KRONOS.exe stays as it is,
for when IT approves it.

It copies the offline copy of BlueForce, another project of Boss's, which already runs on that PC: its launcher,
servers, rules and look. This file is KRONOS's own copy of that recipe, so any session can rebuild it.

**Standing rule (CLAUDE.md):** when Boss wants KRONOS offline, in whatever words, build the offline copy as this file
says and give its download link first.

## 1. Boss's decisions (14-11-2026, D86)

1. It runs in the browser, with no install, like BlueForce's offline copy.
2. The AI step is copy-paste with an AI chat ("Copy for the AI", then paste the reply back), with "Copy the problems
   for the AI" for up to 2 more tries, as the engine does now. (Phase 2.)
3. No password page: KRONOS's code is public and holds no secrets.
4. It looks like BlueForce (section 8).

He also answered: the work PC's browser is **Chrome** (on old Windows that is version 109, its last), and each copy
is handed over as **a GitHub Release link plus the zip as a file in the chat**.

## 2. How it works

```
Start KRONOS (Windows).bat  ->  server/serve.mjs | serve.py | serve.ps1  ->  http://localhost:5181/
  app/index.html + app.js      the page (plain HTML, CSS and JavaScript, no framework, no build step)
  app/worker.js                a classic Web Worker: loads Pyodide from app/pyodide/, unpacks the lark wheel from
                               app/wheels/, writes the toolchain files from app/toolchain/manifest.json into
                               Pyodide's file system under /kronos, and calls kronos_web.check(text)
  compiler/kronos_web.py       the bridge: reads the pasted reply (kronos_engine.read_reply), proves it
                               (kronos_engine.prove), returns plain data as canonical JSON
```

- **One source of truth.** The browser runs the very same `compiler/*.py`, `vcl/semantics.py` and the data files of
  `packaging/kronos.spec`, byte for byte (a test checks it). Never fork or edit the toolchain for the browser: if a
  module needs a change to run there, make it in the shared module, so KRONOS.exe and the browser keep passing the
  same tests. Phase 1 needed no change at all.
- **What can't run in the browser:** tkinter (the window), and subprocess and threading for Codex. Importing
  `kronos_ai` is fine there; only using Codex would fail. The browser's "AI" is Boss's copy and paste (Phase 2),
  using the engine's own message builders and retry logic.
- **Identical results.** `kronos_web.check` returns one canonical JSON text: under Pyodide in Node.js it must equal
  normal Python's, as a string, for the 12 specs, the examples and a few broken replies (validator messages, TEST
  results and the .bas, byte for byte).
- **Nothing loads from the internet.** Pyodide, the wheel, the fonts and the toolchain all come from the local
  server; the page test blocks every request outside localhost and fails on any.

The files in the repo:

| Path | What |
| --- | --- |
| `web/` | the page: `index.html`, `styles.css`, `app.js`, `worker.js`, `logo.svg` |
| `compiler/kronos_web.py` | the bridge the worker and the tests both call |
| `packaging/offline/` | the launcher (committed with CRLF; `.gitattributes` keeps it so), `server/serve.{mjs,py,ps1}`, `licences/` |
| `packaging/build_offline.py` | builds `KRONOS-Offline/` and the zip, standard library only |
| `packaging/offline_pins.json` | every downloaded file, with its URL and SHA-256 |
| `packaging/offline_test/` | Playwright for the page test, pinned with hashes (today's Chromium, Chromium 109, Firefox 115); CI only |
| `tests/test_offline.py`, `tests/offline_pyodide.js` | the tests (section 6) |
| `.github/workflows/build-offline.yml` | CI: build, old browsers, PowerShell on Windows, release |

## 3. The work PC, as BlueForce met it

- Windows, an older version (Explorer looks like Windows 7 or Server 2008 R2, perhaps over Remote Desktop).
- The folder sits on a mapped network drive (a letter such as Y:). **Never write the real folder path into the repo:
  it is public.**
- No Node.js or Python, so the launcher falls back to PowerShell, probably version 2.
- No default web browser: `Start-Process "http://..."` fails with "No application is associated with the specified
  file".
- A browser that runs the app: Chrome or Edge 109+, or Firefox 115+, the last versions for Windows 7 and 8.1.
  Internet Explorer can't.

## 4. The rules (keep every one)

The errors BlueForce met, and the rule for each:

1. **HTTP 500 on every page (PowerShell):** `serve.ps1` found its folder through `$PSScriptRoot`, which was empty on
   that PC. Rule: the .bat passes the app folder in,
   `powershell -NoProfile -ExecutionPolicy Bypass -File "server\serve.ps1" "%CD%\app"`, and `serve.ps1` takes it as
   `param([string]$AppFolder)`. Never rely on `$PSScriptRoot` alone.
2. **A red "Start-Process: No application is associated" error:** the PC has no default browser (the app was running
   anyway). Rule: never a bare `Start-Process "http://..."`. `Open-Browser` uses the default browser only when one is
   set and isn't Internet Explorer; otherwise it looks for Chrome, Edge or Firefox by name in `%LOCALAPPDATA%`,
   Program Files (x86) and Program Files; if none is found it prints one plain line telling the person to open one
   and type the address. Nothing in a server script may end in a red PowerShell error.
3. **A text file shown as one jumbled line:** Notepad before Windows 10 (1809) ignores Unix line endings. Rule: the
   launcher, the servers and the licence texts keep Windows line endings (CRLF); check with `file`: "with CRLF line
   terminators". The app's own files (page, Pyodide, toolchain) stay byte for byte as in the repo and in Pyodide's
   package: the browser reads them, not Notepad, and the toolchain must be the repo's exact files.
4. **A blank page in Internet Explorer.** Rule: the message to Boss says to use Chrome, Edge or Firefox. (The page
   itself says so too, in one line, in a browser without WebAssembly such as Internet Explorer; `index.html` hides
   that line at once everywhere else, so it never flashes while the page loads.)

Also keep:

- The servers listen on localhost / 127.0.0.1 only, never on the network. "Port ... is busy" means the app is already
  running.
- No Node.js advice for old Windows: current Node.js needs Windows 10 or later. On older Windows the PowerShell server
  is the way.
- Boss never runs commands in a terminal: everything is a double-click.
- **KRONOS's own finding (D86): the worker loads `pyodide.asm.js` itself, first.** BlueForce's `serve.py` and
  `serve.ps1` answer one request at a time. Pyodide starts downloading its two big files (the Python library and the
  `.wasm`, 9 MB each) and then loads `pyodide.asm.js` synchronously, so the worker stops reading while the server is
  still sending a big file, and the server never gets to `pyodide.asm.js`: the page never got past "starting Python"
  (seen with `serve.py` here and with `serve.ps1` on Windows in CI). `web/worker.js` now does
  `importScripts('pyodide/pyodide.js', 'pyodide/pyodide.asm.js')`, and Pyodide skips its own load. BlueForce's four
  files stay exactly as they are apart from the name, the port and the version line; because `serve.py` answers one
  request at a time like `serve.ps1`, every page test through it also guards this.
- **A one-request-at-a-time server waits for whoever it is sending to.** So a page that stops reading holds up
  everything else: before the fix above, the launcher step in CI waited four minutes for a 404 behind the Edge window
  that `Open-Browser` had opened; after it, the same requests are answered in under a second while Edge loads the
  page. Never add a synchronous load (`importScripts`, a synchronous XMLHttpRequest) while a download is under way.

## 5. The zip

```
KRONOS-Offline-v<N>-<DD-MM-YYYY>.zip          the version and the build date, India time; Windows only
  KRONOS-Offline/                             always this name, so Boss's folder keeps its name
    Start KRONOS (Windows).bat                node -> py -3 -> python -> PowerShell, in that order (CRLF)
    server/serve.mjs                          Node.js: http://localhost:5181/, 127.0.0.1 only
    server/serve.py                           Python 3: the same
    server/serve.ps1                          PowerShell only: the same (rules 1 and 2)
    app/                                      the page, its scripts and fonts, Pyodide, the lark wheel, the
                                              toolchain with its manifest, and licences/
```

- **Port 5181**, not BlueForce's 5180: Boss may run both at once, and a different port keeps each app's browser
  storage apart.
- No guide inside the zip and no Mac or Linux launcher: what Boss does is said in the chat (section 7). The Node.js
  and Python servers stay, because the .bat falls back to them.
- **Versions:** N starts at 1 and goes up by one with every rebuild that is handed over. The version and date show in
  the page's top bar ("Version 1 · 14-11-2026") and in the server window's start message ("KRONOS Offline, version
  1, built 14-11-2026"). The version list at the end of this file has one row per copy; `build_offline.py` builds the
  version of its last row.
- **Third-party files are never committed.** `build_offline.py` downloads each once into `build/offline-cache/`
  (git-ignored) and refuses any whose SHA-256 differs from `packaging/offline_pins.json`. Their licence texts are
  committed in `packaging/offline/licences/` and go into `app/licences/`.
- The zip is written with sorted names and fixed times (the build day at noon), so the same inputs give the same
  bytes; beside it go `<zip>.sha256.txt` and `build-info.json`.

### The parts and their versions

| Part | Version | From | Licence |
| --- | --- | --- | --- |
| Pyodide (Python 3.11.3) | 0.24.1 | the npm package `pyodide`: `pyodide.js`, `pyodide.asm.js`, `pyodide.asm.wasm`, `python_stdlib.zip`, `pyodide-lock.json` | MPL-2.0 (Pyodide's repository at tag 0.24.1; its npm `package.json` says Apache-2.0); Python itself PSF |
| lark | 1.3.1 | PyPI, the same wheel and hash as `requirements.txt` | MIT |
| Inter | 5.3.0 | the npm package `@fontsource/inter`: Latin 400, 600 and 700, woff2 | OFL-1.1 |

**Why Pyodide 0.24.1.** Pyodide's own docs (`docs/usage/index.md` at each release tag) give these minimum browsers:

| Pyodide | Python | Chrome | Edge | Firefox |
| --- | --- | --- | --- | --- |
| 0.21.3 to 0.24.1 | 3.10.2, then 3.11.2 and 3.11.3 | 71 | 80 | 70 |
| 0.25.1, 0.26.4, 0.27.8, 0.28.3, 0.29.5 | 3.11.3 to 3.13.2 | 112 | not listed | 112 |

Every release from 0.25 on documents a minimum above Chrome/Edge 109, so 0.24.1 is the last that covers Chrome 109,
and it covers Firefox 115. It runs Python 3.11.3, the Python this repo's CI uses. Don't move to a newer Pyodide
while the work PC may be on Windows 7 or 8.1.

**Notes that save time.** In Node.js, pass `new Uint8Array(buffer)` to `pyodide.unpackArchive` and `FS.writeFile`
(a Node `Buffer` fails with "Unknown typed array type 'Buffer'"); `require('app/pyodide/pyodide.js')` works there, so
the Node test runs the zip's own files. In the page's classic worker, `importScripts('pyodide/pyodide.js')`. Every
server must send `.wasm` as `application/wasm`; the other files may be `application/octet-stream`.

## 6. Building, checking and releasing

**Build** (from the repo root, with Python 3.11 and the network for the first build):

```
python packaging/build_offline.py
```

It writes `build/offline/KRONOS-Offline/`, the zip, its `.sha256.txt` and `build-info.json`.

**A new copy for Boss:** make the change, add a row for version N+1 to the version list below (with "(added after
the release)" in its commit column and what changed in a few words), and open the PR. When CI is green, merge, then
release: GitHub, Actions, **Build KRONOS Offline**, Run workflow, branch `main`, tick "release". It publishes
`kronos-offline-v<N>` with the zip and its SHA-256, as `build-exe.yml` does for KRONOS.exe (D79). GitHub's "Latest"
badge stays on the newest KGC KRONOS (KRONOS.exe) release (`--latest=false`), so give Boss this release's own link:
`https://github.com/4ptech1/KGC-KRONOS/releases/tag/kronos-offline-v<N>`. Then put the release's commit into the
version list in a docs-only PR.

**The checks** (`tests/test_offline.py`; in CI they run with `KRONOS_OFFLINE_REQUIRED=1`, so they fail instead of
skipping):

| Check | Where |
| --- | --- |
| The launcher is CRLF, tries node, py -3, python, PowerShell in that order, and passes `"%CD%\app"` | always |
| `serve.ps1` keeps `param([string]$AppFolder)`, `Open-Browser` and no bare `Start-Process "http`; all three servers use 5181 on localhost only and send `.wasm` as `application/wasm`; `serve.py` answers one request at a time, like `serve.ps1` | always |
| The pins are well formed, lark's equals `requirements.txt`, every licence is committed; the test tools are pinned with hashes | always |
| The manifest holds every module `kronos_web` imports, every `compiler/*.py` and every `kronos.spec` data file | always |
| The page uses only what Chrome 109 and Firefox 115 have and names no outside address; the worker loads `pyodide.asm.js` first; the look is BlueForce's values | always |
| `kronos_web` answers a proven recipe, a broken one, a failing TEST, a question, a reply with no recipe and a whole AI reply | always |
| The zip's name and file list, nothing stray; CRLF and the version line where they belong; the toolchain byte for byte; the same inputs give the same bytes | needs the downloads |
| Under Pyodide in Node.js, `kronos_web.check` equals normal Python's text, with the same steps | needs the downloads and Node.js |
| From a freshly unpacked zip, `serve.mjs` and `serve.py` (and on Windows `serve.ps1`, started as the .bat starts it but opening no browser) serve every file, `.wasm` as `application/wasm`, and nothing outside `app/` (404; Windows itself answers 403 to a path that climbs out) | needs the downloads |
| With Playwright and every request outside localhost blocked: the version line, Cancel while Python starts, Python ready, the example PROVEN with exactly normal Python's answer, a broken recipe NOT PROVEN with its problems, an empty box, no page errors | needs Playwright; CI: today's Chromium, Chromium 109 and Firefox 115 on Ubuntu 22.04, and Chromium on Windows through all three servers |
| The launcher itself on Windows with no Node.js or Python on the PATH: it falls back to PowerShell, which serves the page and the `.wasm` file and prints the version | CI, `powershell` job |

What CI can't check: Windows PowerShell 2 (GitHub's Windows machines have 5.1, and version 2 can't be run there), and
Windows 7 itself. Boss's own PC is that check. If an old browser can't be installed in CI, the job says so as a
warning and that browser was not checked.

## 7. What to tell Boss at each handover

Start with the download link and "This is version N, built DD-MM-YYYY.", then:

1. Close the black window if it's open.
2. Delete the old KRONOS-Offline folder.
3. Right-click the zip, choose Extract All..., then Extract.
4. Double-click Start KRONOS (Windows).
5. If the browser doesn't open by itself, open Chrome, Edge or Firefox and type localhost:5181.

Mention anything new in one line, say that Internet Explorer can't show KRONOS, say what he can do now, and send the
zip and screenshots in the chat.

## 8. The look (BlueForce's values)

BlueForce is built with React and Ant Design; KRONOS Offline recreates the look with CSS variables (`web/styles.css`).

- **Font:** Inter, bundled in `app/fonts` (no Google Fonts), falling back to "Segoe UI", system-ui, sans-serif. Text
  black #000000; secondary text #595959.
- **Top bar:** 96 px high, navy #0B2A6F, white title, small lines in #D6DEEF; it stays at the top while the page
  scrolls. **Logo:** a navy #0B2A6F circle with a bold white "K" (`web/logo.svg`).
- **Left sidebar** (if a layout has one): #F7F7F7, black text at 15 px, items 40 px high with radius 8, hover #EFE6D4,
  the current item underlined.
- **Page:** white #FFFFFF. **Cards and panels:** light grey #F5F5F5, no border, no line under the title, radius 8.
- **Buttons, all of them:** light orange #FFD8A8, bold (700) black text, no border, radius 8; on hover or press navy
  #0B2A6F with white text. Main and secondary buttons look the same.
- **Boxes** (text boxes, text areas, dropdowns, date fields): sand-beige fill #EFE6D4, no border; hover #E5D9C2;
  pressed #DBCDB2; while typing, white with a mid-grey edge #595959. An open dropdown list is white, its chosen row
  #EFE6D4, its hovered row #F7F2E8.
- **Tables:** 12 px text; 12 px semibold headers on a light band #F7F2E8; a thin black line under every row and under
  the header, none between columns; cells 8 px top and bottom and 4 px each side; every other row #EFE6D4; hover
  #E5D9C2.
- **Status colours, for status only:** green #1E8E5A (the PROVEN banner #00870E with white text), amber #B26A00, red
  #C8372D. KRONOS's window colours map onto these: accent to navy, ok to green, warn to amber, note to red, bg to
  white, panel to #F5F5F5, line to #000000 for table lines and #BFBFBF for soft edges.
- **Pop-ups with fields:** centred, the label on the left and its box on the right, 12 px between fields, one main
  button, no Cancel (close with the X). The one exception is the "working" panel, which keeps its Cancel button.
- **Corners:** radius 8 everywhere. **Keyboard focus:** a visible navy outline.
- **Words:** plain and short, as the KRONOS window writes them. Keep its EXAMPLE, TIPS and NEVER texts
  (`app/kronos_app.py`).
- Works from 1280 px down to about 900 px wide; phone width is not needed, but nothing may scroll sideways.

## 9. The phases

Each phase is one pull request with green CI and screenshots for Boss, and hands over a new version.

**Phase 1, the proof (done, version 1, D86):** the decision; this file; `build_offline.py`; the four server files; a
one-screen page in BlueForce's look that loads Pyodide offline and checks, tests and builds a pasted recipe (or a
whole AI reply), showing PROVEN or the problems, with the practice example and Cancel; the CI checks; the release
workflow.

**Phase 2, the wizard** (Boss's brief, 14-11-2026): KRONOS.exe's four-step wizard, in BlueForce's look.

- Top bar: the logo, "KGC KRONOS", the steps (1 Ask, 2 Questions, 3 Review, 4 Save) with the current one marked,
  History and Help on the right, and the version line.
- Step 1, Ask: the request box, with KRONOS's EXAMPLE and TIPS. Then "Copy request for the Orchestrator" (Boss's
  claude.ai Orchestrator project has KRONOS's instructions as its project instructions, so this copies the request
  and the conversation so far only; if `vcl/orchestrator_prompt.txt` is newer than what Boss last put in that
  project, tell him in one line, with the exact clicks to update it) and "Copy request with instructions (ChatGPT or
  a new chat)" (the Orchestrator prompt and the request as one message, exactly what KRONOS sends Codex today). One
  short line of steps: paste it into the AI chat, copy the AI's whole reply, paste it below. A large "Paste the AI's
  reply" box, and "Check the reply". The message with instructions is about 48 KB: call
  `navigator.clipboard.writeText` inside the click handler, and fall back to selecting the text in a box for Ctrl+C.
- Step 2, Questions: when the reply is a question (the engine tells them apart), show it. Boss's answer goes into the
  next "Copy..." (the conversation continues), and he pastes the next reply.
- While the toolchain works: a panel covers the step with what is happening and a Cancel button (Cancel stops the
  worker and starts a fresh one). The page never freezes.
- Step 3, Review: the PROVEN (green) or NOT PROVEN (amber) banner, and the same parts as the KRONOS window: what the
  macro does, its assumptions with the tick, what it changes, what it never does, how it was tested, and "Show the
  recipe (advanced)". When the checks fail: the problems in plain words, and "Copy the problems for the AI" (the
  engine's own repair message), up to 2 more tries; after that, the closest attempt with its problems, as KRONOS
  does.
- Step 4, Save: the real sheet names (as KRONOS asks today), then "Download the macro (.bas)" (the browser saves it to
  Downloads), with the numbered import steps and the starter macro's name.
- Its tests: a practice request through all four steps with a scripted "AI reply" (a proven spec pasted in), and the
  downloaded .bas byte for byte what CPython's engine writes for the same reply.

**Phase 3, the rest:** History (past runs kept in this browser on this PC in IndexedDB, every read and write in
try/catch, newest first, with "Download the run record", like the runs folder KRONOS keeps; say plainly that clearing
the browser's data removes them); the practice run (download its module, and a box to paste Excel's result back for
`vcl_pilot.check`, with the verdict); "Suggest for the corpus" (D70: downloads the candidate file for Boss to send to
Claude Code; nothing goes into `data/` from the page); and Help.

**Never**, in any phase: run VBA, open Excel, touch other files, or call the internet; the page loads only from
localhost. No change to KRONOS.exe, the Codex adapter or KRONOS.exe's release; no password page; no real KGC data
anywhere.

Boss's whole brief, word for word, with BlueForce's original four files, is in the git history:
`docs/kronos-offline-plan.md` at commit 3657aff (removed when Phase 1 landed).

## Versions

| Version | Date | Built from (main) | What changed |
| --- | --- | --- | --- |
| 1 | 14-11-2026 | (added after the release) | The proof: one screen that checks, tests and builds a pasted recipe in the browser, offline, with the practice example and Cancel. |
