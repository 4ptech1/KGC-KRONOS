# KGC KRONOS: setup and first use

For Boss, in plain steps. KRONOS is a window where you type what Excel should do in plain words. ChatGPT
Codex writes a recipe, KRONOS tests it on practice data, and you get a macro file to import into Excel.

**You need, in this order:** IT's OK for two programs (step 1), the Codex tool signed in (steps 2 and 3),
KRONOS.exe (step 4). Then you try it (steps 5 to 7).

Practice requests with invented data only, for now. KRONOS never opens, changes or runs your workbooks.

## 1. Ask IT to approve two installs

Send IT this message. Both are needed; nothing else is installed, and nothing about Excel's security changes.

> **To IT: two programs for one user's work PC**
>
> 1. **KGC KRONOS (KRONOS.exe)**, a small Windows app (one file, about 15 to 20 MB) we build ourselves from
>    public source code (github.com/4ptech1/KGC-KRONOS). The version to approve is on the release page, with its
>    SHA-256 checksum. GitHub builds it and runs its self-test before publishing it.
>    - It is **not digitally signed**, so Windows may warn about it.
>    - It unpacks itself into a temporary folder each time it starts.
>    - It writes only to `Documents\KGC KRONOS` (its records and the macro files it saves) and that temporary folder.
>    - It makes **no network connection of its own**. It never opens, changes or runs Excel files, and never
>      deletes or moves files.
> 2. **The OpenAI Codex CLI**, OpenAI's own command-line tool, signed in with the user's ChatGPT account.
>    KRONOS starts it only to turn a typed request into a recipe. This is the only part that uses the internet:
>    the typed request (never a file) goes to OpenAI, over HTTPS. If the PC reaches the internet through a proxy,
>    the Codex CLI needs to be allowed through it. KRONOS was checked with Codex CLI version 0.166.0.
>
> Please tell us where KRONOS.exe may be saved and run from, and whether the Codex CLI may be installed and
> reach OpenAI's servers.

## 2. Install the Codex CLI and sign in

1. Install the Codex CLI the way IT approves, following **OpenAI's current instructions for the Codex CLI on
   Windows** (search for "Codex CLI" on developers.openai.com). OpenAI changes these steps from time to time, and
   I could not read their site while writing this guide, so I have not copied them here.
2. Check it is installed. Click **Start**, type `cmd`, press **Enter** (a black window opens), type
   `codex --version` and press **Enter**.
   *Expected:* a version number. If it says the command is not found, close the black window, open a new one and
   try again; if it still isn't found, the install did not finish.
3. Sign in. In the same black window type `codex login` and press **Enter**. A browser page opens: sign in with
   **your own ChatGPT account** (Plus or Pro).
   *Expected:* the page, or the black window, says you are signed in.
4. Check it. Type `codex login status` and press **Enter**.
   *Expected:* a line saying you are logged in. If it says **Not logged in**, do step 3 again.

## 3. Turn off ChatGPT's "Improve the model for everyone"

1. Open chatgpt.com and sign in.
2. Click your name or picture, then **Settings**, then **Data controls**.
3. Turn **Improve the model for everyone** to **off**.

*Expected:* the switch shows off. (ChatGPT sometimes renames settings, so the words may differ slightly.)

## 4. Download KRONOS.exe

1. Open https://github.com/4ptech1/KGC-KRONOS/releases and click the newest release (named "KGC KRONOS" and a
   version number).
2. Under **Assets**, click **KRONOS.exe**. Save it in the folder IT approved.
3. Optional check that the file is the one GitHub built: in the black window go to that folder and type
   `certutil -hashfile KRONOS.exe SHA256`. *Expected:* a long code that matches the **SHA-256** on the release page.

## 5. Start it for the first time

1. Double-click **KRONOS.exe**. The first start takes a few seconds, because it unpacks itself.
2. If Windows shows **Windows protected your PC**, click **More info**, then **Run anyway**. (If IT blocks it,
   stop and tell IT; see "If something goes wrong".)
3. A window **Before you start** opens. *Expected:* a tick next to "Codex is installed", "Its command line works
   with this version" and "Codex is signed in". (A cross next to the last one: sign in again as in section 2,
   then reopen KRONOS.)
4. Optional: click **Test Codex**. It makes one short call. *Expected:* "Codex answered, so it is installed and
   signed in."
5. Tick **I turned off "Improve the model for everyone"** and **I'll use practice requests with invented data
   only, for now**, then click **Done**.

*Expected:* the main window, "What should Excel do?", with 1 Ask highlighted.

## 6. Your first practice request

Make a practice workbook first, so nothing real is involved: in Excel, make a new workbook, name the first sheet
**Orders**, type the headers **Order No**, **Order Date**, **Customer**, **Status** in row 1, add six made-up rows
(some with Status "Cancelled"), and save it as `practice.xlsx`.

1. In KRONOS, type: *In the Orders sheet, delete every row whose Status says Cancelled, then sort what's left by
   Order Date, newest first.* and click **Ask Codex**.
2. Watch the working panel. It says what it is doing and each try. **Cancel** stops it at any time.
   *Expected:* after a minute or so, the **Review** step.
3. If Codex asks a question, the **Questions** step appears: type your answer and click **Send answer**.
4. On **Review**, read **PROVEN** (green) and then what the macro does, **assumes**, **changes** and **never does**.
   Tick **I checked these assumptions**, then click **Next: Save the macro**.
   If it says **NOT PROVEN** (amber), click **Try again** or **Change my request**.
5. On **Save**, type the real name of your sheet (**Orders**), click **Save macro file...** and click **Save**.
   *Expected:* "Saved:" and the file's location, and the steps for using it in Excel.

## 7. Use the macro in Excel (on a copy)

KRONOS's **Save** screen lists these steps with your macro's exact name. In short, on a **copy** of your workbook,
in Excel 2016:

1. Press **Alt+F11** (the macro editor opens).
2. Click **File > Import File...**, choose the `.bas` file KRONOS saved, click **Open**.
3. Close the editor, press **Alt+F8**, choose the macro named on KRONOS's Save screen, click **Run**.
   *Expected:* a message saying it finished, with its result (for a job like this one, how many rows it removed).
4. To keep the macro, save the workbook as an **Excel Macro-Enabled Workbook (.xlsm)**.

Use one KRONOS macro per workbook (each carries its own copy of the shared helpers, and two would clash).

## Where KRONOS keeps things

Everything is in `Documents\KGC KRONOS`:

| Folder or file | What it holds |
| --- | --- |
| `runs\` | one folder per request: your words, the whole conversation, each recipe, the report |
| `macros\` | the macro files you saved |
| `corpus_candidates\` | requests you offered as training examples ("Suggest for the corpus") |
| `settings.json` | your settings |
| `crash.log` | only if KRONOS ever stops on an error |

**History** (top right) lists past runs, shows their reports, and can bring a past recipe back to Review with
**Use this macro again** (it is tested again first). Nothing here goes to GitHub. To erase your records, close KRONOS and delete the folder.

## If something goes wrong

| You see | What it means | Do this |
| --- | --- | --- |
| Windows won't run KRONOS.exe | IT has blocked it, or it is not allowed to run from that folder | Tell IT; ask where it may run |
| "Codex isn't installed on this PC" | KRONOS can't find the `codex` command | Step 2; then close and reopen KRONOS |
| "Codex isn't signed in" | You need to sign in again | In the black window, `codex login` |
| "Codex can't reach OpenAI over the internet" | The PC is offline, or a firewall or proxy blocks Codex (KRONOS stops it after about 15 seconds instead of waiting) | Check the internet works; on the work PC, tell IT Codex must reach OpenAI over HTTPS |
| "Codex took longer than 5 minutes" | The request was big, or the connection is slow | Try again, or make the request smaller; **Settings** can allow more time |
| "Windows would not start Codex" | A security rule blocks it | Tell IT |
| NOT PROVEN after 3 tries | Codex couldn't write a recipe that passed every test | **Change my request** (name the sheets and headers); nothing is ever offered as a macro unless it is proven |
| Anything else | | Look in `Documents\KGC KRONOS\crash.log`, and tell Claude Code what the window said |

## For the pilot (Phase 4)

Do three practice requests of your own, with invented data. After each, open **History**, select the run, and
copy its report to Claude Code, with a word about anything confusing. Claude Code reads them and fixes what the
real Codex and your real PC show.

## Updating and removing

- **Update:** download the newer KRONOS.exe from the Releases page and replace the old one. Your records stay.
- **Remove:** delete KRONOS.exe, and `Documents\KGC KRONOS` if you want the records gone too.
