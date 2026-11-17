# The new PC's 300 GB: KRONOS's home from March 2027 (D91)

Status: **planned on 17 Nov 2026, from Boss's answers.** The storage arrives with the new PC in March 2027, so this
file says what is prepared before then (in the cloud, as now) and how KRONOS fills the drive from its first day. Each
software piece below is its own PR.

## 1. What Boss decided

Boss's manager gave him 300 GB of storage to be used entirely for KRONOS. Boss answered (17 Nov 2026):

1. **Where:** the disk of the new PC that arrives in March 2027, the PC that will train our own model (D25, D90).
2. **Real data:** yes. It is company storage, so KRONOS may keep real runs, before-and-after copies of real workbooks,
   and a private library of real proven recipes there. They never go into the public repo or a Claude Code session.
3. **Users:** only Boss for now. The team version comes after the pilot, as planned.
4. **"Entirely for KRONOS":** the 300 GB is reserved for KRONOS and fills as KRONOS grows. Each purpose gets a budget,
   and KRONOS writes a monthly report Boss can show his manager. Nothing is filled just to use space.

## 2. What KRONOS needs, measured

| What | Size today | How it was measured |
| --- | --- | --- |
| One KRONOS run (request, conversation, tries, report, the .bas and its practice run) | 84 KB | a wizard run in the cloud |
| KRONOS Offline, one version | 5.8 MB zipped, 20 MB unzipped | the `kronos-offline-v5` release |
| KRONOS.exe, one version | 14 MB | the `kronos-v0.5` release |
| The proven recipe library (`data/`) | under 3 MB | `du` on the repo |
| The whole repo with its history | 24 MB | `du` on the repo (9 MB without its history) |
| The demo workbook (40 rows) | 8 KB | `docs/demo/NCR_Register_Demo.xlsx` |

So KRONOS uses well under 0.1 GB today. Three things will fill the drive:

1. **Training our own model (from March 2027):** datasets, checkpoints and released models. This is the biggest user.
2. **Real workbook records:** a copy of each workbook before and after every KRONOS change. This is the traceability
   record a fabrication quality system values.
3. **Offline copies and backups:** every tool the PC needs (Python, PyTorch, the GPU's libraries), every KRONOS
   version, and weekly safety copies.

## 3. The layout

One folder, `KRONOS`, at the top of the drive. Its letter is known in March, for example `D:\KRONOS`. Names are short,
because Windows limits paths to 260 characters (the same reason as D30).

```
KRONOS\
  README.txt       what each folder is for and its budget (written by the setup)
  apps\            every KRONOS.exe and KRONOS Offline version, with SHA-256 (go back to any version)
  repo\            a clone of the public repo: the toolchain, the simulator, the training code. Public data only
  work\            REAL DATA. Company records; never pushed, never in a Claude Code session
    runs\          every KRONOS run: request, conversation, tries, recipe, .bas, report
    books\         each workbook change: <workbook>\<yyyymmdd-hhmmss>\ holds the original, the changed copy,
                   what changed, and their SHA-256
    recipes\       the private library: real requests, their proven recipes and their VBA
    practice\      results of practice runs in Excel (D81)
  memory\          the AI's memory in VCL (D90 Phase 2): request, recipe and VBA triples, and the
                   similar-recipe index
  sim\             the VBA simulator's results (D90 Phase 1): which VBA passed which recipe's TESTs
  train\           from March 2027
    data\          training datasets: generated, the public library, and the private library (shards)
    runs\          one folder per training run: settings, logs, checkpoints, results
  models\          released model versions: weights, model card, test results, SHA-256. Never pushed
  mirror\          offline copies of every tool, each checked by SHA-256 (Python, Git, PyTorch, GPU libraries,
                   Pyodide, the Python wheels)
  backup\          weekly safety copies of work\runs, work\recipes, memory\ and reports\; repo bundles
  reports\         the monthly storage report (text, and CSV for Excel)
```

## 4. The budget

Each budget is a ceiling, not a target. The report shows the real numbers each month.

| Folder | Budget | Why this much (estimates) |
| --- | --- | --- |
| `apps\` | 3 GB | every version kept: about 50 versions of each app |
| `repo\` | 2 GB | the repo, its build outputs and caches |
| `work\runs\` | 5 GB | 84 KB a run; 10 runs a working day is about 0.2 GB a year |
| `work\books\` | 50 GB | an original and a copy per change; 2 MB workbooks, 10 changes a day: about 10 GB a year, so 5 years |
| `work\recipes\` | 2 GB | about 70 KB a recipe with its VBA; room for 25,000 |
| `practice\`, `sim\` | 5 GB | results and summaries |
| `memory\` | 5 GB | about 10 KB a triple; 100,000 triples and their index |
| `train\data\` | 40 GB | about 8 KB a triple; 1 million triples are about 8 GB raw, about 3 GB prepared; the last 3 versions kept |
| `train\runs\` | 100 GB | for a model of up to about 350 million parameters: about 4.2 GB per full checkpoint (the weights and the optimiser's two copies). Each run keeps its best and its last 3; older runs keep only their final weights |
| `models\` | 10 GB | about 0.7 GB per released model; about 10 releases with their test results |
| `mirror\` | 20 GB | PyTorch with the GPU libraries is several GB per version; two versions, plus Python, Git, Pyodide and wheels |
| `backup\` | 25 GB | 52 weekly copies of the small folders, and repo bundles |
| `reports\` and logs | 3 GB | small; kept for years |
| **Reserve, kept free** | **30 GB** | Windows and the file system need free space; never planned for |
| **Total** | **300 GB** | |

How it is expected to fill (estimates, to be replaced by the monthly report):

| When | Used | Mostly |
| --- | --- | --- |
| Day one (March 2027) | about 12 GB | `mirror\` (PyTorch and the GPU libraries), `apps\`, `repo\` |
| Three months later | about 60 GB | the first datasets and training runs, the first real runs and workbook records |
| End of 2027 | about 150 GB | datasets, checkpoints, released models, a year of workbook records |
| 2029 | about 250 GB | records keep growing; old checkpoints are thinned |

The model's size depends on the PC's graphics card, which is known in March. The budgets above assume a model of up to
about 350 million parameters. A smaller model needs less, and `train\runs\` is resized in March, before training starts.

## 5. The rules

**Privacy (CLAUDE.md, D25, D87):**
- Real data lives only in `work\`, in the private parts of `train\data\` and `memory\`, and in models trained on them.
- Nothing from those places goes to GitHub, into `repo\`, or into a Claude Code session. `repo\` holds only the public
  repo. A check warns if a file from `work\` or a workbook appears there.
- **A model trained on real data stays on this PC** (and in company backup). The public repo gets code and invented
  data only, never weights.
- What Boss copies into an AI chat stays his choice (D87). Once our own model runs on the PC, nothing needs to leave it.

**What may be deleted:**
- **Real records are never deleted by KRONOS:** `work\runs`, `work\books`, `work\recipes`, `work\practice`, released
  `models\` and `reports\`. When one nears its budget, the report says so and Boss decides (for example, moving old
  years to the company archive).
- **What KRONOS can make again** may be thinned by fixed rules, and the report says what was thinned:
  - training checkpoints: the best and the last 3 per run; older runs keep only their final weights;
  - datasets: the last 3 versions;
  - simulator results: summaries after 6 months;
  - weekly backups: the last 52;
  - `mirror\`: the last 2 versions of each tool.

**Space:**
- The report warns when a folder passes 80% of its budget.
- The training scripts refuse to start a run when the drive would fall below its 30 GB reserve, so training can never
  crowd out the real records.

**Integrity:**
- `apps\`, `models\`, `mirror\` and every `work\books\` record carry SHA-256 lists.
- A monthly check confirms that nothing changed by accident.

**Backups:**
- The copies in `backup\` are on the same disk. They protect against mistakes, not against a broken disk.
- Ask IT to back up `KRONOS\work`, `KRONOS\models` and `KRONOS\reports` with the company backup.
- The public code is already safe on GitHub.

## 6. The monthly report (for Boss's manager)

`reports\KRONOS-storage-YYYY-MM.txt`, plus a `.csv` that Excel 2016 opens. It holds:

- the drive's total, used and free space;
- each folder's size against its budget (with a warning past 80%), and its growth since last month;
- counts: KRONOS runs, workbook changes recorded, recipes in the private library, training runs, released models;
- the integrity check's result;
- what was thinned, if anything;
- one plain sentence saying what KRONOS used the space for this month.

## 7. The work, in order

**Before March 2027, in the cloud (one PR each):**

| # | Piece | What it does |
| --- | --- | --- |
| S0 | This plan and D91 | the layout, the budgets and the rules |
| S1 | `compiler/kronos_drive.py` and `vclc drive init / report / check` | `init` makes the folders, `README.txt` and `budgets.json`; `report` writes the monthly report; `check` verifies the SHA-256 lists, the privacy rule and the space rule. Standard library only, tested on a temporary folder in CI, Windows included |
| S2 | KRONOS saves into `work\` | KRONOS.exe's data folder points to `work\`. KRONOS Offline's Phase 3 History saves each run into `work\runs`, and each workbook change into `work\books`, through a narrow save in its local server: localhost only, inside `work\` only, never overwriting (part of KRONOS Offline Phase 3) |
| S3 | The simulator and the memory write to the drive | D90 Phases 1 and 2 write their results to `sim\` and `memory\` when given the drive, and to `build\` in the cloud |
| S4 | The day-one kit | a double-click `KRONOS-Drive-Setup.bat` (no installs): it runs `drive init`, copies the newest apps into `apps\` and writes the first report. Also `docs/kronos-drive-guide.md`: Boss's steps, and a message for IT asking for Python, Git, PyTorch and the GPU driver, and backup of `work\` |

**From March 2027, on the new PC:**

| # | Piece | What it does |
| --- | --- | --- |
| S5 | Day one | IT installs what the message asks for. Boss runs the setup, and KRONOS checks itself and writes its first report. `mirror\` is filled with the tools pinned for this PC's graphics card |
| S6 | Training | datasets into `train\data\`, runs into `train\runs\` under the thinning rule, released models into `models\` with a model card and SHA-256 (D25, D90 Phase 3) |
| S7 | KRONOS uses its own model | KRONOS on this PC asks our model instead of an AI chat, so nothing leaves the PC. Its VBA is used only when the simulator proves it (D90) |
| S8 | The team, later | `work\` grows one folder per person; this needs its own plan after the pilot |

## 8. What is needed from Boss in March

1. The drive's letter, and whether the 300 GB is a drive of its own or shares the Windows disk.
2. The PC's graphics card and memory, to size the model and `train\runs\`.
3. IT's answer on installing Python 3.11, Git, PyTorch and the GPU driver, and on running KRONOS.exe there.
4. IT's backup for `KRONOS\work`, `KRONOS\models` and `KRONOS\reports`.

## 9. What does not change

- The public repo holds invented data only. Real data never goes into it, its tests or a Claude Code session.
- No training before the new PC arrives (CLAUDE.md). D25: the model is trained from zero, on in-house data only.
- D90: the AI's VBA is used only when the simulator proves it against the recipe's TESTs; the compiler's VBA is the
  fallback.
- Running anything in Excel stays optional and never blocks work.
