# Excel pilot (optional): ForgeLib and the router in Excel 2016

Run this only when you choose; nothing waits for it. It takes about 15 minutes.

**What it proves:**
1. ForgeLib, the VBA runtime every generated macro calls, behaves in Excel 2016 the way our
   Python reference says (226 cases).
2. The router spec (task_12, RouteJobPacks) does in Excel what its TESTs say.

**Why it's safe:**
- It works in a new, empty workbook, with made-up data only.
- It makes its own folders below `%TEMP%`, in a new folder named `KGC_KRONOS_SANDBOX_<date>_<n>`,
  and never touches any other file or folder.
- It never deletes anything.

## What you need

- Excel 2016 with macros allowed (you confirmed they run on your PC). No Trust Center changes.
- A browser with GitHub open on `4ptech1/KGC-KRONOS`, branch `main`.
- Three files from the repo:
  - `dist/forge_module.bas`: ForgeLib plus every generated macro
  - `dist/forge_selftest.bas`: the ForgeLib self-test
  - `dist/forge_pilot.bas`: the router pilot

## Steps

**1. Make a new workbook.**
In Excel, click **File → New → Blank workbook**. Then **File → Save As → Browse**, set
**Save as type** to **Excel Macro-Enabled Workbook (\*.xlsm)**, name it `VBAForge_Pilot` and
click **Save**.
*Expected:* the title bar shows `VBAForge_Pilot.xlsm`.

**2. Open the VBA editor.** Press **Alt+F11**.
*Expected:* the "Microsoft Visual Basic for Applications" window opens, and the left panel shows
`VBAProject (VBAForge_Pilot.xlsm)`.
If nothing opens, your PC blocks the editor: tell Claude and stop here.

**3. Add the three modules, one at a time.** For each file, in this order: `forge_module.bas`,
then `forge_selftest.bas`, then `forge_pilot.bas`:
1. On GitHub, open the file (for example `dist/forge_module.bas`) and click **Copy raw file**
   (the two-squares icon at the top right of the file).
2. In the VBA editor, click `VBAProject (VBAForge_Pilot.xlsm)`, then **Insert → Module**. A
   white code window opens.
3. Click inside the code window, press **Ctrl+A** and **Delete** (this removes an
   `Option Explicit` line Excel may have added), then press **Ctrl+V**.

*Expected:* the left panel shows Module1, Module2 and Module3 under Modules.

**4. Compile.** In the VBA editor, click **Debug → Compile VBAProject**.
*Expected:* nothing happens, no message.
If a message appears, copy its text, note the highlighted line, send both to Claude and stop
here.

**5. Run the self-test.** Go back to Excel (click its window), press **Alt+F8**, click
**ForgeSelfTest**, then **Run**.
*Expected:*
- A message says `ForgeLib self-test: 226 of 226 cases pass.`
- A new sheet, ForgeSelfTest, has one row per case.

Click cell **H4** on that sheet, press **Ctrl+C**, and paste it to Claude. It looks like
`SELFTEST 226/226 FAILED: none`.

**6. Run the router pilot.** Press **Alt+F8**, click **ForgePilotRun**, then **Run**.
*Expected:*
- After a few seconds, a message says `Pilot done: <n> lines on the sheet ForgePilot.`
- A new sheet, ForgePilot, has the lines in column A, from `PILOT 1 spec(s)` down to `END`.

Click the **A** column header on that sheet, press **Ctrl+C**, and paste everything to Claude.

**7. Claude checks the results.** Claude saves your paste to a file and runs:
```
python compiler/vclc.py pilot-check <results file> specs/task_12.vcl
```
It compares every line with the reference interpreter, on the date Excel wrote, and tells you
whether they agree. Any difference is a bug that Claude fixes.

## Afterwards

- Keep or delete the workbook; it holds only made-up data.
- If Excel asks about macros when you reopen it, click **Enable Content**.
- The sandbox folders stay in `%TEMP%`, because VBA Forge never deletes anything. To remove them
  yourself:
  1. Press **Windows key + R**, type `%TEMP%` and press **Enter**.
  2. Delete the folders whose names start with `KGC_KRONOS_SANDBOX_`.
