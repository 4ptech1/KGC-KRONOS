# ForgeTest pilot: running batch 000 in Excel 2016

This runs the 4 TESTs of `specs/task_01.vcl` (BuildShortageList) in Excel. It takes about
10 minutes. You only paste and run: no Trust Center changes, no References, no file import,
and nothing is saved.

Before you start, save and close any other work in Excel. If a test hangs, you may need to
stop Excel's macro with Ctrl+Break.

## 1. Copy the batch

1. Open `batches/batch_000_pilot.bas` on GitHub (the link is in the pull request).
2. Click the **Copy raw file** button (two overlapping squares, top right of the file).
   If you can't find it: click **Raw**, then press **Ctrl+A** and **Ctrl+C**.

   *Expected:* nothing visible happens; the code is now on the clipboard.

## 2. Open a blank workbook

1. In Excel: **File > New > Blank workbook**.

   *Expected:* a new empty workbook called **Book1** (or similar). Never save it.

## 3. Paste the code into one new module

1. Press **Alt+F11** (on some laptops **Alt+Fn+F11**).
   *Expected:* the Visual Basic editor opens.
2. If you don't see the Project pane on the left, press **Ctrl+R**. Click **VBAProject (Book1)**.
3. Menu **Insert > Module**.
   *Expected:* an empty white code window titled **Book1 - Module1 (Code)**.
4. If the window already shows `Option Explicit`, press **Ctrl+A** and **Delete** so it's empty.
5. Click in the window and press **Ctrl+V**.
   *Expected:* about 1,080 lines; the first line is `Option Explicit`.
6. Menu **Tools > Options**, tab **General**, box **Error Trapping**: select
   **Break on Unhandled Errors** (it usually is already), then **OK**.
7. Menu **Debug > Compile VBAProject**.
   *Expected:* nothing happens. That means it compiled. If a message appears, go to section 7.

## 4. Run it

1. Press **Alt+F11** again to go back to Excel.
2. Press **Alt+F8** (or **Alt+Fn+F8**).
   *Expected:* the Macro window lists **RunBatch** and **ForgeReset**.
3. Click **RunBatch**, then **Run**.

*Expected:* for a few seconds the status bar at the bottom says
`VBA Forge 000_pilot: running B000-T01 (1 of 4)` and so on. Test workbooks (Book2, Book3, ...)
may flash open and close. Then a message box says:

```
VBA Forge batch 000_pilot finished.
4 tests: 4 PASS, 0 FAIL, 0 ERROR.
```

Click **OK**. Book1 now has two new sheets: **Results** and **ForgeLog**.

## 5. What you should see

On the **Results** sheet, one row per test:

| Test ID  | Pass/Fail | Description starts with |
|----------|-----------|-------------------------|
| B000-T01 | PASS | `RESULT=26; ERROR_COUNT=0; COUNTA(Shortages![Part No])=27; COUNTIF(Shortages![Shortage], "<0")=0` |
| B000-T02 | PASS | `RESULT=26; ERROR_COUNT=0; COUNTIF(Shortages![Remark], "Required qty missing")=3` |
| B000-T03 | PASS | `RESULT=26; ERROR_COUNT=0; PRE_SUM(BOM![Required Qty])=986; PRE_SUM(GRN![Received Qty])=661; ...` |
| B000-T04 | PASS | `RESULT=-1; ERROR_COUNT=1; ... last ForgeLog: BuildShortageList: PRE failed: ...` |

On the **ForgeLog** sheet there is one entry, from B000-T04. That one is expected: that test
calls the macro on the wrong sheet on purpose.

The full list of expected numbers is in `batches/batch_000_pilot.expected.csv`.

## 6. Send the results back

1. Click the **Results** sheet tab.
2. Click cell **A1**, press **Ctrl+A**, then **Ctrl+C**.
3. Paste it into the chat with Claude (**Ctrl+V**). It arrives as tab-separated text; that's fine.
4. If the ForgeLog sheet has more than one entry, copy it the same way and paste it too.
5. Close Book1 and click **Don't Save**.

Send the results even if everything passed: they become `data/results/pilot.csv`.

## 7. If a compile error appears

Excel shows a box like **Compile error: Variable not defined** and highlights a line or a word.

1. Write down the message exactly.
2. Note the highlighted line. If the line is selected, **Ctrl+C** copies it.
3. Note the procedure name from the drop-down at the top right of the code window.
4. Click **OK**, then in the editor menu click **Run > Reset**.
5. Send Claude the message, the highlighted line and the procedure name. Don't fix it yourself.

## 8. If Excel freezes or a run-time error box appears

**Excel freezes** (nothing happens for more than a minute):

1. Press **Ctrl+Break**. On laptops without a Break key try **Ctrl+Fn+B**, **Ctrl+Fn+Pause**
   or **Ctrl+ScrLk**.
2. A box says *Code execution has been interrupted*. Click **End**.

**A run-time error box appears** (for example *Run-time error '1004'*):

1. Write down the error number and message.
2. Click **Debug**. The line where it stopped is highlighted in yellow: note it.
3. In the editor menu click **Run > Reset**.

**After either of these:**

1. Go back to Excel (**Alt+F11**), press **Alt+F8**, click **ForgeReset**, then **Run**.
   *Expected:* a message says Excel's settings are back to normal and leftover test workbooks
   are closed. This turns screen updating, automatic calculation, alerts and events back on.
2. Copy the **Results** sheet as in section 6. The test that hung shows **RUNNING**.
3. Send Claude the Results, the error message and the highlighted line.

If nothing runs at all and Excel says macros are disabled, stop and tell Claude: that means the
work PC blocks even pasted macros, and we need another plan.
