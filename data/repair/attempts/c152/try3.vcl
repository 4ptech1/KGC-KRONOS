VCL 0.6
BUILD
# ASSUME: Calibration Tracker headers (row 1) include "Instrument ID", "Type", "Next Calibration Due" and "Status"
# ASSUME: ITP Tracker headers include "ITP No", "Instrument ID", "Inspection Status" and "Cal Flag"
# ASSUME: Overdue means the due date is before today; Due soon means today up to 14 days ahead; a due date that is not a real date gives No date
# ASSUME: Instrument IDs are matched ignoring case and stray spaces; if an ID appears twice in the Calibration Tracker its first row counts
# ASSUME: Closed points (Inspection Status Closed, any case) are never flagged and get their Cal Flag cleared; a non-Closed point with a blank Instrument ID gets Unknown instrument
# ASSUME: rows with a blank Instrument ID in the Calibration Tracker or a blank ITP No in the ITP Tracker are left alone
# ASSUME: the real sheets are "Calibration Tracker", "ITP Tracker" and "Instruments To Calibrate" (it exists and is rebuilt on every run); the TESTs use Cal, Itp and Due
GOAL     Set each instrument's calibration status, flag ITP points that use overdue or unknown instruments, list the instruments to calibrate and report how many points are out of cal.
NAME     InstrumentControl
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modSteps
SIG      modMain.Run()
SIG      modMain.RunCalStatus()
SIG      modMain.RunItpFlags()
SIG      modMain.RunDueList()
SIG      modMain.RunCalHeadless(cal As Worksheet) As Long
SIG      modMain.RunItpHeadless(itp As Worksheet, cal As Worksheet) As Long
SIG      modMain.RunDueHeadless(cal As Worksheet, out As Worksheet) As Long
SIG      modMain.RunHeadless(cal As Worksheet, itp As Worksheet, out As Worksheet) As Long
SIG      modSteps.SetCalStatus(cal As Worksheet) As Long
SIG      modSteps.FlagItpPoints(itp As Worksheet, cal As Worksheet) As Long
SIG      modSteps.ListDueInstruments(cal As Worksheet, out As Worksheet) As Long
FLOW     modMain.Run -> modMain.RunHeadless -> modSteps.SetCalStatus
FLOW     modMain.RunHeadless -> modSteps.FlagItpPoints
FLOW     modMain.RunHeadless -> modSteps.ListDueInstruments
FLOW     modMain.RunCalStatus -> modMain.RunCalHeadless -> modSteps.SetCalStatus
FLOW     modMain.RunItpFlags -> modMain.RunItpHeadless -> modSteps.FlagItpPoints
FLOW     modMain.RunDueList -> modMain.RunDueHeadless -> modSteps.ListDueInstruments
FIXTURE  cal_01.xlsx sheet=Cal cols=A:Instrument ID,B:Type,C:Next Calibration Due,D:Status rows=11 data=A:values("I-1"x1," I-2 "x1,"I-3"x1,"I-4"x1,"I-5"x1,"I-6"x1,"I-7"x1,"I-8"x1,"I-9"x1,""x1,"i-1"x1)+boundary("I-1"," I-2 ","I-3","I-4","I-5","I-6","I-7","I-8","I-9","","i-1");B:text(6);C:reldates(-30,40)+boundary(-1,0,14,15,"",-30,5,40,-5,3,20);D:values("old"x11)
FIXTURE  cal_01.xlsx sheet=Itp cols=A:ITP No,B:Point,C:Instrument ID,D:Inspection Status,E:Cal Flag rows=10 data=A:seq(1,1)+boundary(1,2,3,"",5,6,7,8,9,10);B:text(5);C:values("i-1"x1," I-6 "x1,"I-9"x1,"X-1"x2,"I-2"x1,"I-4"x1,""x1," x-2 "x1,"I-1"x1)+boundary("i-1"," I-6 ","I-9","X-1","I-2","I-4","X-1",""," x-2 ","I-1");D:values("Open"x6,"Hold"x1,"Closed"x1," closed "x1,"CLOSED"x1)+boundary("Open","Hold","Closed","Open","Open","Open","Open","Open"," closed ","CLOSED");E:values("old"x10)
FIXTURE  cal_01.xlsx sheet=Due cols=A:Instrument ID,B:Type,C:Next Calibration Due,D:Status rows=0 stale=3
TEST     fixture=cal_01.xlsx call=modMain.RunHeadless(Cal, Itp, Due) runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND COUNTIF(Cal![Status], "Overdue") = 3 AND COUNTIF(Cal![Status], "Due soon") = 3 AND COUNTIF(Cal![Status], "OK") = 3 AND COUNTIF(Cal![Status], "No date") = 1 AND COUNTIF(Cal![Status], "old") = 1 AND COUNTIF(Itp![Cal Flag], "Out of cal") = 2 AND COUNTIF(Itp![Cal Flag], "Unknown instrument") = 2 AND COUNTIF(Itp![Cal Flag], "old") = 1 AND COUNTA(Itp![Cal Flag]) = 6 AND ROWS_USED(Due) = 7 AND COUNTIF(Due![Status], "stale") = 0
TEST     fixture=cal_01.xlsx call=modMain.RunHeadless(Cal, Itp, Due) assert=COUNTA(Cal![Instrument ID]) = PRE_COUNTA(Cal![Instrument ID]) AND COUNTA(Itp![ITP No]) = PRE_COUNTA(Itp![ITP No]) AND COUNTIF(Due![Status], "Overdue") = 3 AND COUNTIF(Due![Status], "Due soon") = 3 AND VALUE(Due![Instrument ID]2) = "I-6" AND VALUE(Due![Status]2) = "Overdue" AND VALUE(Due![Status]7) = "Due soon" AND VALUE(Due![Next Calibration Due]2) < VALUE(Due![Next Calibration Due]3) AND VALUE(Due![Next Calibration Due]3) < VALUE(Due![Next Calibration Due]4) AND VALUE(Due![Next Calibration Due]6) < VALUE(Due![Next Calibration Due]7) AND VALUE(Itp![Cal Flag]2) = "Out of cal" AND VALUE(Itp![Cal Flag]8) = "Unknown instrument"
TEST     fixture=cal_01.xlsx call=modMain.RunHeadless(Itp, Cal, Due) assert=RESULT = -1 AND ERROR_COUNT = 1
---
VCL 0.6
TASK
GOAL     Start the instrument control tool on the workbook's own sheets and tell the user how many ITP points are out of cal.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "Calibration Tracker"), SHEET(THIS_BOOK, "ITP Tracker"), SHEET(THIS_BOOK, "Instruments To Calibrate"), into=n)
OP       MSG("ITP points Out of cal: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run only step 1 on the Calibration Tracker and tell the user how many instruments are overdue.
PROC     RunCalStatus
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunCalHeadless, SHEET(THIS_BOOK, "Calibration Tracker"), into=n)
OP       MSG("Instruments Overdue: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run only step 2 on the ITP Tracker and tell the user how many ITP points are out of cal.
PROC     RunItpFlags
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunItpHeadless, SHEET(THIS_BOOK, "ITP Tracker"), SHEET(THIS_BOOK, "Calibration Tracker"), into=n)
OP       MSG("ITP points Out of cal: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run only step 3 and tell the user how many instruments were listed on the Instruments To Calibrate sheet.
PROC     RunDueList
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunDueHeadless, SHEET(THIS_BOOK, "Calibration Tracker"), SHEET(THIS_BOOK, "Instruments To Calibrate"), into=n)
OP       MSG("Instruments listed: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run all three steps on the sheets it is given and return how many ITP points are out of cal.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     cal As Worksheet, itp As Worksheet, out As Worksheet
RET      Long
VAR      n As Long
VAR      listed As Long
VAR      overdue As Long
PRE      COL($cal, "Instrument ID") > 0; COL($cal, "Type") > 0; COL($cal, "Next Calibration Due") > 0; COL($cal, "Status") > 0; COL($itp, "ITP No") > 0; COL($itp, "Instrument ID") > 0; COL($itp, "Inspection Status") > 0; COL($itp, "Cal Flag") > 0
MODIFIES $cal["Status"]
MODIFIES $itp["Cal Flag"]
MODIFIES $out
OP       CALL(modSteps.SetCalStatus, $cal, into=overdue)
OP       CALL(modSteps.FlagItpPoints, $itp, $cal, into=n)
OP       CALL(modSteps.ListDueInstruments, $cal, $out, into=listed)
OP       RETURN(n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run step 1 on the Calibration Tracker sheet it is given and return how many instruments are overdue.
PROC     RunCalHeadless
KIND     Function
MODULE   modMain
ARGS     cal As Worksheet
RET      Long
VAR      n As Long
PRE      COL($cal, "Instrument ID") > 0; COL($cal, "Next Calibration Due") > 0; COL($cal, "Status") > 0
MODIFIES $cal["Status"]
OP       CALL(modSteps.SetCalStatus, $cal, into=n)
OP       RETURN(n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run step 2 on the ITP Tracker and Calibration Tracker sheets it is given and return how many ITP points are out of cal.
PROC     RunItpHeadless
KIND     Function
MODULE   modMain
ARGS     itp As Worksheet, cal As Worksheet
RET      Long
VAR      n As Long
PRE      COL($cal, "Instrument ID") > 0; COL($cal, "Next Calibration Due") > 0; COL($itp, "ITP No") > 0; COL($itp, "Instrument ID") > 0; COL($itp, "Inspection Status") > 0; COL($itp, "Cal Flag") > 0
MODIFIES $itp["Cal Flag"]
OP       CALL(modSteps.FlagItpPoints, $itp, $cal, into=n)
OP       RETURN(n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run step 3 on the Calibration Tracker and the output sheet it is given and return how many instruments were listed.
PROC     RunDueHeadless
KIND     Function
MODULE   modMain
ARGS     cal As Worksheet, out As Worksheet
RET      Long
VAR      n As Long
PRE      COL($cal, "Instrument ID") > 0; COL($cal, "Type") > 0; COL($cal, "Next Calibration Due") > 0
MODIFIES $out
OP       CALL(modSteps.ListDueInstruments, $cal, $out, into=n)
OP       RETURN(n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write Overdue, Due soon, OK or No date in the Status of every instrument in the Calibration Tracker and return how many are overdue.
PROC     SetCalStatus
KIND     Function
MODULE   modSteps
ARGS     cal As Worksheet
RET      Long
VAR      cId As Long
VAR      cDue As Long
VAR      cStatus As Long
VAR      n As Long
PRE      COL($cal, "Instrument ID") > 0; COL($cal, "Next Calibration Due") > 0; COL($cal, "Status") > 0
MODIFIES $cal["Status"]
OP       LET(cId, COL($cal, "Instrument ID"))
OP       LET(cDue, COL($cal, "Next Calibration Due"))
OP       LET(cStatus, COL($cal, "Status"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($cal, cId) var=r
OP         SKIP_IF IS_EMPTY(CELL($cal, r, cId))
OP         IF IS_DATE(CELL($cal, r, cDue))
OP           IF CELL($cal, r, cDue) < TODAY
OP             SET_VALUE(CELL($cal, r, cStatus), "Overdue")
OP             LET(n, n + 1)
OP           ELSE
OP             IF CELL($cal, r, cDue) <= TODAY + 14
OP               SET_VALUE(CELL($cal, r, cStatus), "Due soon")
OP             ELSE
OP               SET_VALUE(CELL($cal, r, cStatus), "OK")
OP         ELSE
OP           SET_VALUE(CELL($cal, r, cStatus), "No date")
OP       RETURN(n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write Out of cal or Unknown instrument in Cal Flag for ITP points that are not Closed and clear it on the others, and return how many are Out of cal.
PROC     FlagItpPoints
KIND     Function
MODULE   modSteps
ARGS     itp As Worksheet, cal As Worksheet
RET      Long
VAR      cCalId As Long
VAR      cDue As Long
VAR      cNo As Long
VAR      cInst As Long
VAR      cStat As Long
VAR      cFlag As Long
VAR      known As Dictionary
VAR      got As String
VAR      n As Long
PRE      COL($cal, "Instrument ID") > 0; COL($cal, "Next Calibration Due") > 0; COL($itp, "ITP No") > 0; COL($itp, "Instrument ID") > 0; COL($itp, "Inspection Status") > 0; COL($itp, "Cal Flag") > 0
MODIFIES $itp["Cal Flag"]
OP       LET(cCalId, COL($cal, "Instrument ID"))
OP       LET(cDue, COL($cal, "Next Calibration Due"))
OP       LET(cNo, COL($itp, "ITP No"))
OP       LET(cInst, COL($itp, "Instrument ID"))
OP       LET(cStat, COL($itp, "Inspection Status"))
OP       LET(cFlag, COL($itp, "Cal Flag"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($cal, cCalId) var=r
OP         SKIP_IF IS_EMPTY(CELL($cal, r, cCalId))
OP         IF IS_DATE(CELL($cal, r, cDue))
OP           IF CELL($cal, r, cDue) < TODAY
OP             DICT_ADD(known, CLEAN_TEXT(CELL($cal, r, cCalId)), "Overdue", mode=keep)
OP           ELSE
OP             DICT_ADD(known, CLEAN_TEXT(CELL($cal, r, cCalId)), "Fine", mode=keep)
OP         ELSE
OP           DICT_ADD(known, CLEAN_TEXT(CELL($cal, r, cCalId)), "Fine", mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($itp, cNo) var=r
OP         SKIP_IF IS_EMPTY(CELL($itp, r, cNo))
OP         IF SAME(CELL($itp, r, cStat), "Closed")
OP           CLEAR(CELL($itp, r, cFlag), what=contents)
OP         ELSE
OP           IF EXISTS(known, CLEAN_TEXT(CELL($itp, r, cInst)))
OP             DICT_GET(known, CLEAN_TEXT(CELL($itp, r, cInst)), into=got)
OP             IF got = "Overdue"
OP               SET_VALUE(CELL($itp, r, cFlag), "Out of cal")
OP               LET(n, n + 1)
OP             ELSE
OP               CLEAR(CELL($itp, r, cFlag), what=contents)
OP           ELSE
OP             SET_VALUE(CELL($itp, r, cFlag), "Unknown instrument")
OP       RETURN(n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Instruments To Calibrate sheet with the overdue and due soon instruments, earliest due date first, and return how many were listed.
PROC     ListDueInstruments
KIND     Function
MODULE   modSteps
ARGS     cal As Worksheet, out As Worksheet
RET      Long
VAR      cId As Long
VAR      cType As Long
VAR      cDue As Long
VAR      label As String
VAR      outRow As Long
VAR      n As Long
PRE      COL($cal, "Instrument ID") > 0; COL($cal, "Type") > 0; COL($cal, "Next Calibration Due") > 0
MODIFIES $out
OP       LET(cId, COL($cal, "Instrument ID"))
OP       LET(cType, COL($cal, "Type"))
OP       LET(cDue, COL($cal, "Next Calibration Due"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Instrument ID", "Type", "Next Calibration Due", "Status")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($cal, cId) var=r
OP         SKIP_IF IS_EMPTY(CELL($cal, r, cId))
OP         SKIP_IF NOT IS_DATE(CELL($cal, r, cDue))
OP         LET(label, "")
OP         IF CELL($cal, r, cDue) < TODAY
OP           LET(label, "Overdue")
OP         ELSE
OP           IF CELL($cal, r, cDue) <= TODAY + 14
OP             LET(label, "Due soon")
OP         IF label <> ""
OP           WRITE_ROW($out, outRow, CELL($cal, r, cId), CELL($cal, r, cType), CELL($cal, r, cDue), label)
OP           LET(outRow, outRow + 1)
OP           LET(n, n + 1)
OP       SORT(USED_RANGE($out), key=3, order=asc, header=yes)
OP       RETURN(n)
ERR      standard
