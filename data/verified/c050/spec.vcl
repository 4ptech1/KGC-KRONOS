VCL 0.6
BUILD
# ASSUME: Weld Log headers (row 1) include Weld Seam No, Welder ID, Weld Date, Process, Visual and Remark; Welder Quals has Welder ID, Process and Expiry Date
# ASSUME: the real sheets are "Weld Log", "Welder Quals", "Exceptions", "Welder Summary" and "Month Summary"; the TESTs use WeldLog, Quals, Exceptions, WelderSum and MonthSum
# ASSUME: Exceptions, Welder Summary and Month Summary exist already and are rebuilt from scratch on every run
# ASSUME: a weld is qualified when its welder holds a qualification for its process that expires on or after the weld date; of several rows the latest expiry counts
# ASSUME: welder IDs and processes are trimmed and matched ignoring case; a blank welder or process, or a weld date or expiry that is not a real date, means not qualified
# ASSUME: a blank Visual is not Accepted, so that weld is an exception; rows with a blank Weld Seam No are skipped
# ASSUME: Remark is owned by this tool: it is overwritten on every weld, and cleared on qualified welds as asked
# ASSUME: a weld with no Welder ID is summed under (no ID); rates are fractions shown as percent
GOAL     Mark welds whose welder is not qualified, copy the exceptions to their own sheet and write a summary for each welder and for the month.
NAME     WeldQualityReview
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modQual
MODULE   modReport
SIG      modMain.Run()
SIG      modMain.RunHeadless(weld As Worksheet, quals As Worksheet, exc As Worksheet, wsum As Worksheet, msum As Worksheet) As Long
SIG      modQual.MarkQualification(weld As Worksheet, quals As Worksheet)
SIG      modReport.CopyExceptions(weld As Worksheet, exc As Worksheet) As Long
SIG      modReport.WriteWelderSummary(weld As Worksheet, exc As Worksheet, wsum As Worksheet)
SIG      modReport.WriteMonthSummary(weld As Worksheet, msum As Worksheet, excCount As Long)
FLOW     modMain.Run -> modMain.RunHeadless -> modQual.MarkQualification
FLOW     modMain.RunHeadless -> modReport.CopyExceptions
FLOW     modMain.RunHeadless -> modReport.WriteWelderSummary
FLOW     modMain.RunHeadless -> modReport.WriteMonthSummary
FIXTURE  weld_01.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Welder ID,C:Weld Date,D:Process,E:Visual,F:Remark rows=17 data=A:seq(1,1)+boundary(1,2,3,4,5,6,7,"",8,9,10,11,12,13,14,15,16);B:values("W-01"x3,"W-02"x2,"W-03"x1," W-01 "x1,""x2,"W-05"x3,"W-04"x1,"W-06"x4)+boundary(" W-01 ","W-01","W-02","W-02","W-03","W-01","","","W-05","W-05","W-05","W-04","W-01","W-06","W-06","W-06","W-06");C:dates(2026-03-01,2026-06-30)+boundary(2026-03-10,2026-03-10,2026-03-15,2026-03-16,2026-03-10,2026-03-10,2026-03-10,"",2026-03-10,"",2026-03-10,2026-03-10,2026-06-30,2026-03-10,2026-03-10,2026-03-10,2026-03-10);D:values("GTAW"x9,"SMAW"x4," gtaw "x1,"TIG"x1,""x2)+boundary(" gtaw ","SMAW","GTAW","GTAW","SMAW","GTAW","GTAW","","SMAW","SMAW","TIG","","GTAW","GTAW","GTAW","GTAW","GTAW");E:values("Accepted"x13,"accepted"x1,"Rejected"x1,""x2)+boundary("Rejected","Accepted","Accepted","Accepted","Accepted","Accepted","Accepted","","","Accepted","Accepted","accepted","Accepted","Accepted","Accepted","Accepted","Accepted");F:values("Rework"x2,"Not qualified"x2,""x13)+boundary("","","","Not qualified","","Rework","","","","Rework","","","","Not qualified","","","")
FIXTURE  weld_01.xlsx sheet=Quals cols=A:Welder ID,B:Process,C:Expiry Date rows=10 data=A:values("W-06"x2,"W-01"x2,"W-02"x2,"W-03"x1,""x1,"W-04"x1,"W-05"x1)+boundary("W-06","W-01","W-01","W-02","W-02","W-03","","W-04","W-05","W-06");B:values("GTAW"x6,"SMAW"x3,""x1)+boundary("GTAW","GTAW","SMAW","GTAW","GTAW","SMAW","GTAW","","SMAW","GTAW");C:dates(2025-01-01,2027-01-01)+boundary(2025-01-01,2026-06-30,2026-01-31,2026-03-15,2025-12-31,"",2027-01-01,2027-01-01,2026-12-31,2026-12-31)
FIXTURE  weld_01.xlsx sheet=Exceptions cols=A:Weld Seam No,B:Welder ID,C:Weld Date,D:Process,E:Visual,F:Remark rows=0 stale=2
FIXTURE  weld_01.xlsx sheet=WelderSum cols=A:Welder ID,B:Welds,C:Exceptions,D:Exception Rate rows=0 stale=2
FIXTURE  weld_01.xlsx sheet=MonthSum cols=A:Welds,B:Exceptions,C:Rate rows=0 stale=2
FIXTURE  weld_02.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Welder ID,C:Weld Date,D:Process,E:Visual,F:Remark rows=3 data=A:seq(1,1);B:values("W-01"x3);C:dates(2026-03-01,2026-03-31);D:values("GTAW"x3);E:values("Accepted"x3);F:values("Not qualified"x1,""x2)+boundary("Not qualified","","")
FIXTURE  weld_02.xlsx sheet=Quals cols=A:Welder ID,B:Process,C:Expiry Date rows=1 data=A:values("W-01"x1);B:values("GTAW"x1);C:dates(2026-12-31,2026-12-31)
FIXTURE  weld_02.xlsx sheet=Exceptions cols=A:Weld Seam No,B:Welder ID,C:Weld Date,D:Process,E:Visual,F:Remark rows=0 stale=2
FIXTURE  weld_02.xlsx sheet=WelderSum cols=A:Welder ID,B:Welds,C:Exceptions,D:Exception Rate rows=0 stale=2
FIXTURE  weld_02.xlsx sheet=MonthSum cols=A:Welds,B:Exceptions,C:Rate rows=0 stale=2
FIXTURE  weld_03.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Welder ID,C:Weld Date,D:Process,E:Visual,F:Remark rows=0
FIXTURE  weld_03.xlsx sheet=Quals cols=A:Welder ID,B:Process,C:Expiry Date rows=0
FIXTURE  weld_03.xlsx sheet=Exceptions cols=A:Weld Seam No,B:Welder ID,C:Weld Date,D:Process,E:Visual,F:Remark rows=0 stale=2
FIXTURE  weld_03.xlsx sheet=WelderSum cols=A:Welder ID,B:Welds,C:Exceptions,D:Exception Rate rows=0 stale=2
FIXTURE  weld_03.xlsx sheet=MonthSum cols=A:Welds,B:Exceptions,C:Rate rows=0 stale=2
FIXTURE  weld_04.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Welder ID,C:Weld Date,D:Process,E:Visual,F:Remark rows=1 data=A:seq(1,1);B:values("W-01"x1);C:dates(2026-03-01,2026-03-31);D:values("GTAW"x1);E:values("Rejected"x1);F:values(""x1)
FIXTURE  weld_04.xlsx sheet=Quals cols=A:Welder ID,B:Process,C:Expiry Date rows=1 data=A:values("W-01"x1);B:values("GTAW"x1);C:dates(2026-12-31,2026-12-31)
FIXTURE  weld_04.xlsx sheet=Exceptions cols=A:Weld Seam No,B:Welder ID,C:Weld Date,D:Process,E:Visual,F:Remark rows=0 stale=2
FIXTURE  weld_04.xlsx sheet=WelderSum cols=A:Welder ID,B:Welds,C:Exceptions,D:Exception Rate rows=0 stale=2
FIXTURE  weld_04.xlsx sheet=MonthSum cols=A:Welds,B:Exceptions,C:Rate rows=0 stale=2
TEST     fixture=weld_01.xlsx call=modMain.RunHeadless(WeldLog, Quals, Exceptions, WelderSum, MonthSum) runs=2 assert=RESULT = 9 AND ERROR_COUNT = 0 AND COUNTIF(WeldLog![Remark], "Not qualified") = 7 AND COUNTA(WeldLog![Remark]) = 8 AND COUNTIF(Exceptions![Remark], "Not qualified") = 7 AND ROWS_USED(Exceptions) = 10 AND COUNTIF(Exceptions![Weld Seam No], "stale") = 0 AND FILL_COUNT(Exceptions![Weld Seam No], #FF0000) = 0
TEST     fixture=weld_01.xlsx call=modMain.RunHeadless(WeldLog, Quals, Exceptions, WelderSum, MonthSum) assert=COUNTA(WeldLog![Weld Seam No]) = PRE_COUNTA(WeldLog![Weld Seam No]) AND COUNTA(WeldLog![Visual]) = PRE_COUNTA(WeldLog![Visual]) AND COUNTA(Quals![Welder ID]) = PRE_COUNTA(Quals![Welder ID]) AND VALUE(WeldLog![Remark]2) = "" AND VALUE(WeldLog![Remark]3) = "Not qualified" AND VALUE(WeldLog![Remark]4) = "" AND VALUE(WeldLog![Remark]5) = "Not qualified" AND VALUE(WeldLog![Remark]7) = "" AND VALUE(WeldLog![Remark]10) = "" AND VALUE(WeldLog![Remark]14) = "" AND VALUE(WeldLog![Remark]15) = ""
TEST     fixture=weld_01.xlsx call=modMain.RunHeadless(WeldLog, Quals, Exceptions, WelderSum, MonthSum) assert=ROWS_USED(WelderSum) = 8 AND SUM(WelderSum![Welds]) = 16 AND SUM(WelderSum![Exceptions]) = 9 AND VALUE(WelderSum![Welder ID]2) = "W-01" AND VALUE(WelderSum![Welds]2) = 4 AND VALUE(WelderSum![Exceptions]2) = 2 AND VALUE(WelderSum![Exception Rate]2) = 0.5 AND VALUE(WelderSum![Welder ID]3) = "W-02" AND VALUE(WelderSum![Exception Rate]3) = 0.5 AND VALUE(WelderSum![Welder ID]5) = "(no ID)" AND VALUE(WelderSum![Exception Rate]5) = 1 AND VALUE(WelderSum![Welds]6) = 3 AND VALUE(WelderSum![Exceptions]6) = 3 AND VALUE(WelderSum![Welder ID]8) = "W-06" AND VALUE(WelderSum![Welds]8) = 4 AND VALUE(WelderSum![Exception Rate]8) = 0 AND COUNTIF(WelderSum![Welder ID], "stale") = 0 AND FILL_COUNT(WelderSum![Welder ID], #FF0000) = 0
TEST     fixture=weld_01.xlsx call=modMain.RunHeadless(WeldLog, Quals, Exceptions, WelderSum, MonthSum) assert=ROWS_USED(MonthSum) = 2 AND VALUE(MonthSum![Welds]2) = 16 AND VALUE(MonthSum![Exceptions]2) = 9 AND VALUE(MonthSum![Rate]2) = 0.5625 AND COUNTIF(MonthSum![Welds], "stale") = 0 AND FILL_COUNT(MonthSum![Welds], #FF0000) = 0
TEST     fixture=weld_02.xlsx call=modMain.RunHeadless(WeldLog, Quals, Exceptions, WelderSum, MonthSum) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTA(WeldLog![Remark]) = 1 AND ROWS_USED(Exceptions) = 1 AND ROWS_USED(WelderSum) = 2 AND VALUE(WelderSum![Welds]2) = 3 AND VALUE(WelderSum![Exception Rate]2) = 0 AND VALUE(MonthSum![Welds]2) = 3 AND VALUE(MonthSum![Exceptions]2) = 0 AND VALUE(MonthSum![Rate]2) = 0
TEST     fixture=weld_03.xlsx call=modMain.RunHeadless(WeldLog, Quals, Exceptions, WelderSum, MonthSum) assert=RESULT = 0 AND ERROR_COUNT = 0 AND ROWS_USED(Exceptions) = 1 AND ROWS_USED(WelderSum) = 1 AND VALUE(MonthSum![Welds]2) = 0 AND VALUE(MonthSum![Exceptions]2) = 0 AND VALUE(MonthSum![Rate]2) = 0
TEST     fixture=weld_04.xlsx call=modMain.RunHeadless(WeldLog, Quals, Exceptions, WelderSum, MonthSum) assert=RESULT = 1 AND ERROR_COUNT = 0 AND COUNTIF(WeldLog![Remark], "Not qualified") = 0 AND ROWS_USED(Exceptions) = 2 AND ROWS_USED(WelderSum) = 2 AND VALUE(WelderSum![Exception Rate]2) = 1 AND VALUE(MonthSum![Welds]2) = 1 AND VALUE(MonthSum![Exceptions]2) = 1 AND VALUE(MonthSum![Rate]2) = 1
TEST     fixture=weld_01.xlsx call=modMain.RunHeadless(Quals, WeldLog, Exceptions, WelderSum, MonthSum) assert=RESULT = -1 AND ERROR_COUNT = 1
---
VCL 0.6
TASK
GOAL     Start the weld quality review on the workbook's own sheets and tell the user how many welds are exceptions.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "Weld Log"), SHEET(THIS_BOOK, "Welder Quals"), SHEET(THIS_BOOK, "Exceptions"), SHEET(THIS_BOOK, "Welder Summary"), SHEET(THIS_BOOK, "Month Summary"), into=n)
OP       IF n < 0
OP         MSG("The weld review did not run: check the sheet headings, see ForgeLog.")
OP       ELSE
OP         MSG("Weld exceptions: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run the four steps of the weld quality review on the five sheets it is given and return how many welds are exceptions.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     weld As Worksheet, quals As Worksheet, exc As Worksheet, wsum As Worksheet, msum As Worksheet
RET      Long
VAR      nExc As Long
PRE      COL($weld, "Weld Seam No") > 0; COL($weld, "Welder ID") > 0; COL($weld, "Weld Date") > 0; COL($weld, "Process") > 0; COL($weld, "Visual") > 0; COL($weld, "Remark") > 0; COL($quals, "Welder ID") > 0; COL($quals, "Process") > 0; COL($quals, "Expiry Date") > 0
MODIFIES $weld["Remark"]
MODIFIES $exc
MODIFIES $wsum
MODIFIES $msum
OP       CALL(modQual.MarkQualification, $weld, $quals)
OP       CALL(modReport.CopyExceptions, $weld, $exc, into=nExc)
OP       CALL(modReport.WriteWelderSummary, $weld, $exc, $wsum)
OP       CALL(modReport.WriteMonthSummary, $weld, $msum, nExc)
OP       RETURN(nExc)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write Not qualified in Remark for each weld whose welder has no qualification for its process valid on the weld date, and clear Remark on the others.
PROC     MarkQualification
KIND     Sub
MODULE   modQual
ARGS     weld As Worksheet, quals As Worksheet
VAR      cSeam As Long
VAR      cWelder As Long
VAR      cDate As Long
VAR      cProc As Long
VAR      cRemark As Long
VAR      qWelder As Long
VAR      qProc As Long
VAR      qExpiry As Long
VAR      latest As Dictionary
VAR      lookup As String
VAR      known As Variant
VAR      qualified As Boolean
PRE      COL($weld, "Weld Seam No") > 0; COL($weld, "Welder ID") > 0; COL($weld, "Weld Date") > 0; COL($weld, "Process") > 0; COL($weld, "Remark") > 0; COL($quals, "Welder ID") > 0; COL($quals, "Process") > 0; COL($quals, "Expiry Date") > 0
MODIFIES $weld["Remark"]
OP       LET(cSeam, COL($weld, "Weld Seam No"))
OP       LET(cWelder, COL($weld, "Welder ID"))
OP       LET(cDate, COL($weld, "Weld Date"))
OP       LET(cProc, COL($weld, "Process"))
OP       LET(cRemark, COL($weld, "Remark"))
OP       LET(qWelder, COL($quals, "Welder ID"))
OP       LET(qProc, COL($quals, "Process"))
OP       LET(qExpiry, COL($quals, "Expiry Date"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($quals, qWelder) var=q
OP         SKIP_IF IS_EMPTY(CELL($quals, q, qWelder)) OR IS_EMPTY(CELL($quals, q, qProc))
OP         SKIP_IF NOT IS_DATE(CELL($quals, q, qExpiry))
OP         LET(lookup, TRIM(CELL($quals, q, qWelder)) & "|" & TRIM(CELL($quals, q, qProc)))
OP         IF EXISTS(latest, lookup)
OP           DICT_GET(latest, lookup, into=known)
OP           IF IS_DATE(known)
OP             IF CELL($quals, q, qExpiry) > known
OP               DICT_ADD(latest, lookup, CELL($quals, q, qExpiry), mode=overwrite)
OP         ELSE
OP           DICT_ADD(latest, lookup, CELL($quals, q, qExpiry), mode=overwrite)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($weld, cSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($weld, r, cSeam))
OP         LET(qualified, FALSE)
OP         LET(lookup, TRIM(CELL($weld, r, cWelder)) & "|" & TRIM(CELL($weld, r, cProc)))
OP         IF EXISTS(latest, lookup)
OP           DICT_GET(latest, lookup, into=known)
OP           IF IS_DATE(known)
OP             IF IS_DATE(CELL($weld, r, cDate))
OP               IF known >= CELL($weld, r, cDate)
OP                 LET(qualified, TRUE)
OP         IF qualified
OP           CLEAR(CELL($weld, r, cRemark), what=contents)
OP         ELSE
OP           SET_VALUE(CELL($weld, r, cRemark), "Not qualified")
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Exceptions sheet with the Weld Log headings and a copy of every weld whose Visual is not Accepted or whose Remark is Not qualified.
PROC     CopyExceptions
KIND     Function
MODULE   modReport
ARGS     weld As Worksheet, exc As Worksheet
RET      Long
VAR      cSeam As Long
VAR      cVisual As Long
VAR      cRemark As Long
VAR      nCols As Long
VAR      outRow As Long
PRE      COL($weld, "Weld Seam No") > 0; COL($weld, "Visual") > 0; COL($weld, "Remark") > 0
MODIFIES $exc
OP       LET(cSeam, COL($weld, "Weld Seam No"))
OP       LET(cVisual, COL($weld, "Visual"))
OP       LET(cRemark, COL($weld, "Remark"))
OP       LET(nCols, LASTCOL($weld, 1))
OP       CLEAR(USED_RANGE($exc), what=all)
OP       COPY_ROW($weld, 1, $exc, 1, cols=nCols)
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($weld, cSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($weld, r, cSeam))
OP         IF NOT SAME(CELL($weld, r, cVisual), "Accepted") OR SAME(CELL($weld, r, cRemark), "Not qualified")
OP           COPY_ROW($weld, r, $exc, outRow, cols=nCols)
OP           LET(outRow, outRow + 1)
OP       RETURN(outRow - 2)
ERR      standard
---
VCL 0.6
TASK
# ASSUME: the Exceptions sheet is a copy of the Weld Log rows, so its Welder ID and Weld Seam No sit in the same columns as in the Weld Log
GOAL     Rebuild the Welder Summary sheet with one row for each welder: the welds, the exceptions and the exception rate.
PROC     WriteWelderSummary
KIND     Sub
MODULE   modReport
ARGS     weld As Worksheet, exc As Worksheet, wsum As Worksheet
VAR      cSeam As Long
VAR      cWelder As Long
VAR      welds As Dictionary
VAR      excs As Dictionary
VAR      who As String
VAR      nW As Variant
VAR      nE As Variant
VAR      outRow As Long
PRE      COL($weld, "Weld Seam No") > 0; COL($weld, "Welder ID") > 0
MODIFIES $wsum
OP       LET(cSeam, COL($weld, "Weld Seam No"))
OP       LET(cWelder, COL($weld, "Welder ID"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($weld, cSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($weld, r, cSeam))
OP         LET(who, IIF(IS_EMPTY(CELL($weld, r, cWelder)), "(no ID)", TRIM(CELL($weld, r, cWelder))))
OP         DICT_ADD(welds, who, 1, mode=sum)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($exc, cSeam) var=e
OP         LET(who, IIF(IS_EMPTY(CELL($exc, e, cWelder)), "(no ID)", TRIM(CELL($exc, e, cWelder))))
OP         DICT_ADD(excs, who, 1, mode=sum)
OP       CLEAR(USED_RANGE($wsum), what=all)
OP       WRITE_ROW($wsum, 1, "Welder ID", "Welds", "Exceptions", "Exception Rate")
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(welds)
OP         DICT_GET(welds, k, into=nW)
OP         LET(nE, 0)
OP         IF EXISTS(excs, k)
OP           DICT_GET(excs, k, into=nE)
OP         IF IS_NUMBER(nW)
OP           IF IS_NUMBER(nE)
OP             WRITE_ROW($wsum, outRow, k, nW, nE, nE / nW)
OP             FORMAT(CELL($wsum, outRow, 4), numfmt="0.0%")
OP             LET(outRow, outRow + 1)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Month Summary sheet with the month's total welds, total exceptions and exception rate.
PROC     WriteMonthSummary
KIND     Sub
MODULE   modReport
ARGS     weld As Worksheet, msum As Worksheet, excCount As Long
VAR      cSeam As Long
VAR      welds As Long
VAR      rate As Double
PRE      COL($weld, "Weld Seam No") > 0
MODIFIES $msum
OP       LET(cSeam, COL($weld, "Weld Seam No"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($weld, cSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($weld, r, cSeam))
OP         LET(welds, welds + 1)
OP       IF welds > 0 THEN LET(rate, $excCount / welds)
OP       CLEAR(USED_RANGE($msum), what=all)
OP       WRITE_ROW($msum, 1, "Welds", "Exceptions", "Rate")
OP       WRITE_ROW($msum, 2, welds, $excCount, rate)
OP       FORMAT(CELL($msum, 2, 3), numfmt="0.0%")
ERR      standard
