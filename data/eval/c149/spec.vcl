VCL 0.6
BUILD
# ASSUME: Welder Quals headers (row 1) include "Welder ID", "Process", "Expiry Date" and "Remark"; Weld Log has "Weld Seam No", "Welder ID", "Process", "Weld Date" and "Qual Check"
# ASSUME: the real sheets are "Welder Quals", "Weld Log" and "Welder Problems"; the TESTs use Quals, Welds and Problems; Welder Problems already exists and is rebuilt on every run
# ASSUME: Expiry Date and Weld Date are real dates; a qualification expired when its Expiry Date is before today and expires soon when it falls from today to 30 days from today
# ASSUME: a qualification with a blank or non-date Expiry Date has its Remark cleared, and a qualification with a blank Welder ID is ignored when checking welds
# ASSUME: a qualification is valid on the Weld Date through its Expiry Date; a weld with a blank or non-date Weld Date or a blank Welder ID is Not qualified
# ASSUME: IDs and processes are compared ignoring case, stray spaces at the ends and repeated inner spaces; Welder Problems skips welds with a blank Welder ID
# ASSUME: each numbered step is its own procedure (MarkQuals, CheckWelds, CountProblems) that modMain.RunHeadless runs in order
GOAL     Mark expired and soon-to-expire welder qualifications, flag welds without a valid qualification and list the problem welders with their counts.
NAME     WelderQualAudit
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modSteps
SIG      modMain.Run()
SIG      modMain.RunHeadless(quals As Worksheet, welds As Worksheet, out As Worksheet) As Long
SIG      modSteps.MarkQuals(quals As Worksheet)
SIG      modSteps.CheckWelds(quals As Worksheet, welds As Worksheet) As Long
SIG      modSteps.CountProblems(welds As Worksheet, out As Worksheet)
FLOW     modMain.Run -> modMain.RunHeadless -> modSteps.MarkQuals
FLOW     modMain.RunHeadless -> modSteps.CheckWelds
FLOW     modMain.RunHeadless -> modSteps.CountProblems
FIXTURE  audit_01.xlsx sheet=Quals cols=A:Welder ID,B:Process,C:Expiry Date,D:Remark rows=10 data=A:values("W1"x3,"W2"x2,""x1,"W3"x1," w4 "x1,"W5"x2)+boundary("W1","W1","W2","W2","","W3"," w4 ","W1","W5","W5");B:values("SMAW"x3,"GTAW"x3," smaw "x1,"smaw"x1,"FCAW"x2)+boundary("SMAW","GTAW","SMAW","GTAW","GTAW","SMAW"," smaw ","smaw","FCAW","FCAW");C:reldates(-100,365)+boundary(-1,0,30,31,200,"",90,-40,-30,365);D:values("old"x10)
FIXTURE  audit_01.xlsx sheet=Welds cols=A:Weld Seam No,B:Welder ID,C:Process,D:Weld Date,E:Qual Check rows=13 data=A:seq(1,1);B:values(" w1 "x1,"W2"x3,""x1,"W3"x3,"W5"x3,"W4"x1,"W1"x1)+boundary(" w1 ","W2","W2","W2","","W3","W3","W3","W5","W5","W4","W1","W5");C:values("smaw"x1,"SMAW"x5,"GTAW"x3,"FCAW"x3," smaw "x1)+boundary("smaw","SMAW","SMAW","GTAW","GTAW","SMAW","SMAW","FCAW","FCAW","SMAW"," smaw ","GTAW","FCAW");D:reldates(-100,100)+boundary(-10,30,31,31,0,-5,-3,0,"",0,50,0,20);E:values("old"x13)
FIXTURE  audit_01.xlsx sheet=Problems cols=A:Welder ID,B:Not qualified welds rows=0 stale=2
FIXTURE  quiet_01.xlsx sheet=Quals cols=A:Welder ID,B:Process,C:Expiry Date,D:Remark rows=2 data=A:values("W1"x2);B:values("SMAW"x2);C:reldates(100,200);D:values("old"x2)
FIXTURE  quiet_01.xlsx sheet=Welds cols=A:Weld Seam No,B:Welder ID,C:Process,D:Weld Date,E:Qual Check rows=2 data=A:seq(1,1);B:values("W1"x2);C:values("SMAW"x2);D:reldates(-50,-10);E:values("old"x2)
FIXTURE  quiet_01.xlsx sheet=Problems cols=A:Welder ID,B:Not qualified welds rows=0 stale=2
TEST     fixture=audit_01.xlsx call=modMain.RunHeadless(Quals, Welds, Problems) runs=2 assert=RESULT = 7 AND ERROR_COUNT = 0 AND COUNTIF(Quals![Remark], "Expired") = 3 AND COUNTIF(Quals![Remark], "Expires soon") = 2 AND COUNTIF(Quals![Remark], "old") = 0 AND COUNTA(Quals![Remark]) = 6 AND COUNTIF(Welds![Qual Check], "Not qualified") = 7 AND COUNTIF(Welds![Qual Check], "old") = 0 AND COUNTA(Welds![Qual Check]) = 8 AND ROWS_USED(Problems) = 4 AND COUNTIF(Problems![Welder ID], "stale") = 0
TEST     fixture=audit_01.xlsx call=modMain.RunHeadless(Quals, Welds, Problems) assert=RESULT = 7 AND COUNTA(Quals![Welder ID]) = PRE_COUNTA(Quals![Welder ID]) AND COUNTA(Welds![Weld Seam No]) = PRE_COUNTA(Welds![Weld Seam No]) AND VALUE(Quals![Remark]2) = "Expired" AND VALUE(Quals![Remark]3) = "Expires soon" AND VALUE(Quals![Remark]4) = "Expires soon" AND VALUE(Quals![Remark]9) = "Expired" AND VALUE(Quals![Remark]10) = "Expired" AND COUNTIFS(Quals![Welder ID], "W5", Quals![Remark], "Expires soon") = 0 AND VALUE(Problems![Welder ID]2) = "W3" AND VALUE(Problems![Not qualified welds]2) = 3 AND VALUE(Problems![Welder ID]3) = "W5" AND VALUE(Problems![Not qualified welds]3) = 2 AND VALUE(Problems![Welder ID]4) = "W2" AND VALUE(Problems![Not qualified welds]4) = 1
TEST     fixture=audit_01.xlsx call=modMain.RunHeadless(Quals, Welds, Problems) assert=COUNTIFS(Welds![Weld Seam No], 1, Welds![Qual Check], "Not qualified") = 0 AND COUNTIFS(Welds![Weld Seam No], 2, Welds![Qual Check], "Not qualified") = 0 AND COUNTIFS(Welds![Weld Seam No], 3, Welds![Qual Check], "Not qualified") = 1 AND COUNTIFS(Welds![Weld Seam No], 4, Welds![Qual Check], "Not qualified") = 0 AND COUNTIFS(Welds![Weld Seam No], 5, Welds![Qual Check], "Not qualified") = 1 AND COUNTIFS(Welds![Weld Seam No], 6, Welds![Qual Check], "Not qualified") = 1 AND COUNTIFS(Welds![Weld Seam No], 7, Welds![Qual Check], "Not qualified") = 1 AND COUNTIFS(Welds![Weld Seam No], 8, Welds![Qual Check], "Not qualified") = 1 AND COUNTIFS(Welds![Weld Seam No], 9, Welds![Qual Check], "Not qualified") = 1 AND COUNTIFS(Welds![Weld Seam No], 10, Welds![Qual Check], "Not qualified") = 1 AND COUNTIFS(Welds![Weld Seam No], 11, Welds![Qual Check], "Not qualified") = 0 AND COUNTIFS(Welds![Weld Seam No], 12, Welds![Qual Check], "Not qualified") = 0 AND COUNTIFS(Welds![Weld Seam No], 13, Welds![Qual Check], "Not qualified") = 0
TEST     fixture=quiet_01.xlsx call=modMain.RunHeadless(Quals, Welds, Problems) assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTA(Quals![Remark]) = 1 AND COUNTA(Welds![Qual Check]) = 1 AND ROWS_USED(Problems) = 1 AND COUNTIF(Problems![Welder ID], "stale") = 0
TEST     fixture=audit_01.xlsx call=modMain.RunHeadless(Welds, Quals, Problems) assert=RESULT = -1 AND ERROR_COUNT = 1
---
VCL 0.6
TASK
GOAL     Start the welder qualification audit on the workbook's own sheets and tell the user how many welds are Not qualified.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "Welder Quals"), SHEET(THIS_BOOK, "Weld Log"), SHEET(THIS_BOOK, "Welder Problems"), into=n)
OP       MSG("Welds Not qualified: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run the three audit steps on the sheets it is given and return how many welds are Not qualified.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     quals As Worksheet, welds As Worksheet, out As Worksheet
RET      Long
VAR      n As Long
PRE      COL($quals, "Welder ID") > 0; COL($quals, "Process") > 0; COL($quals, "Expiry Date") > 0; COL($quals, "Remark") > 0; COL($welds, "Weld Seam No") > 0; COL($welds, "Welder ID") > 0; COL($welds, "Process") > 0; COL($welds, "Weld Date") > 0; COL($welds, "Qual Check") > 0
MODIFIES $quals["Remark"]
MODIFIES $welds["Qual Check"]
MODIFIES $out
OP       CALL(modSteps.MarkQuals, $quals)
OP       CALL(modSteps.CheckWelds, $quals, $welds, into=n)
OP       CALL(modSteps.CountProblems, $welds, $out)
OP       RETURN(n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write Expired or Expires soon in the Remark column of each welder qualification and clear it on the others.
PROC     MarkQuals
KIND     Sub
MODULE   modSteps
ARGS     quals As Worksheet
VAR      cId As Long
VAR      cExp As Long
VAR      cRemark As Long
PRE      COL($quals, "Welder ID") > 0; COL($quals, "Expiry Date") > 0; COL($quals, "Remark") > 0
MODIFIES $quals["Remark"]
OP       LET(cId, COL($quals, "Welder ID"))
OP       LET(cExp, COL($quals, "Expiry Date"))
OP       LET(cRemark, COL($quals, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($quals, cId) var=r
OP         IF IS_DATE(CELL($quals, r, cExp))
OP           IF CELL($quals, r, cExp) < TODAY
OP             SET_VALUE(CELL($quals, r, cRemark), "Expired")
OP           ELSE
OP             IF CELL($quals, r, cExp) <= TODAY + 30
OP               SET_VALUE(CELL($quals, r, cRemark), "Expires soon")
OP             ELSE
OP               CLEAR(CELL($quals, r, cRemark), what=contents)
OP         ELSE
OP           CLEAR(CELL($quals, r, cRemark), what=contents)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write Not qualified in the Qual Check column of each weld whose welder has no qualification for its process still valid on the Weld Date, clear it on the others, and return how many were flagged.
PROC     CheckWelds
KIND     Function
MODULE   modSteps
ARGS     quals As Worksheet, welds As Worksheet
RET      Long
VAR      cqId As Long
VAR      cqProc As Long
VAR      cqExp As Long
VAR      cwSeam As Long
VAR      cwId As Long
VAR      cwProc As Long
VAR      cwDate As Long
VAR      cwCheck As Long
VAR      latest As Dictionary
VAR      best As Variant
VAR      pair As String
VAR      bad As Long
PRE      COL($quals, "Welder ID") > 0; COL($quals, "Process") > 0; COL($quals, "Expiry Date") > 0; COL($welds, "Weld Seam No") > 0; COL($welds, "Welder ID") > 0; COL($welds, "Process") > 0; COL($welds, "Weld Date") > 0; COL($welds, "Qual Check") > 0
MODIFIES $welds["Qual Check"]
OP       LET(cqId, COL($quals, "Welder ID"))
OP       LET(cqProc, COL($quals, "Process"))
OP       LET(cqExp, COL($quals, "Expiry Date"))
OP       LET(cwSeam, COL($welds, "Weld Seam No"))
OP       LET(cwId, COL($welds, "Welder ID"))
OP       LET(cwProc, COL($welds, "Process"))
OP       LET(cwDate, COL($welds, "Weld Date"))
OP       LET(cwCheck, COL($welds, "Qual Check"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($quals, cqProc) var=q
OP         SKIP_IF IS_EMPTY(CELL($quals, q, cqId))
OP         SKIP_IF NOT IS_DATE(CELL($quals, q, cqExp))
OP         LET(pair, CLEAN_TEXT(CELL($quals, q, cqId)) & "|" & CLEAN_TEXT(CELL($quals, q, cqProc)))
OP         IF EXISTS(latest, pair)
OP           DICT_GET(latest, pair, into=best)
OP           IF IS_DATE(best)
OP             IF CELL($quals, q, cqExp) > best
OP               DICT_ADD(latest, pair, CELL($quals, q, cqExp), mode=overwrite)
OP         ELSE
OP           DICT_ADD(latest, pair, CELL($quals, q, cqExp), mode=overwrite)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($welds, cwSeam) var=w
OP         SET_VALUE(CELL($welds, w, cwCheck), "Not qualified")
OP         LET(bad, bad + 1)
OP         IF IS_DATE(CELL($welds, w, cwDate))
OP           LET(pair, CLEAN_TEXT(CELL($welds, w, cwId)) & "|" & CLEAN_TEXT(CELL($welds, w, cwProc)))
OP           IF EXISTS(latest, pair)
OP             DICT_GET(latest, pair, into=best)
OP             IF IS_DATE(best)
OP               IF best >= CELL($welds, w, cwDate)
OP                 CLEAR(CELL($welds, w, cwCheck), what=contents)
OP                 LET(bad, bad - 1)
OP       RETURN(bad)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Welder Problems sheet with each Welder ID and its number of Not qualified welds, the most first.
PROC     CountProblems
KIND     Sub
MODULE   modSteps
ARGS     welds As Worksheet, out As Worksheet
VAR      cId As Long
VAR      cCheck As Long
VAR      counts As Dictionary
VAR      got As Long
VAR      outRow As Long
PRE      COL($welds, "Welder ID") > 0; COL($welds, "Qual Check") > 0
MODIFIES $out
OP       LET(cId, COL($welds, "Welder ID"))
OP       LET(cCheck, COL($welds, "Qual Check"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Welder ID", "Not qualified welds")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($welds, cId) var=r
OP         SKIP_IF IS_EMPTY(CELL($welds, r, cId))
OP         IF SAME(CELL($welds, r, cCheck), "Not qualified")
OP           DICT_ADD(counts, CLEAN_TEXT(CELL($welds, r, cId)), 1, mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(counts)
OP         DICT_GET(counts, k, into=got)
OP         WRITE_ROW($out, outRow, k, got)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=2, order=desc, header=yes)
ERR      standard
