VCL 0.7
BUILD
# ASSUME: the real sheets are "Welder Quals", "Weld Log" and "Welder Problems"; the TESTs use Quals, Welds and Out
# ASSUME: Welder Problems exists and is rebuilt on every run; its headings are Welder ID and Not Qualified Welds
# ASSUME: a qualification expires on its Expiry Date and is still valid on that day; Expires soon means today up to 30 days ahead
# ASSUME: a welder's qualification for a process is valid on the Weld Date when the latest Expiry Date for that welder and process is on or after it
# ASSUME: Welder ID and Process are matched trimmed, ignoring case and extra inner spaces; a blank or non-date Expiry Date or Weld Date is never marked
# ASSUME: a weld row with a blank Weld Seam No is skipped; the welder IDs on Welder Problems keep the spelling of their first weld
# ASSUME: one macro runs all four steps; VCL lets only modMain.Run be started alone, so the three steps are separate procedures that Run chains, and can be started alone only by calling them from the VBA window
GOAL     Mark expired and soon-expiring welder qualifications, mark welds made without a valid qualification, list the welders with the most unqualified welds, and report how many welds are not qualified.
NAME     WelderQualAudit
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modSteps
SIG      modMain.Run()
SIG      modMain.RunHeadless(quals As Worksheet, welds As Worksheet, out As Worksheet) As Long
SIG      modSteps.MarkQualExpiry(quals As Worksheet)
SIG      modSteps.CheckWeldQuals(quals As Worksheet, welds As Worksheet) As Long
SIG      modSteps.WriteWelderProblems(welds As Worksheet, out As Worksheet)
FLOW     modMain.Run -> modMain.RunHeadless -> modSteps.MarkQualExpiry
FLOW     modMain.RunHeadless -> modSteps.CheckWeldQuals
FLOW     modMain.RunHeadless -> modSteps.WriteWelderProblems
FIXTURE  wq_01.xlsx sheet=Quals cols=A:Welder ID,B:Process,C:Expiry Date,D:Remark rows=7 data=A:values("W-1"x2,"W-2"x2,"w-3 "x1,"W-4"x1,"W-3"x1)+boundary("W-1","W-1","W-2","W-2","w-3 ","W-4","W-3");B:values("GTAW"x4,"SMAW"x2,"gtaw "x1)+boundary("GTAW","SMAW","GTAW","SMAW","gtaw ","GTAW","GTAW");C:reldates(-60,120)+boundary(-1,0,30,31,-5,"",100);D:values("old"x7)
FIXTURE  wq_01.xlsx sheet=Welds cols=A:Weld Seam No,B:Welder ID,C:Process,D:Weld Date,E:Qual Check rows=10 data=A:seq(1,1)+boundary(1,2,3,4,"",6,7,8,9,10);B:values("W-1"x5,"W-9"x1,"w-2 "x1,"W-3"x1,"W-4"x1,"W-5"x1)+boundary("W-1","W-1","W-1","W-1","W-9","w-2 ","W-3","W-4","W-5","W-1");C:values("GTAW"x6,"SMAW"x3,"gtaw"x1)+boundary("GTAW","GTAW","SMAW","SMAW","GTAW","gtaw","GTAW","GTAW","SMAW","GTAW");D:reldates(-60,120)+boundary(-1,0,0,1,0,-40,50,-3,-3,"");E:values("old"x10)
FIXTURE  wq_01.xlsx sheet=Out cols=A:Welder ID,B:Not Qualified Welds rows=0 stale=2
TEST     fixture=wq_01.xlsx call=modMain.RunHeadless(Quals, Welds, Out) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTIF(Quals![Remark], "Expired") = 2 AND COUNTIF(Quals![Remark], "Expires soon") = 2 AND COUNTIF(Quals![Remark], "old") = 0 AND VALUE(Quals![Remark]2) = "Expired" AND VALUE(Quals![Remark]3) = "Expires soon" AND VALUE(Quals![Remark]5) = "" AND COUNTIF(Welds![Qual Check], "Not qualified") = 4 AND COUNTIF(Welds![Qual Check], "old") = 1 AND VALUE(Welds![Qual Check]3) = "Not qualified" AND VALUE(Welds![Qual Check]2) = "" AND VALUE(Welds![Qual Check]8) = ""
TEST     fixture=wq_01.xlsx call=modMain.RunHeadless(Quals, Welds, Out) assert=RESULT = 4 AND ROWS_USED(Out) = 4 AND VALUE(Out![Not Qualified Welds]2) = 2 AND VALUE(Out![Not Qualified Welds]3) = 1 AND VALUE(Out![Not Qualified Welds]4) = 1 AND COUNTIF(Out![Welder ID], "stale") = 0 AND COUNTA(Quals![Welder ID]) = PRE_COUNTA(Quals![Welder ID]) AND COUNTA(Welds![Weld Seam No]) = PRE_COUNTA(Welds![Weld Seam No]) AND COUNTA(Welds![Welder ID]) = PRE_COUNTA(Welds![Welder ID])
TEST     fixture=wq_01.xlsx call=modMain.RunHeadless(Welds, Quals, Out) assert=RESULT = -1 AND ERROR_COUNT = 1
---
VCL 0.7
TASK
GOAL     Start the welder qualification audit on the workbook's own sheets and tell the user how many welds are not qualified.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "Welder Quals"), SHEET(THIS_BOOK, "Weld Log"), SHEET(THIS_BOOK, "Welder Problems"), into=n)
OP       MSG("Welds not qualified: " & n)
ERR      standard
---
VCL 0.7
TASK
GOAL     Run the three steps of the welder qualification audit on the three sheets it is given and return how many welds are not qualified.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     quals As Worksheet, welds As Worksheet, out As Worksheet
RET      Long
VAR      bad As Long
PRE      COL($quals, "Welder ID") > 0; COL($quals, "Process") > 0; COL($quals, "Expiry Date") > 0; COL($quals, "Remark") > 0; COL($welds, "Weld Seam No") > 0; COL($welds, "Welder ID") > 0; COL($welds, "Process") > 0; COL($welds, "Weld Date") > 0; COL($welds, "Qual Check") > 0
MODIFIES $quals["Remark"]
MODIFIES $welds["Qual Check"]
MODIFIES $out
OP       CALL(modSteps.MarkQualExpiry, $quals)
OP       CALL(modSteps.CheckWeldQuals, $quals, $welds, into=bad)
OP       CALL(modSteps.WriteWelderProblems, $welds, $out)
OP       RETURN(bad)
ERR      standard
---
VCL 0.7
TASK
GOAL     Write Expired or Expires soon in each qualification's Remark from its expiry date, and clear the Remark on the others.
PROC     MarkQualExpiry
KIND     Sub
MODULE   modSteps
ARGS     quals As Worksheet
VAR      cId As Long
VAR      cExp As Long
VAR      cRemark As Long
VAR      mark As Long
PRE      COL($quals, "Welder ID") > 0; COL($quals, "Process") > 0; COL($quals, "Expiry Date") > 0; COL($quals, "Remark") > 0
MODIFIES $quals["Remark"]
OP       LET(cId, COL($quals, "Welder ID"))
OP       LET(cExp, COL($quals, "Expiry Date"))
OP       LET(cRemark, COL($quals, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($quals, cId) var=r
OP         SKIP_IF IS_EMPTY(CELL($quals, r, cId))
OP         LET(mark, 0)
OP         IF IS_DATE(CELL($quals, r, cExp))
OP           IF CELL($quals, r, cExp) < TODAY
OP             LET(mark, 1)
OP           ELSE
OP             IF CELL($quals, r, cExp) <= TODAY + 30
OP               LET(mark, 2)
OP         IF mark = 1
OP           SET_VALUE(CELL($quals, r, cRemark), "Expired")
OP         ELSE
OP           IF mark = 2
OP             SET_VALUE(CELL($quals, r, cRemark), "Expires soon")
OP           ELSE
OP             CLEAR(CELL($quals, r, cRemark), what=contents)
ERR      standard
---
VCL 0.7
TASK
GOAL     Write Not qualified in each weld's Qual Check when its welder had no valid qualification for the process on the weld date, and count them.
PROC     CheckWeldQuals
KIND     Function
MODULE   modSteps
ARGS     quals As Worksheet, welds As Worksheet
RET      Long
VAR      qId As Long
VAR      qProc As Long
VAR      qExp As Long
VAR      wSeam As Long
VAR      wId As Long
VAR      wProc As Long
VAR      wDate As Long
VAR      wCheck As Long
VAR      best As Dictionary
VAR      k As String
VAR      d As Long
VAR      cur As Long
VAR      mark As Long
VAR      bad As Long
PRE      COL($quals, "Welder ID") > 0; COL($quals, "Process") > 0; COL($quals, "Expiry Date") > 0; COL($quals, "Remark") > 0; COL($welds, "Weld Seam No") > 0; COL($welds, "Welder ID") > 0; COL($welds, "Process") > 0; COL($welds, "Weld Date") > 0; COL($welds, "Qual Check") > 0
MODIFIES $welds["Qual Check"]
OP       LET(qId, COL($quals, "Welder ID"))
OP       LET(qProc, COL($quals, "Process"))
OP       LET(qExp, COL($quals, "Expiry Date"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($quals, qId) var=r
OP         SKIP_IF IS_EMPTY(CELL($quals, r, qId))
OP         SKIP_IF NOT IS_DATE(CELL($quals, r, qExp))
OP         LET(d, CELL($quals, r, qExp) - TODAY)
OP         LET(k, CLEAN_TEXT(CELL($quals, r, qId)) & "|" & CLEAN_TEXT(CELL($quals, r, qProc)))
OP         IF EXISTS(best, k)
OP           DICT_GET(best, k, into=cur)
OP           IF d > cur
OP             DICT_ADD(best, k, d, mode=overwrite)
OP         ELSE
OP           DICT_ADD(best, k, d, mode=overwrite)
OP       LET(wSeam, COL($welds, "Weld Seam No"))
OP       LET(wId, COL($welds, "Welder ID"))
OP       LET(wProc, COL($welds, "Process"))
OP       LET(wDate, COL($welds, "Weld Date"))
OP       LET(wCheck, COL($welds, "Qual Check"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($welds, wSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($welds, r, wSeam))
OP         LET(mark, 0)
OP         IF IS_DATE(CELL($welds, r, wDate))
OP           LET(k, CLEAN_TEXT(CELL($welds, r, wId)) & "|" & CLEAN_TEXT(CELL($welds, r, wProc)))
OP           IF EXISTS(best, k)
OP             DICT_GET(best, k, into=cur)
OP             IF CELL($welds, r, wDate) - TODAY > cur
OP               LET(mark, 1)
OP           ELSE
OP             LET(mark, 1)
OP         IF mark = 1
OP           SET_VALUE(CELL($welds, r, wCheck), "Not qualified")
OP           LET(bad, bad + 1)
OP         ELSE
OP           CLEAR(CELL($welds, r, wCheck), what=contents)
OP       RETURN(bad)
ERR      standard
---
VCL 0.7
TASK
GOAL     Rebuild the Welder Problems sheet with each welder's number of Not qualified welds, the most first.
PROC     WriteWelderProblems
KIND     Sub
MODULE   modSteps
ARGS     welds As Worksheet, out As Worksheet
VAR      wId As Long
VAR      wCheck As Long
VAR      counts As Dictionary
VAR      got As Long
VAR      outRow As Long
PRE      COL($welds, "Welder ID") > 0; COL($welds, "Qual Check") > 0
MODIFIES $out
OP       LET(wId, COL($welds, "Welder ID"))
OP       LET(wCheck, COL($welds, "Qual Check"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($welds, wId) var=r
OP         SKIP_IF IS_EMPTY(CELL($welds, r, wId))
OP         IF SAME(CELL($welds, r, wCheck), "Not qualified")
OP           DICT_ADD(counts, CLEAN_TEXT(CELL($welds, r, wId)), 1, mode=sum)
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Welder ID", "Not Qualified Welds")
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(counts)
OP         DICT_GET(counts, k, into=got)
OP         WRITE_ROW($out, outRow, k, got)
OP         LET(outRow, outRow + 1)
OP       IF outRow > 2
OP         SORT(USED_RANGE($out), key=2, order=desc, header=yes)
ERR      standard
