VCL 0.6
BUILD
# ASSUME: NDE Report headers (row 1) are "Weld Seam No", "Method", "Result" and "Report Date"; ITP Tracker headers are "ITP No", "Point", "Inspection Status" and "Planned Date"
# ASSUME: the real sheets are "NDE Report", "ITP Tracker", "Rejects", "Method Summary" and "Inspection Summary" and all five exist; the TESTs use Nde, Itp, Rejects, MethodSummary and InspSummary
# ASSUME: a Result counts as Reject, Accept or Repair when it is that word ignoring case and extra spaces; every Reject row is listed with its cells as typed, even with a blank seam, method or date
# ASSUME: Report Date holds real dates; Rejects is sorted newest first and a Reject with no date goes last
# ASSUME: an ITP Tracker row is a point when its Point is not blank and other rows are left alone; a blank status counts as not Closed
# ASSUME: a point is Overdue only when its Planned Date is a real date before today (today itself is not); a blank or non-date Planned Date is not overdue
# ASSUME: the Flag column is found by its header; when the tracker has none, the header Flag is added after the last used header of row 1
# ASSUME: Flag holds only the word Overdue, so every other point gets an empty Flag
# ASSUME: methods are matched after TRIM, ignoring case, and shown as first typed; a row with a blank Method is left out of Method Summary; other results (Pending, blank) are counted nowhere
# ASSUME: the output sheets are cleared and rewritten on every run; Method Summary is sorted A to Z; Inspection Summary has a Measure and a Value column and counts the Rejects rows, the Overdue points and the Method Summary rows
GOAL     List the rejected NDE results newest first, flag overdue ITP points, and rebuild the method summary and inspection summary sheets.
NAME     InspectionFollowUp
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modRejects
MODULE   modTracker
MODULE   modSummary
SIG      modMain.Run()
SIG      modMain.RunHeadless(nde As Worksheet, itp As Worksheet, rejects As Worksheet, methodSum As Worksheet, inspSum As Worksheet) As Long
SIG      modRejects.ListRejects(nde As Worksheet, out As Worksheet) As Long
SIG      modTracker.FlagOverdue(itp As Worksheet) As Long
SIG      modSummary.WriteMethodSummary(nde As Worksheet, out As Worksheet) As Long
SIG      modSummary.WriteInspectionSummary(out As Worksheet, rejected As Long, overdue As Long, methodCount As Long)
FLOW     modMain.Run -> modMain.RunHeadless -> modRejects.ListRejects
FLOW     modMain.RunHeadless -> modTracker.FlagOverdue
FLOW     modMain.RunHeadless -> modSummary.WriteMethodSummary
FLOW     modMain.RunHeadless -> modSummary.WriteInspectionSummary
FIXTURE  insp_01.xlsx sheet=Nde cols=A:Weld Seam No,B:Method,C:Result,D:Report Date rows=14 data=A:values("S-01"x1,"S-02"x1,"S-03"x1,"S-04"x1,"S-05"x1,"S-06"x1,"S-07"x1,""x1,"S-09"x1,"S-10"x1,"S-11"x1,"S-12"x1,"S-13"x1,"S-14"x1)+boundary("S-01","S-02","S-03","S-04","S-05","S-06","S-07","","S-09","S-10","S-11","S-12","S-13","S-14");B:values("RT"x3,"UT"x3,"MT"x4,"PT"x1," ut "x1,""x1,"rt"x1)+boundary("RT","UT","RT","MT","UT","PT","RT","MT"," ut ","UT","MT","","rt","MT");C:values("Accept"x4,"Reject"x5,"Repair"x2," reject "x1,"Pending"x1,""x1)+boundary("Accept","Reject","Reject","Repair"," reject ","Accept","Reject","Reject","Repair","Accept","Pending","Reject","Accept","");D:dates(2026-01-01,2026-09-30)+boundary(2026-03-01,2026-05-10,2026-08-20,2026-02-02,2026-07-01,2026-06-06,"",2026-04-15,2026-09-09,2026-01-20,2026-05-05,2026-03-30,2026-08-01,2026-04-04)
FIXTURE  insp_01.xlsx sheet=Itp cols=A:ITP No,B:Point,C:Inspection Status,D:Planned Date,E:Flag rows=12 data=A:values("ITP-1"x2,"ITP-2"x2,"ITP-3"x2,"ITP-4"x2,"ITP-5"x2,"ITP-6"x2)+boundary("ITP-1","ITP-1","ITP-2","ITP-2","ITP-3","ITP-3","ITP-4","ITP-4","ITP-5","ITP-5","ITP-6","ITP-6");B:values("P-01"x1,"P-02"x1,"P-03"x1,"P-04"x1,"P-05"x1,"P-06"x1,"P-07"x1,""x1,"P-09"x1,"P-10"x1,"P-11"x1,"P-12"x1)+boundary("P-01","P-02","P-03","P-04","P-05","P-06","P-07","","P-09","P-10","P-11","P-12");C:values("Open"x6,"Closed"x2," closed "x1,"In Progress"x1,""x1,"Hold"x1)+boundary("Open","Open","Closed"," closed ","In Progress","","Open","Open","Open","Open","Hold","Closed");D:reldates(-40,30)+boundary(-1,0,-30,-5,-10,-3,"",-20,1,30,-40,-2);E:values(""x6,"Overdue"x5,"Note"x1)+boundary("","Overdue","Overdue","Overdue","","","Overdue","Note","","","","Overdue")
FIXTURE  insp_01.xlsx sheet=Rejects cols=A:Weld Seam No,B:Method,C:Result,D:Report Date rows=0 stale=8
FIXTURE  insp_01.xlsx sheet=MethodSummary cols=A:Method,B:Accept,C:Repair,D:Reject rows=0 stale=6
FIXTURE  insp_01.xlsx sheet=InspSummary cols=A:Measure,B:Value rows=0 stale=5
FIXTURE  insp_02.xlsx sheet=Nde cols=A:Weld Seam No,B:Method,C:Result,D:Report Date rows=3 data=A:values("S-A1"x1,"S-A2"x1,"S-A3"x1)+boundary("S-A1","S-A2","S-A3");B:values("UT"x2,"RT"x1)+boundary("UT","RT","UT");C:values("Reject"x1,"Accept"x1,"Repair"x1)+boundary("Reject","Accept","Repair");D:dates(2026-01-01,2026-09-30)+boundary(2026-05-05,2026-05-06,2026-05-07)
FIXTURE  insp_02.xlsx sheet=Itp cols=A:ITP No,B:Point,C:Inspection Status,D:Planned Date rows=3 data=A:values("ITP-A"x2,"ITP-B"x1)+boundary("ITP-A","ITP-A","ITP-B");B:values("P-A1"x1,"P-A2"x1,"P-A3"x1)+boundary("P-A1","P-A2","P-A3");C:values("Open"x2,"Closed"x1)+boundary("Open","Closed","Open");D:reldates(-9,9)+boundary(-2,-2,3)
FIXTURE  insp_02.xlsx sheet=Rejects cols=A:Weld Seam No,B:Method,C:Result,D:Report Date rows=0 stale=1
FIXTURE  insp_02.xlsx sheet=MethodSummary cols=A:Method,B:Accept,C:Repair,D:Reject rows=0 stale=1
FIXTURE  insp_02.xlsx sheet=InspSummary cols=A:Measure,B:Value rows=0 stale=1
TEST     fixture=insp_01.xlsx call=modMain.RunHeadless(Nde, Itp, Rejects, MethodSummary, InspSummary) runs=2 assert=RESULT = 6 AND ERROR_COUNT = 0 AND COUNTIF(Itp![Flag], "Overdue") = 4 AND ROWS_USED(Rejects) = 7 AND ROWS_USED(MethodSummary) = 5 AND ROWS_USED(InspSummary) = 4 AND COUNTIF(Rejects![Weld Seam No], "stale") = 0 AND COUNTIF(MethodSummary![Method], "stale") = 0 AND COUNTIF(InspSummary![Measure], "stale") = 0
TEST     fixture=insp_01.xlsx call=modMain.RunHeadless(Nde, Itp, Rejects, MethodSummary, InspSummary) assert=RESULT = 6 AND ERROR_COUNT = 0 AND COUNTIFS(Itp![Point], "P-01", Itp![Flag], "Overdue") = 1 AND COUNTIFS(Itp![Point], "P-02", Itp![Flag], "Overdue") = 0 AND COUNTIFS(Itp![Point], "P-03", Itp![Flag], "Overdue") = 0 AND COUNTIFS(Itp![Point], "P-04", Itp![Flag], "Overdue") = 0 AND COUNTIFS(Itp![Point], "P-05", Itp![Flag], "Overdue") = 1 AND COUNTIFS(Itp![Point], "P-06", Itp![Flag], "Overdue") = 1 AND COUNTIFS(Itp![Point], "P-07", Itp![Flag], "Overdue") = 0 AND COUNTIFS(Itp![Point], "P-09", Itp![Flag], "Overdue") = 0 AND COUNTIFS(Itp![Point], "P-10", Itp![Flag], "Overdue") = 0 AND COUNTIFS(Itp![Point], "P-11", Itp![Flag], "Overdue") = 1 AND COUNTIFS(Itp![Point], "P-12", Itp![Flag], "Overdue") = 0 AND COUNTIF(Itp![Flag], "Note") = 1 AND COUNTA(Itp![Flag]) = 6
TEST     fixture=insp_01.xlsx call=modMain.RunHeadless(Nde, Itp, Rejects, MethodSummary, InspSummary) assert=ERROR_COUNT = 0 AND VALUE(Rejects![Weld Seam No]2) = "S-03" AND VALUE(Rejects![Method]2) = "RT" AND VALUE(Rejects![Weld Seam No]3) = "S-05" AND VALUE(Rejects![Result]3) = " reject " AND VALUE(Rejects![Weld Seam No]4) = "S-02" AND VALUE(Rejects![Method]5) = "MT" AND VALUE(Rejects![Weld Seam No]6) = "S-12" AND VALUE(Rejects![Weld Seam No]7) = "S-07" AND VALUE(Rejects![Method]7) = "RT" AND COUNTA(Rejects![Report Date]) = 6 AND COUNTIF(Rejects![Result], "Accept") = 0 AND COUNTIF(Rejects![Result], "Repair") = 0
TEST     fixture=insp_01.xlsx call=modMain.RunHeadless(Nde, Itp, Rejects, MethodSummary, InspSummary) assert=ERROR_COUNT = 0 AND VALUE(MethodSummary![Method]2) = "MT" AND VALUE(MethodSummary![Accept]2) = 0 AND VALUE(MethodSummary![Repair]2) = 1 AND VALUE(MethodSummary![Reject]2) = 1 AND VALUE(MethodSummary![Method]3) = "PT" AND VALUE(MethodSummary![Accept]3) = 1 AND VALUE(MethodSummary![Repair]3) = 0 AND VALUE(MethodSummary![Reject]3) = 0 AND VALUE(MethodSummary![Method]4) = "RT" AND VALUE(MethodSummary![Accept]4) = 2 AND VALUE(MethodSummary![Repair]4) = 0 AND VALUE(MethodSummary![Reject]4) = 2 AND VALUE(MethodSummary![Method]5) = "UT" AND VALUE(MethodSummary![Accept]5) = 1 AND VALUE(MethodSummary![Repair]5) = 1 AND VALUE(MethodSummary![Reject]5) = 2
TEST     fixture=insp_01.xlsx call=modMain.RunHeadless(Nde, Itp, Rejects, MethodSummary, InspSummary) assert=ERROR_COUNT = 0 AND VALUE(InspSummary![Measure]2) = "Rejects" AND VALUE(InspSummary![Value]2) = 6 AND VALUE(InspSummary![Measure]3) = "Overdue points" AND VALUE(InspSummary![Value]3) = 4 AND VALUE(InspSummary![Measure]4) = "Methods" AND VALUE(InspSummary![Value]4) = 4
TEST     fixture=insp_01.xlsx call=modMain.RunHeadless(Nde, Itp, Rejects, MethodSummary, InspSummary) assert=ERROR_COUNT = 0 AND COUNTA(Nde![Weld Seam No]) = PRE_COUNTA(Nde![Weld Seam No]) AND COUNTA(Nde![Method]) = PRE_COUNTA(Nde![Method]) AND COUNTA(Nde![Result]) = PRE_COUNTA(Nde![Result]) AND COUNTA(Nde![Report Date]) = PRE_COUNTA(Nde![Report Date]) AND COUNTA(Itp![ITP No]) = PRE_COUNTA(Itp![ITP No]) AND COUNTA(Itp![Point]) = PRE_COUNTA(Itp![Point]) AND COUNTA(Itp![Inspection Status]) = PRE_COUNTA(Itp![Inspection Status]) AND COUNTA(Itp![Planned Date]) = PRE_COUNTA(Itp![Planned Date])
TEST     fixture=insp_02.xlsx call=modMain.RunHeadless(Nde, Itp, Rejects, MethodSummary, InspSummary) assert=RESULT = 1 AND ERROR_COUNT = 0 AND VALUE(Itp![Flag]1) = "Flag" AND COUNTA(Itp!E:E) = 2 AND COUNTIF(Itp![Flag], "Overdue") = 1 AND COUNTIFS(Itp![Point], "P-A1", Itp![Flag], "Overdue") = 1 AND ROWS_USED(Rejects) = 2 AND VALUE(Rejects![Weld Seam No]2) = "S-A1" AND ROWS_USED(MethodSummary) = 3 AND VALUE(MethodSummary![Method]2) = "RT" AND VALUE(MethodSummary![Method]3) = "UT" AND VALUE(MethodSummary![Repair]3) = 1 AND VALUE(InspSummary![Value]2) = 1 AND VALUE(InspSummary![Value]3) = 1 AND VALUE(InspSummary![Value]4) = 2
TEST     fixture=insp_01.xlsx call=modMain.RunHeadless(Itp, Nde, Rejects, MethodSummary, InspSummary) assert=RESULT = -1 AND ERROR_COUNT = 1
---
VCL 0.6
TASK
GOAL     Start the inspection follow-up on the workbook's own sheets and tell the user how many rejects were listed.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "NDE Report"), SHEET(THIS_BOOK, "ITP Tracker"), SHEET(THIS_BOOK, "Rejects"), SHEET(THIS_BOOK, "Method Summary"), SHEET(THIS_BOOK, "Inspection Summary"), into=n)
OP       MSG("Rejects listed: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run every step of the inspection follow-up on the five sheets it is given and return how many rejects were listed.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     nde As Worksheet, itp As Worksheet, rejects As Worksheet, methodSum As Worksheet, inspSum As Worksheet
RET      Long
VAR      nRejects As Long
VAR      nOverdue As Long
VAR      nMethods As Long
PRE      COL($nde, "Weld Seam No") > 0; COL($nde, "Method") > 0; COL($nde, "Result") > 0; COL($nde, "Report Date") > 0; COL($itp, "Point") > 0; COL($itp, "Inspection Status") > 0; COL($itp, "Planned Date") > 0
MODIFIES $itp
MODIFIES $rejects
MODIFIES $methodSum
MODIFIES $inspSum
OP       CALL(modRejects.ListRejects, $nde, $rejects, into=nRejects)
OP       CALL(modTracker.FlagOverdue, $itp, into=nOverdue)
OP       CALL(modSummary.WriteMethodSummary, $nde, $methodSum, into=nMethods)
OP       CALL(modSummary.WriteInspectionSummary, $inspSum, nRejects, nOverdue, nMethods)
OP       RETURN(nRejects)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Rejects sheet with every NDE row whose result is Reject, newest report first, and return how many were listed.
PROC     ListRejects
KIND     Function
MODULE   modRejects
ARGS     nde As Worksheet, out As Worksheet
RET      Long
VAR      cSeam As Long
VAR      cMethod As Long
VAR      cResult As Long
VAR      cDate As Long
VAR      outRow As Long
VAR      listed As Long
PRE      COL($nde, "Weld Seam No") > 0; COL($nde, "Method") > 0; COL($nde, "Result") > 0; COL($nde, "Report Date") > 0
MODIFIES $out
OP       LET(cSeam, COL($nde, "Weld Seam No"))
OP       LET(cMethod, COL($nde, "Method"))
OP       LET(cResult, COL($nde, "Result"))
OP       LET(cDate, COL($nde, "Report Date"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Weld Seam No", "Method", "Result", "Report Date")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($nde, cResult) var=r
OP         IF SAME(CELL($nde, r, cResult), "Reject")
OP           WRITE_ROW($out, outRow, CELL($nde, r, cSeam), CELL($nde, r, cMethod), CELL($nde, r, cResult), CELL($nde, r, cDate))
OP           LET(outRow, outRow + 1)
OP           LET(listed, listed + 1)
OP       SORT(USED_RANGE($out), key=4, order=desc, header=yes)
OP       RETURN(listed)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write Overdue in the Flag column for every point that is not Closed and is planned before today, and clear Flag on the others.
PROC     FlagOverdue
KIND     Function
MODULE   modTracker
ARGS     itp As Worksheet
RET      Long
VAR      cPoint As Long
VAR      cStatus As Long
VAR      cPlanned As Long
VAR      cFlag As Long
VAR      late As Boolean
VAR      flagged As Long
PRE      COL($itp, "Point") > 0; COL($itp, "Inspection Status") > 0; COL($itp, "Planned Date") > 0
MODIFIES $itp
OP       LET(cPoint, COL($itp, "Point"))
OP       LET(cStatus, COL($itp, "Inspection Status"))
OP       LET(cPlanned, COL($itp, "Planned Date"))
OP       LET(cFlag, COL($itp, "Flag"))
OP       IF cFlag = 0
OP         LET(cFlag, LASTCOL($itp, 1) + 1)
OP         SET_VALUE(CELL($itp, 1, cFlag), "Flag")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($itp, cPoint) var=r
OP         SKIP_IF IS_EMPTY(CELL($itp, r, cPoint))
OP         LET(late, FALSE)
OP         IF NOT SAME(CELL($itp, r, cStatus), "Closed")
OP           IF IS_DATE(CELL($itp, r, cPlanned))
OP             IF CELL($itp, r, cPlanned) < TODAY
OP               LET(late, TRUE)
OP         IF late
OP           SET_VALUE(CELL($itp, r, cFlag), "Overdue")
OP           LET(flagged, flagged + 1)
OP         ELSE
OP           CLEAR(CELL($itp, r, cFlag), what=contents)
OP       RETURN(flagged)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the method summary with each NDE method and how many of its results are Accept, Repair and Reject, A to Z.
PROC     WriteMethodSummary
KIND     Function
MODULE   modSummary
ARGS     nde As Worksheet, out As Worksheet
RET      Long
VAR      cMethod As Long
VAR      cResult As Long
VAR      accCounts As Dictionary
VAR      repCounts As Dictionary
VAR      rejCounts As Dictionary
VAR      nAccept As Long
VAR      nRepair As Long
VAR      nReject As Long
VAR      outRow As Long
PRE      COL($nde, "Method") > 0; COL($nde, "Result") > 0
MODIFIES $out
OP       LET(cMethod, COL($nde, "Method"))
OP       LET(cResult, COL($nde, "Result"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Method", "Accept", "Repair", "Reject")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($nde, cMethod) var=r
OP         SKIP_IF IS_EMPTY(CELL($nde, r, cMethod))
OP         DICT_ADD(accCounts, TRIM(CELL($nde, r, cMethod)), IIF(SAME(CELL($nde, r, cResult), "Accept"), 1, 0), mode=sum)
OP         DICT_ADD(repCounts, TRIM(CELL($nde, r, cMethod)), IIF(SAME(CELL($nde, r, cResult), "Repair"), 1, 0), mode=sum)
OP         DICT_ADD(rejCounts, TRIM(CELL($nde, r, cMethod)), IIF(SAME(CELL($nde, r, cResult), "Reject"), 1, 0), mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(accCounts)
OP         DICT_GET(accCounts, k, into=nAccept)
OP         DICT_GET(repCounts, k, into=nRepair)
OP         DICT_GET(rejCounts, k, into=nReject)
OP         WRITE_ROW($out, outRow, k, nAccept, nRepair, nReject)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=1, order=asc, header=yes)
OP       RETURN(COUNT(accCounts))
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the inspection summary sheet with labelled counts of rejects, overdue points and methods.
PROC     WriteInspectionSummary
KIND     Sub
MODULE   modSummary
ARGS     out As Worksheet, rejected As Long, overdue As Long, methodCount As Long
MODIFIES $out
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Measure", "Value")
OP       WRITE_ROW($out, 2, "Rejects", $rejected)
OP       WRITE_ROW($out, 3, "Overdue points", $overdue)
OP       WRITE_ROW($out, 4, "Methods", $methodCount)
ERR      standard
