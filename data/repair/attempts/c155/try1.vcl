VCL 0.6
BUILD
# ASSUME: Weld Log headers (row 1) include "Weld Seam No", "Welder ID" and "NDE Status"; NDE Report headers include "Weld Seam No", "Method" and "Result"
# ASSUME: Seam numbers are matched ignoring case and stray spaces; a Result of Reject is matched ignoring case and spaces
# ASSUME: Weld Log rows with a blank Weld Seam No are left alone; welds with a blank Welder ID are not counted per welder
# ASSUME: the Repair List and Welder Rejects sheets exist and are rebuilt on every run; the real sheets are "Weld Log", "NDE Report", "Repair List" and "Welder Rejects", the TESTs use Wlog, Nde, Repair and Rej
# ASSUME: the lists are built from the NDE Status the first step has just written; the three steps run together in one macro, because a BUILD cannot start a step on its own from a sheet argument
GOAL     Set each weld's NDE status from the NDE report, list the rejected welds, count rejected welds per welder and report how many welds are not tested.
NAME     WeldNdeFollowUp
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modSteps
SIG      modMain.Run()
SIG      modMain.RunHeadless(wlog As Worksheet, nde As Worksheet, repair As Worksheet, rej As Worksheet) As Long
SIG      modSteps.SetNdeStatus(wlog As Worksheet, nde As Worksheet) As Long
SIG      modSteps.ListRepairs(wlog As Worksheet, repair As Worksheet) As Long
SIG      modSteps.CountRejects(wlog As Worksheet, rej As Worksheet) As Long
FLOW     modMain.Run -> modMain.RunHeadless -> modSteps.SetNdeStatus
FLOW     modMain.RunHeadless -> modSteps.ListRepairs
FLOW     modMain.RunHeadless -> modSteps.CountRejects
FIXTURE  weld_01.xlsx sheet=Wlog cols=A:Weld Seam No,B:Welder ID,C:NDE Status rows=11 data=A:values("W-1"x1,"W-2"x1,"W-3"x1,"W-4"x1,"W-5"x1,""x1,"W-7"x1,"W-8"x1,"W-9"x1,"W-10"x1,"W-11"x1)+boundary("W-1","W-2","W-3","W-4","W-5","","W-7","W-8","W-9","W-10","W-11");B:values("A1"x3,"B2"x3,"C3"x2,"c3"x1,""x1," b2 "x1)+boundary("A1","B2","A1","C3","A1","B2","B2","c3",""," b2 ","C3");C:values("old"x11)
FIXTURE  weld_01.xlsx sheet=Nde cols=A:Weld Seam No,B:Method,C:Result rows=12 data=A:values("W-1"x1,"w-1"x1," W-2 "x1,"W-3"x1,"w-5"x1,"W-5"x1,"W-7"x1,"W-9"x1,"W-10"x1,"W-11"x1,""x1,"W-99"x1)+boundary("W-1","w-1"," W-2 ","W-3","w-5","W-5","W-7","W-9","W-10","W-11","","W-99");B:text(4);C:values("Pass"x5,"Reject"x6," REJECT "x1)+boundary("Pass","Reject","Pass"," REJECT ","Reject","Pass","Reject","Reject","Reject","Pass","Reject","Pass")
FIXTURE  weld_01.xlsx sheet=Repair cols=A:Weld Seam No,B:Welder ID,C:NDE Status rows=0 stale=3
FIXTURE  weld_01.xlsx sheet=Rej cols=A:Welder ID,B:Rejected rows=0 stale=3
TEST     fixture=weld_01.xlsx call=modMain.RunHeadless(Wlog, Nde, Repair, Rej) runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND COUNTIF(Wlog![NDE Status], "Rejected") = 6 AND COUNTIF(Wlog![NDE Status], "Accepted") = 2 AND COUNTIF(Wlog![NDE Status], "Not tested") = 2 AND COUNTIF(Wlog![NDE Status], "old") = 1 AND ROWS_USED(Repair) = 7 AND ROWS_USED(Rej) = 3 AND COUNTIF(Repair![NDE Status], "Rejected") = 6 AND COUNTIF(Repair![Weld Seam No], "stale") = 0 AND COUNTIF(Rej![Welder ID], "stale") = 0
TEST     fixture=weld_01.xlsx call=modMain.RunHeadless(Wlog, Nde, Repair, Rej) assert=COUNTA(Wlog![Weld Seam No]) = PRE_COUNTA(Wlog![Weld Seam No]) AND COUNTA(Nde![Result]) = PRE_COUNTA(Nde![Result]) AND VALUE(Wlog![NDE Status]2) = "Rejected" AND VALUE(Wlog![NDE Status]3) = "Accepted" AND VALUE(Wlog![NDE Status]5) = "Not tested" AND VALUE(Rej![Welder ID]2) = "A1" AND VALUE(Rej![Rejected]2) = 3 AND VALUE(Rej![Rejected]3) = 2 AND VALUE(Repair![Weld Seam No]2) = "W-1" AND VALUE(Repair![Weld Seam No]3) = "W-10" AND VALUE(Repair![Weld Seam No]7) = "W-9"
TEST     fixture=weld_01.xlsx call=modMain.RunHeadless(Nde, Wlog, Repair, Rej) assert=RESULT = -1 AND ERROR_COUNT = 1
---
VCL 0.6
TASK
GOAL     Start the weld NDE follow-up on the workbook's own sheets and tell the user how many welds are not tested.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "Weld Log"), SHEET(THIS_BOOK, "NDE Report"), SHEET(THIS_BOOK, "Repair List"), SHEET(THIS_BOOK, "Welder Rejects"), into=n)
OP       MSG("Welds Not tested: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run all three steps on the sheets it is given and return how many welds are not tested.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     wlog As Worksheet, nde As Worksheet, repair As Worksheet, rej As Worksheet
RET      Long
VAR      n As Long
VAR      listed As Long
VAR      welders As Long
PRE      COL($wlog, "Weld Seam No") > 0; COL($wlog, "Welder ID") > 0; COL($wlog, "NDE Status") > 0; COL($nde, "Weld Seam No") > 0; COL($nde, "Result") > 0
MODIFIES $wlog["NDE Status"]
MODIFIES $repair
MODIFIES $rej
OP       CALL(modSteps.SetNdeStatus, $wlog, $nde, into=n)
OP       CALL(modSteps.ListRepairs, $wlog, $repair, into=listed)
OP       CALL(modSteps.CountRejects, $wlog, $rej, into=welders)
OP       RETURN(n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write Rejected, Accepted or Not tested in NDE Status of every weld from the NDE report and return how many welds are Not tested.
PROC     SetNdeStatus
KIND     Function
MODULE   modSteps
ARGS     wlog As Worksheet, nde As Worksheet
RET      Long
VAR      cSeam As Long
VAR      cStatus As Long
VAR      nSeam As Long
VAR      nResult As Long
VAR      tested As Dictionary
VAR      rejected As Dictionary
VAR      n As Long
PRE      COL($wlog, "Weld Seam No") > 0; COL($wlog, "NDE Status") > 0; COL($nde, "Weld Seam No") > 0; COL($nde, "Result") > 0
MODIFIES $wlog["NDE Status"]
OP       LET(cSeam, COL($wlog, "Weld Seam No"))
OP       LET(cStatus, COL($wlog, "NDE Status"))
OP       LET(nSeam, COL($nde, "Weld Seam No"))
OP       LET(nResult, COL($nde, "Result"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($nde, nSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($nde, r, nSeam))
OP         DICT_ADD(tested, CLEAN_TEXT(CELL($nde, r, nSeam)), 1, mode=keep)
OP         IF SAME(CELL($nde, r, nResult), "Reject")
OP           DICT_ADD(rejected, CLEAN_TEXT(CELL($nde, r, nSeam)), 1, mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($wlog, cSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($wlog, r, cSeam))
OP         IF EXISTS(rejected, CLEAN_TEXT(CELL($wlog, r, cSeam)))
OP           SET_VALUE(CELL($wlog, r, cStatus), "Rejected")
OP         ELSE
OP           IF EXISTS(tested, CLEAN_TEXT(CELL($wlog, r, cSeam)))
OP             SET_VALUE(CELL($wlog, r, cStatus), "Accepted")
OP           ELSE
OP             SET_VALUE(CELL($wlog, r, cStatus), "Not tested")
OP             LET(n, n + 1)
OP       RETURN(n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Repair List sheet with the rejected welds and the same headings, sorted by weld seam number A to Z, and return how many were listed.
PROC     ListRepairs
KIND     Function
MODULE   modSteps
ARGS     wlog As Worksheet, repair As Worksheet
RET      Long
VAR      cSeam As Long
VAR      cWelder As Long
VAR      cStatus As Long
VAR      outRow As Long
VAR      n As Long
PRE      COL($wlog, "Weld Seam No") > 0; COL($wlog, "Welder ID") > 0; COL($wlog, "NDE Status") > 0
MODIFIES $repair
OP       LET(cSeam, COL($wlog, "Weld Seam No"))
OP       LET(cWelder, COL($wlog, "Welder ID"))
OP       LET(cStatus, COL($wlog, "NDE Status"))
OP       CLEAR(USED_RANGE($repair), what=all)
OP       WRITE_ROW($repair, 1, "Weld Seam No", "Welder ID", "NDE Status")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($wlog, cSeam) var=r
OP         SKIP_IF NOT SAME(CELL($wlog, r, cStatus), "Rejected")
OP         WRITE_ROW($repair, outRow, CELL($wlog, r, cSeam), CELL($wlog, r, cWelder), CELL($wlog, r, cStatus))
OP         LET(outRow, outRow + 1)
OP         LET(n, n + 1)
OP       SORT(USED_RANGE($repair), key=1, order=asc, header=yes)
OP       RETURN(n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Welder Rejects sheet with each welder ID and its number of rejected welds, the most first, and return how many welders were listed.
PROC     CountRejects
KIND     Function
MODULE   modSteps
ARGS     wlog As Worksheet, rej As Worksheet
RET      Long
VAR      cSeam As Long
VAR      cWelder As Long
VAR      cStatus As Long
VAR      counts As Dictionary
VAR      got As Long
VAR      outRow As Long
VAR      n As Long
PRE      COL($wlog, "Weld Seam No") > 0; COL($wlog, "Welder ID") > 0; COL($wlog, "NDE Status") > 0
MODIFIES $rej
OP       LET(cSeam, COL($wlog, "Weld Seam No"))
OP       LET(cWelder, COL($wlog, "Welder ID"))
OP       LET(cStatus, COL($wlog, "NDE Status"))
OP       CLEAR(USED_RANGE($rej), what=all)
OP       WRITE_ROW($rej, 1, "Welder ID", "Rejected")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($wlog, cSeam) var=r
OP         SKIP_IF NOT SAME(CELL($wlog, r, cStatus), "Rejected")
OP         SKIP_IF IS_EMPTY(CELL($wlog, r, cWelder))
OP         DICT_ADD(counts, CLEAN_TEXT(CELL($wlog, r, cWelder)), 1, mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(counts)
OP         DICT_GET(counts, k, into=got)
OP         WRITE_ROW($rej, outRow, k, got)
OP         LET(outRow, outRow + 1)
OP         LET(n, n + 1)
OP       SORT(USED_RANGE($rej), key=2, order=desc, header=yes)
OP       RETURN(n)
ERR      standard
