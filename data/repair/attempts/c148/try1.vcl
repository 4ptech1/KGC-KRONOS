VCL 0.6
BUILD
# ASSUME: ITP Tracker headers (row 1) include "ITP No", "Point", "Weld Seam No", "Inspection Status" and "Flag"
# ASSUME: NDE Report has "Weld Seam No" and "Result"; NCR Register has "NCR No", "ITP No" and "Status"
# ASSUME: the real sheets are "ITP Tracker", "NDE Report", "NCR Register", "Flagged Points" and "NCR Count"; the TESTs use Itp, Nde, Ncr, Flagged and Counts
# ASSUME: Flagged Points and NCR Count already exist and are rebuilt on every run
# ASSUME: IDs are trimmed and compared ignoring case; a Result of exactly Reject (any case) is a reject; an NCR with a blank Status is not Closed
# ASSUME: a point with a blank ITP No is skipped and its Flag is left alone; blank Weld Seam No or NCR ITP No never match
# ASSUME: NCR Count lists every ITP No that has at least one NCR not Closed, most first
GOAL     Flag ITP points with an NDE reject or an open NCR, list the flagged points sorted by ITP No, and count open NCRs per ITP No.
NAME     WeeklyInspection
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modSteps
SIG      modMain.Run()
SIG      modMain.RunFlagPoints()
SIG      modMain.RunListFlagged()
SIG      modMain.RunCountNcrs()
SIG      modMain.RunHeadless(itp As Worksheet, nde As Worksheet, ncr As Worksheet, flagged As Worksheet, counts As Worksheet) As Long
SIG      modSteps.FlagPoints(itp As Worksheet, nde As Worksheet, ncr As Worksheet) As Long
SIG      modSteps.ListFlagged(itp As Worksheet, out As Worksheet)
SIG      modSteps.CountNcrs(ncr As Worksheet, out As Worksheet)
FLOW     modMain.Run -> modMain.RunHeadless -> modSteps.FlagPoints
FLOW     modMain.RunHeadless -> modSteps.ListFlagged
FLOW     modMain.RunHeadless -> modSteps.CountNcrs
FLOW     modMain.RunFlagPoints -> modSteps.FlagPoints
FLOW     modMain.RunListFlagged -> modSteps.ListFlagged
FLOW     modMain.RunCountNcrs -> modSteps.CountNcrs
FIXTURE  week_01.xlsx sheet=Itp cols=A:ITP No,B:Point,C:Weld Seam No,D:Inspection Status,E:Flag rows=8 data=A:values("ITP-3"x1,"ITP-1"x1,"ITP-2"x1,"ITP-8"x1,""x1,"ITP-4"x1,"ITP-5"x1,"ITP-6"x1)+boundary("ITP-3","ITP-1","ITP-2","ITP-8","","ITP-4","ITP-5","ITP-6");B:seq(1,1);C:values("W-01"x2,"W-02"x1,"W-03"x1," w-04 "x1,"W-05"x1,""x1,"W-06"x1)+boundary("W-01","W-02","W-03"," w-04 ","W-01","W-05","","W-06");D:values("Done"x5,"Pending"x3);E:values("old"x8)
FIXTURE  week_01.xlsx sheet=Nde cols=A:Weld Seam No,B:Result rows=8 data=A:values("W-01"x2,"W-02"x1,"W-03"x1,"W-04"x1,"W-05"x1,"W-06"x1,""x1)+boundary("W-01","W-02","W-03","W-04","W-05","W-06","W-01","");B:values("Reject"x3,"Pass"x3,"Rejected"x1," reject "x1)+boundary("Reject","Pass","Reject"," reject ","Pass","Rejected","Pass","Reject")
FIXTURE  week_01.xlsx sheet=Ncr cols=A:NCR No,B:ITP No,C:Status rows=10 data=A:seq(1,1);B:values("ITP-1"x2,"ITP-2"x3,"ITP-4"x2,"ITP-7"x1,"ITP-3"x1,""x1)+boundary("ITP-1","ITP-2","ITP-2","ITP-4","ITP-4","","ITP-7"," itp-1 ","ITP-2","ITP-3");C:values("Open"x5,"Closed"x3,"In Review"x1,""x1)+boundary("Open","Open","In Review","Closed"," closed ","Open","","Open","Open","Closed")
FIXTURE  week_01.xlsx sheet=Flagged cols=A:ITP No,B:Point,C:Weld Seam No,D:Inspection Status,E:Flag rows=0 stale=2
FIXTURE  week_01.xlsx sheet=Counts cols=A:ITP No,B:NCR Count rows=0 stale=2
FIXTURE  quiet_01.xlsx sheet=Itp cols=A:ITP No,B:Point,C:Weld Seam No,D:Inspection Status,E:Flag rows=3 data=A:values("ITP-1"x3);B:seq(1,1);C:values("W-01"x3);D:values("Done"x3);E:values("old"x3)
FIXTURE  quiet_01.xlsx sheet=Nde cols=A:Weld Seam No,B:Result rows=2 data=A:values("W-01"x2);B:values("Pass"x2)
FIXTURE  quiet_01.xlsx sheet=Ncr cols=A:NCR No,B:ITP No,C:Status rows=2 data=A:seq(1,1);B:values("ITP-1"x2);C:values("Closed"x2)
FIXTURE  quiet_01.xlsx sheet=Flagged cols=A:ITP No,B:Point,C:Weld Seam No,D:Inspection Status,E:Flag rows=0 stale=2
FIXTURE  quiet_01.xlsx sheet=Counts cols=A:ITP No,B:NCR Count rows=0 stale=2
TEST     fixture=week_01.xlsx call=modMain.RunHeadless(Itp, Nde, Ncr, Flagged, Counts) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTIF(Itp![Flag], "NDE reject") = 3 AND COUNTIF(Itp![Flag], "Open NCR") = 1 AND COUNTIF(Itp![Flag], "old") = 1 AND COUNTA(Itp![Flag]) = 6 AND ROWS_USED(Flagged) = 5 AND COUNTIF(Flagged![ITP No], "stale") = 0 AND ROWS_USED(Counts) = 4 AND COUNTIF(Counts![ITP No], "stale") = 0
TEST     fixture=week_01.xlsx call=modMain.RunHeadless(Itp, Nde, Ncr, Flagged, Counts) assert=RESULT = 4 AND COUNTA(Itp![ITP No]) = PRE_COUNTA(Itp![ITP No]) AND COUNTA(Nde![Weld Seam No]) = PRE_COUNTA(Nde![Weld Seam No]) AND COUNTA(Ncr![NCR No]) = PRE_COUNTA(Ncr![NCR No]) AND VALUE(Flagged![ITP No]2) = "ITP-1" AND VALUE(Flagged![ITP No]3) = "ITP-2" AND VALUE(Flagged![ITP No]4) = "ITP-3" AND VALUE(Flagged![ITP No]5) = "ITP-8" AND VALUE(Flagged![Flag]2) = "Open NCR" AND VALUE(Flagged![Flag]3) = "NDE reject" AND VALUE(Flagged![Flag]4) = "NDE reject" AND VALUE(Flagged![Flag]5) = "NDE reject" AND VALUE(Flagged![Point]2) = 2 AND VALUE(Flagged![Point]5) = 4
TEST     fixture=week_01.xlsx call=modMain.RunHeadless(Itp, Nde, Ncr, Flagged, Counts) assert=VALUE(Counts![ITP No]2) = "ITP-2" AND VALUE(Counts![NCR Count]2) = 3 AND VALUE(Counts![ITP No]3) = "ITP-1" AND VALUE(Counts![NCR Count]3) = 2 AND VALUE(Counts![ITP No]4) = "ITP-7" AND VALUE(Counts![NCR Count]4) = 1
TEST     fixture=quiet_01.xlsx call=modMain.RunHeadless(Itp, Nde, Ncr, Flagged, Counts) assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTA(Itp![Flag]) = 1 AND ROWS_USED(Flagged) = 1 AND ROWS_USED(Counts) = 1
TEST     fixture=week_01.xlsx call=modMain.RunHeadless(Nde, Itp, Ncr, Flagged, Counts) assert=RESULT = -1 AND ERROR_COUNT = 1
---
VCL 0.6
TASK
GOAL     Start the weekly inspection on the workbook's own sheets and tell the user how many points were flagged.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "ITP Tracker"), SHEET(THIS_BOOK, "NDE Report"), SHEET(THIS_BOOK, "NCR Register"), SHEET(THIS_BOOK, "Flagged Points"), SHEET(THIS_BOOK, "NCR Count"), into=n)
OP       MSG("Points flagged: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run only step 1 on the workbook's own sheets: flag the ITP points and tell the user how many were flagged.
PROC     RunFlagPoints
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modSteps.FlagPoints, SHEET(THIS_BOOK, "ITP Tracker"), SHEET(THIS_BOOK, "NDE Report"), SHEET(THIS_BOOK, "NCR Register"), into=n)
OP       MSG("Points flagged: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run only step 2 on the workbook's own sheets: rebuild the Flagged Points sheet from the current Flag column.
PROC     RunListFlagged
KIND     Sub
MODULE   modMain
OP       CALL(modSteps.ListFlagged, SHEET(THIS_BOOK, "ITP Tracker"), SHEET(THIS_BOOK, "Flagged Points"))
ERR      standard
---
VCL 0.6
TASK
GOAL     Run only step 3 on the workbook's own sheets: rebuild the NCR Count sheet from the NCR Register.
PROC     RunCountNcrs
KIND     Sub
MODULE   modMain
OP       CALL(modSteps.CountNcrs, SHEET(THIS_BOOK, "NCR Register"), SHEET(THIS_BOOK, "NCR Count"))
ERR      standard
---
VCL 0.6
TASK
GOAL     Run all three steps on the sheets it is given and return how many ITP points were flagged.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     itp As Worksheet, nde As Worksheet, ncr As Worksheet, flagged As Worksheet, counts As Worksheet
RET      Long
VAR      n As Long
PRE      COL($itp, "ITP No") > 0; COL($itp, "Point") > 0; COL($itp, "Weld Seam No") > 0; COL($itp, "Inspection Status") > 0; COL($itp, "Flag") > 0; COL($nde, "Weld Seam No") > 0; COL($nde, "Result") > 0; COL($ncr, "ITP No") > 0; COL($ncr, "Status") > 0
MODIFIES $itp["Flag"]
MODIFIES $flagged
MODIFIES $counts
OP       CALL(modSteps.FlagPoints, $itp, $nde, $ncr, into=n)
OP       CALL(modSteps.ListFlagged, $itp, $flagged)
OP       CALL(modSteps.CountNcrs, $ncr, $counts)
OP       RETURN(n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write NDE reject or Open NCR in the Flag column of every ITP point that needs it, clear it on the others, and return how many were flagged.
PROC     FlagPoints
KIND     Function
MODULE   modSteps
ARGS     itp As Worksheet, nde As Worksheet, ncr As Worksheet
RET      Long
VAR      cItp As Long
VAR      cSeam As Long
VAR      cFlag As Long
VAR      cNdeSeam As Long
VAR      cNdeRes As Long
VAR      cNcrItp As Long
VAR      cNcrStat As Long
VAR      rejects As Dictionary
VAR      openNcr As Dictionary
VAR      flagged As Long
PRE      COL($itp, "ITP No") > 0; COL($itp, "Weld Seam No") > 0; COL($itp, "Flag") > 0; COL($nde, "Weld Seam No") > 0; COL($nde, "Result") > 0; COL($ncr, "ITP No") > 0; COL($ncr, "Status") > 0
MODIFIES $itp["Flag"]
OP       LET(cItp, COL($itp, "ITP No"))
OP       LET(cSeam, COL($itp, "Weld Seam No"))
OP       LET(cFlag, COL($itp, "Flag"))
OP       LET(cNdeSeam, COL($nde, "Weld Seam No"))
OP       LET(cNdeRes, COL($nde, "Result"))
OP       LET(cNcrItp, COL($ncr, "ITP No"))
OP       LET(cNcrStat, COL($ncr, "Status"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($nde, cNdeSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($nde, r, cNdeSeam))
OP         IF SAME(CELL($nde, r, cNdeRes), "Reject")
OP           DICT_ADD(rejects, TRIM(CELL($nde, r, cNdeSeam)), 1, mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ncr, cNcrItp) var=r
OP         SKIP_IF IS_EMPTY(CELL($ncr, r, cNcrItp))
OP         IF NOT SAME(CELL($ncr, r, cNcrStat), "Closed")
OP           DICT_ADD(openNcr, TRIM(CELL($ncr, r, cNcrItp)), 1, mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($itp, cItp) var=r
OP         SKIP_IF IS_EMPTY(CELL($itp, r, cItp))
OP         IF EXISTS(rejects, TRIM(CELL($itp, r, cSeam)))
OP           SET_VALUE(CELL($itp, r, cFlag), "NDE reject")
OP           LET(flagged, flagged + 1)
OP         ELSE
OP           IF EXISTS(openNcr, TRIM(CELL($itp, r, cItp)))
OP             SET_VALUE(CELL($itp, r, cFlag), "Open NCR")
OP             LET(flagged, flagged + 1)
OP           ELSE
OP             CLEAR(CELL($itp, r, cFlag), what=contents)
OP       RETURN(flagged)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Flagged Points sheet with every point flagged NDE reject or Open NCR, using the same headings, sorted by ITP No A to Z.
PROC     ListFlagged
KIND     Sub
MODULE   modSteps
ARGS     itp As Worksheet, out As Worksheet
VAR      cItp As Long
VAR      cPoint As Long
VAR      cSeam As Long
VAR      cStat As Long
VAR      cFlag As Long
VAR      outRow As Long
PRE      COL($itp, "ITP No") > 0; COL($itp, "Point") > 0; COL($itp, "Weld Seam No") > 0; COL($itp, "Inspection Status") > 0; COL($itp, "Flag") > 0
MODIFIES $out
OP       LET(cItp, COL($itp, "ITP No"))
OP       LET(cPoint, COL($itp, "Point"))
OP       LET(cSeam, COL($itp, "Weld Seam No"))
OP       LET(cStat, COL($itp, "Inspection Status"))
OP       LET(cFlag, COL($itp, "Flag"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "ITP No", "Point", "Weld Seam No", "Inspection Status", "Flag")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($itp, cItp) var=r
OP         IF SAME(CELL($itp, r, cFlag), "NDE reject") OR SAME(CELL($itp, r, cFlag), "Open NCR")
OP           WRITE_ROW($out, outRow, CELL($itp, r, cItp), CELL($itp, r, cPoint), CELL($itp, r, cSeam), CELL($itp, r, cStat), CELL($itp, r, cFlag))
OP           LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=1, order=asc, header=yes)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the NCR Count sheet with each ITP No and its number of NCRs that are not Closed, the most first.
PROC     CountNcrs
KIND     Sub
MODULE   modSteps
ARGS     ncr As Worksheet, out As Worksheet
VAR      cItp As Long
VAR      cStat As Long
VAR      counts As Dictionary
VAR      got As Long
VAR      outRow As Long
PRE      COL($ncr, "ITP No") > 0; COL($ncr, "Status") > 0
MODIFIES $out
OP       LET(cItp, COL($ncr, "ITP No"))
OP       LET(cStat, COL($ncr, "Status"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "ITP No", "NCR Count")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ncr, cItp) var=r
OP         SKIP_IF IS_EMPTY(CELL($ncr, r, cItp))
OP         IF NOT SAME(CELL($ncr, r, cStat), "Closed")
OP           DICT_ADD(counts, TRIM(CELL($ncr, r, cItp)), 1, mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(counts)
OP         DICT_GET(counts, k, into=got)
OP         WRITE_ROW($out, outRow, k, got)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=2, order=desc, header=yes)
ERR      standard
