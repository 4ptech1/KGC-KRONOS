VCL 0.6
BUILD
# ASSUME: the real sheets are "Calibration Tracker", "NCR Register", "Overdue Instruments" and "NCR Instruments"; the TESTs use Cal, NCR, Over and Inst
# ASSUME: the two output sheets already exist and are rebuilt, headings included, on every run
# ASSUME: Overdue means the due date is before today; due today up to 14 days ahead is Due soon; a due date that is not a real date is No date
# ASSUME: Instrument IDs are matched after trimming and ignoring case; the Instrument Counts sheet shows them in upper case
# ASSUME: an NCR with a blank Instrument ID gets no flag; Unknown instrument is written for any NCR, Closed or not, whose ID is not in the tracker
# ASSUME: the number reported is the NCRs given either flag; a Closed NCR is never flagged Overdue instrument
# ASSUME: NCRs are counted per instrument whatever their status, and a blank row in the Calibration Tracker's ID column is not an instrument
GOAL     Set each instrument's calibration status, flag NCRs on overdue or unknown instruments, list the overdue instruments by due date and count NCRs per instrument.
NAME     CalibrationNcrCheck
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modSteps
SIG      modMain.Run()
SIG      modMain.RunHeadless(cal As Worksheet, ncr As Worksheet, over As Worksheet, inst As Worksheet) As Long
SIG      modSteps.SetCalStatus(cal As Worksheet)
SIG      modSteps.FlagNcrs(cal As Worksheet, ncr As Worksheet) As Long
SIG      modSteps.ListOverdue(cal As Worksheet, over As Worksheet)
SIG      modSteps.CountByInstrument(ncr As Worksheet, inst As Worksheet)
FLOW     modMain.Run -> modMain.RunHeadless -> modSteps.SetCalStatus
FLOW     modMain.RunHeadless -> modSteps.FlagNcrs
FLOW     modMain.RunHeadless -> modSteps.ListOverdue
FLOW     modMain.RunHeadless -> modSteps.CountByInstrument
FIXTURE  calib_01.xlsx sheet=Cal cols=A:Instrument ID,B:Type,C:Next Calibration Due,D:Status rows=10 data=A:values("G-1"x1,"G-2"x1," g-3 "x1,"G-4"x1,"G-5"x1,"G-6"x1,""x1,"G-8"x1,"G-9"x1,"G-10"x1)+boundary("G-1","G-2"," g-3 ","G-4","G-5","G-6","","G-8","G-9","G-10");B:values("Gauge"x4,"Meter"x3,"Scale"x3)+boundary("Gauge","Meter","Scale","Gauge","Meter","Scale","Gauge","Meter","Scale","Gauge");C:reldates(-30,60)+boundary(-5,-1,0,14,15,"",30,40,-20,-9);D:values("old"x10)
FIXTURE  calib_01.xlsx sheet=NCR cols=A:NCR No,B:Instrument ID,C:Status,D:Cal Flag rows=13 data=A:values("N-1"x1,"N-2"x1,"N-3"x1,"N-4"x1,"N-5"x1,"N-6"x1,"N-7"x1,"N-8"x1,"N-9"x1,"N-10"x1,"N-11"x1,"N-12"x1,"N-13"x1)+boundary("N-1","N-2","N-3","N-4","N-5","N-6","N-7","N-8","N-9","N-10","N-11","N-12","N-13");B:values(" g-1 "x1,"G-1"x3,"G-2"x1,"g-3"x1,"X-9"x1,""x1,"G-9"x3,"G-10"x1,"x-9"x1)+boundary(" g-1 ","G-1","G-2","g-3","X-9","","G-9","G-9","G-10","x-9","G-9","G-1","G-1");C:values("Open"x8,"Closed"x3," closed "x1,"Hold"x1)+boundary("Open","Closed","Open","Open","Open","Open"," closed ","Open","Open","Closed","Open","Hold","Closed");D:values("old"x13)
FIXTURE  calib_01.xlsx sheet=Over cols=A:Instrument ID,B:Type,C:Next Calibration Due,D:Status rows=0 stale=2
FIXTURE  calib_01.xlsx sheet=Inst cols=A:Instrument ID,B:NCR Count rows=0 stale=2
FIXTURE  calib_02.xlsx sheet=Cal cols=A:Instrument ID,B:Type,C:Next Calibration Due,D:Status rows=3 data=A:values("G-1"x1,"G-2"x1,"G-3"x1);B:values("Gauge"x3);C:reldates(20,60)+boundary(15,"",20);D:values("old"x3)
FIXTURE  calib_02.xlsx sheet=NCR cols=A:NCR No,B:Instrument ID,C:Status,D:Cal Flag rows=2 data=A:values("N-1"x1,"N-2"x1);B:values(""x2);C:values("Open"x2);D:values("old"x2)
FIXTURE  calib_02.xlsx sheet=Over cols=A:Instrument ID,B:Type,C:Next Calibration Due,D:Status rows=0 stale=2
FIXTURE  calib_02.xlsx sheet=Inst cols=A:Instrument ID,B:NCR Count rows=0 stale=2
TEST     fixture=calib_01.xlsx call=modMain.RunHeadless(Cal, NCR, Over, Inst) runs=2 assert=RESULT = 8 AND ERROR_COUNT = 0 AND COUNTIF(Cal![Status], "Overdue") = 4 AND COUNTIF(Cal![Status], "Due soon") = 2 AND COUNTIF(Cal![Status], "OK") = 3 AND COUNTIF(Cal![Status], "No date") = 1 AND COUNTIF(Cal![Status], "old") = 0 AND COUNTIF(NCR![Cal Flag], "Overdue instrument") = 6 AND COUNTIF(NCR![Cal Flag], "Unknown instrument") = 2 AND COUNTA(NCR![Cal Flag]) = 9 AND COUNTA(Over![Instrument ID]) = 5 AND COUNTIF(Over![Instrument ID], "stale") = 0 AND COUNTIF(Over![Status], "Overdue") = 4 AND VALUE(Over![Instrument ID]2) = "G-9" AND VALUE(Over![Instrument ID]3) = "G-10" AND VALUE(Over![Instrument ID]4) = "G-1" AND VALUE(Over![Instrument ID]5) = "G-2" AND VALUE(Over![Type]2) = "Scale" AND VALUE(Over![Type]5) = "Meter" AND VALUE(Over![Next Calibration Due]2) < VALUE(Over![Next Calibration Due]3) AND VALUE(Over![Next Calibration Due]4) < VALUE(Over![Next Calibration Due]5) AND COUNTIF(Inst![Instrument ID], "stale") = 0 AND ROWS_USED(Inst) = 7
TEST     fixture=calib_01.xlsx call=modMain.RunHeadless(Cal, NCR, Over, Inst) assert=RESULT = 8 AND COUNTA(Cal![Instrument ID]) = PRE_COUNTA(Cal![Instrument ID]) AND COUNTA(NCR![NCR No]) = PRE_COUNTA(NCR![NCR No]) AND COUNTA(NCR![Instrument ID]) = PRE_COUNTA(NCR![Instrument ID]) AND COUNTA(NCR![Status]) = PRE_COUNTA(NCR![Status]) AND COUNTIFS(NCR![NCR No], "N-1", NCR![Cal Flag], "Overdue instrument") = 1 AND COUNTIFS(NCR![NCR No], "N-12", NCR![Cal Flag], "Overdue instrument") = 1 AND COUNTIFS(NCR![NCR No], "N-2", NCR![Cal Flag], "<>") = 0 AND COUNTIFS(NCR![NCR No], "N-4", NCR![Cal Flag], "<>") = 0 AND COUNTIFS(NCR![NCR No], "N-6", NCR![Cal Flag], "<>") = 0 AND COUNTIFS(NCR![NCR No], "N-7", NCR![Cal Flag], "<>") = 0 AND COUNTIFS(NCR![NCR No], "N-13", NCR![Cal Flag], "<>") = 0 AND COUNTIFS(NCR![NCR No], "N-10", NCR![Cal Flag], "Unknown instrument") = 1 AND COUNTIFS(NCR![NCR No], "N-5", NCR![Cal Flag], "Unknown instrument") = 1 AND VALUE(Inst![Instrument ID]2) = "G-1" AND VALUE(Inst![NCR Count]2) = 4 AND VALUE(Inst![Instrument ID]3) = "G-9" AND VALUE(Inst![NCR Count]3) = 3 AND VALUE(Inst![Instrument ID]4) = "X-9" AND VALUE(Inst![NCR Count]4) = 2 AND SUM(Inst![NCR Count]) = 12
TEST     fixture=calib_02.xlsx call=modMain.RunHeadless(Cal, NCR, Over, Inst) assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTIF(Cal![Status], "OK") = 2 AND COUNTIF(Cal![Status], "No date") = 1 AND COUNTA(NCR![Cal Flag]) = PRE_COUNTA(NCR![Cal Flag]) - 2 AND COUNTA(Over![Instrument ID]) = 1 AND COUNTA(Inst![Instrument ID]) = 1
TEST     fixture=calib_01.xlsx call=modMain.RunHeadless(NCR, Cal, Over, Inst) assert=RESULT = -1 AND ERROR_COUNT = 1
---
VCL 0.6
TASK
GOAL     Start the calibration and NCR check on the workbook's own four sheets and tell the user how many NCRs were flagged.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "Calibration Tracker"), SHEET(THIS_BOOK, "NCR Register"), SHEET(THIS_BOOK, "Overdue Instruments"), SHEET(THIS_BOOK, "NCR Instruments"), into=n)
OP       MSG("NCRs flagged: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run all four steps of the calibration and NCR check on the sheets it is given and return how many NCRs were flagged.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     cal As Worksheet, ncr As Worksheet, over As Worksheet, inst As Worksheet
RET      Long
VAR      flagged As Long
PRE      COL($cal, "Instrument ID") > 0; COL($cal, "Type") > 0; COL($cal, "Next Calibration Due") > 0; COL($cal, "Status") > 0; COL($ncr, "NCR No") > 0; COL($ncr, "Instrument ID") > 0; COL($ncr, "Status") > 0; COL($ncr, "Cal Flag") > 0
MODIFIES $cal["Status"]
MODIFIES $ncr["Cal Flag"]
MODIFIES $over
MODIFIES $inst
OP       CALL(modSteps.SetCalStatus, $cal)
OP       CALL(modSteps.FlagNcrs, $cal, $ncr, into=flagged)
OP       CALL(modSteps.ListOverdue, $cal, $over)
OP       CALL(modSteps.CountByInstrument, $ncr, $inst)
OP       RETURN(flagged)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write Overdue, Due soon, OK or No date in each instrument's Status from its next calibration due date.
PROC     SetCalStatus
KIND     Sub
MODULE   modSteps
ARGS     cal As Worksheet
VAR      cId As Long
VAR      cDue As Long
VAR      cStat As Long
PRE      COL($cal, "Instrument ID") > 0; COL($cal, "Next Calibration Due") > 0; COL($cal, "Status") > 0
MODIFIES $cal["Status"]
OP       LET(cId, COL($cal, "Instrument ID"))
OP       LET(cDue, COL($cal, "Next Calibration Due"))
OP       LET(cStat, COL($cal, "Status"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($cal, cId) var=r
OP         IF IS_DATE(CELL($cal, r, cDue))
OP           IF CELL($cal, r, cDue) < TODAY
OP             SET_VALUE(CELL($cal, r, cStat), "Overdue")
OP           ELSE
OP             IF CELL($cal, r, cDue) <= TODAY + 14
OP               SET_VALUE(CELL($cal, r, cStat), "Due soon")
OP             ELSE
OP               SET_VALUE(CELL($cal, r, cStat), "OK")
OP         ELSE
OP           SET_VALUE(CELL($cal, r, cStat), "No date")
ERR      standard
---
VCL 0.6
TASK
GOAL     Flag each open NCR on an overdue instrument and each NCR on an instrument missing from the tracker, clear the flag on the rest and return how many were flagged.
PROC     FlagNcrs
KIND     Function
MODULE   modSteps
ARGS     cal As Worksheet, ncr As Worksheet
RET      Long
VAR      cCalId As Long
VAR      cCalStat As Long
VAR      cNo As Long
VAR      cNcrId As Long
VAR      cNcrStat As Long
VAR      cFlag As Long
VAR      known As Dictionary
VAR      late As Dictionary
VAR      id As String
VAR      n As Long
PRE      COL($cal, "Instrument ID") > 0; COL($cal, "Status") > 0; COL($ncr, "NCR No") > 0; COL($ncr, "Instrument ID") > 0; COL($ncr, "Status") > 0; COL($ncr, "Cal Flag") > 0
MODIFIES $ncr["Cal Flag"]
OP       LET(cCalId, COL($cal, "Instrument ID"))
OP       LET(cCalStat, COL($cal, "Status"))
OP       LET(cNo, COL($ncr, "NCR No"))
OP       LET(cNcrId, COL($ncr, "Instrument ID"))
OP       LET(cNcrStat, COL($ncr, "Status"))
OP       LET(cFlag, COL($ncr, "Cal Flag"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($cal, cCalId) var=r
OP         DICT_ADD(known, TRIM(CELL($cal, r, cCalId)), 1, mode=keep)
OP         IF SAME(CELL($cal, r, cCalStat), "Overdue") THEN DICT_ADD(late, TRIM(CELL($cal, r, cCalId)), 1, mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ncr, cNo) var=q
OP         CLEAR(CELL($ncr, q, cFlag), what=contents)
OP         SKIP_IF IS_EMPTY(CELL($ncr, q, cNcrId))
OP         LET(id, TRIM(CELL($ncr, q, cNcrId)))
OP         IF EXISTS(known, id)
OP           IF EXISTS(late, id) AND NOT SAME(CELL($ncr, q, cNcrStat), "Closed")
OP             SET_VALUE(CELL($ncr, q, cFlag), "Overdue instrument")
OP             LET(n, n + 1)
OP         ELSE
OP           SET_VALUE(CELL($ncr, q, cFlag), "Unknown instrument")
OP           LET(n, n + 1)
OP       RETURN(n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Overdue Instruments sheet with every overdue instrument and its details, the earliest due date first.
PROC     ListOverdue
KIND     Sub
MODULE   modSteps
ARGS     cal As Worksheet, over As Worksheet
VAR      cId As Long
VAR      cType As Long
VAR      cDue As Long
VAR      cStat As Long
VAR      outRow As Long
PRE      COL($cal, "Instrument ID") > 0; COL($cal, "Type") > 0; COL($cal, "Next Calibration Due") > 0; COL($cal, "Status") > 0
MODIFIES $over
OP       LET(cId, COL($cal, "Instrument ID"))
OP       LET(cType, COL($cal, "Type"))
OP       LET(cDue, COL($cal, "Next Calibration Due"))
OP       LET(cStat, COL($cal, "Status"))
OP       CLEAR(USED_RANGE($over), what=all)
OP       WRITE_ROW($over, 1, "Instrument ID", "Type", "Next Calibration Due", "Status")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($cal, cId) var=r
OP         IF SAME(CELL($cal, r, cStat), "Overdue")
OP           WRITE_ROW($over, outRow, CELL($cal, r, cId), CELL($cal, r, cType), CELL($cal, r, cDue), CELL($cal, r, cStat))
OP           LET(outRow, outRow + 1)
OP       IF outRow > 2 THEN SORT(USED_RANGE($over), key=3, order=asc, header=yes)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the NCR Instruments sheet with each instrument named in the NCR Register and its number of NCRs, the most first.
PROC     CountByInstrument
KIND     Sub
MODULE   modSteps
ARGS     ncr As Worksheet, inst As Worksheet
VAR      cId As Long
VAR      counts As Dictionary
VAR      got As Long
VAR      outRow As Long
PRE      COL($ncr, "Instrument ID") > 0
MODIFIES $inst
OP       LET(cId, COL($ncr, "Instrument ID"))
OP       CLEAR(USED_RANGE($inst), what=all)
OP       WRITE_ROW($inst, 1, "Instrument ID", "NCR Count")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ncr, cId) var=r
OP         SKIP_IF IS_EMPTY(CELL($ncr, r, cId))
OP         DICT_ADD(counts, TRIM(CELL($ncr, r, cId)), 1, mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(counts)
OP         DICT_GET(counts, k, into=got)
OP         WRITE_ROW($inst, outRow, UPPER(k), got)
OP         LET(outRow, outRow + 1)
OP       IF outRow > 2 THEN SORT(USED_RANGE($inst), key=2, order=desc, header=yes)
ERR      standard
