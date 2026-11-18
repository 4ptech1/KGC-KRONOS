VCL 0.6
TASK
# ASSUME: NDE Report headers in row 1 include "Weld Seam No", "Welder ID" and "Result"; Welder Quals headers include "Welder ID", "Process", "Expiry Date" and "Remark"
# ASSUME: the real sheets are "NDE Report", "Welder Performance", "Welder Quals" and "Expired Quals"; the TESTs use NdeReport, WelderPerf, WelderQuals and ExpiredQuals
# ASSUME: the output sheets exist and are rebuilt on every run; Welder Performance gets the headings Welder ID, Welds Tested, Rejects
# ASSUME: every NDE Report row with a Welder ID counts as one tested weld; a row with a blank Welder ID is ignored; Reject is matched ignoring case and spaces
# ASSUME: a welder is expired when Expiry Date is a real date before today; a blank or non-date Expiry Date is not expired
# ASSUME: Welder Quals rows with a blank Welder ID are left alone; a welder with no NDE rows has 0 rejects
# ASSUME: Expired Quals lists every expired row, whatever its Remark, oldest expiry first, with the Remark as written by this run
# ASSUME: Welder IDs are trimmed and compared ignoring case; welders with equally many rejects may come in any order
GOAL     Rebuild the welder performance and expired qualification sheets from the NDE report and mark welders with more than 2 rejects as Review in Welder Quals.
PROC     RebuildWelderPerformance
KIND     Function
MODULE   modWelders
ARGS     nde As Worksheet, perf As Worksheet, quals As Worksheet, expired As Worksheet
RET      Long
VAR      cNdeWelder As Long
VAR      cNdeResult As Long
VAR      cQWelder As Long
VAR      cQProcess As Long
VAR      cQExpiry As Long
VAR      cQRemark As Long
VAR      tested As Dictionary
VAR      rejects As Dictionary
VAR      wid As String
VAR      remark As String
VAR      rej As Long
VAR      nTested As Long
VAR      nRej As Long
VAR      outRow As Long
VAR      expRow As Long
VAR      marked As Long
PRE      COL($nde, "Welder ID") > 0; COL($nde, "Result") > 0; COL($quals, "Welder ID") > 0; COL($quals, "Process") > 0; COL($quals, "Expiry Date") > 0; COL($quals, "Remark") > 0
MODIFIES $perf
MODIFIES $expired
MODIFIES $quals["Remark"]
OP       LET(cNdeWelder, COL($nde, "Welder ID"))
OP       LET(cNdeResult, COL($nde, "Result"))
OP       LET(cQWelder, COL($quals, "Welder ID"))
OP       LET(cQProcess, COL($quals, "Process"))
OP       LET(cQExpiry, COL($quals, "Expiry Date"))
OP       LET(cQRemark, COL($quals, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($nde, cNdeWelder) var=r
OP         SKIP_IF IS_EMPTY(CELL($nde, r, cNdeWelder))
OP         DICT_ADD(tested, TRIM(CELL($nde, r, cNdeWelder)), 1, mode=sum)
OP         DICT_ADD(rejects, TRIM(CELL($nde, r, cNdeWelder)), IIF(SAME(CELL($nde, r, cNdeResult), "Reject"), 1, 0), mode=sum)
OP       CLEAR(USED_RANGE($perf), what=all)
OP       WRITE_ROW($perf, 1, "Welder ID", "Welds Tested", "Rejects")
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(tested)
OP         DICT_GET(tested, k, into=nTested)
OP         DICT_GET(rejects, k, into=nRej)
OP         WRITE_ROW($perf, outRow, k, nTested, nRej)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($perf), key=3, order=desc, header=yes)
OP       CLEAR(USED_RANGE($expired), what=all)
OP       WRITE_ROW($expired, 1, "Welder ID", "Process", "Expiry Date", "Remark")
OP       LET(expRow, 1)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($quals, cQWelder) var=r
OP         SKIP_IF IS_EMPTY(CELL($quals, r, cQWelder))
OP         LET(wid, TRIM(CELL($quals, r, cQWelder)))
OP         LET(rej, 0)
OP         IF EXISTS(rejects, wid) THEN DICT_GET(rejects, wid, into=rej)
OP         LET(remark, IIF(rej > 2, "Review", ""))
OP         IF IS_DATE(CELL($quals, r, cQExpiry))
OP           IF CELL($quals, r, cQExpiry) < TODAY
OP             IF remark = "" THEN LET(remark, "Expired")
OP             LET(expRow, expRow + 1)
OP             WRITE_ROW($expired, expRow, CELL($quals, r, cQWelder), CELL($quals, r, cQProcess), CELL($quals, r, cQExpiry), remark)
OP         IF remark = "" THEN CLEAR(CELL($quals, r, cQRemark), what=contents)
OP         IF remark <> "" THEN SET_VALUE(CELL($quals, r, cQRemark), remark)
OP         IF remark = "Review" THEN LET(marked, marked + 1)
OP       SORT(USED_RANGE($expired), key=3, order=asc, header=yes)
OP       MSG("Rows marked Review: " & marked)
OP       RETURN(marked)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  welder_perf_01.xlsx sheet=NdeReport cols=A:Weld Seam No,B:Welder ID,C:Result rows=12 data=A:seq(1,1);B:values("W-1"x4,"W-2"x3,"W-3"x2,"W-4"x1," W-1 "x1,""x1)+boundary("W-1","W-2","W-1","W-3"," W-1 ","W-2","W-1","W-2","W-3","W-1","","W-4");C:values("Reject"x6,"reject "x1,"Pass"x5)+boundary("Reject","Reject","Reject","Pass","Reject","reject ","Pass","Pass","Reject","Pass","Reject","Pass")
FIXTURE  welder_perf_01.xlsx sheet=WelderPerf cols=A:Welder ID,B:Welds Tested,C:Rejects rows=0 stale=2
FIXTURE  welder_perf_01.xlsx sheet=WelderQuals cols=A:Welder ID,B:Process,C:Expiry Date,D:Remark rows=8 data=A:values("W-1"x1,"W-2"x2,"W-3"x1,"W-4"x1," W-1 "x1,"W-9"x1,""x1)+boundary("W-1","W-2","W-3","W-4"," W-1 ","W-9","","W-2");B:values("TIG"x5,"MIG"x3)+boundary("TIG","MIG","TIG","MIG","MIG","TIG","TIG","TIG");C:reldates(-400,100)+boundary(-30,-1,0,5,40,-400,-50,"");D:values("old"x1,""x3,"Review"x1,"Expired"x2,"keep"x1)+boundary("old","","Review","Expired","","","keep","Expired")
FIXTURE  welder_perf_01.xlsx sheet=ExpiredQuals cols=A:Welder ID,B:Process,C:Expiry Date,D:Remark rows=0 stale=2
TEST     fixture=welder_perf_01.xlsx call=RebuildWelderPerformance(NdeReport, WelderPerf, WelderQuals, ExpiredQuals) runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND COUNTIF(WelderQuals![Remark], "Review") = 2 AND COUNTIF(WelderQuals![Remark], "Expired") = 2 AND COUNTIF(WelderQuals![Remark], "keep") = 1 AND COUNTA(WelderQuals![Remark]) = 6 AND ROWS_USED(WelderPerf) = 5 AND COUNTIF(WelderPerf![Welder ID], "stale") = 0 AND ROWS_USED(ExpiredQuals) = 4 AND COUNTIF(ExpiredQuals![Welder ID], "stale") = 0
TEST     fixture=welder_perf_01.xlsx call=RebuildWelderPerformance(NdeReport, WelderPerf, WelderQuals, ExpiredQuals) assert=VALUE(WelderPerf![Welder ID]2) = "W-1" AND VALUE(WelderPerf![Welds Tested]2) = 5 AND VALUE(WelderPerf![Rejects]2) = 3 AND VALUE(WelderPerf![Welder ID]3) = "W-2" AND VALUE(WelderPerf![Welds Tested]3) = 3 AND VALUE(WelderPerf![Rejects]3) = 2 AND VALUE(WelderPerf![Welder ID]4) = "W-3" AND VALUE(WelderPerf![Welds Tested]4) = 2 AND VALUE(WelderPerf![Rejects]4) = 1 AND VALUE(WelderPerf![Welder ID]5) = "W-4" AND VALUE(WelderPerf![Welds Tested]5) = 1 AND VALUE(WelderPerf![Rejects]5) = 0
TEST     fixture=welder_perf_01.xlsx call=RebuildWelderPerformance(NdeReport, WelderPerf, WelderQuals, ExpiredQuals) assert=VALUE(ExpiredQuals![Welder ID]2) = "W-9" AND VALUE(ExpiredQuals![Remark]2) = "Expired" AND VALUE(ExpiredQuals![Welder ID]3) = "W-1" AND VALUE(ExpiredQuals![Process]3) = "TIG" AND VALUE(ExpiredQuals![Remark]3) = "Review" AND VALUE(ExpiredQuals![Welder ID]4) = "W-2" AND VALUE(ExpiredQuals![Remark]4) = "Expired" AND VALUE(WelderQuals![Remark]2) = "Review" AND VALUE(WelderQuals![Remark]6) = "Review" AND VALUE(WelderQuals![Remark]9) = "" AND VALUE(WelderQuals![Remark]5) = "" AND VALUE(WelderQuals![Remark]4) = "" AND VALUE(WelderQuals![Remark]8) = "keep"
TEST     fixture=welder_perf_01.xlsx call=RebuildWelderPerformance(NdeReport, WelderPerf, WelderQuals, ExpiredQuals) assert=ERROR_COUNT = 0 AND COUNTA(NdeReport![Welder ID]) = PRE_COUNTA(NdeReport![Welder ID]) AND COUNTA(WelderQuals![Welder ID]) = PRE_COUNTA(WelderQuals![Welder ID]) AND COUNTA(WelderQuals![Process]) = PRE_COUNTA(WelderQuals![Process])
TEST     fixture=welder_perf_01.xlsx call=RebuildWelderPerformance(WelderQuals, WelderPerf, NdeReport, ExpiredQuals) assert=RESULT = -1 AND ERROR_COUNT = 1
