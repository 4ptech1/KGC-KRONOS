VCL 0.6
TASK
# ASSUME: the real sheets are "NCR Register" and "NCR Aging"; the TESTs use NcrRegister and NcrAging
# ASSUME: NCR Aging exists already; it is cleared and rebuilt from the register on every run
# ASSUME: the register has headers "NCR No", "Issue Date" and "Status" in row 1; Open is matched ignoring case and extra spaces
# ASSUME: age in days is today minus Issue Date; an Open NCR with no real Issue Date is listed last, with its date as typed and no age
GOAL     Rebuild the NCR Aging sheet with every Open NCR from the register and its age in days, oldest first, and return how many were listed.
PROC     RebuildNcrAging
KIND     Function
MODULE   modNcrAging
ARGS     reg As Worksheet, out As Worksheet
RET      Long
VAR      cNo As Long
VAR      cDate As Long
VAR      cStatus As Long
VAR      outRow As Long
PRE      COL($reg, "NCR No") > 0; COL($reg, "Issue Date") > 0; COL($reg, "Status") > 0
MODIFIES $out
OP       LET(cNo, COL($reg, "NCR No"))
OP       LET(cDate, COL($reg, "Issue Date"))
OP       LET(cStatus, COL($reg, "Status"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "NCR No", "Issue Date", "Age Days")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($reg, cStatus) var=r
OP         SKIP_IF NOT SAME(CELL($reg, r, cStatus), "Open")
OP         SKIP_IF NOT IS_DATE(CELL($reg, r, cDate))
OP         WRITE_ROW($out, outRow, CELL($reg, r, cNo), CELL($reg, r, cDate), TODAY - CELL($reg, r, cDate))
OP         LET(outRow, outRow + 1)
OP       IF outRow > 2
OP         SORT(USED_RANGE($out), key=3, order=desc, header=yes)
OP       LOOP_ROWS dir=down from=2 to=outRow - 1 var=q
OP         FORMAT(CELL($out, q, 2), numfmt="yyyy-mm-dd")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($reg, cStatus) var=r
OP         SKIP_IF NOT SAME(CELL($reg, r, cStatus), "Open")
OP         SKIP_IF IS_DATE(CELL($reg, r, cDate))
OP         WRITE_ROW($out, outRow, CELL($reg, r, cNo), CELL($reg, r, cDate))
OP         LET(outRow, outRow + 1)
OP       RETURN(outRow - 2)
ERR      standard
FIXTURE  ncr_01.xlsx sheet=NcrRegister cols=A:NCR No,B:Issue Date,C:Status,D:Owner rows=12 data=A:seq(1,1);B:reldates(-80,-1)+boundary(-45,-90,"","",-1,"",-30);C:values("Open"x5," open "x1,"Closed"x3,"Hold"x1,""x2)+boundary("Open"," open ","Open","Closed","Open","Open","Hold");D:text(6)
FIXTURE  ncr_01.xlsx sheet=NcrAging cols=A:NCR No,B:Issue Date,C:Age Days rows=0 stale=3
FIXTURE  ncr_02.xlsx sheet=NcrRegister cols=A:NCR No,B:Issue Date,C:Status,D:Owner rows=4 data=A:seq(1,1);B:reldates(-80,-1)+boundary("");C:values("Open"x1,"Closed"x2,"Hold"x1)+boundary("Open","Closed");D:text(6)
FIXTURE  ncr_02.xlsx sheet=NcrAging cols=A:NCR No,B:Issue Date,C:Age Days rows=0 stale=3
TEST     fixture=ncr_01.xlsx call=RebuildNcrAging(NcrRegister, NcrAging) runs=2 assert=RESULT = 6 AND ERROR_COUNT = 0 AND ROWS_USED(NcrAging) = 7 AND COUNTIF(NcrAging![NCR No], "stale") = 0 AND VALUE(NcrAging![NCR No]2) = 2 AND VALUE(NcrAging![Age Days]2) = 90 AND VALUE(NcrAging![Age Days]5) = 1 AND VALUE(NcrAging![NCR No]6) = 3 AND VALUE(NcrAging![NCR No]7) = 6 AND COUNTA(NcrAging![Age Days]) = 5
TEST     fixture=ncr_01.xlsx call=RebuildNcrAging(NcrRegister, NcrAging) assert=RESULT = 6 AND VALUE(NcrAging![Age Days]2) >= VALUE(NcrAging![Age Days]3) AND VALUE(NcrAging![Age Days]3) >= VALUE(NcrAging![Age Days]4) AND VALUE(NcrAging![Age Days]4) >= VALUE(NcrAging![Age Days]5) AND COUNTIF(NcrAging![NCR No], 4) = 0 AND COUNTIF(NcrAging![NCR No], 7) = 0 AND COUNTA(NcrRegister![NCR No]) = PRE_COUNTA(NcrRegister![NCR No]) AND COUNTIF(NcrRegister![Status], "Open") = PRE_COUNTIF(NcrRegister![Status], "Open")
TEST     fixture=ncr_02.xlsx call=RebuildNcrAging(NcrRegister, NcrAging) runs=2 assert=RESULT = 1 AND ERROR_COUNT = 0 AND ROWS_USED(NcrAging) = 2 AND VALUE(NcrAging![NCR No]2) = 1 AND COUNTIF(NcrAging![NCR No], "stale") = 0 AND COUNTA(NcrAging![Age Days]) = 1
TEST     fixture=ncr_01.xlsx call=RebuildNcrAging(NcrAging, NcrRegister) assert=RESULT = -1 AND ERROR_COUNT = 1
