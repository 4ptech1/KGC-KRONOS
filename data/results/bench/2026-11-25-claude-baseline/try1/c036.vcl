VCL 0.7
TASK
# ASSUME: NCR Register headers in row 1 include "Category" and "Status"; Category Summary is an existing sheet rebuilt on every run
# ASSUME: a status counts as Open or Closed when it is SAME as that word; any other status is counted in neither column
# ASSUME: categories are matched ignoring case and stray spaces and written in alphabetical order; rows with a blank category are ignored
GOAL     Rebuild the Category Summary sheet with, for each NCR category in alphabetical order, how many NCRs are Open and how many are Closed.
PROC     BuildCategorySummary
KIND     Function
MODULE   modNcr
ARGS     ncr As Worksheet, out As Worksheet
RET      Long
VAR      cCat As Long
VAR      cStatus As Long
VAR      openN As Dictionary
VAR      closedN As Dictionary
VAR      nOpen As Long
VAR      nClosed As Long
VAR      outRow As Long
VAR      written As Long
PRE      COL($ncr, "Category") > 0; COL($ncr, "Status") > 0
MODIFIES $out
OP       LET(cCat, COL($ncr, "Category"))
OP       LET(cStatus, COL($ncr, "Status"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Category", "Open", "Closed")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ncr, cStatus) var=r
OP         SKIP_IF IS_EMPTY(CELL($ncr, r, cCat))
OP         DICT_ADD(openN, TRIM(CELL($ncr, r, cCat)), IIF(SAME(CELL($ncr, r, cStatus), "Open"), 1, 0), mode=sum)
OP         DICT_ADD(closedN, TRIM(CELL($ncr, r, cCat)), IIF(SAME(CELL($ncr, r, cStatus), "Closed"), 1, 0), mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(openN)
OP         DICT_GET(openN, k, into=nOpen)
OP         DICT_GET(closedN, k, into=nClosed)
OP         WRITE_ROW($out, outRow, k, nOpen, nClosed)
OP         LET(outRow, outRow + 1)
OP         LET(written, written + 1)
OP       SORT(USED_RANGE($out), key=1, order=asc, header=yes)
OP       RETURN(written)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  ncr_01.xlsx sheet=Ncr cols=A:NCR No,B:Category,C:Issue Date,D:Status rows=10 data=A:seq(1,1);B:values("Welding"x4,"Coating"x3,"Material"x2,""x1)+boundary("Welding","Coating","Welding","Material","Coating","","Welding","Material","Coating","Welding");C:dates(2026-01-01,2026-09-30);D:values("Open"x6,"Closed"x4)+boundary("Open","Closed","Open","Open","Open","Open","Closed","Closed","Open","Closed")
FIXTURE  ncr_01.xlsx sheet=Summary cols=A:Category,B:Open,C:Closed rows=0 stale=2
FIXTURE  ncr_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=ncr_01.xlsx call=BuildCategorySummary(Ncr, Summary) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTA(Summary![Category]) = 4 AND COUNTIF(Summary![Category], "stale") = 0 AND VALUE(Summary![Category]2) = "Coating" AND VALUE(Summary![Open]2) = 2 AND VALUE(Summary![Closed]2) = 1
TEST     fixture=ncr_01.xlsx call=BuildCategorySummary(Ncr, Summary) assert=COUNTA(Ncr![NCR No]) = PRE_COUNTA(Ncr![NCR No]) AND VALUE(Summary![Category]3) = "Material" AND VALUE(Summary![Open]3) = 1 AND VALUE(Summary![Closed]3) = 1 AND VALUE(Summary![Category]4) = "Welding" AND VALUE(Summary![Open]4) = 2 AND VALUE(Summary![Closed]4) = 2
TEST     fixture=ncr_01.xlsx call=BuildCategorySummary(Notes, Summary) assert=RESULT = -1 AND ERROR_COUNT = 1
