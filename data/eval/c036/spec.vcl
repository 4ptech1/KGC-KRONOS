VCL 0.6
TASK
# ASSUME: NCR Register headers in row 1 include "Category" and "Status"; the real sheets are "NCR Register" and "Category Summary", the TESTs use Ncr and Summary
# ASSUME: the Category Summary sheet already exists and is rebuilt from scratch on every run; its headings are Category, Open and Closed
# ASSUME: categories are trimmed and matched ignoring case; the first spelling met is the one written
# ASSUME: Status is matched with SAME, so Open and Closed ignore case and extra spaces; any other status is not counted, and a category with only other statuses is listed with 0 and 0
# ASSUME: a row with no category is ignored; categories are sorted A to Z
GOAL     Rebuild the Category Summary sheet with one row for each NCR category, counting its Open and Closed NCRs, sorted alphabetically, and return how many categories were listed.
PROC     BuildCategorySummary
KIND     Function
MODULE   modNcrSummary
ARGS     ncr As Worksheet, out As Worksheet
RET      Long
VAR      cCat As Long
VAR      cStatus As Long
VAR      openN As Dictionary
VAR      closedN As Dictionary
VAR      catName As String
VAR      nOpen As Long
VAR      nClosed As Long
VAR      outRow As Long
PRE      COL($ncr, "Category") > 0; COL($ncr, "Status") > 0
MODIFIES $out
OP       LET(cCat, COL($ncr, "Category"))
OP       LET(cStatus, COL($ncr, "Status"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ncr, cCat) var=r
OP         SKIP_IF IS_EMPTY(CELL($ncr, r, cCat))
OP         LET(catName, TRIM(CELL($ncr, r, cCat)))
OP         DICT_ADD(openN, catName, 0, mode=keep)
OP         IF SAME(CELL($ncr, r, cStatus), "Open")
OP           DICT_ADD(openN, catName, 1, mode=sum)
OP         IF SAME(CELL($ncr, r, cStatus), "Closed")
OP           DICT_ADD(closedN, catName, 1, mode=sum)
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Category", "Open", "Closed")
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(openN)
OP         DICT_GET(openN, k, into=nOpen)
OP         LET(nClosed, 0)
OP         IF EXISTS(closedN, k)
OP           DICT_GET(closedN, k, into=nClosed)
OP         WRITE_ROW($out, outRow, k, nOpen, nClosed)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=1, order=asc, header=yes)
OP       RETURN(COUNT(openN))
PERF     screen_off
ERR      standard
FIXTURE  ncr_01.xlsx sheet=Ncr cols=A:NCR No,B:Category,C:Issue Date,D:Status rows=10 data=A:seq(1,1);B:values("Weld"x2,"Material"x2,"Dim"x2,""x1," Weld "x1,"Coating"x1,"weld"x1)+boundary("Weld","Material","Weld","Dim",""," Weld ","Coating","weld","Material","Dim");C:dates(2026-01-01,2026-09-30);D:values("Open"x5,"Closed"x4,"Overdue"x1)+boundary("Open","Closed","Open","Open","Open","Closed","Overdue","Closed","Open","Closed")
FIXTURE  ncr_01.xlsx sheet=Summary cols=A:Category,B:Open,C:Closed rows=0 stale=6
TEST     fixture=ncr_01.xlsx call=BuildCategorySummary(Ncr, Summary) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTA(Summary![Category]) = 5 AND ROWS_USED(Summary) = 5 AND COUNTIF(Summary![Category], "stale") = 0 AND VALUE(Summary![Category]2) = "Coating" AND VALUE(Summary![Category]3) = "Dim" AND VALUE(Summary![Category]4) = "Material" AND VALUE(Summary![Category]5) = "Weld" AND VALUE(Summary![Open]5) = 2 AND VALUE(Summary![Closed]5) = 2 AND VALUE(Summary![Open]2) = 0 AND VALUE(Summary![Closed]2) = 0
TEST     fixture=ncr_01.xlsx call=BuildCategorySummary(Ncr, Summary) assert=SUM(Summary![Open]) = 4 AND SUM(Summary![Closed]) = 4 AND VALUE(Summary![Open]3) = 1 AND VALUE(Summary![Closed]3) = 1 AND VALUE(Summary![Open]4) = 1 AND VALUE(Summary![Closed]4) = 1 AND COUNTA(Ncr![NCR No]) = PRE_COUNTA(Ncr![NCR No]) AND COUNTA(Ncr![Category]) = PRE_COUNTA(Ncr![Category]) AND COUNTIF(Ncr![Status], "Open") = PRE_COUNTIF(Ncr![Status], "Open")
TEST     fixture=ncr_01.xlsx call=BuildCategorySummary(Summary, Ncr) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(Summary![Category]) = PRE_COUNTA(Summary![Category])
