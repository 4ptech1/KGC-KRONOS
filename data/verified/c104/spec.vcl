VCL 0.6
TASK
# ASSUME: ITP Tracker headers in row 1 include "ITP No", "Point" and "Inspection Status"
# ASSUME: each row is one point; rows with a blank ITP No are ignored; ITP No is trimmed and matched ignoring case
# ASSUME: Open, Hold and Closed are matched ignoring case and extra spaces; any other status is not counted but the ITP still gets a row
# ASSUME: ITP numbers are shown in capitals
# ASSUME: the ITP Status sheet exists, is passed in and is rebuilt on every run, ITP numbers sorted A to Z
GOAL     Rebuild the ITP Status sheet with one row per ITP number and how many of its points are Open, Hold and Closed, sorted A to Z.
PROC     SummariseItpStatus
KIND     Function
MODULE   modItp
ARGS     itp As Worksheet, out As Worksheet
RET      Long
VAR      cItp As Long
VAR      cStatus As Long
VAR      nOpen As Dictionary
VAR      nHold As Dictionary
VAR      nClosed As Dictionary
VAR      a As Long
VAR      b As Long
VAR      c As Long
VAR      outRow As Long
PRE      COL($itp, "ITP No") > 0; COL($itp, "Inspection Status") > 0
MODIFIES $out
OP       LET(cItp, COL($itp, "ITP No"))
OP       LET(cStatus, COL($itp, "Inspection Status"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "ITP No", "Open", "Hold", "Closed")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($itp, cItp) var=r
OP         SKIP_IF IS_EMPTY(CELL($itp, r, cItp))
OP         DICT_ADD(nOpen, TRIM(CELL($itp, r, cItp)), 0, mode=sum)
OP         DICT_ADD(nHold, TRIM(CELL($itp, r, cItp)), 0, mode=sum)
OP         DICT_ADD(nClosed, TRIM(CELL($itp, r, cItp)), 0, mode=sum)
OP         IF SAME(CELL($itp, r, cStatus), "Open")
OP           DICT_ADD(nOpen, TRIM(CELL($itp, r, cItp)), 1, mode=sum)
OP         IF SAME(CELL($itp, r, cStatus), "Hold")
OP           DICT_ADD(nHold, TRIM(CELL($itp, r, cItp)), 1, mode=sum)
OP         IF SAME(CELL($itp, r, cStatus), "Closed")
OP           DICT_ADD(nClosed, TRIM(CELL($itp, r, cItp)), 1, mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(nOpen)
OP         DICT_GET(nOpen, k, into=a)
OP         DICT_GET(nHold, k, into=b)
OP         DICT_GET(nClosed, k, into=c)
OP         WRITE_ROW($out, outRow, UPPER(k), a, b, c)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=1, order=asc, header=yes)
OP       RETURN(COUNT(nOpen))
PERF     screen_off
ERR      standard
FIXTURE  itp_01.xlsx sheet=Tracker cols=A:ITP No,B:Point,C:Inspection Status rows=14 data=A:values("ITP-2"x4,"ITP-1"x3,"itp-3"x2,"ITP-4"x2,""x1,"ITP-5"x1,"ITP-1 "x1)+boundary("ITP-2","ITP-1","itp-3","","ITP-2","ITP-4","ITP-1","ITP-5","ITP-2","itp-3","ITP-4","ITP-2","ITP-1 ","ITP-1");B:seq(1,1);C:values("Open"x3,"hold "x2,"CLOSED"x3,"Closed"x1,"Pending"x2,"Open "x1,"Hold"x1,""x1)+boundary("Open","CLOSED","Pending","Open","CLOSED","hold ","Closed","","Open ","Open","Hold","hold ","Pending","CLOSED")
FIXTURE  itp_01.xlsx sheet=Summary cols=A:ITP No,B:Open,C:Hold,D:Closed rows=0 stale=2
TEST     fixture=itp_01.xlsx call=SummariseItpStatus(Tracker, Summary) runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND COUNTA(Summary![ITP No]) = 6 AND COUNTIF(Summary![ITP No], "stale") = 0 AND VALUE(Summary![ITP No]2) = "ITP-1" AND VALUE(Summary![Open]2) = 0 AND VALUE(Summary![Hold]2) = 0 AND VALUE(Summary![Closed]2) = 3 AND VALUE(Summary![ITP No]3) = "ITP-2" AND VALUE(Summary![Open]3) = 2 AND VALUE(Summary![Hold]3) = 1 AND VALUE(Summary![Closed]3) = 1 AND VALUE(Summary![ITP No]4) = "ITP-3" AND VALUE(Summary![Open]4) = 1 AND VALUE(Summary![Hold]4) = 0 AND VALUE(Summary![Closed]4) = 0 AND VALUE(Summary![ITP No]5) = "ITP-4" AND VALUE(Summary![Open]5) = 0 AND VALUE(Summary![Hold]5) = 2 AND VALUE(Summary![Closed]5) = 0 AND VALUE(Summary![ITP No]6) = "ITP-5" AND VALUE(Summary![Open]6) = 0 AND VALUE(Summary![Hold]6) = 0 AND VALUE(Summary![Closed]6) = 0
TEST     fixture=itp_01.xlsx call=SummariseItpStatus(Tracker, Summary) assert=SUM(Summary![Open]) + SUM(Summary![Hold]) + SUM(Summary![Closed]) = 10 AND COUNTA(Tracker![Point]) = PRE_COUNTA(Tracker![Point]) AND COUNTA(Tracker![Inspection Status]) = PRE_COUNTA(Tracker![Inspection Status])
TEST     fixture=itp_01.xlsx call=SummariseItpStatus(Summary, Tracker) assert=RESULT = -1 AND ERROR_COUNT = 1
