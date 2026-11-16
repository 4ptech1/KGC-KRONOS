VCL 0.6
TASK
# ASSUME: the NCR Register headers are in row 1 and include "NCR No", "Department" and "Status"; data starts in row 2
# ASSUME: the summary goes on a new sheet named NCR Summary after the last sheet; when NCR Summary is already there the run changes nothing and returns 0
# ASSUME: a Status that is Open, Closed or On Hold apart from case and extra spaces is counted as Open, Closed or On Hold; any other Status is counted under its own trimmed text
# ASSUME: a row with a blank NCR No is not an NCR and is not counted
GOAL     Add an NCR Summary sheet that counts the NCRs by status and by department, and return how many NCRs were counted.
PROC     SummariseNcrs
KIND     Function
MODULE   modNcrReview
ARGS     book As Workbook, ws As Worksheet
RET      Long
VAR      cNcr As Long
VAR      cDept As Long
VAR      cStatus As Long
VAR      out As Worksheet
VAR      byStatus As Dictionary
VAR      byDept As Dictionary
VAR      label As String
VAR      got As Long
VAR      outRow As Long
VAR      total As Long
PRE      COL($ws, "NCR No") > 0; COL($ws, "Department") > 0; COL($ws, "Status") > 0
MODIFIES $book sheets
OP       LET(cNcr, COL($ws, "NCR No"))
OP       LET(cDept, COL($ws, "Department"))
OP       LET(cStatus, COL($ws, "Status"))
OP       EXIT_IF SHEET_EXISTS($book, "NCR Summary")
OP       LET(total, 0)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cNcr) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cNcr))
OP         LET(label, TRIM(CELL($ws, r, cStatus)))
OP         IF SAME(label, "Open") THEN LET(label, "Open")
OP         IF SAME(label, "Closed") THEN LET(label, "Closed")
OP         IF SAME(label, "On Hold") THEN LET(label, "On Hold")
OP         DICT_ADD(byStatus, label, 1, mode=sum)
OP         DICT_ADD(byDept, TRIM(CELL($ws, r, cDept)), 1, mode=sum)
OP         LET(total, total + 1)
OP       ADD_SHEET($book, name="NCR Summary", into=out)
OP       WRITE_ROW(out, 1, "Status", "NCRs")
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(byStatus)
OP         DICT_GET(byStatus, k, into=got)
OP         WRITE_ROW(out, outRow, k, got)
OP         LET(outRow, outRow + 1)
OP       LET(outRow, outRow + 1)
OP       WRITE_ROW(out, outRow, "Department", "NCRs")
OP       LET(outRow, outRow + 1)
OP       FOR_EACH var=k in=KEYS(byDept)
OP         DICT_GET(byDept, k, into=got)
OP         WRITE_ROW(out, outRow, k, got)
OP         LET(outRow, outRow + 1)
OP       LET(outRow, outRow + 1)
OP       WRITE_ROW(out, outRow, "Total", total)
OP       RETURN(total)
ERR      standard
FIXTURE  ncr_03.xlsx sheet=NCR cols=A:NCR No,B:Department,C:Status rows=9 data=A:values("NCR-1"x1,"NCR-2"x1,"NCR-3"x1,""x1,"NCR-5"x1,"NCR-6"x1,"NCR-7"x1,"NCR-8"x1,"NCR-9"x1)+boundary("NCR-1","NCR-2","NCR-3","","NCR-5","NCR-6","NCR-7","NCR-8","NCR-9");B:values("QA"x3,"Welding"x3,"Machining"x2,"Stores"x1)+boundary("QA","Welding","QA","Stores","Machining","Welding","QA","Machining","Welding");C:values(" open"x1,"Open"x3,"CLOSED "x1,"on hold"x1,"Closed"x2,"On Hold"x1)+boundary(" open","Open","CLOSED ","Open","on hold","Closed","Open","Closed","On Hold")
FIXTURE  ncr_03.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=ncr_03.xlsx call=SummariseNcrs(ncr_03.xlsx, NCR) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND SHEET_EXISTS("NCR Summary") AND VALUE('NCR Summary'!A1) = "Status" AND SUMIFS('NCR Summary'![NCRs], 'NCR Summary'![Status], "Open") = 3 AND SUMIFS('NCR Summary'![NCRs], 'NCR Summary'![Status], "Closed") = 3 AND SUMIFS('NCR Summary'![NCRs], 'NCR Summary'![Status], "On Hold") = 2
TEST     fixture=ncr_03.xlsx call=SummariseNcrs(ncr_03.xlsx, NCR) assert=RESULT = 8 AND ERROR_COUNT = 0 AND ROWS_USED('NCR Summary') = 11 AND VALUE('NCR Summary'!A1) = "Status" AND VALUE('NCR Summary'!B1) = "NCRs" AND VALUE('NCR Summary'!A2) = "Open" AND VALUE('NCR Summary'!B2) = 3 AND VALUE('NCR Summary'!A3) = "Closed" AND VALUE('NCR Summary'!B3) = 3 AND VALUE('NCR Summary'!A4) = "On Hold" AND VALUE('NCR Summary'!B4) = 2 AND COUNTA('NCR Summary'!A5) = 0 AND VALUE('NCR Summary'!A6) = "Department" AND VALUE('NCR Summary'!B6) = "NCRs" AND VALUE('NCR Summary'!A7) = "QA" AND VALUE('NCR Summary'!B7) = 3 AND VALUE('NCR Summary'!A8) = "Welding" AND VALUE('NCR Summary'!B8) = 3 AND VALUE('NCR Summary'!A9) = "Machining" AND VALUE('NCR Summary'!B9) = 2 AND COUNTA('NCR Summary'!A10) = 0 AND VALUE('NCR Summary'!A11) = "Total" AND VALUE('NCR Summary'!B11) = 8 AND SUMIFS('NCR Summary'![NCRs], 'NCR Summary'![Status], "Stores") = 0 AND COUNTA(NCR![Status]) = PRE_COUNTA(NCR![Status])
TEST     fixture=ncr_03.xlsx call=SummariseNcrs(ncr_03.xlsx, Notes) assert=RESULT = -1 AND ERROR_COUNT = 1 AND SHEET_EXISTS("NCR Summary") = 0
