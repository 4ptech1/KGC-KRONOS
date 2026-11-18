VCL 0.6
TASK
# ASSUME: PQR List headers in row 1 include "PQR No", "Process", "Test Date", "Result" and "Remark"
# ASSUME: Result is matched ignoring case and extra spaces; Test Date must be a real date, otherwise the row is left alone
# ASSUME: more than 90 days ago means the Test Date is before today minus 90 days, so exactly 90 days ago is not due
# ASSUME: every other row with a real Test Date has its Remark cleared, including Fail rows that are not old enough
# ASSUME: the macro returns the number of retests due
GOAL     Write Retest due in Remark for every failed PQR tested more than 90 days ago, clear Remark on other rows with a real test date and return how many retests are due.
PROC     MarkPqrRetests
KIND     Function
MODULE   modPqr
ARGS     ws As Worksheet
RET      Long
VAR      cDate As Long
VAR      cResult As Long
VAR      cRemark As Long
VAR      due As Long
PRE      COL($ws, "PQR No") > 0; COL($ws, "Process") > 0; COL($ws, "Test Date") > 0; COL($ws, "Result") > 0; COL($ws, "Remark") > 0
MODIFIES $ws["Remark"]
OP       LET(cDate, COL($ws, "Test Date"))
OP       LET(cResult, COL($ws, "Result"))
OP       LET(cRemark, COL($ws, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, COL($ws, "PQR No")) var=r
OP         SKIP_IF NOT IS_DATE(CELL($ws, r, cDate))
OP         IF SAME(CELL($ws, r, cResult), "Fail") AND CELL($ws, r, cDate) < TODAY - 90
OP           SET_VALUE(CELL($ws, r, cRemark), "Retest due")
OP           LET(due, due + 1)
OP         ELSE
OP           CLEAR(CELL($ws, r, cRemark), what=contents)
OP       RETURN(due)
PERF     screen_off
ERR      standard
FIXTURE  pqr_01.xlsx sheet=Pqr cols=A:PQR No,B:Process,C:Test Date,D:Result,E:Remark rows=10 data=A:seq(1,1);B:values("GTAW"x5,"SMAW"x5);C:reldates(-200,0)+boundary(-91,-90,-89,-300,-120,"",-10,-100,-95,-5);D:values("Fail"x2,"FAIL"x1,"fail "x1,"Pass"x3,"Fail"x3)+boundary("Fail","Fail","Fail","FAIL","fail ","Fail","Fail","Pass","Pass","Pass");E:values("x"x10)
FIXTURE  pqr_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=pqr_01.xlsx call=MarkPqrRetests(Pqr) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(Pqr![Remark], "Retest due") = 3 AND COUNTIF(Pqr![Remark], "x") = 1
TEST     fixture=pqr_01.xlsx call=MarkPqrRetests(Pqr) assert=VALUE(Pqr![Remark]2) = "Retest due" AND VALUE(Pqr![Remark]3) = "" AND VALUE(Pqr![Remark]4) = "" AND VALUE(Pqr![Remark]5) = "Retest due" AND VALUE(Pqr![Remark]6) = "Retest due" AND VALUE(Pqr![Remark]7) = "x" AND COUNTA(Pqr![PQR No]) = PRE_COUNTA(Pqr![PQR No]) AND COUNTA(Pqr![Result]) = PRE_COUNTA(Pqr![Result])
TEST     fixture=pqr_01.xlsx call=MarkPqrRetests(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
