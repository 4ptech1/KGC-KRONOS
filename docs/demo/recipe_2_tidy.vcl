VCL 0.6
TASK
# ASSUME: the NCR Register headers are in row 1 and include "NCR No", "Date Raised" and "Status"; data starts in row 2
# ASSUME: a Status that is Open, Closed or On Hold apart from case and extra spaces is rewritten exactly as Open, Closed or On Hold; any other Status is left as it is
# ASSUME: a repeated NCR No is a copy typed twice: the first row stays and each later row with the same NCR No is deleted; DEDUPE ignores case but does not trim, and a blank NCR No is never removed
# ASSUME: Date Raised holds real dates; rows with a blank date sort to the bottom
GOAL     Write every NCR status in one clean form, remove repeated NCR numbers keeping the first, and sort the register by Date Raised, newest first.
PROC     TidyNcrRegister
KIND     Function
MODULE   modNcrReview
ARGS     ws As Worksheet
RET      Long
VAR      cNcr As Long
VAR      cDate As Long
VAR      cStatus As Long
VAR      removed As Long
PRE      COL($ws, "NCR No") > 0; COL($ws, "Date Raised") > 0; COL($ws, "Status") > 0
MODIFIES $ws
OP       LET(cNcr, COL($ws, "NCR No"))
OP       LET(cDate, COL($ws, "Date Raised"))
OP       LET(cStatus, COL($ws, "Status"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cNcr) var=r
OP         IF SAME(CELL($ws, r, cStatus), "Open")
OP           SET_VALUE(CELL($ws, r, cStatus), "Open")
OP         IF SAME(CELL($ws, r, cStatus), "Closed")
OP           SET_VALUE(CELL($ws, r, cStatus), "Closed")
OP         IF SAME(CELL($ws, r, cStatus), "On Hold")
OP           SET_VALUE(CELL($ws, r, cStatus), "On Hold")
OP       DEDUPE(USED_RANGE($ws), cols=cNcr, header=yes, into=removed)
OP       SORT(USED_RANGE($ws), key=cDate, order=desc, header=yes)
OP       RETURN(removed)
PERF     screen_off
ERR      standard
FIXTURE  ncr_02.xlsx sheet=NCR cols=A:NCR No,B:Date Raised,C:Status rows=8 data=A:values("NCR-1"x2,"NCR-2"x1,"NCR-3"x1,"NCR-4"x2,"NCR-5"x1,"NCR-6"x1)+boundary("NCR-1","NCR-2","NCR-1","NCR-3","NCR-4","NCR-5","NCR-4","NCR-6");B:dates(2026-03-01,2026-03-31)+boundary(2026-03-01,2026-03-05,2026-03-02,2026-03-08,2026-03-03,2026-03-07,2026-03-04,2026-03-06);C:values(" open"x1,"CLOSED "x1,"Open"x1,"on hold"x1,"Closed"x2,"OPEN"x1,"On Hold"x1)+boundary(" open","CLOSED ","Open","on hold","Closed","OPEN","Closed","On Hold")
FIXTURE  ncr_02.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=ncr_02.xlsx call=TidyNcrRegister(NCR) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTA(NCR![NCR No]) = 7 AND COUNTIF(NCR![Status], "Open") = 2 AND COUNTIF(NCR![Status], "Closed") = 2 AND COUNTIF(NCR![Status], "On Hold") = 2
TEST     fixture=ncr_02.xlsx call=TidyNcrRegister(NCR) assert=RESULT = 2 AND ERROR_COUNT = 0 AND COUNTA(NCR![NCR No]) = PRE_COUNTA(NCR![NCR No]) - RESULT AND VALUE(NCR![NCR No]2) = "NCR-3" AND VALUE(NCR![Status]2) = "On Hold" AND VALUE(NCR![NCR No]3) = "NCR-5" AND VALUE(NCR![Status]3) = "Open" AND VALUE(NCR![Status]4) = "On Hold" AND VALUE(NCR![NCR No]5) = "NCR-2" AND VALUE(NCR![Status]5) = "Closed" AND VALUE(NCR![Status]6) = "Closed" AND VALUE(NCR![NCR No]7) = "NCR-1" AND VALUE(NCR![Status]7) = "Open"
TEST     fixture=ncr_02.xlsx call=TidyNcrRegister(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
