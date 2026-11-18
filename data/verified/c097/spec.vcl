VCL 0.6
TASK
# ASSUME: Dispatch headers in row 1 include "Customer"
# ASSUME: the customer name is matched ignoring case and stray spaces
# ASSUME: a blank customer name given to the macro changes nothing and returns 0
# ASSUME: FILTER shows only the matching rows and shows the rest again, so a re-run for another customer starts from the full list
GOAL     Show only the Dispatch rows of the given customer by hiding all others and return how many rows are visible below the headings.
PROC     ShowCustomerDispatches
KIND     Function
MODULE   modDispatch
ARGS     ws As Worksheet, customer As String
RET      Long
VAR      cCust As Long
VAR      shown As Long
PRE      COL($ws, "Customer") > 0
MODIFIES $ws
OP       EXIT_IF IS_EMPTY($customer)
OP       LET(cCust, COL($ws, "Customer"))
OP       FILTER(USED_RANGE($ws), col=cCust, crit=$customer)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cCust) var=r
OP         IF SAME(CELL($ws, r, cCust), $customer) THEN LET(shown, shown + 1)
OP       RETURN(shown)
PERF     screen_off
ERR      standard
FIXTURE  dispatch_01.xlsx sheet=Dispatch cols=A:Dispatch No,B:Customer,C:Planned Date,D:Actual Date rows=9 data=A:seq(1,1);B:values("Acme"x3,"  ACME "x1,"acme"x1,"Borg"x2,"Borgs"x1,""x1)+boundary("Acme","Borgs","  ACME ","Borg","","acme","Acme");C:dates(2026-01-01,2026-12-31);D:dates(2026-01-01,2026-12-31)+blank(2)
FIXTURE  dispatch_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=dispatch_01.xlsx call=ShowCustomerDispatches(Dispatch, "ACME") runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND VISIBLE_COUNT(Dispatch![Dispatch No]) = 6 AND COUNTA(Dispatch![Dispatch No]) = PRE_COUNTA(Dispatch![Dispatch No])
TEST     fixture=dispatch_01.xlsx call=ShowCustomerDispatches(Dispatch, " borg") assert=RESULT = 2 AND VISIBLE_COUNT(Dispatch![Dispatch No]) = 3 AND VISIBLE_COUNT(Dispatch![Customer]) = 3
TEST     fixture=dispatch_01.xlsx call=ShowCustomerDispatches(Dispatch, "Zeta") assert=RESULT = 0 AND ERROR_COUNT = 0 AND VISIBLE_COUNT(Dispatch![Dispatch No]) = 1
TEST     fixture=dispatch_01.xlsx call=ShowCustomerDispatches(Dispatch, "") assert=RESULT = 0 AND VISIBLE_COUNT(Dispatch![Dispatch No]) = PRE_VISIBLE_COUNT(Dispatch![Dispatch No])
TEST     fixture=dispatch_01.xlsx call=ShowCustomerDispatches(Notes, "Acme") assert=RESULT = -1 AND ERROR_COUNT = 1
