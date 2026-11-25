VCL 0.7
TASK
# ASSUME: the Job Tracker headings are in row 1; rows with a blank Job ID are skipped
# ASSUME: Status is matched with SAME; Percent Complete is a number from 0 to 100, and a row whose percent is not a number is not flagged (its Flag is cleared)
# ASSUME: the Flag heading is added in the next free column when it is missing
GOAL     Flag each job whose status disagrees with its percent complete, clear the Flag on the others, and count the jobs flagged.
PROC     FlagJobStatus
KIND     Function
MODULE   modJobs
ARGS     ws As Worksheet
RET      Long
VAR      cJob As Long
VAR      cStatus As Long
VAR      cPct As Long
VAR      cFlag As Long
VAR      mark As Long
VAR      flagged As Long
PRE      COL($ws, "Job ID") > 0; COL($ws, "Status") > 0; COL($ws, "Percent Complete") > 0
MODIFIES $ws
OP       LET(cJob, COL($ws, "Job ID"))
OP       LET(cStatus, COL($ws, "Status"))
OP       LET(cPct, COL($ws, "Percent Complete"))
OP       LET(cFlag, COL($ws, "Flag"))
OP       IF cFlag = 0
OP         LET(cFlag, LASTCOL($ws, 1) + 1)
OP         SET_VALUE(CELL($ws, 1, cFlag), "Flag")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cJob))
OP         LET(mark, 0)
OP         IF IS_NUMBER(CELL($ws, r, cPct))
OP           IF CELL($ws, r, cPct) = 100 AND NOT SAME(CELL($ws, r, cStatus), "Complete")
OP             LET(mark, 1)
OP           IF CELL($ws, r, cPct) < 100 AND SAME(CELL($ws, r, cStatus), "Complete")
OP             LET(mark, 1)
OP         IF mark = 1
OP           SET_VALUE(CELL($ws, r, cFlag), "Check status")
OP           LET(flagged, flagged + 1)
OP         ELSE
OP           CLEAR(CELL($ws, r, cFlag), what=contents)
OP       RETURN(flagged)
ERR      standard
FIXTURE  jobs_01.xlsx sheet=Jobs cols=A:Job ID,B:Customer,C:Status,D:Percent Complete,E:Flag rows=7 data=A:seq(1,1);B:text(6);C:values("Open"x2,"Complete"x4,"In Progress"x1)+boundary("Open","Complete","Complete","Open","Complete","In Progress","Complete");D:int(0,100)+boundary(100,100,99,99,"",100,50);E:values("old"x7)
FIXTURE  jobs_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  jobs_02.xlsx sheet=Jobs cols=A:Job ID,B:Customer,C:Status,D:Percent Complete rows=7 data=A:seq(1,1);B:text(6);C:values("Open"x2,"Complete"x4,"In Progress"x1)+boundary("Open","Complete","Complete","Open","Complete","In Progress","Complete");D:int(0,100)+boundary(100,100,99,99,"",100,50)
FIXTURE  jobs_02.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=jobs_01.xlsx call=FlagJobStatus(Jobs) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTIF(Jobs![Flag], "Check status") = 4 AND COUNTIF(Jobs![Flag], "old") = 0 AND VALUE(Jobs![Flag]2) = "Check status" AND VALUE(Jobs![Flag]3) = "" AND VALUE(Jobs![Flag]5) = "" AND VALUE(Jobs![Flag]6) = ""
TEST     fixture=jobs_01.xlsx call=FlagJobStatus(Jobs) assert=COUNTA(Jobs![Job ID]) = PRE_COUNTA(Jobs![Job ID]) AND COUNTA(Jobs![Status]) = PRE_COUNTA(Jobs![Status]) AND COUNTA(Jobs![Percent Complete]) = PRE_COUNTA(Jobs![Percent Complete])
TEST     fixture=jobs_02.xlsx call=FlagJobStatus(Jobs) assert=RESULT = 4 AND ERROR_COUNT = 0 AND VALUE(Jobs![Flag]1) = "Flag" AND COUNTIF(Jobs![Flag], "Check status") = 4
TEST     fixture=jobs_01.xlsx call=FlagJobStatus(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
