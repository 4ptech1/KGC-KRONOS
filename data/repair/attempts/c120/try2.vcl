VCL 0.6
TASK
# ASSUME: Dispatch headers in row 1 include "Dispatch No" and "Job ID"; Job Tracker headers include "Job ID" and "Status"; the real sheets are "Dispatch" and "Job Tracker"
# ASSUME: Job IDs are matched with CLEAN_TEXT, ignoring case; a Dispatch row with no Job ID is an Unknown job
# ASSUME: a Dispatch row with no Dispatch No is left as it is; a job found twice in the tracker uses its first row
# ASSUME: Job closed means the tracker Status is Closed, ignoring case and spaces; the Job Check heading is added after the last column when missing
# ASSUME: the number needing a look is the rows marked Unknown job plus the rows marked Job closed
GOAL     Mark each Dispatch row whose job is not in the Job Tracker or is closed there, clear the mark on the others, and return how many need a look.
PROC     CheckDispatchJobs
KIND     Function
MODULE   modDispatch
ARGS     dispatch As Worksheet, tracker As Worksheet
RET      Long
VAR      cNo As Long
VAR      cJob As Long
VAR      cCheck As Long
VAR      tJob As Long
VAR      tStatus As Long
VAR      states As Dictionary
VAR      state As String
VAR      look As Long
PRE      COL($dispatch, "Dispatch No") > 0; COL($dispatch, "Job ID") > 0; COL($tracker, "Job ID") > 0; COL($tracker, "Status") > 0
MODIFIES $dispatch
OP       LET(cNo, COL($dispatch, "Dispatch No"))
OP       LET(cJob, COL($dispatch, "Job ID"))
OP       LET(tJob, COL($tracker, "Job ID"))
OP       LET(tStatus, COL($tracker, "Status"))
OP       LET(cCheck, COL($dispatch, "Job Check"))
OP       IF cCheck = 0
OP         LET(cCheck, LASTCOL($dispatch, 1) + 1)
OP         SET_VALUE(CELL($dispatch, 1, cCheck), "Job Check")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($tracker, tJob) var=t
OP         SKIP_IF IS_EMPTY(CELL($tracker, t, tJob))
OP         DICT_ADD(states, CLEAN_TEXT(CELL($tracker, t, tJob)), IIF(SAME(CELL($tracker, t, tStatus), "Closed"), "Closed", "Other"), mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($dispatch, cNo) var=r
OP         SKIP_IF IS_EMPTY(CELL($dispatch, r, cNo))
OP         SET_VALUE(CELL($dispatch, r, cCheck), "")
OP         IF IS_EMPTY(CELL($dispatch, r, cJob))
OP           SET_VALUE(CELL($dispatch, r, cCheck), "Unknown job")
OP           LET(look, look + 1)
OP         ELSE
OP           IF EXISTS(states, CLEAN_TEXT(CELL($dispatch, r, cJob)))
OP             DICT_GET(states, CLEAN_TEXT(CELL($dispatch, r, cJob)), into=state)
OP             IF state = "Closed"
OP               SET_VALUE(CELL($dispatch, r, cCheck), "Job closed")
OP               LET(look, look + 1)
OP           ELSE
OP             SET_VALUE(CELL($dispatch, r, cCheck), "Unknown job")
OP             LET(look, look + 1)
OP       RETURN(look)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  dispatch_01.xlsx sheet=Dispatch cols=A:Dispatch No,B:Job ID,C:Planned Date,D:Job Check rows=10 stale=2 data=A:seq(1,1)+boundary("");B:values("J-999"x1," j-100 "x1,"J-101"x2,"j-102"x1,"J-998"x1,""x1,"J-103"x1,"J-100"x1,"J-104"x1)+boundary("J-999"," j-100 ","J-101","j-102","J-998","","J-103","J-100","J-104","J-101");C:dates(2026-01-01,2026-03-31);D:values("old"x9,"keep"x1)+boundary("keep")
FIXTURE  dispatch_01.xlsx sheet=Tracker cols=A:Job ID,B:Customer,C:Status rows=7 stale=1 data=A:values("J-100"x1,"J-101"x1,"J-102"x1,"J-103"x1,"J-104"x1,"j-100"x1,""x1)+boundary("J-100"," J-101 ","J-102","J-103","J-104","j-100","");B:text(8);C:values("Open"x1,"closed "x1,"Closed"x4,""x1)+boundary("Open","closed ","Closed","","Closed","Closed","Closed")
FIXTURE  dispatch_02.xlsx sheet=Dispatch cols=A:Dispatch No,B:Job ID,C:Planned Date rows=3 stale=1 data=A:seq(1,1);B:values("J-100"x1,"J-101"x1,"J-555"x1)+boundary("J-100","J-101","J-555");C:dates(2026-01-01,2026-03-31)
FIXTURE  dispatch_02.xlsx sheet=Tracker cols=A:Job ID,B:Customer,C:Status rows=2 stale=1 data=A:values("J-100"x1,"J-101"x1)+boundary("J-100","J-101");B:text(8);C:values("Open"x1,"Closed"x1)+boundary("Open","Closed")
TEST     fixture=dispatch_01.xlsx call=CheckDispatchJobs(Dispatch, Tracker) runs=2 assert=RESULT = 6 AND ERROR_COUNT = 0 AND COUNTIF(Dispatch![Job Check], "Unknown job") = 2 AND COUNTIF(Dispatch![Job Check], "Job closed") = 4 AND COUNTIF(Dispatch![Job Check], "old") = 0 AND COUNTIF(Dispatch![Job Check], "keep") = 1 AND VALUE(Dispatch![Job Check]3) = "" AND VALUE(Dispatch![Job Check]4) = "Job closed" AND VALUE(Dispatch![Job Check]6) = "Unknown job" AND VALUE(Dispatch![Job Check]8) = "" AND VALUE(Dispatch![Job Check]9) = ""
TEST     fixture=dispatch_01.xlsx call=CheckDispatchJobs(Dispatch, Tracker) assert=COUNTA(Dispatch![Dispatch No]) = PRE_COUNTA(Dispatch![Dispatch No]) AND COUNTA(Dispatch![Job ID]) = PRE_COUNTA(Dispatch![Job ID]) AND COUNTA(Tracker![Job ID]) = PRE_COUNTA(Tracker![Job ID]) AND COUNTIF(Tracker![Status], "Closed") = PRE_COUNTIF(Tracker![Status], "Closed")
TEST     fixture=dispatch_02.xlsx call=CheckDispatchJobs(Dispatch, Tracker) assert=RESULT = 2 AND ERROR_COUNT = 0 AND VALUE(Dispatch![Job Check]1) = "Job Check" AND VALUE(Dispatch![Job Check]2) = "" AND VALUE(Dispatch![Job Check]3) = "Job closed" AND VALUE(Dispatch![Job Check]4) = "Unknown job" AND COUNTA(Dispatch![Dispatch No]) = PRE_COUNTA(Dispatch![Dispatch No])
TEST     fixture=dispatch_01.xlsx call=CheckDispatchJobs(Tracker, Dispatch) assert=RESULT = -1 AND ERROR_COUNT = 1
