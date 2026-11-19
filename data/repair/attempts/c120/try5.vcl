VCL 0.6
TASK
# ASSUME: Dispatch headers in row 1 include "Dispatch No" and "Job ID"; Job Tracker headers include "Job ID" and "Status"
# ASSUME: Job IDs are matched ignoring case and stray spaces; a blank Job ID in Dispatch counts as Unknown job
# ASSUME: if a Job ID appears twice in the Job Tracker, its first row decides; Closed is matched ignoring case and spaces
# ASSUME: Job Check is added after the last used column of Dispatch when its heading is missing; the result is the number of rows marked
GOAL     Mark each dispatch row Unknown job or Job closed by checking its Job ID against the Job Tracker, clear the others and return how many need a look.
PROC     CheckDispatchJobs
KIND     Function
MODULE   modDispatch
ARGS     dispatch As Worksheet, tracker As Worksheet
RET      Long
VAR      cNo As Long
VAR      cJob As Long
VAR      cCheck As Long
VAR      cTJob As Long
VAR      cTStatus As Long
VAR      jobs As Dictionary
VAR      st As String
VAR      needLook As Long
PRE      COL($dispatch, "Dispatch No") > 0; COL($dispatch, "Job ID") > 0; COL($tracker, "Job ID") > 0; COL($tracker, "Status") > 0
MODIFIES $dispatch
OP       LET(cNo, COL($dispatch, "Dispatch No"))
OP       LET(cJob, COL($dispatch, "Job ID"))
OP       LET(cCheck, COL($dispatch, "Job Check"))
OP       LET(cTJob, COL($tracker, "Job ID"))
OP       LET(cTStatus, COL($tracker, "Status"))
OP       IF cCheck = 0
OP         LET(cCheck, LASTCOL($dispatch, 1) + 1)
OP         SET_VALUE(CELL($dispatch, 1, cCheck), "Job Check")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($tracker, cTJob) var=t
OP         SKIP_IF IS_EMPTY(CELL($tracker, t, cTJob))
OP         DICT_ADD(jobs, TRIM(CELL($tracker, t, cTJob)), TRIM(CELL($tracker, t, cTStatus)), mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($dispatch, cNo) var=r
OP         IF EXISTS(jobs, TRIM(CELL($dispatch, r, cJob)))
OP           DICT_GET(jobs, TRIM(CELL($dispatch, r, cJob)), into=st)
OP           IF SAME(st, "Closed")
OP             SET_VALUE(CELL($dispatch, r, cCheck), "Job closed")
OP             LET(needLook, needLook + 1)
OP           ELSE
OP             CLEAR(CELL($dispatch, r, cCheck), what=contents)
OP         ELSE
OP           SET_VALUE(CELL($dispatch, r, cCheck), "Unknown job")
OP           LET(needLook, needLook + 1)
OP       RETURN(needLook)
PERF     screen_off
ERR      standard
FIXTURE  disp_01.xlsx sheet=Dispatch cols=A:Dispatch No,B:Job ID,C:Planned Date,D:Job Check rows=8 data=A:seq(1,1);B:values("J-1"x1," j-2 "x1,"J-3"x1,""x1,"J-9"x1,"j-1"x1,"J-2"x1,"J-8"x1)+boundary("J-1"," j-2 ","J-3","","J-9","j-1","J-2","J-8");C:dates(2026-01-01,2026-09-30);D:values("old"x8)
FIXTURE  disp_01.xlsx sheet=Tracker cols=A:Job ID,B:Customer,C:Status rows=6 data=A:values("J-1"x1," J-2 "x1,""x1,"J-3"x1,"J-4"x1,"j-1"x1)+boundary("J-1"," J-2 ","","J-3","J-4","j-1");B:text(8);C:values("Open"x2," closed "x1,"x"x1,"Closed"x2)+boundary("Open"," closed ","x","Closed","Open","Closed")
FIXTURE  disp_02.xlsx sheet=Dispatch cols=A:Dispatch No,B:Job ID,C:Planned Date rows=3 data=A:seq(1,1);B:values("J-1"x1,"J-2"x1,"J-5"x1);C:dates(2026-01-01,2026-09-30)
FIXTURE  disp_02.xlsx sheet=Tracker cols=A:Job ID,B:Customer,C:Status rows=3 data=A:values("J-1"x1,"J-2"x1,"J-3"x1);B:text(8);C:values("Open"x1,"Closed"x1,"Open"x1)
TEST     fixture=disp_01.xlsx call=CheckDispatchJobs(Dispatch, Tracker) runs=2 assert=ERROR_COUNT = 0 AND RESULT = 6 AND COUNTIF(Dispatch![Job Check], "Unknown job") = 3 AND COUNTIF(Dispatch![Job Check], "Job closed") = 3 AND COUNTIF(Dispatch![Job Check], "old") = 0 AND COUNTA(Dispatch![Job Check]) = 7
TEST     fixture=disp_01.xlsx call=CheckDispatchJobs(Dispatch, Tracker) assert=COUNTA(Dispatch![Dispatch No]) = PRE_COUNTA(Dispatch![Dispatch No]) AND COUNTA(Tracker![Job ID]) = PRE_COUNTA(Tracker![Job ID]) AND VALUE(Dispatch![Job Check]2) = "" AND VALUE(Dispatch![Job Check]3) = "Job closed" AND VALUE(Dispatch![Job Check]5) = "Unknown job" AND VALUE(Dispatch![Job Check]7) = ""
TEST     fixture=disp_02.xlsx call=CheckDispatchJobs(Dispatch, Tracker) assert=RESULT = 2 AND ERROR_COUNT = 0 AND VALUE(Dispatch![Job Check]1) = "Job Check" AND VALUE(Dispatch![Job Check]3) = "Job closed" AND VALUE(Dispatch![Job Check]4) = "Unknown job" AND COUNTA(Dispatch![Job ID]) = PRE_COUNTA(Dispatch![Job ID])
TEST     fixture=disp_01.xlsx call=CheckDispatchJobs(Tracker, Dispatch) assert=RESULT = -1 AND ERROR_COUNT = 1
