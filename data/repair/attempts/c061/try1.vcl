VCL 0.6
TASK
# ASSUME: Dispatch headers in row 1 include "Dispatch No", "Job ID" and "Remark"; Job Tracker headers include "Job ID" and "Status"
# ASSUME: the real sheets are "Dispatch" and "Job Tracker"; the TESTs use Dispatch and JobTracker
# ASSUME: the Remark column already exists on the Dispatch sheet; if it is missing the macro stops and writes a log entry
# ASSUME: Job IDs are cleaned before matching and compared ignoring case; a job is Closed only when every Job Tracker row for it says Closed
# ASSUME: a dispatch with a blank Job ID, or one whose job is not in the Job Tracker, gets the remark
# ASSUME: rows with a blank Dispatch No are not dispatches and are left untouched; on every other row the Remark becomes the remark or is cleared, whatever it held
GOAL     Write Job not closed in the Remark column of each dispatch whose job is not Closed or is missing from the Job Tracker, clear the other remarks and report how many got it.
PROC     FlagDispatchesWithOpenJobs
KIND     Function
MODULE   modDispatch
ARGS     dispatch As Worksheet, tracker As Worksheet
RET      Long
VAR      cNo As Long
VAR      cJob As Long
VAR      cRemark As Long
VAR      cTrackJob As Long
VAR      cStatus As Long
VAR      closed As Dictionary
VAR      notClosed As Dictionary
VAR      flagged As Long
PRE      COL($dispatch, "Dispatch No") > 0; COL($dispatch, "Job ID") > 0; COL($dispatch, "Remark") > 0; COL($tracker, "Job ID") > 0; COL($tracker, "Status") > 0
MODIFIES $dispatch["Remark"]
OP       LET(cNo, COL($dispatch, "Dispatch No"))
OP       LET(cJob, COL($dispatch, "Job ID"))
OP       LET(cRemark, COL($dispatch, "Remark"))
OP       LET(cTrackJob, COL($tracker, "Job ID"))
OP       LET(cStatus, COL($tracker, "Status"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($tracker, cTrackJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($tracker, r, cTrackJob))
OP         IF SAME(CELL($tracker, r, cStatus), "Closed")
OP           DICT_ADD(closed, CLEAN_TEXT(CELL($tracker, r, cTrackJob)), 1, mode=keep)
OP         ELSE
OP           DICT_ADD(notClosed, CLEAN_TEXT(CELL($tracker, r, cTrackJob)), 1, mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($dispatch, cNo) var=r
OP         SKIP_IF IS_EMPTY(CELL($dispatch, r, cNo))
OP         IF EXISTS(notClosed, CLEAN_TEXT(CELL($dispatch, r, cJob))) OR NOT EXISTS(closed, CLEAN_TEXT(CELL($dispatch, r, cJob)))
OP           SET_VALUE(CELL($dispatch, r, cRemark), "Job not closed")
OP           LET(flagged, flagged + 1)
OP         ELSE
OP           CLEAR(CELL($dispatch, r, cRemark), what=contents)
OP       MSG("Dispatches marked Job not closed: " & flagged)
OP       RETURN(flagged)
PERF     screen_off
ERR      standard
FIXTURE  dispatch_01.xlsx sheet=Dispatch cols=A:Dispatch No,B:Job ID,C:Customer,D:Remark rows=12 data=A:values("D-1"x1,"D-2"x1,"D-3"x1,"D-4"x1,"D-5"x1,"D-6"x1,"D-7"x1,"D-8"x1,""x1,"D-10"x1,"D-11"x1,"D-12"x1)+boundary("D-1","D-2","D-3","D-4","D-5","D-6","D-7","D-8","","D-10","D-11","D-12");B:values("J-1"x3," J-2"x1,"J-3"x2,"j-4"x1,"J-5"x1,"J-6"x1,"J-8"x1,"J-9"x1,""x1)+boundary("J-1"," J-2","J-3","j-4","J-5","J-6","J-8","J-9","J-3","","J-1","J-1");C:text(10);D:values("Job not closed"x4,"ok"x1,""x5,"old"x1,"keep"x1)+boundary("Job not closed","ok","","Job not closed","Job not closed","","","old","keep","","Job not closed","")
FIXTURE  dispatch_01.xlsx sheet=JobTracker cols=A:Job ID,B:Status rows=10 data=A:values("J-1"x1," j-2 "x1,"J-3"x1,"J-4"x1,""x1,"J-5"x1,"J-6"x2,"J-8"x2)+boundary("J-1"," j-2 ","J-3","J-4","","J-5","J-6","J-6","J-8","J-8");B:values("Closed"x5,"Open"x3,"closed "x1,""x1)+boundary("Closed","Closed","Open","closed ","Closed","","Closed","Open","Open","Closed")
TEST     fixture=dispatch_01.xlsx call=FlagDispatchesWithOpenJobs(Dispatch, JobTracker) runs=2 assert=RESULT = 6 AND ERROR_COUNT = 0 AND COUNTIF(Dispatch![Remark], "Job not closed") = 6 AND COUNTIF(Dispatch![Remark], "keep") = 1 AND COUNTA(Dispatch![Remark]) = 8 AND VALUE(Dispatch![Remark]2) = "" AND VALUE(Dispatch![Remark]4) = "Job not closed" AND VALUE(Dispatch![Remark]11) = "Job not closed" AND VALUE(Dispatch![Remark]12) = ""
TEST     fixture=dispatch_01.xlsx call=FlagDispatchesWithOpenJobs(Dispatch, JobTracker) assert=COUNTIFS(Dispatch![Remark], "Job not closed", Dispatch![Job ID], "J-1") = 0 AND COUNTIFS(Dispatch![Remark], "Job not closed", Dispatch![Job ID], "J-9") = 1 AND COUNTA(Dispatch![Dispatch No]) = PRE_COUNTA(Dispatch![Dispatch No]) AND COUNTA(Dispatch![Job ID]) = PRE_COUNTA(Dispatch![Job ID]) AND COUNTA(Dispatch![Customer]) = PRE_COUNTA(Dispatch![Customer]) AND COUNTA(JobTracker![Status]) = PRE_COUNTA(JobTracker![Status])
TEST     fixture=dispatch_01.xlsx call=FlagDispatchesWithOpenJobs(JobTracker, Dispatch) assert=RESULT = -1 AND ERROR_COUNT = 1
