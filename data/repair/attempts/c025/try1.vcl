VCL 0.6
TASK
# ASSUME: Job Tracker headers in row 1 include "Job ID", "Customer", "Due Date" and "Status"
# ASSUME: Late Jobs is an existing sheet; it is rebuilt from scratch on every run, headings included
# ASSUME: a job is late when its Due Date is a real date before today; a job due today is not late
# ASSUME: any Status other than Closed (ignoring case and spaces) counts as not Closed, a blank Status too
# ASSUME: a row with a blank Job ID is not a job and is skipped; a row without a real Due Date is not listed
# ASSUME: jobs with the same Due Date may appear in any order
GOAL     List every job that is not closed and is past its due date on the Late Jobs sheet, oldest due date first.
PROC     ListLateJobs
KIND     Function
MODULE   modLateJobs
ARGS     tracker As Worksheet, late As Worksheet
RET      Long
VAR      cId As Long
VAR      cCust As Long
VAR      cDue As Long
VAR      cStatus As Long
VAR      outRow As Long
PRE      COL($tracker, "Job ID") > 0; COL($tracker, "Customer") > 0; COL($tracker, "Due Date") > 0; COL($tracker, "Status") > 0
MODIFIES $late
OP       LET(cId, COL($tracker, "Job ID"))
OP       LET(cCust, COL($tracker, "Customer"))
OP       LET(cDue, COL($tracker, "Due Date"))
OP       LET(cStatus, COL($tracker, "Status"))
OP       CLEAR(USED_RANGE($late), what=all)
OP       WRITE_ROW($late, 1, "Job ID", "Customer", "Due Date", "Status")
OP       LET(outRow, 1)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($tracker, cId) var=r
OP         SKIP_IF IS_EMPTY(CELL($tracker, r, cId))
OP         IF NOT SAME(CELL($tracker, r, cStatus), "Closed")
OP           IF IS_DATE(CELL($tracker, r, cDue))
OP             IF CELL($tracker, r, cDue) < TODAY
OP               LET(outRow, outRow + 1)
OP               WRITE_ROW($late, outRow, CELL($tracker, r, cId), CELL($tracker, r, cCust), CELL($tracker, r, cDue), CELL($tracker, r, cStatus))
OP       IF outRow > 2 THEN SORT(USED_RANGE($late), key=3, order=asc, header=yes)
OP       RETURN(outRow - 1)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  jobs_01.xlsx sheet=JobTracker cols=A:Job ID,B:Customer,C:Due Date,D:Status rows=10 data=A:values("J-2"x1,"J-3"x1,"J-4"x1,"J-5"x1,"J-6"x1,"J-7"x1,"J-8"x1,"J-9"x1,"J-10"x1,""x1)+boundary("J-2","","J-3","J-4","J-5","J-6","J-7","J-8","J-9","J-10");B:text(8);C:reldates(-30,30)+boundary(-1,-5,0,1,-10,-3,-20,"",-2,-7);D:values("Open"x6,"Closed"x1," closed "x1,""x1,"On Hold"x1)+boundary("Open","Open","Open","Open","Closed"," closed ","","Open","On Hold","Open")
FIXTURE  jobs_01.xlsx sheet=LateJobs cols=A:Job ID,B:Customer,C:Due Date,D:Status rows=0 stale=3
FIXTURE  jobs_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  jobs_02.xlsx sheet=JobTracker cols=A:Job ID,B:Customer,C:Due Date,D:Status rows=3 data=A:values("J-1"x1,"J-2"x1,"J-3"x1)+boundary("J-1","J-2","J-3");B:text(8);C:reldates(-30,30)+boundary(-2,-9,5);D:values("Open"x3)
FIXTURE  jobs_02.xlsx sheet=LateJobs cols=A:Job ID,B:Customer,C:Due Date,D:Status rows=0 stale=3
TEST     fixture=jobs_01.xlsx call=ListLateJobs(JobTracker, LateJobs) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTA(LateJobs![Job ID]) = 5 AND COUNTIF(LateJobs![Job ID], "stale") = 0 AND VALUE(LateJobs![Job ID]2) = "J-7" AND VALUE(LateJobs![Job ID]3) = "J-10" AND VALUE(LateJobs![Job ID]4) = "J-9" AND VALUE(LateJobs![Job ID]5) = "J-2"
TEST     fixture=jobs_01.xlsx call=ListLateJobs(JobTracker, LateJobs) assert=COUNTA(JobTracker![Job ID]) = PRE_COUNTA(JobTracker![Job ID]) AND COUNTA(JobTracker![Status]) = PRE_COUNTA(JobTracker![Status]) AND COUNTIF(JobTracker![Status], "Open") = PRE_COUNTIF(JobTracker![Status], "Open") AND COUNTIF(LateJobs![Status], "On Hold") = 1 AND COUNTIF(LateJobs![Status], "Open") = 2 AND COUNTIF(LateJobs![Status], "Closed") = 0 AND ROWS_USED(LateJobs) = 5
TEST     fixture=jobs_02.xlsx call=ListLateJobs(JobTracker, LateJobs) assert=RESULT = 2 AND ERROR_COUNT = 0 AND COUNTIF(LateJobs![Job ID], "stale") = 0 AND VALUE(LateJobs![Job ID]2) = "J-2" AND VALUE(LateJobs![Job ID]3) = "J-1" AND ROWS_USED(LateJobs) = 3
TEST     fixture=jobs_01.xlsx call=ListLateJobs(Notes, LateJobs) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTIF(LateJobs![Job ID], "stale") = 3
