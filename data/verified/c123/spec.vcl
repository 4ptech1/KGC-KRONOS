VCL 0.6
TASK
# ASSUME: Job Tracker headers in row 1 include "Job ID", "Customer" and "Status"
# ASSUME: spaces means spaces, tabs and line breaks; a Job ID that is already clean is left untouched and not counted
# ASSUME: blank Job ID cells, and cells holding only spaces, are skipped and left as they are
GOAL     Make every Job ID in the Job Tracker trimmed, upper case and with single spaces inside, and return how many cells were changed.
PROC     CleanJobIds
KIND     Function
MODULE   modJobs
ARGS     ws As Worksheet
RET      Long
VAR      cJob As Long
VAR      clean As String
VAR      changed As Long
PRE      COL($ws, "Job ID") > 0
MODIFIES $ws["Job ID"]
OP       LET(cJob, COL($ws, "Job ID"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cJob))
OP         LET(clean, UPPER(CLEAN_TEXT(CELL($ws, r, cJob))))
OP         IF CELL($ws, r, cJob) <> clean
OP           SET_VALUE(CELL($ws, r, cJob), clean)
OP           LET(changed, changed + 1)
OP       RETURN(changed)
PERF     screen_off
ERR      standard
FIXTURE  jobs_01.xlsx sheet=Jobs cols=A:Job ID,B:Customer,C:Status rows=10 data=A:values("job  101 "x1,"JOB 102"x1,"job 103"x1," Job   104"x1,"JOB 105"x1,""x1,"   "x1,"Job 106"x1,"a  b  c"x1,"JOB 107"x1)+boundary("job  101 ","JOB 102","job 103"," Job   104","JOB 105","","   ","Job 106","a  b  c","JOB 107");B:values(" acme  ltd "x10);C:values("open "x10)
FIXTURE  jobs_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=jobs_01.xlsx call=CleanJobIds(Jobs) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND VALUE(Jobs![Job ID]2) = "JOB 101" AND VALUE(Jobs![Job ID]4) = "JOB 103" AND VALUE(Jobs![Job ID]5) = "JOB 104" AND VALUE(Jobs![Job ID]10) = "A B C"
TEST     fixture=jobs_01.xlsx call=CleanJobIds(Jobs) assert=RESULT = 5 AND COUNTA(Jobs![Job ID]) = PRE_COUNTA(Jobs![Job ID]) AND COUNTA(Jobs![Customer]) = PRE_COUNTA(Jobs![Customer]) AND VALUE(Jobs![Customer]2) = " acme  ltd " AND VALUE(Jobs![Status]3) = "open " AND VALUE(Jobs![Job ID]3) = "JOB 102" AND VALUE(Jobs![Job ID]7) = "" AND VALUE(Jobs![Job ID]9) = "JOB 106" AND VALUE(Jobs![Job ID]8) = "   "
TEST     fixture=jobs_01.xlsx call=CleanJobIds(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(Jobs![Job ID]) = PRE_COUNTA(Jobs![Job ID])
