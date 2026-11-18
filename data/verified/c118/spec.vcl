VCL 0.6
TASK
# ASSUME: Job Tracker headers in row 1 include "Job ID" and "Status"; the real sheets are "Job Tracker" and "Status Count"
# ASSUME: the Status Count sheet exists and is rebuilt on every run; a row with no Job ID is not a job
# ASSUME: statuses are matched ignoring case and stray spaces, shown as first typed (trimmed); equal counts keep no set order
GOAL     Rebuild the Status Count sheet with each job status and its number of jobs, most first, then a Total row.
PROC     CountJobsByStatus
KIND     Function
MODULE   modJobs
ARGS     jobs As Worksheet, out As Worksheet
RET      Long
VAR      cJob As Long
VAR      cStatus As Long
VAR      counts As Dictionary
VAR      got As Long
VAR      outRow As Long
VAR      total As Long
PRE      COL($jobs, "Job ID") > 0; COL($jobs, "Status") > 0
MODIFIES $out
OP       LET(cJob, COL($jobs, "Job ID"))
OP       LET(cStatus, COL($jobs, "Status"))
OP       LET(total, 0)
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Status", "Count")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($jobs, cJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($jobs, r, cJob))
OP         LET(total, total + 1)
OP         IF IS_EMPTY(CELL($jobs, r, cStatus))
OP           DICT_ADD(counts, "Blank", 1, mode=sum)
OP         ELSE
OP           DICT_ADD(counts, TRIM(CELL($jobs, r, cStatus)), 1, mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(counts)
OP         DICT_GET(counts, k, into=got)
OP         WRITE_ROW($out, outRow, k, got)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=2, order=desc, header=yes)
OP       WRITE_ROW($out, outRow, "Total", total)
OP       RETURN(total)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  jobs_01.xlsx sheet=Jobs cols=A:Job ID,B:Customer,C:Status rows=11 data=A:seq(1,1)+boundary(1,"");B:text(8);C:values("Open"x4,"Closed"x1,"closed"x1,"Closed "x1,""x2,"Hold"x1,"Ghost"x1)+boundary("Hold","Ghost","Open","Closed ","closed","","")
FIXTURE  jobs_01.xlsx sheet=Counts cols=A:Status,B:Count rows=0 stale=3
TEST     fixture=jobs_01.xlsx call=CountJobsByStatus(Jobs, Counts) runs=2 assert=RESULT = 10 AND ERROR_COUNT = 0 AND ROWS_USED(Counts) = 6 AND VALUE(Counts![Status]2) = "Open" AND VALUE(Counts![Count]2) = 4 AND VALUE(Counts![Status]3) = "Closed" AND VALUE(Counts![Count]3) = 3 AND VALUE(Counts![Status]4) = "Blank" AND VALUE(Counts![Count]4) = 2 AND VALUE(Counts![Status]5) = "Hold" AND VALUE(Counts![Count]5) = 1 AND VALUE(Counts![Status]6) = "Total" AND VALUE(Counts![Count]6) = 10 AND COUNTIF(Counts![Status], "stale") = 0 AND COUNTIF(Counts![Status], "Ghost") = 0
TEST     fixture=jobs_01.xlsx call=CountJobsByStatus(Jobs, Counts) assert=COUNTA(Jobs![Job ID]) = PRE_COUNTA(Jobs![Job ID]) AND COUNTIF(Jobs![Status], "Closed") = PRE_COUNTIF(Jobs![Status], "Closed") AND COUNTA(Jobs![Customer]) = PRE_COUNTA(Jobs![Customer])
TEST     fixture=jobs_01.xlsx call=CountJobsByStatus(Counts, Jobs) assert=RESULT = -1 AND ERROR_COUNT = 1
