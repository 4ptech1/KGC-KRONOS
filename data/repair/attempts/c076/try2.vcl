VCL 0.6
TASK
# ASSUME: the Job Tracker headers in row 1 include "Job ID", "Start Date", "Due Date" and "Status"; the TESTs call the sheet Jobs
# ASSUME: a missing "Flag" heading is added in the first free column of row 1
# ASSUME: Bad dates is written only when both dates are real Excel dates and Due Date is earlier than Start Date, whatever the Status
# ASSUME: No due date is written only when the Due Date cell is blank and Status is not Closed (ignoring case and extra spaces); a blank Status counts as not Closed
# ASSUME: a Due Date or Start Date that is text and not a real date gets no flag
# ASSUME: a row with a blank Job ID is not a job: its Flag is left as it is
# ASSUME: the number of flagged jobs is returned and shown in a message
GOAL     Flag jobs in the Job Tracker whose Due Date is before the Start Date or that are not Closed and have no Due Date, clear every other Flag and count the flagged jobs.
PROC     FlagJobDates
KIND     Function
MODULE   modJobTracker
ARGS     ws As Worksheet
RET      Long
VAR      cJob As Long
VAR      cStart As Long
VAR      cDue As Long
VAR      cStatus As Long
VAR      cFlag As Long
VAR      mark As String
VAR      n As Long
PRE      COL($ws, "Job ID") > 0; COL($ws, "Start Date") > 0; COL($ws, "Due Date") > 0; COL($ws, "Status") > 0
MODIFIES $ws
OP       LET(cJob, COL($ws, "Job ID"))
OP       LET(cStart, COL($ws, "Start Date"))
OP       LET(cDue, COL($ws, "Due Date"))
OP       LET(cStatus, COL($ws, "Status"))
OP       LET(cFlag, COL($ws, "Flag"))
OP       IF cFlag = 0
OP         LET(cFlag, LASTCOL($ws, 1) + 1)
OP         SET_VALUE(CELL($ws, 1, cFlag), "Flag")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cJob))
OP         LET(mark, "")
OP         IF IS_EMPTY(CELL($ws, r, cDue))
OP           IF NOT SAME(CELL($ws, r, cStatus), "Closed") THEN LET(mark, "No due date")
OP         ELSE
OP           IF IS_DATE(CELL($ws, r, cDue))
OP             IF IS_DATE(CELL($ws, r, cStart))
OP               IF CELL($ws, r, cDue) < CELL($ws, r, cStart) THEN LET(mark, "Bad dates")
OP         SET_VALUE(CELL($ws, r, cFlag), mark)
OP         IF mark <> "" THEN LET(n, n + 1)
OP       MSG("Jobs flagged: " & n)
OP       RETURN(n)
PERF     screen_off
ERR      standard
FIXTURE  jobs_01.xlsx sheet=Jobs cols=A:Job ID,B:Customer,C:Start Date,D:Due Date,E:Status,F:Flag rows=12 data=A:values("J-101"x1,"J-102"x1,"J-103"x1,"J-104"x1,"J-105"x1,"J-106"x1,""x1,"J-108"x1,"J-109"x1,"J-110"x1,"J-111"x1,"J-112"x1)+boundary("J-101","J-102","J-103","J-104","J-105","J-106","","J-108","J-109","J-110","J-111","J-112");B:text(8);C:dates(2026-03-01,2026-03-31)+boundary(2026-03-10,2026-03-10,2026-03-10,2026-03-10,2026-03-10,2026-04-20,2026-03-10,2026-03-10,"",2026-03-10,2026-05-02,"");D:dates(2026-03-01,2026-03-31)+boundary(2026-03-09,2026-03-10,2026-03-11,"","",2026-04-01,"",2026-03-20,2026-03-05,2026-03-10,2026-05-01,"");E:values("Open"x5,"Closed"x2," closed "x1,"In Progress"x3,""x1)+boundary("Open","Open","Closed","Open"," closed ","Closed","Open","In Progress","In Progress","Open","","In Progress");F:values(""x4,"Bad dates"x2,"old"x3,"No due date"x2,"keep"x1)+boundary("","Bad dates","old","","No due date","","keep","old","Bad dates","No due date","old","")
FIXTURE  jobs_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  jobs_02.xlsx sheet=Jobs cols=A:Job ID,B:Customer,C:Start Date,D:Due Date,E:Status rows=4 data=A:seq(201,1);B:text(8);C:dates(2026-06-01,2026-06-30)+boundary(2026-06-10,2026-06-10,2026-06-10,2026-06-10);D:dates(2026-06-01,2026-06-30)+boundary(2026-06-09,2026-06-10,"","");E:values(""x1,"Open"x2,"Closed"x1)+boundary("","Open","Open","Closed")
TEST     fixture=jobs_01.xlsx call=FlagJobDates(Jobs) runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND COUNTIF(Jobs![Flag], "Bad dates") = 3 AND COUNTIF(Jobs![Flag], "No due date") = 2 AND COUNTIF(Jobs![Flag], "keep") = 1 AND COUNTA(Jobs![Flag]) = 7 AND VALUE(Jobs![Flag]1) = "Flag" AND VALUE(Jobs![Flag]2) = "Bad dates" AND VALUE(Jobs![Flag]5) = "No due date" AND VALUE(Jobs![Flag]7) = "Bad dates" AND VALUE(Jobs![Flag]8) = "keep" AND VALUE(Jobs![Flag]12) = "Bad dates" AND VALUE(Jobs![Flag]13) = "No due date" AND COUNTA(Jobs!G:G) = 0
TEST     fixture=jobs_01.xlsx call=FlagJobDates(Jobs) assert=RESULT = 5 AND ERROR_COUNT = 0 AND COUNTA(Jobs![Job ID]) = PRE_COUNTA(Jobs![Job ID]) AND COUNTA(Jobs![Start Date]) = PRE_COUNTA(Jobs![Start Date]) AND COUNTA(Jobs![Due Date]) = PRE_COUNTA(Jobs![Due Date]) AND COUNTA(Jobs![Status]) = PRE_COUNTA(Jobs![Status]) AND COUNTIF(Jobs![Status], "Closed") = PRE_COUNTIF(Jobs![Status], "Closed")
TEST     fixture=jobs_02.xlsx call=FlagJobDates(Jobs) runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND VALUE(Jobs!F1) = "Flag" AND VALUE(Jobs!E1) = "Status" AND VALUE(Jobs!F2) = "Bad dates" AND VALUE(Jobs!F4) = "No due date" AND COUNTA(Jobs!F:F) = 3 AND COUNTA(Jobs!G:G) = 0 AND COUNTA(Jobs![Status]) = PRE_COUNTA(Jobs![Status])
TEST     fixture=jobs_01.xlsx call=FlagJobDates(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
