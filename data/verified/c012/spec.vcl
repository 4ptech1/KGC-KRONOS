VCL 0.6
TASK
# ASSUME: Tracker headers are in row 1 and include "Job ID", "Report File" and "Status"; column order does not matter
# ASSUME: $dump is the folder the test reports arrive in, and $jobs holds one folder per job; both must exist (sandbox folders only until March 2027)
# ASSUME: each report is copied into its job's folder, never moved: the file in $dump always stays, and nothing is ever overwritten
# ASSUME: a row with a blank Job ID is skipped and its Status is left as it is
# ASSUME: a Job ID or Report File that is not a safe file name gets Status "Bad name", and nothing is made or copied for that row
# ASSUME: Status gets the file op's status word (Copied, Already there, Conflict, Locked, Missing, Blocked by file, ...)
# ASSUME: the result counts the rows whose report is now in its job's folder (Copied or Already there), so a re-run gives the same result
GOAL     Copy each job's test report from the dump folder into that job's own folder and write the outcome of each copy in the Status column.
PROC     RouteJobPacks
KIND     Function
MODULE   modRouter
ARGS     tracker As Worksheet, dump As String, jobs As String
RET      Long
VAR      cJob As Long
VAR      cFile As Long
VAR      cStatus As Long
VAR      status As String
VAR      routed As Long
PRE      COL($tracker, "Job ID") > 0; COL($tracker, "Report File") > 0; COL($tracker, "Status") > 0; FOLDER_EXISTS($dump); FOLDER_EXISTS($jobs)
MODIFIES $tracker["Status"]
MODIFIES $jobs files
OP       LET(cJob, COL($tracker, "Job ID"))
OP       LET(cFile, COL($tracker, "Report File"))
OP       LET(cStatus, COL($tracker, "Status"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($tracker, cJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($tracker, r, cJob))
OP         IF SAFE_NAME(CELL($tracker, r, cJob)) = "" OR SAFE_NAME(CELL($tracker, r, cFile)) = ""
OP           SET_VALUE(CELL($tracker, r, cStatus), "Bad name")
OP         ELSE
OP           MAKE_FOLDER(PATH_JOIN($jobs, SAFE_NAME(CELL($tracker, r, cJob))), into=status)
OP           IF status = "Created" OR status = "Exists"
OP             COPY_FILE(PATH_JOIN($dump, SAFE_NAME(CELL($tracker, r, cFile))), PATH_JOIN($jobs, SAFE_NAME(CELL($tracker, r, cJob)), SAFE_NAME(CELL($tracker, r, cFile))), into=status)
OP           SET_VALUE(CELL($tracker, r, cStatus), status)
OP           IF status = "Copied" OR status = "Already there" THEN LET(routed, routed + 1)
OP       RETURN(routed)
PERF     screen_off
ERR      standard
FIXTURE  jobpacks_01.xlsx sheet=Tracker cols=A:Job ID,B:Report File,C:Status rows=10 data=A:values("J-101"x2,"J-102"x1,"J-103"x1,"..\x"x1,"J-104"x1,""x1,"J-105"x1,"J-106"x1,"J-107"x1)+boundary("J-101","J-102","J-103","..\x","J-104","","J-105","J-101","J-106","J-107");B:values("r1.pdf"x2,"r2.pdf"x1,"r3.pdf"x1,"r4.pdf"x2,"r5.pdf"x1,"r9.pdf"x1,"CON.pdf"x1,"r6.pdf"x1)+boundary("r1.pdf","r2.pdf","r3.pdf","r4.pdf","r5.pdf","r1.pdf","r9.pdf","r4.pdf","CON.pdf","r6.pdf");C:values("old"x10)
FIXTURE  jobpacks_01.xlsx sheet=Notes cols=A:Note rows=3 data=A:text(10)
FIXTURE  jobpacks_01.xlsx folder=Dump files=r1.pdf:120,r2.pdf:80,r3.pdf,r4.pdf:64,r5.pdf,r6.pdf:70 locked=r3.pdf
FIXTURE  jobpacks_01.xlsx folder=Jobs files=J-102\r2.pdf:99,J-104:5,J-107\r6.pdf:70
TEST     fixture=jobpacks_01.xlsx call=RouteJobPacks(Tracker, Dump, Jobs) assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(Tracker![Status], "Copied") = 2 AND COUNTIF(Tracker![Status], "Already there") = 1 AND COUNTIF(Tracker![Status], "Conflict") = 1 AND COUNTIF(Tracker![Status], "Locked") = 1 AND COUNTIF(Tracker![Status], "Missing") = 1 AND COUNTIF(Tracker![Status], "Blocked by file") = 1 AND COUNTIF(Tracker![Status], "Bad name") = 2 AND COUNTIF(Tracker![Status], "old") = 1 AND FILE_EXISTS("Jobs\J-101\r1.pdf") AND FILE_EXISTS("Jobs\J-101\r4.pdf") AND FILE_COUNT("Jobs") = PRE_FILE_COUNT("Jobs") + 2 AND FILE_COUNT("Dump") = PRE_FILE_COUNT("Dump")
TEST     fixture=jobpacks_01.xlsx call=RouteJobPacks(Tracker, Dump, Jobs) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(Tracker![Status], "Already there") = 3 AND COUNTIF(Tracker![Status], "Copied") = 0 AND COUNTIF(Tracker![Status], "Locked") = 1 AND COUNTIF(Tracker![Status], "Missing") = 1 AND FILE_COUNT("Jobs") = PRE_FILE_COUNT("Jobs") + 2
TEST     fixture=jobpacks_01.xlsx call=RouteJobPacks(Tracker, Dump, "X:\KRONOS_NO_SUCH_FOLDER") assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTIF(Tracker![Status], "old") = 10 AND FILE_COUNT("Jobs") = PRE_FILE_COUNT("Jobs")
TEST     fixture=jobpacks_01.xlsx call=RouteJobPacks(Notes, Dump, Jobs) assert=RESULT = -1 AND ERROR_COUNT = 1 AND FILE_COUNT("Jobs") = PRE_FILE_COUNT("Jobs") AND FILE_COUNT("Dump") = PRE_FILE_COUNT("Dump")
