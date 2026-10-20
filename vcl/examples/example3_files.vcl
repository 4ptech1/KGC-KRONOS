VCL 0.4
TASK
# ASSUME: Tracker headers in row 1 include "Job ID", "Report File" and "Status"
# ASSUME: both folders exist; each report is copied, never moved, and nothing is overwritten
# ASSUME: a blank Job ID is skipped; an unsafe Job ID or Report File gets Status "Bad name"
PROC     FileJobReports
KIND     Function
MODULE   modReports
ARGS     tracker As Worksheet, dump As String, jobs As String
RET      Long
VAR      cJob As Long
VAR      cFile As Long
VAR      cStatus As Long
VAR      status As String
VAR      filed As Long
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
OP           IF status = "Copied" OR status = "Already there" THEN LET(filed, filed + 1)
OP       RETURN(filed)
PERF     screen_off
ERR      standard
FIXTURE  reports_01.xlsx sheet=Tracker cols=A:Job ID,B:Report File,C:Status rows=7 data=A:values("J-1"x2,"J-2"x1,"..\x"x1,""x1,"J-3"x1,"J-4"x1)+boundary("J-1","J-2","..\x","","J-3","J-1","J-4");B:values("a.pdf"x2,"b.pdf"x2,"c.pdf"x1,"d.pdf"x1,"CON.pdf"x1)+boundary("a.pdf","b.pdf","a.pdf","b.pdf","c.pdf","d.pdf","CON.pdf");C:values("old"x7)
FIXTURE  reports_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  reports_01.xlsx folder=Dump files=a.pdf:100,b.pdf:80,c.pdf locked=c.pdf
FIXTURE  reports_01.xlsx folder=Jobs files=J-2\b.pdf:99
TEST     fixture=reports_01.xlsx call=FileJobReports(Tracker, Dump, Jobs) runs=2 assert=RESULT = 1 AND ERROR_COUNT = 0 AND COUNTIF(Tracker![Status], "Already there") = 1 AND COUNTIF(Tracker![Status], "Conflict") = 1 AND COUNTIF(Tracker![Status], "Locked") = 1 AND COUNTIF(Tracker![Status], "Missing") = 1 AND COUNTIF(Tracker![Status], "Bad name") = 2 AND FILE_EXISTS("Jobs\J-1\a.pdf") AND FILE_COUNT("Jobs") = PRE_FILE_COUNT("Jobs") + 1
TEST     fixture=reports_01.xlsx call=FileJobReports(Tracker, Dump, Jobs) assert=COUNTIF(Tracker![Status], "Copied") = 1 AND COUNTIF(Tracker![Status], "old") = 1 AND FILE_COUNT("Dump") = PRE_FILE_COUNT("Dump")
TEST     fixture=reports_01.xlsx call=FileJobReports(Notes, Dump, Jobs) assert=RESULT = -1 AND ERROR_COUNT = 1 AND FILE_COUNT("Jobs") = PRE_FILE_COUNT("Jobs")
