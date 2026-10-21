VCL 0.6
TASK
# ASSUME: Job Tracker headers in row 1 include "Job ID" and "Folder Status"
# ASSUME: the Jobs folder exists and is given as a full path; each folder is made directly inside it
# ASSUME: Folder Status gets the folder op's own status: Created (made), Exists (already there), or a problem such as Bad name or Blocked by file
# ASSUME: a blank Job ID is skipped and its Folder Status is left as it is; Job IDs are trimmed to make the folder name
# ASSUME: the same Job ID on two rows makes one folder; the later row reports Exists
GOAL     Make a folder inside the Jobs folder for every Job ID in the Job Tracker and write what happened for each row in the Folder Status column.
PROC     MakeJobFolders
KIND     Function
MODULE   modJobFolders
ARGS     tracker As Worksheet, jobs As String
RET      Long
VAR      cId As Long
VAR      cStatus As Long
VAR      status As String
VAR      made As Long
PRE      COL($tracker, "Job ID") > 0; COL($tracker, "Folder Status") > 0; FOLDER_EXISTS($jobs)
MODIFIES $tracker["Folder Status"]
MODIFIES $jobs files
OP       LET(cId, COL($tracker, "Job ID"))
OP       LET(cStatus, COL($tracker, "Folder Status"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($tracker, cId) var=r
OP         SKIP_IF IS_EMPTY(CELL($tracker, r, cId))
OP         IF SAFE_NAME(CELL($tracker, r, cId)) = ""
OP           SET_VALUE(CELL($tracker, r, cStatus), "Bad name")
OP         ELSE
OP           MAKE_FOLDER(PATH_JOIN($jobs, SAFE_NAME(CELL($tracker, r, cId))), into=status)
OP           SET_VALUE(CELL($tracker, r, cStatus), status)
OP           IF status = "Created" THEN LET(made, made + 1)
OP       RETURN(made)
PERF     screen_off
ERR      standard
FIXTURE  jobfolders_01.xlsx sheet=JobTracker cols=A:Job ID,B:Customer,C:Folder Status rows=9 data=A:values("J-1"x2,"J-2"x1,"J-3"x1," J-4 "x1,"J-5"x1,"A:B"x1,""x1,"J-6"x1)+boundary("J-1","J-2","J-1","","A:B","J-3"," J-4 ","J-5","J-6");B:text(8);C:values("old"x9)
FIXTURE  jobfolders_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  jobfolders_01.xlsx folder=Jobs files=J-2\a.txt:10,J-5:20
TEST     fixture=jobfolders_01.xlsx call=MakeJobFolders(JobTracker, Jobs) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTIF(JobTracker![Folder Status], "Exists") = 6 AND COUNTIF(JobTracker![Folder Status], "Bad name") = 1 AND COUNTIF(JobTracker![Folder Status], "Blocked by file") = 1 AND COUNTIF(JobTracker![Folder Status], "old") = 1 AND COUNTIF(JobTracker![Folder Status], "Created") = 0
TEST     fixture=jobfolders_01.xlsx call=MakeJobFolders(JobTracker, Jobs) assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTIF(JobTracker![Folder Status], "Created") = 4 AND COUNTIF(JobTracker![Folder Status], "Exists") = 2 AND FILE_COUNT("Jobs") = PRE_FILE_COUNT("Jobs") AND COUNTA(JobTracker![Job ID]) = PRE_COUNTA(JobTracker![Job ID])
TEST     fixture=jobfolders_01.xlsx call=MakeJobFolders(Notes, Jobs) assert=RESULT = -1 AND ERROR_COUNT = 1 AND FILE_COUNT("Jobs") = PRE_FILE_COUNT("Jobs")
