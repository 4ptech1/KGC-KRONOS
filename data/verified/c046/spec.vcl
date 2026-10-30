VCL 0.6
TASK
# ASSUME: the Job Folder path is the String argument jobFolder and the folder exists; if it does not, nothing is written
# ASSUME: status.txt is read as a text file; B1 of Jobs is overwritten, header cell or not
# ASSUME: if the file cannot be read for another reason (Locked, Too big, Not text), B1 gets that status word instead
# ASSUME: the macro returns 1 when the text was read and 0 when B1 got a status word
GOAL     Read the text of status.txt in the job folder into cell B1 of the Jobs sheet, or write the status Missing there when the file is not found.
PROC     ReadJobStatusFile
KIND     Function
MODULE   modJobs
ARGS     jobs As Worksheet, jobFolder As String
RET      Long
VAR      txt As String
VAR      status As String
VAR      readOk As Long
PRE      FOLDER_EXISTS($jobFolder)
MODIFIES $jobs
OP       READ_TEXT(PATH_JOIN($jobFolder, "status.txt"), txt, into=status)
OP       IF status = "Read"
OP         SET_VALUE(CELL($jobs, 1, 2), txt)
OP         LET(readOk, 1)
OP       ELSE
OP         SET_VALUE(CELL($jobs, 1, 2), status)
OP       RETURN(readOk)
ERR      standard
FIXTURE  jobs_01.xlsx sheet=Jobs cols=A:Job ID,B:Info rows=3 data=A:seq(1,1);B:values("old"x3)
FIXTURE  jobs_01.xlsx folder=Ready files=status.txt:"Weld review passed",other.txt:20
FIXTURE  jobs_01.xlsx folder=Bare files=notes.txt:30
FIXTURE  jobs_01.xlsx folder=Held files=status.txt:"Locked text" locked=status.txt
TEST     fixture=jobs_01.xlsx call=ReadJobStatusFile(Jobs, Ready) runs=2 assert=RESULT = 1 AND ERROR_COUNT = 0 AND VALUE(Jobs!B1) = "Weld review passed" AND COUNTA(Jobs![Job ID]) = PRE_COUNTA(Jobs![Job ID]) AND FILE_COUNT("Ready") = PRE_FILE_COUNT("Ready")
TEST     fixture=jobs_01.xlsx call=ReadJobStatusFile(Jobs, Bare) assert=RESULT = 0 AND ERROR_COUNT = 0 AND VALUE(Jobs!B1) = "Missing" AND FILE_COUNT("Bare") = PRE_FILE_COUNT("Bare")
TEST     fixture=jobs_01.xlsx call=ReadJobStatusFile(Jobs, Held) assert=RESULT = 0 AND ERROR_COUNT = 0 AND VALUE(Jobs!B1) = "Locked"
TEST     fixture=jobs_01.xlsx call=ReadJobStatusFile(Jobs, "C:\KRONOS_NO_SUCH_FOLDER") assert=RESULT = -1 AND ERROR_COUNT = 1 AND VALUE(Jobs!B1) = "Info"
