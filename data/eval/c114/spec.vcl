VCL 0.6
TASK
# ASSUME: the sheet is "Job Tracker" with headers in row 1 including "Job ID", "Status" and "Percent Complete"
# ASSUME: a job counts as complete when its Status is Complete, ignoring case and spaces
# ASSUME: Percent Complete is a whole-percent number; 100 or more counts as 100 percent and below 100 as not done
# ASSUME: a job with a blank or non-number Percent Complete is not flagged and its Flag is cleared
# ASSUME: rows with a blank Job ID are left alone
# ASSUME: if there is no Flag heading it is added in the first free column of row 1
GOAL     Write Check status in the Flag column of each job whose Status and Percent Complete disagree, clear the others, and return how many were flagged.
PROC     FlagJobStatus
KIND     Function
MODULE   modJobs
ARGS     ws As Worksheet
RET      Long
VAR      cJob As Long
VAR      cStat As Long
VAR      cPct As Long
VAR      cFlag As Long
VAR      flagged As Long
PRE      COL($ws, "Job ID") > 0; COL($ws, "Status") > 0; COL($ws, "Percent Complete") > 0
MODIFIES $ws
OP       LET(cJob, COL($ws, "Job ID"))
OP       LET(cStat, COL($ws, "Status"))
OP       LET(cPct, COL($ws, "Percent Complete"))
OP       LET(cFlag, COL($ws, "Flag"))
OP       IF cFlag = 0
OP         LET(cFlag, LASTCOL($ws, 1) + 1)
OP         SET_VALUE(CELL($ws, 1, cFlag), "Flag")
OP       LET(flagged, 0)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cJob))
OP         IF IS_NUMBER(CELL($ws, r, cPct))
OP           IF CELL($ws, r, cPct) >= 100 AND NOT SAME(CELL($ws, r, cStat), "Complete")
OP             SET_VALUE(CELL($ws, r, cFlag), "Check status")
OP             LET(flagged, flagged + 1)
OP           ELSE
OP             IF CELL($ws, r, cPct) < 100 AND SAME(CELL($ws, r, cStat), "Complete")
OP               SET_VALUE(CELL($ws, r, cFlag), "Check status")
OP               LET(flagged, flagged + 1)
OP             ELSE
OP               CLEAR(CELL($ws, r, cFlag), what=contents)
OP         ELSE
OP           CLEAR(CELL($ws, r, cFlag), what=contents)
OP       MSG("Jobs flagged: " & flagged)
OP       RETURN(flagged)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  jobs_01.xlsx sheet=Jobs cols=A:Job ID,B:Customer,C:Status,D:Percent Complete,E:Flag rows=12 data=A:seq(1,1)+boundary(1,2,3,4,5,6,7,8,9,10,"",12);B:text(8);C:values("Complete"x4,"In Progress"x3,"complete "x2,"Not Started"x1,"Hold"x2)+boundary("Complete","Complete","In Progress","In Progress","complete ","complete ","Not Started","Complete","In Progress","Hold","Complete","Hold");D:int(0,100)+boundary("",99,100,99,100,50,0,100,"",100,50,100);E:values("Check status"x4,"old"x2,"keep"x1,""x5)+boundary("Check status","old","","Check status","Check status","","","Check status","","old","keep","")
FIXTURE  jobs_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  jobs_02.xlsx sheet=Jobs cols=A:Job ID,B:Customer,C:Status,D:Percent Complete rows=12 data=A:seq(1,1)+boundary(1,2,3,4,5,6,7,8,9,10,"",12);B:text(8);C:values("Complete"x4,"In Progress"x3,"complete "x2,"Not Started"x1,"Hold"x2)+boundary("Complete","Complete","In Progress","In Progress","complete ","complete ","Not Started","Complete","In Progress","Hold","Complete","Hold");D:int(0,100)+boundary("",99,100,99,100,50,0,100,"",100,50,100)
TEST     fixture=jobs_01.xlsx call=FlagJobStatus(Jobs) runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND COUNTIF(Jobs![Flag], "Check status") = 5 AND COUNTIF(Jobs![Flag], "old") = 0 AND COUNTIF(Jobs![Flag], "keep") = 1 AND VALUE(Jobs![Flag]3) = "Check status" AND VALUE(Jobs![Flag]2) = "" AND VALUE(Jobs![Flag]6) = "" AND VALUE(Jobs![Flag]12) = "keep"
TEST     fixture=jobs_01.xlsx call=FlagJobStatus(Jobs) assert=COUNTA(Jobs![Job ID]) = PRE_COUNTA(Jobs![Job ID]) AND COUNTA(Jobs![Percent Complete]) = PRE_COUNTA(Jobs![Percent Complete]) AND VALUE(Jobs![Flag]4) = "Check status" AND VALUE(Jobs![Flag]5) = "" AND VALUE(Jobs![Flag]7) = "Check status" AND VALUE(Jobs![Flag]9) = "" AND VALUE(Jobs![Flag]11) = "Check status" AND VALUE(Jobs![Flag]13) = "Check status"
TEST     fixture=jobs_02.xlsx call=FlagJobStatus(Jobs) assert=RESULT = 5 AND ERROR_COUNT = 0 AND VALUE(Jobs![Flag]1) = "Flag" AND COUNTIF(Jobs![Flag], "Check status") = 5 AND COUNTA(Jobs![Flag]) = 6
TEST     fixture=jobs_01.xlsx call=FlagJobStatus(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
