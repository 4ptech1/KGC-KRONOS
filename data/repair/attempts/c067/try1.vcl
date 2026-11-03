VCL 0.6
TASK
# ASSUME: Dispatch headers in row 1 include "Dispatch No" and "Actual Date"
# ASSUME: the real sheet is "Dispatch" and the real folder is "Dispatch Docs"; the TESTs use Dispatch and Docs
# ASSUME: cell G1 is free; it holds only the result of the last run, and the other cells of the sheet are not changed
# ASSUME: a row counts as a dispatch only when it has a Dispatch No; its Actual Date is missing when the cell is blank or holds only spaces
# ASSUME: the file holds the count as plain text with no line break; a count of 0 is written as "0"
# ASSUME: G1 shows the file's status: Written, Already there (the file holds that number), or Conflict (it holds something else and is never overwritten); also Locked or Too long
GOAL     Count the dispatches that have no actual date, write that number as text into open_dispatches.txt in the Dispatch Docs folder unless it already holds it, and show what happened in G1.
PROC     WriteOpenDispatchCount
KIND     Function
MODULE   modDispatch
ARGS     ws As Worksheet, docs As String
RET      Long
VAR      cNo As Long
VAR      cActual As Long
VAR      openCount As Long
VAR      status As String
PRE      COL($ws, "Dispatch No") > 0; COL($ws, "Actual Date") > 0; FOLDER_EXISTS($docs)
MODIFIES $ws
MODIFIES $docs files
OP       LET(cNo, COL($ws, "Dispatch No"))
OP       LET(cActual, COL($ws, "Actual Date"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cNo) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cNo))
OP         IF IS_EMPTY(CELL($ws, r, cActual))
OP           LET(openCount, openCount + 1)
OP       WRITE_TEXT(PATH_JOIN($docs, "open_dispatches.txt"), "" & openCount, into=status)
OP       SET_VALUE(RANGE($ws, "G1"), status)
OP       RETURN(openCount)
PERF     screen_off
ERR      standard
FIXTURE  dispatch_01.xlsx sheet=Dispatch cols=A:Dispatch No,B:Customer,C:Planned Date,D:Actual Date rows=8 data=A:seq(1001,1)+boundary(1001,1002,1003,"",1005,1006,1007,1008);B:text(8);C:dates(2026-01-01,2026-06-30);D:dates(2026-01-05,2026-07-10)+boundary("",2026-02-10,"","",2026-03-01,"",2026-03-02,2026-03-03)
FIXTURE  dispatch_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  dispatch_01.xlsx folder=Docs files=notes.txt:20
FIXTURE  dispatch_02.xlsx sheet=Dispatch cols=A:Dispatch No,B:Customer,C:Planned Date,D:Actual Date rows=8 data=A:seq(1001,1)+boundary(1001,1002,1003,"",1005,1006,1007,1008);B:text(8);C:dates(2026-01-01,2026-06-30);D:dates(2026-01-05,2026-07-10)+boundary("",2026-02-10,"","",2026-03-01,"",2026-03-02,2026-03-03)
FIXTURE  dispatch_02.xlsx folder=Docs files=notes.txt:20,open_dispatches.txt:"3"
FIXTURE  dispatch_03.xlsx sheet=Dispatch cols=A:Dispatch No,B:Customer,C:Planned Date,D:Actual Date rows=8 data=A:seq(1001,1)+boundary(1001,1002,1003,"",1005,1006,1007,1008);B:text(8);C:dates(2026-01-01,2026-06-30);D:dates(2026-01-05,2026-07-10)+boundary("",2026-02-10,"","",2026-03-01,"",2026-03-02,2026-03-03)
FIXTURE  dispatch_03.xlsx folder=Docs files=notes.txt:20,open_dispatches.txt:"9"
FIXTURE  dispatch_04.xlsx sheet=Dispatch cols=A:Dispatch No,B:Customer,C:Planned Date,D:Actual Date rows=8 data=A:seq(1001,1)+boundary(1001,1002,1003,"",1005,1006,1007,1008);B:text(8);C:dates(2026-01-01,2026-06-30);D:dates(2026-01-05,2026-07-10)+boundary("",2026-02-10,"","",2026-03-01,"",2026-03-02,2026-03-03)
FIXTURE  dispatch_04.xlsx folder=Docs files=open_dispatches.txt:"3" locked=open_dispatches.txt
FIXTURE  dispatch_05.xlsx sheet=Dispatch cols=A:Dispatch No,B:Customer,C:Planned Date,D:Actual Date rows=3 data=A:seq(2001,1);B:text(8);C:dates(2026-01-01,2026-06-30);D:dates(2026-01-05,2026-07-10)
FIXTURE  dispatch_05.xlsx folder=Docs files=notes.txt:20
TEST     fixture=dispatch_01.xlsx call=WriteOpenDispatchCount(Dispatch, Docs) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND VALUE(Dispatch!G1) = "Already there" AND FILE_TEXT("Docs\open_dispatches.txt") = "3" AND FILE_COUNT("Docs") = PRE_FILE_COUNT("Docs") + 1
TEST     fixture=dispatch_01.xlsx call=WriteOpenDispatchCount(Dispatch, Docs) assert=RESULT = 3 AND ERROR_COUNT = 0 AND VALUE(Dispatch!G1) = "Written" AND FILE_EXISTS("Docs\open_dispatches.txt") AND COUNTA(Dispatch![Dispatch No]) = PRE_COUNTA(Dispatch![Dispatch No]) AND COUNTA(Dispatch![Actual Date]) = PRE_COUNTA(Dispatch![Actual Date])
TEST     fixture=dispatch_01.xlsx call=WriteOpenDispatchCount(Notes, Docs) assert=RESULT = -1 AND ERROR_COUNT = 1 AND FILE_COUNT("Docs") = PRE_FILE_COUNT("Docs")
TEST     fixture=dispatch_01.xlsx call=WriteOpenDispatchCount(Dispatch, "C:\NoSuchFolder") assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(Dispatch![Dispatch No]) = PRE_COUNTA(Dispatch![Dispatch No])
TEST     fixture=dispatch_02.xlsx call=WriteOpenDispatchCount(Dispatch, Docs) assert=RESULT = 3 AND ERROR_COUNT = 0 AND VALUE(Dispatch!G1) = "Already there" AND FILE_TEXT("Docs\open_dispatches.txt") = "3" AND FILE_COUNT("Docs") = PRE_FILE_COUNT("Docs")
TEST     fixture=dispatch_03.xlsx call=WriteOpenDispatchCount(Dispatch, Docs) assert=RESULT = 3 AND ERROR_COUNT = 0 AND VALUE(Dispatch!G1) = "Conflict" AND FILE_TEXT("Docs\open_dispatches.txt") = "9" AND FILE_COUNT("Docs") = PRE_FILE_COUNT("Docs")
TEST     fixture=dispatch_04.xlsx call=WriteOpenDispatchCount(Dispatch, Docs) assert=RESULT = 3 AND ERROR_COUNT = 0 AND VALUE(Dispatch!G1) = "Locked" AND FILE_COUNT("Docs") = PRE_FILE_COUNT("Docs")
TEST     fixture=dispatch_05.xlsx call=WriteOpenDispatchCount(Dispatch, Docs) assert=RESULT = 0 AND ERROR_COUNT = 0 AND VALUE(Dispatch!G1) = "Written" AND FILE_TEXT("Docs\open_dispatches.txt") = "0"
