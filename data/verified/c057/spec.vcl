VCL 0.6
TASK
# ASSUME: Welder Quals headers (row 1) include "Welder ID", "Process" and "Expiry Date"
# ASSUME: each row is one welder holding one process; a welder listed twice for the same process counts once, with the first row's expiry
# ASSUME: rows with a blank Welder ID or Process are skipped; Process and Welder ID are trimmed and compared ignoring case
# ASSUME: expired means an Expiry Date before today; a blank or non-date Expiry Date counts as not expired
# ASSUME: Process Summary is a new sheet after the last one; if it is already there the run changes nothing and returns 0
# ASSUME: the real sheet is "Welder Quals"; the TESTs use Quals
GOAL     Make a Process Summary sheet that lists each welding process from A to Z with how many welders hold it and how many of those have expired.
PROC     MakeProcessSummary
KIND     Function
MODULE   modQuals
ARGS     ws As Worksheet, wb As Workbook
RET      Long
VAR      cId As Long
VAR      cProc As Long
VAR      cExp As Long
VAR      pname As String
VAR      wkey As String
VAR      seen As Dictionary
VAR      have As Dictionary
VAR      gone As Dictionary
VAR      out As Worksheet
VAR      outRow As Long
VAR      nHave As Long
VAR      nGone As Long
PRE      COL($ws, "Welder ID") > 0; COL($ws, "Process") > 0; COL($ws, "Expiry Date") > 0
MODIFIES $wb sheets
OP       IF SHEET_EXISTS($wb, "Process Summary")
OP         LOG("Process Summary is already there; left as it is")
OP         RETURN(0)
OP       LET(cId, COL($ws, "Welder ID"))
OP       LET(cProc, COL($ws, "Process"))
OP       LET(cExp, COL($ws, "Expiry Date"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cId) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cId))
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cProc))
OP         LET(pname, TRIM(CELL($ws, r, cProc)))
OP         LET(wkey, pname & "|" & TRIM(CELL($ws, r, cId)))
OP         SKIP_IF EXISTS(seen, wkey)
OP         DICT_ADD(seen, wkey, 1, mode=overwrite)
OP         DICT_ADD(have, pname, 1, mode=sum)
OP         DICT_ADD(gone, pname, 0, mode=sum)
OP         IF IS_DATE(CELL($ws, r, cExp))
OP           IF CELL($ws, r, cExp) < TODAY
OP             DICT_ADD(gone, pname, 1, mode=sum)
OP       ADD_SHEET($wb, name="Process Summary", into=out)
OP       WRITE_ROW(out, 1, "Process", "Welders", "Expired")
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(have)
OP         DICT_GET(have, k, into=nHave)
OP         DICT_GET(gone, k, into=nGone)
OP         WRITE_ROW(out, outRow, k, nHave, nGone)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE(out), key=1, order=asc, header=yes)
OP       RETURN(outRow - 2)
PERF     screen_off
ERR      standard
FIXTURE  welders_01.xlsx sheet=Quals cols=A:Welder ID,B:Name,C:Process,D:Expiry Date rows=16 data=A:values("W01"x3,"W02"x2,"W03"x1,"W04"x1,"W05"x1,"W06"x1,"W07"x1,"W08"x1,"W09"x1,""x1,"W10"x1,"W11"x1,"W12"x1)+boundary("W01","W02","W03","W04","W05","W06","W01","W01","W07","W08","W09","","W10","W11","W02","W12");B:text(10);C:values("GTAW"x5,"SMAW"x3,"smaw"x1,"FCAW"x2,""x1,"SAW"x1,"GTAW "x1,"gtaw"x1,"GMAW"x1)+boundary("GTAW","GTAW","GTAW","SMAW","SMAW","smaw","SMAW","GTAW","FCAW","FCAW","","GTAW","SAW","GTAW ","gtaw","GMAW");D:reldates(-60,120)+boundary(5,-1,0,-30,100,-2,10,-50,-10,"",-5,-5,20,-3,30,1)
FIXTURE  welders_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=welders_01.xlsx call=MakeProcessSummary(Quals, welders_01.xlsx) assert=RESULT = 5 AND ERROR_COUNT = 0 AND ROWS_USED('Process Summary') = 6 AND VALUE('Process Summary'![Process]2) = "FCAW" AND VALUE('Process Summary'![Welders]2) = 2 AND VALUE('Process Summary'![Expired]2) = 1 AND VALUE('Process Summary'![Process]3) = "GMAW" AND VALUE('Process Summary'![Welders]3) = 1 AND VALUE('Process Summary'![Expired]3) = 0 AND VALUE('Process Summary'![Process]4) = "GTAW" AND VALUE('Process Summary'![Welders]4) = 4 AND VALUE('Process Summary'![Expired]4) = 2 AND VALUE('Process Summary'![Process]5) = "SAW" AND VALUE('Process Summary'![Welders]5) = 1 AND VALUE('Process Summary'![Expired]5) = 0 AND VALUE('Process Summary'![Process]6) = "SMAW" AND VALUE('Process Summary'![Welders]6) = 4 AND VALUE('Process Summary'![Expired]6) = 2
TEST     fixture=welders_01.xlsx call=MakeProcessSummary(Quals, welders_01.xlsx) runs=2 assert=ERROR_COUNT = 0 AND SHEET_EXISTS("Process Summary") AND ROWS_USED('Process Summary') = 6 AND SUM('Process Summary'![Welders]) = 12 AND SUM('Process Summary'![Expired]) = 5
TEST     fixture=welders_01.xlsx call=MakeProcessSummary(Quals, welders_01.xlsx) assert=COUNTA(Quals![Welder ID]) = PRE_COUNTA(Quals![Welder ID]) AND COUNTA(Quals![Process]) = PRE_COUNTA(Quals![Process]) AND COUNTA(Quals![Expiry Date]) = PRE_COUNTA(Quals![Expiry Date]) AND SHEET_EXISTS("Process Summary") AND SUMIFS('Process Summary'![Expired], 'Process Summary'![Process], "GTAW") = 2
TEST     fixture=welders_01.xlsx call=MakeProcessSummary(Notes, welders_01.xlsx) assert=RESULT = -1 AND ERROR_COUNT = 1 AND SHEET_EXISTS("Process Summary") = 0
