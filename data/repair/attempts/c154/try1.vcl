VCL 0.6
TASK
# ASSUME: Welder Quals has headers in row 1: "Welder ID", "Process", "Expiry Date"; the real sheets are "Welder Quals" and "Welder List"
# ASSUME: the Welder List sheet already exists and is passed in; it is cleared and rebuilt on every run
# ASSUME: the process is matched ignoring case and stray spaces at the ends; rows with a blank Process are skipped, and a blank process argument stops the run
# ASSUME: rows are copied as they are, with the same three headings, and sorted by Welder ID A to Z; Welder Quals is only read
GOAL     Rebuild the Welder List sheet with every Welder Quals row for the given process, sorted by Welder ID, and return how many rows were listed.
PROC     ListWeldersForProcess
KIND     Function
MODULE   modWelders
ARGS     quals As Worksheet, out As Worksheet, proc As String
RET      Long
VAR      cId As Long
VAR      cProc As Long
VAR      cExp As Long
VAR      outRow As Long
PRE      COL($quals, "Welder ID") > 0; COL($quals, "Process") > 0; COL($quals, "Expiry Date") > 0; NOT IS_EMPTY($proc)
MODIFIES $out
OP       LET(cId, COL($quals, "Welder ID"))
OP       LET(cProc, COL($quals, "Process"))
OP       LET(cExp, COL($quals, "Expiry Date"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Welder ID", "Process", "Expiry Date")
OP       LET(outRow, 1)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($quals, cId) var=r
OP         SKIP_IF IS_EMPTY(CELL($quals, r, cProc))
OP         IF SAME(CELL($quals, r, cProc), $proc)
OP           LET(outRow, outRow + 1)
OP           WRITE_ROW($out, outRow, CELL($quals, r, cId), CELL($quals, r, cProc), CELL($quals, r, cExp))
OP       IF outRow > 2
OP         SORT(USED_RANGE($out), key=1, order=asc, header=yes)
OP       RETURN(outRow - 1)
PERF     screen_off
ERR      standard
FIXTURE  quals_01.xlsx sheet=Quals cols=A:Welder ID,B:Process,C:Expiry Date rows=9 stale=2 data=A:values("W-05"x1,"W-03"x1,"W-08"x1,"W-01"x1,"W-07"x1,"W-02"x1,"W-06"x1,"W-04"x1,"W-09"x1)+boundary("W-05","W-03","W-08","W-01","W-07","W-02","W-06","W-04","W-09");B:values("GTAW"x1,"gtaw "x1," GTAW"x1,"SMAW"x1,"smaw"x1,"GTAW"x1,"FCAW"x1,""x1,"GTAW2"x1)+boundary("GTAW","gtaw "," GTAW","SMAW","smaw","GTAW","FCAW","","GTAW2");C:dates(2026-01-01,2027-12-31)
FIXTURE  quals_01.xlsx sheet=List cols=A:Welder ID,B:Process,C:Expiry Date rows=0 stale=3
FIXTURE  quals_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=quals_01.xlsx call=ListWeldersForProcess(Quals, List, "GTAW") runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND ROWS_USED(List) = 5 AND COUNTIF(List![Welder ID], "stale") = 0 AND VALUE(List![Welder ID]2) = "W-02" AND VALUE(List![Welder ID]3) = "W-03" AND VALUE(List![Welder ID]4) = "W-05" AND VALUE(List![Welder ID]5) = "W-08"
TEST     fixture=quals_01.xlsx call=ListWeldersForProcess(Quals, List, " smaw ") assert=RESULT = 2 AND ERROR_COUNT = 0 AND VALUE(List![Welder ID]2) = "W-01" AND VALUE(List![Welder ID]3) = "W-07" AND COUNTA(Quals![Welder ID]) = PRE_COUNTA(Quals![Welder ID]) AND COUNTA(Quals![Expiry Date]) = PRE_COUNTA(Quals![Expiry Date])
TEST     fixture=quals_01.xlsx call=ListWeldersForProcess(Quals, List, "TIG") assert=RESULT = 0 AND ERROR_COUNT = 0 AND ROWS_USED(List) = 1
TEST     fixture=quals_01.xlsx call=ListWeldersForProcess(Notes, List, "GTAW") assert=RESULT = -1 AND ERROR_COUNT = 1
