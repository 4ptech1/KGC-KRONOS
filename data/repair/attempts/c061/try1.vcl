VCL 0.6
TASK
# ASSUME: Welder Quals headers in row 1 include "Welder ID", "Process" and "Expiry Date"; WPS Register headers include "Process"
# ASSUME: the "No WPS" sheet already exists; it is rebuilt on every run and gets the Welder Quals headings in row 1
# ASSUME: Process is trimmed and compared ignoring case, so 'gtaw' and 'GTAW' match
# ASSUME: a qualification row with a blank Welder ID is skipped; one with a blank Process is listed, because no WPS covers it
# ASSUME: each listed row is copied as it stands, in Welder Quals order, and a WPS row with a blank Process covers nothing
# ASSUME: the real sheets are "Welder Quals", "WPS Register" and "No WPS"; the TESTs use WelderQuals, WpsReg and NoWps
GOAL     List on the No WPS sheet every welder qualification whose process has no WPS in the register, and return how many were listed.
PROC     ListQualsWithoutWps
KIND     Function
MODULE   modWelderQuals
ARGS     quals As Worksheet, wps As Worksheet, out As Worksheet
RET      Long
VAR      cWelder As Long
VAR      cProc As Long
VAR      cExpiry As Long
VAR      cWpsProc As Long
VAR      covered As Dictionary
VAR      listed As Long
PRE      COL($quals, "Welder ID") > 0; COL($quals, "Process") > 0; COL($quals, "Expiry Date") > 0; COL($wps, "Process") > 0
MODIFIES $out
OP       LET(cWelder, COL($quals, "Welder ID"))
OP       LET(cProc, COL($quals, "Process"))
OP       LET(cExpiry, COL($quals, "Expiry Date"))
OP       LET(cWpsProc, COL($wps, "Process"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($wps, cWpsProc) var=r
OP         SKIP_IF IS_EMPTY(CELL($wps, r, cWpsProc))
OP         DICT_ADD(covered, TRIM(CELL($wps, r, cWpsProc)), 1, mode=keep)
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, CELL($quals, 1, cWelder), CELL($quals, 1, cProc), CELL($quals, 1, cExpiry))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($quals, cWelder) var=q
OP         SKIP_IF IS_EMPTY(CELL($quals, q, cWelder))
OP         IF NOT EXISTS(covered, TRIM(CELL($quals, q, cProc)))
OP           LET(listed, listed + 1)
OP           WRITE_ROW($out, listed + 1, CELL($quals, q, cWelder), CELL($quals, q, cProc), CELL($quals, q, cExpiry))
OP       RETURN(listed)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  quals_01.xlsx sheet=WelderQuals cols=A:Welder ID,B:Process,C:Expiry Date rows=9 data=A:values("W-01"x1,"W-02"x1,"W-03"x1,"W-04"x1,""x1,"W-06"x1,"W-07"x1,"W-08"x1,"W-09"x1)+boundary("W-01","W-02","W-03","W-04","","W-06","W-07","W-08","W-09");B:values("gtaw"x1,"SMAW"x1,"GMAW"x2,""x1,"fcaw"x1,"PAW"x1," GMAW "x1,"SAW"x1)+boundary("gtaw","SMAW","GMAW","","GMAW","fcaw","PAW"," GMAW ","SAW");C:dates(2026-01-01,2027-12-31)+boundary(2026-03-01,2026-04-02,2026-05-03,2026-06-04,2026-07-05,2026-08-06,2026-09-07,2026-10-08,2026-11-09)
FIXTURE  quals_01.xlsx sheet=WpsReg cols=A:WPS No,B:Process rows=6 data=A:seq(1,1);B:values("GTAW"x2,"smaw "x1,"FCAW"x1,""x1,"SAW"x1)+boundary("GTAW","","smaw ","GTAW","FCAW","SAW")
FIXTURE  quals_01.xlsx sheet=NoWps cols=A:Welder ID,B:Process,C:Expiry Date rows=0 stale=3
TEST     fixture=quals_01.xlsx call=ListQualsWithoutWps(WelderQuals, WpsReg, NoWps) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTA(NoWps![Welder ID]) = 5 AND COUNTIF(NoWps![Welder ID], "stale") = 0 AND ROWS_USED(NoWps) = 5
TEST     fixture=quals_01.xlsx call=ListQualsWithoutWps(WelderQuals, WpsReg, NoWps) assert=VALUE(NoWps![Welder ID]1) = "Welder ID" AND VALUE(NoWps![Process]1) = "Process" AND VALUE(NoWps![Expiry Date]1) = "Expiry Date" AND VALUE(NoWps![Welder ID]2) = "W-03" AND VALUE(NoWps![Process]2) = "GMAW" AND VALUE(NoWps![Welder ID]3) = "W-04" AND VALUE(NoWps![Welder ID]4) = "W-07" AND VALUE(NoWps![Process]4) = "PAW" AND VALUE(NoWps![Welder ID]5) = "W-08" AND VALUE(NoWps![Process]5) = " GMAW " AND COUNTA(NoWps![Expiry Date]) = 5
TEST     fixture=quals_01.xlsx call=ListQualsWithoutWps(WelderQuals, WpsReg, NoWps) assert=ERROR_COUNT = 0 AND COUNTA(WelderQuals![Welder ID]) = PRE_COUNTA(WelderQuals![Welder ID]) AND COUNTA(WelderQuals![Process]) = PRE_COUNTA(WelderQuals![Process]) AND COUNTA(WpsReg![Process]) = PRE_COUNTA(WpsReg![Process])
TEST     fixture=quals_01.xlsx call=ListQualsWithoutWps(WpsReg, WelderQuals, NoWps) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(NoWps![Welder ID]) = PRE_COUNTA(NoWps![Welder ID])
