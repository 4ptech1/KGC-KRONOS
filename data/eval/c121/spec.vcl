VCL 0.6
TASK
# ASSUME: ITP Tracker headers in row 1 include "Instrument ID" and "Cal Check"; Calibration Tracker headers include "Instrument ID" and "Next Calibration Due"
# ASSUME: instrument IDs are trimmed and matched ignoring case; if an ID repeats in the Calibration Tracker the first row counts
# ASSUME: overdue means a real due date before today; due today is not overdue, and a blank or non-date due date is treated as not overdue
# ASSUME: the number returned is the points marked Out of cal plus those marked Unknown instrument
GOAL     Mark each ITP point whose instrument is overdue or missing from the Calibration Tracker, clear the others, and return how many need a look.
PROC     CheckItpCalibration
KIND     Function
MODULE   modCalCheck
ARGS     itp As Worksheet, cal As Worksheet
RET      Long
VAR      cInst As Long
VAR      cCheck As Long
VAR      cCalInst As Long
VAR      cDue As Long
VAR      dues As Dictionary
VAR      due As Variant
VAR      flagged As Long
PRE      COL($itp, "Instrument ID") > 0; COL($itp, "Cal Check") > 0; COL($cal, "Instrument ID") > 0; COL($cal, "Next Calibration Due") > 0
MODIFIES $itp["Cal Check"]
OP       LET(cInst, COL($itp, "Instrument ID"))
OP       LET(cCheck, COL($itp, "Cal Check"))
OP       LET(cCalInst, COL($cal, "Instrument ID"))
OP       LET(cDue, COL($cal, "Next Calibration Due"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($cal, cCalInst) var=r
OP         SKIP_IF IS_EMPTY(CELL($cal, r, cCalInst))
OP         DICT_ADD(dues, TRIM(CELL($cal, r, cCalInst)), CELL($cal, r, cDue), mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($itp, cInst) var=s
OP         SKIP_IF IS_EMPTY(CELL($itp, s, cInst))
OP         IF EXISTS(dues, TRIM(CELL($itp, s, cInst)))
OP           DICT_GET(dues, TRIM(CELL($itp, s, cInst)), into=due)
OP           SET_VALUE(CELL($itp, s, cCheck), "")
OP           IF IS_DATE(due)
OP             IF due < TODAY
OP               SET_VALUE(CELL($itp, s, cCheck), "Out of cal")
OP               LET(flagged, flagged + 1)
OP         ELSE
OP           SET_VALUE(CELL($itp, s, cCheck), "Unknown instrument")
OP           LET(flagged, flagged + 1)
OP       RETURN(flagged)
PERF     screen_off
ERR      standard
FIXTURE  itp_01.xlsx sheet=Itp cols=A:ITP No,B:Point,C:Instrument ID,D:Cal Check rows=9 data=A:seq(1,1);B:text(8);C:values("PG-1"x1," pg-1 "x1,"PG-2"x1,"PG-3"x1,"PG-4"x1,"PG-9"x1,""x2,"PG-5"x1)+boundary("PG-1"," pg-1 ","PG-2","PG-3","PG-4","PG-9","","PG-5","");D:values("old"x9)
FIXTURE  itp_01.xlsx sheet=Cal cols=A:Instrument ID,B:Next Calibration Due,C:Remark rows=7 data=A:values("PG-1"x2,"PG-2"x1,"PG-3"x1,"PG-4"x1,"PG-5"x1,""x1)+boundary("PG-1","","PG-2","PG-3","PG-4","PG-5","PG-1");B:reldates(-30,30)+boundary(-1,-3,0,1,"",-5,10);C:text(6)
TEST     fixture=itp_01.xlsx call=CheckItpCalibration(Itp, Cal) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTIF(Itp![Cal Check], "Out of cal") = 3 AND COUNTIF(Itp![Cal Check], "Unknown instrument") = 1 AND COUNTIF(Itp![Cal Check], "old") = 2
TEST     fixture=itp_01.xlsx call=CheckItpCalibration(Itp, Cal) assert=VALUE(Itp![Cal Check]2) = "Out of cal" AND VALUE(Itp![Cal Check]3) = "Out of cal" AND VALUE(Itp![Cal Check]4) = "" AND VALUE(Itp![Cal Check]5) = "" AND VALUE(Itp![Cal Check]6) = "" AND VALUE(Itp![Cal Check]7) = "Unknown instrument" AND VALUE(Itp![Cal Check]8) = "old" AND VALUE(Itp![Cal Check]9) = "Out of cal" AND COUNTA(Itp![Instrument ID]) = PRE_COUNTA(Itp![Instrument ID]) AND COUNTA(Cal![Instrument ID]) = PRE_COUNTA(Cal![Instrument ID])
TEST     fixture=itp_01.xlsx call=CheckItpCalibration(Cal, Itp) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTIF(Itp![Cal Check], "old") = PRE_COUNTIF(Itp![Cal Check], "old")
