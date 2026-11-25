VCL 0.7
TASK
# ASSUME: both sheets have headings in row 1; the Cal Check heading already exists in the ITP Tracker
# ASSUME: an instrument is overdue when its Next Calibration Due is a date before today; a blank or non-date due date is not overdue
# ASSUME: Instrument IDs are matched with SAME (trimmed, case ignored); the first Calibration Tracker row of an ID is the one used
# ASSUME: points need a look means the points marked Out of cal or Unknown instrument
GOAL     Mark each ITP point whose instrument is overdue for calibration or missing from the calibration tracker, clear the others, and count the points marked.
PROC     CheckItpCalibration
KIND     Function
MODULE   modCal
ARGS     itp As Worksheet, cal As Worksheet
RET      Long
VAR      cInst As Long
VAR      cCheck As Long
VAR      calInst As Long
VAR      calDue As Long
VAR      known As Dictionary
VAR      overdue As Dictionary
VAR      id As String
VAR      flagged As Long
PRE      COL($itp, "Instrument ID") > 0; COL($itp, "Cal Check") > 0; COL($cal, "Instrument ID") > 0; COL($cal, "Next Calibration Due") > 0
MODIFIES $itp["Cal Check"]
OP       LET(calInst, COL($cal, "Instrument ID"))
OP       LET(calDue, COL($cal, "Next Calibration Due"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($cal, calInst) var=r
OP         SKIP_IF IS_EMPTY(CELL($cal, r, calInst))
OP         DICT_ADD(known, TRIM(CELL($cal, r, calInst)), 1, mode=keep)
OP         IF IS_DATE(CELL($cal, r, calDue))
OP           IF CELL($cal, r, calDue) < TODAY
OP             DICT_ADD(overdue, TRIM(CELL($cal, r, calInst)), 1, mode=keep)
OP       LET(cInst, COL($itp, "Instrument ID"))
OP       LET(cCheck, COL($itp, "Cal Check"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($itp, cInst) var=r
OP         SKIP_IF IS_EMPTY(CELL($itp, r, cInst))
OP         LET(id, TRIM(CELL($itp, r, cInst)))
OP         IF NOT EXISTS(known, id)
OP           SET_VALUE(CELL($itp, r, cCheck), "Unknown instrument")
OP           LET(flagged, flagged + 1)
OP         ELSE
OP           IF EXISTS(overdue, id)
OP             SET_VALUE(CELL($itp, r, cCheck), "Out of cal")
OP             LET(flagged, flagged + 1)
OP           ELSE
OP             CLEAR(CELL($itp, r, cCheck), what=contents)
OP       RETURN(flagged)
ERR      standard
FIXTURE  cal_01.xlsx sheet=Itp cols=A:ITP No,B:Point,C:Instrument ID,D:Cal Check rows=7 data=A:seq(1,1);B:text(6);C:values("I-1"x1,"I-2"x1,"I-3"x1,"I-4"x1,"I-9"x1,""x1,"i-5"x1)+boundary("I-1","I-2","I-3","I-4","I-9","","i-5");D:values("old"x7)
FIXTURE  cal_01.xlsx sheet=Cal cols=A:Instrument ID,B:Next Calibration Due rows=5 data=A:values("I-1"x1,"I-2"x1,"I-3"x1,"I-4"x1,"I-5"x1)+boundary("I-1","I-2","I-3","I-4","I-5");B:reldates(-30,30)+boundary(-1,0,1,"",-10)
TEST     fixture=cal_01.xlsx call=CheckItpCalibration(Itp, Cal) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(Itp![Cal Check], "Out of cal") = 2 AND COUNTIF(Itp![Cal Check], "Unknown instrument") = 1 AND COUNTIF(Itp![Cal Check], "old") = 1 AND VALUE(Itp![Cal Check]3) = "" AND VALUE(Itp![Cal Check]4) = "" AND VALUE(Itp![Cal Check]5) = ""
TEST     fixture=cal_01.xlsx call=CheckItpCalibration(Itp, Cal) assert=COUNTA(Itp![ITP No]) = PRE_COUNTA(Itp![ITP No]) AND COUNTA(Cal![Instrument ID]) = PRE_COUNTA(Cal![Instrument ID]) AND COUNTA(Cal![Next Calibration Due]) = PRE_COUNTA(Cal![Next Calibration Due])
TEST     fixture=cal_01.xlsx call=CheckItpCalibration(Cal, Itp) assert=RESULT = -1 AND ERROR_COUNT = 1
