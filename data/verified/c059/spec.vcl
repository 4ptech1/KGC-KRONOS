VCL 0.6
TASK
# ASSUME: Calibration Tracker headers in row 1 include "Instrument ID", "Location" and "Next Calibration Due"
# ASSUME: the real sheets are "Calibration Tracker" and "Location Summary"; the TESTs use Calibration and LocationSummary
# ASSUME: the Location Summary sheet already exists; it is rebuilt on every run with the headers Location, Instruments and Earliest Due
# ASSUME: a row is one instrument when it has an Instrument ID and a Location; rows missing either are not counted
# ASSUME: locations are trimmed and compared ignoring case, and listed in the order they first appear
# ASSUME: the earliest date uses real dates only; a location with no real date gets a blank Earliest Due
GOAL     Rebuild the Location Summary sheet with one row per location, its number of instruments and its earliest next calibration due date.
PROC     SummariseCalibrationByLocation
KIND     Function
MODULE   modCalibration
ARGS     tracker As Worksheet, out As Worksheet
RET      Long
VAR      cId As Long
VAR      cLoc As Long
VAR      cDue As Long
VAR      loc As String
VAR      counts As Dictionary
VAR      earliest As Dictionary
VAR      cur As Variant
VAR      got As Long
VAR      outRow As Long
PRE      COL($tracker, "Instrument ID") > 0; COL($tracker, "Location") > 0; COL($tracker, "Next Calibration Due") > 0
MODIFIES $out
OP       LET(cId, COL($tracker, "Instrument ID"))
OP       LET(cLoc, COL($tracker, "Location"))
OP       LET(cDue, COL($tracker, "Next Calibration Due"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Location", "Instruments", "Earliest Due")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($tracker, cLoc) var=r
OP         SKIP_IF IS_EMPTY(CELL($tracker, r, cId)) OR IS_EMPTY(CELL($tracker, r, cLoc))
OP         LET(loc, TRIM(CELL($tracker, r, cLoc)))
OP         DICT_ADD(counts, loc, 1, mode=sum)
OP         IF IS_DATE(CELL($tracker, r, cDue))
OP           IF EXISTS(earliest, loc)
OP             DICT_GET(earliest, loc, into=cur)
OP             IF IS_DATE(cur)
OP               IF CELL($tracker, r, cDue) < cur THEN DICT_ADD(earliest, loc, CELL($tracker, r, cDue), mode=overwrite)
OP           ELSE
OP             DICT_ADD(earliest, loc, CELL($tracker, r, cDue), mode=overwrite)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(counts)
OP         DICT_GET(counts, k, into=got)
OP         WRITE_ROW($out, outRow, k, got)
OP         IF EXISTS(earliest, k)
OP           DICT_GET(earliest, k, into=cur)
OP           SET_VALUE(CELL($out, outRow, 3), cur)
OP         LET(outRow, outRow + 1)
OP       RETURN(COUNT(counts))
PERF     screen_off
ERR      standard
FIXTURE  calibration_01.xlsx sheet=Calibration cols=A:Instrument ID,B:Location,C:Next Calibration Due rows=10 data=A:values("I-1"x1,"I-2"x1,"I-3"x1,"I-4"x1,"I-5"x1,"I-6"x1,""x1,"I-8"x1,"I-9"x1,"I-10"x1)+boundary("I-1","I-2","I-3","I-4","I-5","I-6","","I-8","I-9","I-10");B:values("Lab"x4," lab "x1,"Yard"x3,"Store"x1,""x1)+boundary("Lab","Lab","Yard"," lab ","Yard","Store","Lab","","Yard","Lab");C:dates(2026-01-01,2026-12-31)+boundary(2026-05-10,2026-03-01,2026-07-01,2026-04-15,"","",2026-01-01,2026-02-01,2026-06-30,2026-03-01)
FIXTURE  calibration_01.xlsx sheet=LocationSummary cols=A:Location,B:Instruments,C:Earliest Due rows=0 stale=2
TEST     fixture=calibration_01.xlsx call=SummariseCalibrationByLocation(Calibration, LocationSummary) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTA(LocationSummary![Location]) = 4 AND COUNTIF(LocationSummary![Location], "stale") = 0 AND VALUE(LocationSummary![Location]2) = "Lab" AND VALUE(LocationSummary![Instruments]2) = 4 AND VALUE(LocationSummary![Location]3) = "Yard" AND VALUE(LocationSummary![Instruments]3) = 3 AND VALUE(LocationSummary![Location]4) = "Store" AND VALUE(LocationSummary![Instruments]4) = 1
TEST     fixture=calibration_01.xlsx call=SummariseCalibrationByLocation(Calibration, LocationSummary) assert=COUNTA(LocationSummary![Earliest Due]) = 3 AND COUNTIF(LocationSummary![Earliest Due], "2026-03-01") = 1 AND COUNTIF(LocationSummary![Earliest Due], "2026-06-30") = 1 AND SUM(LocationSummary![Instruments]) = 8
TEST     fixture=calibration_01.xlsx call=SummariseCalibrationByLocation(Calibration, LocationSummary) assert=ERROR_COUNT = 0 AND COUNTA(Calibration![Instrument ID]) = PRE_COUNTA(Calibration![Instrument ID]) AND COUNTA(Calibration![Location]) = PRE_COUNTA(Calibration![Location]) AND COUNTA(Calibration![Next Calibration Due]) = PRE_COUNTA(Calibration![Next Calibration Due])
TEST     fixture=calibration_01.xlsx call=SummariseCalibrationByLocation(LocationSummary, Calibration) assert=RESULT = -1 AND ERROR_COUNT = 1
