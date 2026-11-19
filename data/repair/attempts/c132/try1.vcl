VCL 0.6
TASK
# ASSUME: Welder Quals headers in row 1 include "Welder ID", "Process", "Expiry Date" and "Status"
# ASSUME: Expiry Date holds real dates; an empty or non-date Expiry Date gives No date
# ASSUME: Expired means before today, Expires soon means today up to and including today plus 30 days, Valid is later than that
# ASSUME: Status is overwritten on every row that has a Welder ID, so a re-run gives the same result
# ASSUME: the macro returns the number of rows marked Expired or Expires soon
GOAL     Write Expired, Expires soon, Valid or No date in the Status of every welder qualification by its expiry date and return how many are Expired or Expires soon.
PROC     SetQualStatus
KIND     Function
MODULE   modQuals
ARGS     ws As Worksheet
RET      Long
VAR      cDate As Long
VAR      cStatus As Long
VAR      flagged As Long
PRE      COL($ws, "Welder ID") > 0; COL($ws, "Process") > 0; COL($ws, "Expiry Date") > 0; COL($ws, "Status") > 0
MODIFIES $ws["Status"]
OP       LET(cDate, COL($ws, "Expiry Date"))
OP       LET(cStatus, COL($ws, "Status"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, COL($ws, "Welder ID")) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, COL($ws, "Welder ID")))
OP         IF IS_DATE(CELL($ws, r, cDate))
OP           IF CELL($ws, r, cDate) < TODAY
OP             SET_VALUE(CELL($ws, r, cStatus), "Expired")
OP             LET(flagged, flagged + 1)
OP           ELSE
OP             IF CELL($ws, r, cDate) <= TODAY + 30
OP               SET_VALUE(CELL($ws, r, cStatus), "Expires soon")
OP               LET(flagged, flagged + 1)
OP             ELSE
OP               SET_VALUE(CELL($ws, r, cStatus), "Valid")
OP         ELSE
OP           SET_VALUE(CELL($ws, r, cStatus), "No date")
OP       RETURN(flagged)
PERF     screen_off
ERR      standard
FIXTURE  quals_01.xlsx sheet=Quals cols=A:Welder ID,B:Process,C:Expiry Date,D:Status rows=12 data=A:seq(1,1)+blank(1);B:values("GTAW"x6,"SMAW"x6);C:reldates(-100,100)+boundary(-1,0,1,29,30,31,"",-60,60)+blank(1);D:values("old"x12)
FIXTURE  quals_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=quals_01.xlsx call=SetQualStatus(Quals) runs=2 assert=ERROR_COUNT = 0 AND COUNTIF(Quals![Status], "Expired") = 3 AND COUNTIF(Quals![Status], "Expires soon") = 4 AND COUNTIF(Quals![Status], "Valid") = 3 AND COUNTIF(Quals![Status], "No date") = 2 AND COUNTIF(Quals![Status], "old") = 0
TEST     fixture=quals_01.xlsx call=SetQualStatus(Quals) assert=RESULT = 7 AND COUNTA(Quals![Welder ID]) = PRE_COUNTA(Quals![Welder ID]) AND COUNTA(Quals![Expiry Date]) = PRE_COUNTA(Quals![Expiry Date]) AND COUNTA(Quals![Process]) = PRE_COUNTA(Quals![Process])
TEST     fixture=quals_01.xlsx call=SetQualStatus(Quals) assert=VALUE(Quals![Status]2) = "Expired" AND VALUE(Quals![Status]3) = "Expires soon" AND VALUE(Quals![Status]4) = "Expires soon" AND VALUE(Quals![Status]5) = "Expires soon" AND VALUE(Quals![Status]6) = "Expires soon" AND VALUE(Quals![Status]7) = "Valid" AND VALUE(Quals![Status]8) = "No date"
TEST     fixture=quals_01.xlsx call=SetQualStatus(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
