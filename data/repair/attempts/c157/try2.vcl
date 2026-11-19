VCL 0.6
TASK
# ASSUME: the real sheets are Dispatch (Dispatch No, Job ID, Planned Date, NCR Check) and NCR Register (NCR No, Job ID, Status); the TESTs use Dispatch and NcrReg
# ASSUME: an NCR is not Closed when its Status is anything but Closed (ignoring case and spaces), a blank Status included
# ASSUME: Job IDs are matched ignoring case and stray spaces; NCRs with a blank Job ID are ignored
# ASSUME: a dispatch with a blank Job ID is left alone; the other dispatches get Open NCR or a blank NCR Check
# ASSUME: the NCR Check heading is added after the last used column of row 1 when it is missing; the last Dispatch row has a Job ID
GOAL     Mark each dispatch whose job has a not-closed NCR as Open NCR in the NCR Check column, clear the others and return how many are marked.
PROC     MarkOpenNcrDispatches
KIND     Function
MODULE   modNcrCheck
ARGS     disp As Worksheet, ncr As Worksheet
RET      Long
VAR      cJob As Long
VAR      cChk As Long
VAR      nJob As Long
VAR      nStat As Long
VAR      openJobs As Dictionary
VAR      marked As Long
PRE      COL($disp, "Job ID") > 0; COL($ncr, "Job ID") > 0; COL($ncr, "Status") > 0
MODIFIES $disp
OP       LET(cJob, COL($disp, "Job ID"))
OP       LET(nJob, COL($ncr, "Job ID"))
OP       LET(nStat, COL($ncr, "Status"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ncr, nJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($ncr, r, nJob))
OP         IF NOT SAME(CELL($ncr, r, nStat), "Closed") THEN DICT_ADD(openJobs, CLEAN_TEXT(CELL($ncr, r, nJob)), 1, mode=keep)
OP       IF COL($disp, "NCR Check") = 0 THEN SET_VALUE(CELL($disp, 1, LASTCOL($disp, 1) + 1), "NCR Check")
OP       LET(cChk, COL($disp, "NCR Check"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($disp, cJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($disp, r, cJob))
OP         IF EXISTS(openJobs, CLEAN_TEXT(CELL($disp, r, cJob)))
OP           SET_VALUE(CELL($disp, r, cChk), "Open NCR")
OP           LET(marked, marked + 1)
OP         ELSE
OP           SET_VALUE(CELL($disp, r, cChk), "")
OP       MSG("Dispatches with an Open NCR: " & marked)
OP       RETURN(marked)
ERR      standard
FIXTURE  ncr_nocol_01.xlsx sheet=Dispatch cols=A:Dispatch No,B:Job ID,C:Planned Date rows=8 data=A:seq(1,1);B:values("J-1"x1,"j-1"x1,"J-2"x1,""x1,"J-3"x1,"J-4"x1,"j-2"x1,"J-5"x1)+boundary("J-1"," j-1","J-2","","J-3","J-4","j-2 ","J-5");C:dates(2026-03-01,2026-03-31)
FIXTURE  ncr_nocol_01.xlsx sheet=NcrReg cols=A:NCR No,B:Job ID,C:Status rows=8 data=A:seq(1,1);B:values("J-1"x1,"j-2"x1,"J-3"x2,"J-5"x1,""x1,"J-2"x1,"J-6"x1)+boundary("J-1","j-2","J-3","J-3","J-5 ","","J-2","J-6");C:values("Open"x4,"Closed"x3,""x1)+boundary("Open","Closed","Closed","Open","open","Open"," closed ","")
FIXTURE  ncr_hascol_01.xlsx sheet=Dispatch cols=A:Dispatch No,B:Job ID,C:Planned Date,D:NCR Check rows=8 data=A:seq(1,1);B:values("J-1"x1,"j-1"x1,"J-2"x1,""x1,"J-3"x1,"J-4"x1,"j-2"x1,"J-5"x1)+boundary("J-1"," j-1","J-2","","J-3","J-4","j-2 ","J-5");C:dates(2026-03-01,2026-03-31);D:values("Open NCR"x3,"keep"x1,"idle"x4)+boundary("idle","idle","Open NCR","keep","idle","Open NCR","Open NCR","idle")
FIXTURE  ncr_hascol_01.xlsx sheet=NcrReg cols=A:NCR No,B:Job ID,C:Status rows=8 data=A:seq(1,1);B:values("J-1"x1,"j-2"x1,"J-3"x2,"J-5"x1,""x1,"J-2"x1,"J-6"x1)+boundary("J-1","j-2","J-3","J-3","J-5 ","","J-2","J-6");C:values("Open"x4,"Closed"x3,""x1)+boundary("Open","Closed","Closed","Open","open","Open"," closed ","")
TEST     fixture=ncr_nocol_01.xlsx call=MarkOpenNcrDispatches(Dispatch, NcrReg) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTIF(Dispatch![NCR Check], "Open NCR") = 4 AND COUNTA(Dispatch![NCR Check]) = 5 AND VALUE(Dispatch![NCR Check]1) = "NCR Check"
TEST     fixture=ncr_hascol_01.xlsx call=MarkOpenNcrDispatches(Dispatch, NcrReg) assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTIF(Dispatch![NCR Check], "Open NCR") = 4 AND COUNTIF(Dispatch![NCR Check], "idle") = 0 AND COUNTIF(Dispatch![NCR Check], "keep") = 1 AND COUNTA(Dispatch![Job ID]) = PRE_COUNTA(Dispatch![Job ID])
TEST     fixture=ncr_hascol_01.xlsx call=MarkOpenNcrDispatches(NcrReg, Dispatch) assert=RESULT = -1 AND ERROR_COUNT = 1
