VCL 0.6
TASK
# ASSUME: ITP Tracker headers in row 1 include "ITP No"; NCR Register headers include "ITP No" and "Status"
# ASSUME: the real sheets are "ITP Tracker" and "NCR Register"; the TESTs use Itp and Ncr
# ASSUME: the NCR Register names only an ITP No, so every tracker row with that ITP No is flagged
# ASSUME: an NCR is open when its Status says Open, ignoring case and extra spaces; ITP No is trimmed and matched ignoring case
# ASSUME: NCR Flag belongs to this macro: it is added as the last column if missing, and every run rewrites it, clearing flags whose NCR is no longer open
# ASSUME: NCR rows with a blank ITP No are skipped
GOAL     Add an NCR Flag column to the ITP Tracker if it is missing, mark with Open NCR every point whose ITP has an open NCR, and tell how many points are flagged.
PROC     FlagItpPointsWithOpenNcr
KIND     Function
MODULE   modItp
ARGS     itp As Worksheet, ncr As Worksheet
RET      Long
VAR      cItp As Long
VAR      cFlag As Long
VAR      cNcrItp As Long
VAR      cNcrStatus As Long
VAR      openNcr As Dictionary
VAR      flagged As Long
PRE      COL($itp, "ITP No") > 0; COL($ncr, "ITP No") > 0; COL($ncr, "Status") > 0
MODIFIES $itp
OP       LET(cItp, COL($itp, "ITP No"))
OP       LET(cFlag, COL($itp, "NCR Flag"))
OP       LET(cNcrItp, COL($ncr, "ITP No"))
OP       LET(cNcrStatus, COL($ncr, "Status"))
OP       IF cFlag = 0
OP         LET(cFlag, LASTCOL($itp, 1) + 1)
OP         SET_VALUE(CELL($itp, 1, cFlag), "NCR Flag")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ncr, cNcrItp) var=r
OP         SKIP_IF IS_EMPTY(CELL($ncr, r, cNcrItp))
OP         IF SAME(CELL($ncr, r, cNcrStatus), "Open")
OP           DICT_ADD(openNcr, TRIM(CELL($ncr, r, cNcrItp)), 1, mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($itp, cItp) var=p
OP         IF EXISTS(openNcr, TRIM(CELL($itp, p, cItp)))
OP           SET_VALUE(CELL($itp, p, cFlag), "Open NCR")
OP           LET(flagged, flagged + 1)
OP         ELSE
OP           CLEAR(CELL($itp, p, cFlag), what=contents)
OP       MSG("ITP points flagged: " & flagged)
OP       RETURN(flagged)
PERF     screen_off
ERR      standard
FIXTURE  itp_ncr_01.xlsx sheet=Itp cols=A:ITP No,B:Point,C:Inspection Status rows=8 data=A:values("ITP-01"x1,"ITP-02"x1,"ITP-03"x1,"ITP-04"x1," ITP-02 "x1,"ITP-05"x1,""x1,"ITP-07"x1)+boundary("ITP-01","ITP-02","ITP-03","ITP-04"," ITP-02 ","ITP-05","","ITP-07");B:text(8);C:values("Done"x4,"Open"x4)
FIXTURE  itp_ncr_01.xlsx sheet=Ncr cols=A:NCR No,B:ITP No,C:Status rows=9 data=A:seq(1,1);B:values("ITP-02"x2,"ITP-03"x2," itp-05 "x1,"ITP-04"x1,""x1,"ITP-06"x1,"ITP-01"x1)+boundary("ITP-02","ITP-02","ITP-03"," itp-05 ","ITP-04","","ITP-06","ITP-01","ITP-03");C:values("Open"x5," open "x1,"Closed"x3)+boundary("Open","Open","Closed","Open"," open ","Open","Open","Closed","Closed")
FIXTURE  itp_ncr_02.xlsx sheet=Itp cols=A:ITP No,B:Point,C:NCR Flag,D:Inspection Status rows=8 data=A:values("ITP-01"x1,"ITP-02"x1,"ITP-03"x1,"ITP-04"x1," ITP-02 "x1,"ITP-05"x1,""x1,"ITP-07"x1)+boundary("ITP-01","ITP-02","ITP-03","ITP-04"," ITP-02 ","ITP-05","","ITP-07");B:text(8);C:values("Open NCR"x3,""x5)+boundary("Open NCR","Open NCR","Open NCR","","","","","");D:values("Done"x4,"Open"x4)
FIXTURE  itp_ncr_02.xlsx sheet=Ncr cols=A:NCR No,B:ITP No,C:Status rows=9 data=A:seq(1,1);B:values("ITP-02"x2,"ITP-03"x2," itp-05 "x1,"ITP-04"x1,""x1,"ITP-06"x1,"ITP-01"x1)+boundary("ITP-02","ITP-02","ITP-03"," itp-05 ","ITP-04","","ITP-06","ITP-01","ITP-03");C:values("Open"x5," open "x1,"Closed"x3)+boundary("Open","Open","Closed","Open"," open ","Open","Open","Closed","Closed")
TEST     fixture=itp_ncr_01.xlsx call=FlagItpPointsWithOpenNcr(Itp, Ncr) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND VALUE(Itp![NCR Flag]1) = "NCR Flag" AND COUNTIF(Itp![NCR Flag], "Open NCR") = 4 AND COUNTA(Itp![NCR Flag]) = 5 AND COUNTA(Itp!E:E) = 0 AND VALUE(Itp![NCR Flag]3) = "Open NCR" AND VALUE(Itp![NCR Flag]5) = "Open NCR" AND VALUE(Itp![NCR Flag]6) = "Open NCR" AND VALUE(Itp![NCR Flag]7) = "Open NCR"
TEST     fixture=itp_ncr_01.xlsx call=FlagItpPointsWithOpenNcr(Itp, Ncr) assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTA(Itp![ITP No]) = PRE_COUNTA(Itp![ITP No]) AND COUNTA(Itp![Point]) = PRE_COUNTA(Itp![Point]) AND COUNTA(Ncr![Status]) = PRE_COUNTA(Ncr![Status]) AND COUNTIF(Ncr![Status], "Open") = PRE_COUNTIF(Ncr![Status], "Open")
TEST     fixture=itp_ncr_02.xlsx call=FlagItpPointsWithOpenNcr(Itp, Ncr) assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTIF(Itp![NCR Flag], "Open NCR") = 4 AND COUNTA(Itp![NCR Flag]) = 5 AND COUNTA(Itp!E:E) = 0 AND COUNTA(Itp![Inspection Status]) = PRE_COUNTA(Itp![Inspection Status]) AND VALUE(Itp![NCR Flag]3) = "Open NCR" AND VALUE(Itp![NCR Flag]7) = "Open NCR"
TEST     fixture=itp_ncr_01.xlsx call=FlagItpPointsWithOpenNcr(Ncr, Itp) assert=RESULT = -1 AND ERROR_COUNT = 1
