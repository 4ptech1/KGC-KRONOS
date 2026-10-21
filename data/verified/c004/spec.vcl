VCL 0.6
TASK
# ASSUME: the RT, UT, MPT and DPT sheets each have headers "Weld Seam No" and "Result" in row 1
# ASSUME: seam numbers are compared as trimmed text, ignoring case
# ASSUME: if a seam appears more than once in a sheet (re-shoot), the lowest row is the latest result
# ASSUME: Summary has one row per seam found in any of the four sheets; it is rebuilt every run, sorted by seam
GOAL     Rebuild a Summary sheet with one row for each weld seam and its latest result from the RT, UT, MPT and DPT sheets, sorted by seam number.
PROC     BuildNdeSummary
KIND     Function
MODULE   modNde
ARGS     rt As Worksheet, ut As Worksheet, mpt As Worksheet, dpt As Worksheet, out As Worksheet
RET      Long
VAR      seams As Dictionary
VAR      rtResults As Dictionary
VAR      utResults As Dictionary
VAR      mptResults As Dictionary
VAR      dptResults As Dictionary
VAR      cRtSeam As Long
VAR      cRtRes As Long
VAR      cUtSeam As Long
VAR      cUtRes As Long
VAR      cMptSeam As Long
VAR      cMptRes As Long
VAR      cDptSeam As Long
VAR      cDptRes As Long
VAR      v As Variant
VAR      outRow As Long
PRE      COL($rt, "Weld Seam No") > 0; COL($rt, "Result") > 0; COL($ut, "Weld Seam No") > 0; COL($ut, "Result") > 0; COL($mpt, "Weld Seam No") > 0; COL($mpt, "Result") > 0; COL($dpt, "Weld Seam No") > 0; COL($dpt, "Result") > 0
MODIFIES $out
OP       LET(cRtSeam, COL($rt, "Weld Seam No"))
OP       LET(cRtRes, COL($rt, "Result"))
OP       LET(cUtSeam, COL($ut, "Weld Seam No"))
OP       LET(cUtRes, COL($ut, "Result"))
OP       LET(cMptSeam, COL($mpt, "Weld Seam No"))
OP       LET(cMptRes, COL($mpt, "Result"))
OP       LET(cDptSeam, COL($dpt, "Weld Seam No"))
OP       LET(cDptRes, COL($dpt, "Result"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Weld Seam No", "RT", "UT", "MPT", "DPT")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($rt, cRtSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($rt, r, cRtSeam))
OP         DICT_ADD(seams, TRIM(CELL($rt, r, cRtSeam)), 1, mode=keep)
OP         DICT_ADD(rtResults, TRIM(CELL($rt, r, cRtSeam)), CELL($rt, r, cRtRes), mode=overwrite)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ut, cUtSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($ut, r, cUtSeam))
OP         DICT_ADD(seams, TRIM(CELL($ut, r, cUtSeam)), 1, mode=keep)
OP         DICT_ADD(utResults, TRIM(CELL($ut, r, cUtSeam)), CELL($ut, r, cUtRes), mode=overwrite)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($mpt, cMptSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($mpt, r, cMptSeam))
OP         DICT_ADD(seams, TRIM(CELL($mpt, r, cMptSeam)), 1, mode=keep)
OP         DICT_ADD(mptResults, TRIM(CELL($mpt, r, cMptSeam)), CELL($mpt, r, cMptRes), mode=overwrite)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($dpt, cDptSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($dpt, r, cDptSeam))
OP         DICT_ADD(seams, TRIM(CELL($dpt, r, cDptSeam)), 1, mode=keep)
OP         DICT_ADD(dptResults, TRIM(CELL($dpt, r, cDptSeam)), CELL($dpt, r, cDptRes), mode=overwrite)
OP       LET(outRow, 2)
OP       FOR_EACH var=s in=KEYS(seams)
OP         SET_VALUE(CELL($out, outRow, 1), s)
OP         LET(v, "")
OP         IF EXISTS(rtResults, s) THEN DICT_GET(rtResults, s, into=v)
OP         IF IS_EMPTY(v) THEN LET(v, "Pending")
OP         SET_VALUE(CELL($out, outRow, 2), v)
OP         LET(v, "")
OP         IF EXISTS(utResults, s) THEN DICT_GET(utResults, s, into=v)
OP         IF IS_EMPTY(v) THEN LET(v, "Pending")
OP         SET_VALUE(CELL($out, outRow, 3), v)
OP         LET(v, "")
OP         IF EXISTS(mptResults, s) THEN DICT_GET(mptResults, s, into=v)
OP         IF IS_EMPTY(v) THEN LET(v, "Pending")
OP         SET_VALUE(CELL($out, outRow, 4), v)
OP         LET(v, "")
OP         IF EXISTS(dptResults, s) THEN DICT_GET(dptResults, s, into=v)
OP         IF IS_EMPTY(v) THEN LET(v, "Pending")
OP         SET_VALUE(CELL($out, outRow, 5), v)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=1, order=asc, header=yes)
OP       RETURN(outRow - 2)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  nde_summary_01.xlsx sheet=RT cols=A:Weld Seam No,B:Result rows=30 data=A:seq(33,-1)+boundary(101,111,111)+blank(2);B:values("Accepted"x25,"Repair"x2,""x3)+boundary("Accepted","Accepted","Repair")
FIXTURE  nde_summary_01.xlsx sheet=UT cols=A:Weld Seam No,B:Result rows=20 data=A:seq(1,1)+boundary(102,112,112)+blank(2);B:values("Accepted"x17,"Repair"x1,""x2)+boundary("Accepted","Repair","Accepted")
FIXTURE  nde_summary_01.xlsx sheet=MPT cols=A:Weld Seam No,B:Result rows=25 data=A:seq(6,1)+boundary(103,113,113)+blank(2);B:values("Accepted"x22,"Repair"x1,""x2)+boundary("Accepted","Accepted","Repair")
FIXTURE  nde_summary_01.xlsx sheet=DPT cols=A:Weld Seam No,B:Result rows=10 data=A:seq(1,1)+boundary(104,114,114)+blank(1);B:values("Accepted"x9,"Repair"x1)+boundary("Accepted","Repair","Accepted")
FIXTURE  nde_summary_01.xlsx sheet=Summary cols=A:Weld Seam No,B:RT,C:UT,D:MPT,E:DPT rows=0 stale=40
TEST     fixture=nde_summary_01.xlsx call=BuildNdeSummary(RT, UT, MPT, DPT, Summary) runs=2 assert=RESULT = 35 AND COUNTA(Summary![Weld Seam No]) = 36 AND COUNTIF(Summary![RT], "stale") = 0 AND ERROR_COUNT = 0
TEST     fixture=nde_summary_01.xlsx call=BuildNdeSummary(RT, UT, MPT, DPT, Summary) assert=COUNTIF(Summary![RT], "Pending") = 11 AND COUNTIF(Summary![UT], "Pending") = 20 AND COUNTIF(Summary![MPT], "Pending") = 15 AND COUNTIF(Summary![DPT], "Pending") = 27 AND COUNTIFS(Summary![Weld Seam No], 111, Summary![RT], "Repair") = 1 AND COUNTIFS(Summary![Weld Seam No], 112, Summary![UT], "Accepted") = 1 AND COUNTIFS(Summary![Weld Seam No], 113, Summary![MPT], "Repair") = 1 AND COUNTIFS(Summary![Weld Seam No], 114, Summary![DPT], "Accepted") = 1
TEST     fixture=nde_summary_01.xlsx call=BuildNdeSummary(RT, UT, MPT, DPT, Summary) assert=COUNTA(RT![Result]) = PRE_COUNTA(RT![Result]) AND COUNTA(UT![Result]) = PRE_COUNTA(UT![Result]) AND COUNTA(MPT![Result]) = PRE_COUNTA(MPT![Result]) AND COUNTA(DPT![Result]) = PRE_COUNTA(DPT![Result])
TEST     fixture=nde_summary_01.xlsx call=BuildNdeSummary(RT, Summary, MPT, DPT, Summary) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(Summary![Weld Seam No]) = PRE_COUNTA(Summary![Weld Seam No])
