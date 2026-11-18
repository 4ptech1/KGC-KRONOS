VCL 0.6
TASK
# ASSUME: the sheet is "Welder Quals" with headers in row 1: A "Welder ID", B "Process", C "Expiry Date"; the TESTs use a sheet named Welder Quals
# ASSUME: the formula counts the whole Expiry Date column so it stays right when rows are added; header text is not counted
# ASSUME: if Qual Check already exists it is left exactly as it is
GOAL     Add a Qual Check sheet if it is missing, with a label in A1 and a live formula in B1 counting expired qualifications in Welder Quals.
PROC     AddQualCheckSheet
KIND     Function
MODULE   modQuals
ARGS     wb As Workbook, quals As Worksheet
RET      Long
VAR      qc As Worksheet
VAR      added As Long
PRE      COL($quals, "Expiry Date") = 3; COL($quals, "Welder ID") = 1
MODIFIES $wb sheets
OP       IF NOT SHEET_EXISTS($wb, "Qual Check")
OP         ADD_SHEET($wb, name="Qual Check", into=qc)
OP         SET_VALUE(CELL(qc, 1, 1), "Expired quals")
OP         SET_FORMULA(CELL(qc, 1, 2), "=COUNTIF('Welder Quals'!C:C,""<""&TODAY())", style=a1)
OP         LET(added, 1)
OP       RETURN(added)
ERR      standard
FIXTURE  quals_01.xlsx sheet=Welder_Quals cols=A:Welder ID,B:Process,C:Expiry Date rows=6 data=A:seq(1,1);B:values("GTAW"x3,"SMAW"x3);C:reldates(-30,30)+boundary(-1,0,1)
TEST     fixture=quals_01.xlsx call=AddQualCheckSheet(quals_01.xlsx, Welder Quals) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND SHEET_EXISTS("Qual Check") AND VALUE('Qual Check'!A1) = "Expired quals" AND FORMULA('Qual Check'!B1) = "=COUNTIF('Welder Quals'!C:C,""<""&TODAY())" AND ROWS_USED('Qual Check') = 1
TEST     fixture=quals_01.xlsx call=AddQualCheckSheet(quals_01.xlsx, Welder Quals) assert=RESULT = 1 AND ERROR_COUNT = 0 AND COUNTA(Welder Quals![Welder ID]) = PRE_COUNTA(Welder Quals![Welder ID]) AND COUNTA(Welder Quals![Expiry Date]) = PRE_COUNTA(Welder Quals![Expiry Date])
