VCL 0.6
TASK
# ASSUME: WPS Register headers in row 1 include "WPS No" and "Filler Class"; Consumable Stock headers include "Filler Class" and "Qty On Hand"
# ASSUME: filler classes match after trimming the spaces at both ends, ignoring case; spaces inside a class are kept
# ASSUME: a batch counts as stock only when its Qty On Hand is a number above zero; a blank or text quantity is no stock
# ASSUME: a stock row with a blank Filler Class is ignored, so a WPS with a blank Filler Class gets "No stock"
# ASSUME: a row with a blank WPS No is not a WPS: it is not counted and its Stock Remark is left as it is
# ASSUME: the Stock Remark column belongs to this macro: on every other WPS row whatever it holds is cleared
# ASSUME: a missing "Stock Remark" heading is added in the first free column of row 1
GOAL     Write No stock in the Stock Remark column of every WPS whose filler class has no stock batch with a quantity above zero, clear it on the others and return how many have no stock.
PROC     FlagWpsWithoutStock
KIND     Function
MODULE   modWps
ARGS     wps As Worksheet, stock As Worksheet
RET      Long
VAR      cWps As Long
VAR      cFiller As Long
VAR      cRemark As Long
VAR      sFiller As Long
VAR      sQty As Long
VAR      inStock As Dictionary
VAR      n As Long
PRE      COL($wps, "WPS No") > 0; COL($wps, "Filler Class") > 0; COL($stock, "Filler Class") > 0; COL($stock, "Qty On Hand") > 0
MODIFIES $wps
OP       LET(cWps, COL($wps, "WPS No"))
OP       LET(cFiller, COL($wps, "Filler Class"))
OP       LET(sFiller, COL($stock, "Filler Class"))
OP       LET(sQty, COL($stock, "Qty On Hand"))
OP       LET(cRemark, COL($wps, "Stock Remark"))
OP       IF cRemark = 0
OP         LET(cRemark, LASTCOL($wps, 1) + 1)
OP         SET_VALUE(CELL($wps, 1, cRemark), "Stock Remark")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($stock, sFiller) var=s
OP         SKIP_IF IS_EMPTY(CELL($stock, s, sFiller))
OP         IF IS_NUMBER(CELL($stock, s, sQty))
OP           IF CELL($stock, s, sQty) > 0
OP             DICT_ADD(inStock, TRIM(CELL($stock, s, sFiller)), 1, mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($wps, cWps) var=r
OP         SKIP_IF IS_EMPTY(CELL($wps, r, cWps))
OP         IF EXISTS(inStock, TRIM(CELL($wps, r, cFiller)))
OP           SET_VALUE(CELL($wps, r, cRemark), "")
OP         ELSE
OP           SET_VALUE(CELL($wps, r, cRemark), "No stock")
OP           LET(n, n + 1)
OP       RETURN(n)
PERF     screen_off
ERR      standard
FIXTURE  wps_stock_01.xlsx sheet=Wps cols=A:WPS No,B:Process,C:Filler Class rows=10 data=A:values("W-1"x1,"W-2"x1,"W-3"x1,""x1,"W-5"x1,"W-6"x1,"W-7"x1,"W-8"x1,"W-10"x1,"W-9"x1)+boundary("W-1","W-2","W-3","","W-5","W-6","W-7","W-8","W-10","W-9");B:values("SMAW"x4,"GTAW"x4,"FCAW"x1,"SAW"x1);C:values("E7018"x1," er70s-6"x1,"E308L"x1,"E9999"x2," e6013 "x1,""x1,"E316L"x1,"ER308L"x1,"e7018"x1)+boundary("E7018"," er70s-6","E308L","E9999"," e6013 ","","E316L","ER308L","e7018","E9999")
FIXTURE  wps_stock_01.xlsx sheet=Stock cols=A:Batch No,B:Filler Class,C:Qty On Hand rows=9 data=A:seq(101,1);B:values("E6013"x1,"E7018"x1," e7018 "x1,"ER70S-6"x2,"E308L"x1,""x1,"E316L"x1,"er308l "x1)+boundary("E6013","E7018"," e7018 ","ER70S-6","E308L","","ER70S-6","E316L","er308l ");C:int(1,20)+boundary(1,0,12,0,"n/a",5,-3,"",7)
FIXTURE  wps_stock_02.xlsx sheet=Wps cols=A:WPS No,B:Stock Remark,C:Process,D:Filler Class rows=10 data=A:values("W-1"x1,"W-2"x1,"W-3"x1,""x1,"W-5"x1,"W-6"x1,"W-7"x1,"W-8"x1,"W-10"x1,"W-9"x1)+boundary("W-1","W-2","W-3","","W-5","W-6","W-7","W-8","W-10","W-9");B:values("No stock"x3,""x5,"keep"x1,"old"x1)+boundary("No stock","","No stock","keep","old","","","No stock","","");C:values("SMAW"x4,"GTAW"x4,"FCAW"x1,"SAW"x1);D:values("E7018"x1," er70s-6"x1,"E308L"x1,"E9999"x2," e6013 "x1,""x1,"E316L"x1,"ER308L"x1,"e7018"x1)+boundary("E7018"," er70s-6","E308L","E9999"," e6013 ","","E316L","ER308L","e7018","E9999")
FIXTURE  wps_stock_02.xlsx sheet=Stock cols=A:Batch No,B:Filler Class,C:Qty On Hand rows=9 data=A:seq(101,1);B:values("E6013"x1,"E7018"x1," e7018 "x1,"ER70S-6"x2,"E308L"x1,""x1,"E316L"x1,"er308l "x1)+boundary("E6013","E7018"," e7018 ","ER70S-6","E308L","","ER70S-6","E316L","er308l ");C:int(1,20)+boundary(1,0,12,0,"n/a",5,-3,"",7)
TEST     fixture=wps_stock_01.xlsx call=FlagWpsWithoutStock(Wps, Stock) runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND VALUE(Wps!D1) = "Stock Remark" AND COUNTA(Wps!A1:Z1) = 4 AND COUNTIF(Wps![Stock Remark], "No stock") = 5 AND COUNTA(Wps![Stock Remark]) = 6 AND VALUE(Wps![Stock Remark]3) = "No stock" AND VALUE(Wps![Stock Remark]4) = "No stock" AND VALUE(Wps![Stock Remark]7) = "No stock" AND VALUE(Wps![Stock Remark]8) = "No stock" AND VALUE(Wps![Stock Remark]11) = "No stock"
TEST     fixture=wps_stock_02.xlsx call=FlagWpsWithoutStock(Wps, Stock) runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND COUNTA(Wps!A1:Z1) = 4 AND COUNTIF(Wps![Stock Remark], "No stock") = 5 AND COUNTA(Wps![Stock Remark]) = 7 AND VALUE(Wps![Stock Remark]2) = "" AND VALUE(Wps![Stock Remark]5) = "keep" AND VALUE(Wps![Stock Remark]6) = "" AND VALUE(Wps![Stock Remark]9) = "" AND VALUE(Wps![Stock Remark]11) = "No stock"
TEST     fixture=wps_stock_02.xlsx call=FlagWpsWithoutStock(Wps, Stock) assert=COUNTA(Wps![WPS No]) = PRE_COUNTA(Wps![WPS No]) AND COUNTA(Wps![Filler Class]) = PRE_COUNTA(Wps![Filler Class]) AND COUNTA(Stock![Filler Class]) = PRE_COUNTA(Stock![Filler Class]) AND SUM(Stock![Qty On Hand]) = PRE_SUM(Stock![Qty On Hand])
TEST     fixture=wps_stock_01.xlsx call=FlagWpsWithoutStock(Stock, Wps) assert=RESULT = -1 AND ERROR_COUNT = 1
