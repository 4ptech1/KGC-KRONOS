VCL 0.7
TASK
# ASSUME: both sheets have their headings in row 1; the Enq Summary sheet is the one named "Enq Summary" of the workbook read from
# ASSUME: the first item row is row 2 of Enq Summary; its Enquiry and Date are written on every new register row as that text
# ASSUME: an item whose Revised Tag No is already in the register is not added again, so a re-run adds nothing
# ASSUME: rows are added after the last row that holds an Item; the Sr.Nos. and other columns of a new row are left blank
GOAL     Append each item of the estimation workbook's Enq Summary sheet to the enquiry register after its last row and return how many rows were added.
PROC     AddEnquiryItems
KIND     Function
MODULE   modRegister
ARGS     reg As Worksheet, est As Workbook
RET      Long
VAR      src As Worksheet
VAR      known As Dictionary
VAR      cRegItem As Long
VAR      cRegTag As Long
VAR      cSrcItem As Long
VAR      cSrcTag As Long
VAR      outRow As Long
VAR      added As Long
PRE      SHEET_EXISTS($est, "Enq Summary"); COL($reg, "Item") > 0; COL($reg, "Revised Tag No") > 0; COL($reg, "Enquiry") > 0; COL($reg, "Date") > 0; COL($reg, "Type") > 0; COL($reg, "U Stamp") > 0; COL($reg, "Old Tag No") > 0; COL($reg, "Qty") > 0; COL($reg, "MOC") > 0; COL($reg, "Shell ID mm") > 0; COL($reg, "Shell THK mm") > 0; COL($reg, "TL To TL mm") > 0
PRE      COL(SHEET($est, "Enq Summary"), "Item") > 0; COL(SHEET($est, "Enq Summary"), "Revised Tag No") > 0; COL(SHEET($est, "Enq Summary"), "Enquiry") > 0; COL(SHEET($est, "Enq Summary"), "Date") > 0; COL(SHEET($est, "Enq Summary"), "Type") > 0; COL(SHEET($est, "Enq Summary"), "U Stamp") > 0; COL(SHEET($est, "Enq Summary"), "Old Tag No") > 0; COL(SHEET($est, "Enq Summary"), "Qty") > 0; COL(SHEET($est, "Enq Summary"), "MOC") > 0; COL(SHEET($est, "Enq Summary"), "Shell ID mm") > 0; COL(SHEET($est, "Enq Summary"), "Shell THK mm") > 0; COL(SHEET($est, "Enq Summary"), "TL To TL mm") > 0
MODIFIES $reg["Enquiry", "Date", "Item", "Type", "U Stamp", "Old Tag No", "Revised Tag No", "Qty", "MOC", "Shell ID mm", "Shell THK mm", "TL To TL mm"]
OP       LET(src, SHEET($est, "Enq Summary"))
OP       LET(cRegItem, COL($reg, "Item"))
OP       LET(cRegTag, COL($reg, "Revised Tag No"))
OP       LET(cSrcItem, COL(src, "Item"))
OP       LET(cSrcTag, COL(src, "Revised Tag No"))
OP       LET(outRow, LASTROW($reg, cRegItem) + 1)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($reg, cRegItem) var=q
OP         DICT_ADD(known, TRIM(CELL($reg, q, cRegTag)), 1, mode=overwrite)
OP       LOOP_ROWS dir=down from=2 to=LASTROW(src, cSrcItem) var=r
OP         SKIP_IF IS_EMPTY(CELL(src, r, cSrcItem))
OP         SKIP_IF EXISTS(known, TRIM(CELL(src, r, cSrcTag)))
OP         SET_VALUE(CELL($reg, outRow, COL($reg, "Enquiry")), CELL(src, 2, COL(src, "Enquiry")))
OP         SET_VALUE(CELL($reg, outRow, COL($reg, "Date")), CELL(src, 2, COL(src, "Date")))
OP         SET_VALUE(CELL($reg, outRow, cRegItem), CELL(src, r, cSrcItem))
OP         SET_VALUE(CELL($reg, outRow, COL($reg, "Type")), CELL(src, r, COL(src, "Type")))
OP         SET_VALUE(CELL($reg, outRow, COL($reg, "U Stamp")), CELL(src, r, COL(src, "U Stamp")))
OP         SET_VALUE(CELL($reg, outRow, COL($reg, "Old Tag No")), CELL(src, r, COL(src, "Old Tag No")))
OP         SET_VALUE(CELL($reg, outRow, cRegTag), CELL(src, r, cSrcTag))
OP         SET_VALUE(CELL($reg, outRow, COL($reg, "Qty")), CELL(src, r, COL(src, "Qty")))
OP         SET_VALUE(CELL($reg, outRow, COL($reg, "MOC")), CELL(src, r, COL(src, "MOC")))
OP         SET_VALUE(CELL($reg, outRow, COL($reg, "Shell ID mm")), CELL(src, r, COL(src, "Shell ID mm")))
OP         SET_VALUE(CELL($reg, outRow, COL($reg, "Shell THK mm")), CELL(src, r, COL(src, "Shell THK mm")))
OP         SET_VALUE(CELL($reg, outRow, COL($reg, "TL To TL mm")), CELL(src, r, COL(src, "TL To TL mm")))
OP         DICT_ADD(known, TRIM(CELL(src, r, cSrcTag)), 1, mode=overwrite)
OP         LET(outRow, outRow + 1)
OP         LET(added, added + 1)
OP       RETURN(added)
PERF     screen_off
ERR      standard
FIXTURE  register_01.xlsx template=enquiry_register
FIXTURE  estimation_01.xlsx template=estimation
TEST     fixture=register_01.xlsx read=estimation_01.xlsx call=AddEnquiryItems(Sheet1, estimation_01.xlsx) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND ROWS_USED(Sheet1) = 9
TEST     fixture=register_01.xlsx read=estimation_01.xlsx call=AddEnquiryItems(Sheet1, estimation_01.xlsx) assert=RESULT = 2 AND ERROR_COUNT = 0 AND COUNTA(Sheet1![Item]) = PRE_COUNTA(Sheet1![Item]) + 2 AND VALUE(Sheet1![Enquiry]8) = "E-9100 (TBA)" AND VALUE(Sheet1![Enquiry]9) = "E-9100 (TBA)" AND VALUE(Sheet1![Date]9) = "02.11.2026" AND VALUE(Sheet1![Item]8) = "Channel Head Assembly" AND VALUE(Sheet1![Item]9) = "Floating Head Assembly" AND VALUE(Sheet1![Type]9) = "HE Part" AND VALUE(Sheet1![U Stamp]8) = "No"
TEST     fixture=register_01.xlsx read=estimation_01.xlsx call=AddEnquiryItems(Sheet1, estimation_01.xlsx) assert=VALUE(Sheet1![Old Tag No]8) = "DM-1001A" AND VALUE(Sheet1![Revised Tag No]9) = "DM-1001B" AND VALUE(Sheet1![Qty]8) = 2 AND VALUE(Sheet1![MOC]9) = "SA 516 Gr 70" AND VALUE(Sheet1![Shell ID mm]8) = 700 AND VALUE(Sheet1![Shell THK mm]9) = 18 AND VALUE(Sheet1![TL To TL mm]8) = 1200 AND VALUE(Sheet1![Shell ID mm]9) = 640 AND VALUE(Sheet1![Date]8) = "02.11.2026" AND VALUE(Sheet1![TL To TL mm]9) = 900 AND VALUE(Sheet1![Shell THK mm]8) = 16 AND VALUE(Sheet1![Old Tag No]9) = "DM-1001B" AND VALUE(Sheet1![Qty]9) = 2 AND VALUE(Sheet1![MOC]8) = "SA 516 Gr 70" AND VALUE(Sheet1![Type]8) = "HE Part" AND VALUE(Sheet1![U Stamp]9) = "No" AND VALUE(Sheet1![Revised Tag No]8) = "DM-1001A"
