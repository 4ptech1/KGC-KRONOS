"""D94: <, >, <= and >= on numbers typed as text compare as numbers, in the interpreter and in the VBA.

Before D94 the interpreter compared two text values as text ("10" > "9" was FALSE) while the compiled VBA
put both through ForgeNum and compared numbers, so a TEST could pass on an answer Excel would not give.
"""
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
sys.path.insert(0, str(ROOT / "vcl"))

import semantics as S  # noqa: E402
from vcl_codegen import compile_text  # noqa: E402
from vcl_interp import run_tests  # noqa: E402

SPEC = """VCL 0.6
TASK
# ASSUME: WPS headers in row 1 include "WPS No", "Thickness Min", "Thickness Max" and "Remark"
GOAL     Write Range? in Remark where Thickness Min is bigger than Thickness Max and return how many.
PROC     MarkBadRanges
KIND     Function
MODULE   modWps
ARGS     ws As Worksheet
RET      Long
VAR      cNo As Long
VAR      cMin As Long
VAR      cMax As Long
VAR      cRemark As Long
VAR      n As Long
PRE      COL($ws, "WPS No") > 0; COL($ws, "Thickness Min") > 0; COL($ws, "Thickness Max") > 0; COL($ws, "Remark") > 0
MODIFIES $ws["Remark"]
OP       LET(cNo, COL($ws, "WPS No"))
OP       LET(cMin, COL($ws, "Thickness Min"))
OP       LET(cMax, COL($ws, "Thickness Max"))
OP       LET(cRemark, COL($ws, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cNo) var=r
OP         IF IS_NUMBER(CELL($ws, r, cMin))
OP           IF IS_NUMBER(CELL($ws, r, cMax))
OP             IF CELL($ws, r, cMin) > CELL($ws, r, cMax)
OP               SET_VALUE(CELL($ws, r, cRemark), "Range?")
OP               LET(n, n + 1)
OP       RETURN(n)
PERF     screen_off
ERR      standard
FIXTURE  wps_01.xlsx sheet=Wps cols=A:WPS No,B:Thickness Min,C:Thickness Max,D:Remark rows=4 data=A:seq(1,1);B:values("9"x1,"10"x1,"3"x1,"20"x1)+boundary("9","10","3","20");C:values("10"x1,"9"x1,"3"x1,"5"x1)+boundary("10","9","3","5");D:values(""x4)
TEST     fixture=wps_01.xlsx call=MarkBadRanges(Wps) assert=RESULT = 2 AND ERROR_COUNT = 0 AND VALUE(Wps![Remark]3) = "Range?" AND VALUE(Wps![Remark]5) = "Range?" AND VALUE(Wps![Remark]2) = ""
"""


class TextNumbers(unittest.TestCase):
    def test_the_rule(self):
        self.assertEqual([S.order("9", "10"), S.order("10", " 9 "), S.order("W-10", "W-9"), S.order("abc", "ABC")],
                         [-1, 1, -1, 1])

    def test_the_interpreter_compares_text_numbers_as_numbers(self):
        results = run_tests(SPEC)
        self.assertTrue(all(r.passed for r in results), [r.failed for r in results])

    def test_the_vba_calls_forge_compare_for_two_workbook_values(self):
        vba = compile_text(SPEC)
        self.assertIn("ForgeCompare(", vba)
        self.assertNotIn("ForgeNum(ws.Cells", vba.split("ForgeCompare(")[1].split("\n")[0])


if __name__ == "__main__":
    unittest.main()
