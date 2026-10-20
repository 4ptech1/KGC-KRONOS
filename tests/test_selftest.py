"""compiler/vcl_selftest.py: the ForgeLib self-test that runs every conformance case in Excel (D34)."""
import datetime as dt
import re
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
sys.path.insert(0, str(ROOT / "vcl"))
sys.path.insert(0, str(ROOT / "tests"))
from conformance_cases import CASES  # noqa: E402
from semantics import XlError  # noqa: E402
from vba_lint import lint, public_helpers  # noqa: E402
from vcl_selftest import selftest_module, vba_text, vba_value  # noqa: E402

FORGELIB = (ROOT / "vba" / "ForgeLib.bas").read_text(encoding="utf-8")


class Values(unittest.TestCase):
    def test_vba_literals(self):
        cases = [(None, "Empty"), (True, "True"), (25, "25#"), (-7, "-7#"), (25.0, "25#"), (2.5, "2.5"),
                 ("W-01", '"W-01"'), ('J"1', '"J""1"'), ("", '""'), (" x\n", 'ChrW(160) & "x" & ChrW(10)'),
                 ("café", '"caf" & ChrW(233)'), (dt.date(2026, 10, 15), "DateSerial(2026, 10, 15)"),
                 (dt.datetime(2026, 10, 15, 8, 30), "(DateSerial(2026, 10, 15) + TimeSerial(8, 30, 0))"),
                 (XlError("#N/A"), "CVErr(xlErrNA)"), (XlError("#DIV/0!"), "CVErr(xlErrDiv0)"),
                 (["a", None], 'Array("a", Empty)')]
        for value, vba in cases:
            with self.subTest(value=value):
                self.assertEqual(vba_value(value), vba)

    def test_text_is_plain_ascii_in_vba(self):
        self.assertTrue(all(32 <= ord(ch) <= 126 for ch in vba_text("\x01\x07\tßé ok")))

    def test_no_plain_literal_for_odd_numbers(self):
        with self.assertRaises(ValueError):
            vba_value(1e-20)


class Module(unittest.TestCase):
    MODULE = selftest_module()

    def test_every_case_in_order(self):
        found = re.findall(r"^    SelfTestCase (\d+), ", self.MODULE, re.M)
        self.assertEqual([int(n) for n in found], list(range(1, len(CASES) + 1)))
        self.assertEqual(len(re.findall(r"^    SelfTestCheck ", self.MODULE, re.M)), len(CASES))
        self.assertIn(f"    total = {len(CASES)}\n", self.MODULE)

    def test_each_meaning_is_checked_through_forgelib(self):
        checks = re.findall(r"^    SelfTestCheck (\w+)", self.MODULE, re.M)
        used = {name for name in checks if name.startswith("Forge")}
        self.assertTrue(used <= public_helpers(FORGELIB), used - public_helpers(FORGELIB))
        for name in ("ForgeSafeName", "ForgePathJoin", "ForgeParseDate", "ForgeSame", "ForgeNeedsTextGuard"):
            self.assertIn(name, used)
        self.assertIn("ForgeCol(gScratch, header)", self.MODULE)        # COL and LASTROW on a scratch sheet
        self.assertIn("ForgeLastRow(gScratch, 1)", self.MODULE)
        self.assertIn('ForgeDictAdd d, key, 1, "overwrite"', self.MODULE)  # S10 through a real dictionary

    def test_lint_clean_next_to_forgelib(self):
        self.assertEqual(lint(self.MODULE, public_helpers(FORGELIB)), [])
        unknown = lint(self.MODULE)   # alone it calls helpers it doesn't define: the check is live
        self.assertTrue(unknown and all("unknown helper Forge" in msg for _, msg in unknown))

    def test_a_case_that_raises_does_not_stop_the_run(self):
        parts = re.findall(r"Private Sub SelfTestPart\d+\(\)\n.*?End Sub", self.MODULE, re.S)
        self.assertEqual(len(parts), -(-len(CASES) // 40))
        for part in parts:
            self.assertIn("    On Error Resume Next\n", part)
            self.assertTrue(part.endswith("    SelfTestFlush\nEnd Sub"))

    def test_the_committed_module_is_up_to_date(self):
        committed = (ROOT / "dist" / "forge_selftest.bas").read_text(encoding="utf-8")
        self.assertTrue(committed == self.MODULE, "dist/forge_selftest.bas is out of date: run "
                                                  "python compiler/vclc.py all specs/*.vcl -o dist and commit it")


if __name__ == "__main__":
    unittest.main()
