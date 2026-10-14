"""Unit tests for harness/assemble.py and the pilot batch. Run: python -m unittest discover tests"""
import csv
import io
import os
import re
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path[:0] = [os.path.join(ROOT, "compiler"), os.path.join(ROOT, "harness")]

import assemble  # noqa: E402


def read(path):
    with open(os.path.join(ROOT, path), encoding="utf-8") as f:
        return f.read()


class PilotBatch(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.bas, cls.csv = assemble.build()

    def test_committed_files_are_up_to_date(self):
        self.assertEqual(read(assemble.BATCH["bas"]), self.bas, "run python harness/assemble.py")
        self.assertEqual(read(assemble.BATCH["csv"]), self.csv, "run python harness/assemble.py")

    def test_batch_is_safe_to_paste(self):
        self.assertEqual(assemble.paste_problems(self.bas), [])

    def test_every_test_procedure_is_run(self):
        defined = re.findall(r"^Private Sub (ForgeTest_\w+)\(\)", self.bas, re.M)
        listed = re.findall(r"^    (ForgeTest_\w+)$", self.bas.split("Private Sub ForgeBatchTests()")[1], re.M)
        self.assertEqual(defined, ["ForgeTest_B000_T01", "ForgeTest_B000_T02", "ForgeTest_B000_T03",
                                   "ForgeTest_B000_T04"])
        self.assertEqual(listed, defined)
        self.assertIn(f"ForgeBatchTestCount = {len(defined)}", self.bas)

    def test_alt_f8_lists_only_runbatch_and_forgereset(self):
        macros = re.findall(r"^Public Sub (\w+)\(\)", self.bas, re.M)
        self.assertEqual(macros, ["RunBatch", "ForgeReset"])

    def test_expected_results(self):
        rows = list(csv.DictReader(io.StringIO(self.csv)))
        self.assertEqual({r["expected_outcome"] for r in rows}, {"PASS"})
        value = {(r["test_id"], r["measure"]): int(r["expected_value"]) for r in rows}
        result = value[("B000-T01", "RESULT")]
        parts = sum(value[("B000-T01", f"Shortages rows: {kind}")]
                    for kind in ("Not received", "Short", "Required qty missing"))
        self.assertEqual(result, parts)
        self.assertEqual(value[("B000-T01", "COUNTA(Shortages![Part No])")], result + 1)
        self.assertEqual(value[("B000-T02", 'COUNTIF(Shortages![Remark], "Required qty missing")')], 3)
        self.assertEqual(value[("B000-T04", "RESULT")], -1)
        self.assertEqual(value[("B000-T04", "ERROR_COUNT")], 1)


class PasteChecks(unittest.TestCase):
    """The checks must catch the mistakes that would stop a pasted module from compiling."""

    @classmethod
    def setUpClass(cls):
        cls.bas, _ = assemble.build()

    def assertCaught(self, text, fragment):
        self.assertNotEqual(text, self.bas)
        problems = assemble.paste_problems(text)
        self.assertTrue(any(fragment in p for p in problems), problems)

    def test_faults_are_caught(self):
        body = self.bas[len("Option Explicit\n"):]
        cases = {
            "is not declared in BuildShortageList": self.bas.replace("330         got = 0", "330         gotx = 0"),
            "ForgeEndTst is not defined": self.bas.replace("ForgeEndTest\n    Exit Sub", "ForgeEndTst\n    Exit Sub", 1),
            "no Attribute lines": "Option Explicit\nAttribute VB_Name = \"X\"\n" + body,
            "must start with Option Explicit": body,
            "no Select or Activate": self.bas.replace("out.UsedRange.Clear", "out.UsedRange.Select"),
            "non-ASCII character": self.bas.replace('"Remark"', '"Remark–"', 1),
            "missing label Donee": self.bas.replace("Resume Done", "Resume Donee", 1),
            "line number 330 is used 2 times": self.bas.replace("340         If received", "330         If received"),
            "needs ' _' at the very end": self.bas.replace('True, 2, _', 'True, 2,_', 1),
            "only CreateObject": self.bas.replace('"Scripting.Dictionary"', '"WScript.Shell"'),
            "no Kill, RmDir, Shell": self.bas.replace("out.UsedRange.Clear", 'Kill "x.txt"'),
            "characters (VBA allows 1023)": self.bas.replace("' Pilot code", "' " + "x" * 1100, 1),
        }
        for fragment, text in cases.items():
            with self.subTest(fragment=fragment):
                self.assertCaught(text, fragment)


if __name__ == "__main__":
    unittest.main()
