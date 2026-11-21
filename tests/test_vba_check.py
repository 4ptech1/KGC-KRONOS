"""The compiler's VBA, run in the VBA simulator, ends where the reference interpreter does (D90 Phase 1, D99, D100).

Every committed spec, example and corpus pair, BUILDs and folder FIXTUREs included, is checked TEST by TEST: RESULT, ERROR_COUNT, MSG and LOG, and every sheet's cells, fills,
number formats and hidden rows. A bug planted in compiled VBA must be caught, so the check is not just agreeing.
"""
import datetime as dt
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))

import vba_check  # noqa: E402
import vba_sim as V  # noqa: E402
from vba_lint import lint  # noqa: E402
from vcl_codegen import assemble, compile_build, compile_text  # noqa: E402

TODAY = dt.date(2026, 11, 19)
SPECS = (sorted((ROOT / "specs").glob("*.vcl")) + sorted((ROOT / "vcl" / "examples").glob("*.vcl"))
         + sorted((ROOT / "data").glob("verified/c*/spec.vcl")) + sorted((ROOT / "data").glob("eval/c*/spec.vcl")))


def name_of(path):
    return path.parent.name if path.name == "spec.vcl" else path.name


class EverySpec(unittest.TestCase):
    def test_the_vba_ends_where_the_interpreter_does(self):
        proven, skipped = 0, []
        for path in SPECS:
            with self.subTest(spec=name_of(path)):
                out = vba_check.check_text(path.read_text(encoding="utf-8"), name_of(path), TODAY)
                self.assertEqual(out.differences, [])
                proven += out.proven
                if out.skipped is not None:
                    skipped.append(path.name)
        self.assertGreaterEqual(proven, 165)   # every spec, example and corpus pair with TESTs on 19 Nov 2026
        # build_example.vcl doesn't parse on purpose; Example 5 is a DOCUMENT, which makes no VBA (D115)
        self.assertEqual(sorted(skipped), ["build_example.vcl", "example5_document.vcl"])


class PlantedBugs(unittest.TestCase):
    def module(self, path):
        return assemble([compile_text(path.read_text(encoding="utf-8"), path.name)])

    def check(self, path, old, new):
        text = path.read_text(encoding="utf-8")
        module = self.module(path)
        self.assertEqual(module.count(old), 1, old)
        return vba_check.check_text(text, path.name, TODAY, module=module.replace(old, new))

    def test_a_wrong_sort_order_is_caught(self):
        out = self.check(ROOT / "specs" / "task_09.vcl", "ForgeSort ForgeUsedRange(out), cDate_v, True, True",
                         "ForgeSort ForgeUsedRange(out), cDate_v, False, True")
        self.assertTrue(any("cells differ" in d for d in out.differences), out.differences)

    def test_a_wrong_result_is_caught(self):
        out = self.check(ROOT / "specs" / "task_09.vcl", "BuildOpenNcrSummary = (outRow - 2)",
                         "BuildOpenNcrSummary = (outRow - 1)")
        self.assertTrue(any("RESULT is" in d for d in out.differences), out.differences)

    def test_a_wrong_pre_result_is_caught(self):
        path = ROOT / "specs" / "task_01.vcl"
        line = next(ln for ln in self.module(path).splitlines() if "PRE failed: (COL($bom,\"\"Required Qty\"\")>0)" in ln)
        out = self.check(path, line, line.replace("BuildShortageList = -1", "BuildShortageList = 0"))
        self.assertTrue(any("RESULT is" in d for d in out.differences), out.differences)

    def test_a_date_written_where_a_number_belongs_is_caught(self):
        # the bug D99 fixed: Date - ForgeNum(cell) is a Date in VBA, so Excel showed 06/10/1900, not 280
        out = self.check(ROOT / "specs" / "task_09.vcl", "(Date - CDate(ncr.Cells(r, cDate_v).Value))",
                         "(Date - ForgeNum(ncr.Cells(r, cDate_v).Value))")
        self.assertTrue(any("'1900-10-06'" in d for d in out.differences), out.differences)


class FilesAndBuilds(unittest.TestCase):
    """Part 3 (D100): the file ops on the interpreter's mock file system, and a BUILD through RunHeadless."""

    def check(self, path, old, new, build=False):
        text = path.read_text(encoding="utf-8")
        module = compile_build(text, path.name)[1] if build else assemble([compile_text(text, path.name)])
        self.assertIn(old, module)
        return vba_check.check_text(text, path.name, TODAY, module=module.replace(old, new, 1)).differences

    def test_a_wrong_copy_status_is_caught(self):
        found = self.check(ROOT / "specs" / "task_12.vcl", 'ForgeCopyFile = "Copied"', 'ForgeCopyFile = "Already there"')
        self.assertTrue(any("cells differ" in d for d in found), found)

    def test_a_copy_that_never_happens_is_caught(self):
        found = self.check(ROOT / "specs" / "task_12.vcl", "fso.CopyFile src, dst, False", "")
        self.assertTrue(found)

    def test_a_wrong_text_file_is_caught(self):
        found = self.check(ROOT / "data" / "verified" / "c067" / "spec.vcl", "ts.Write text", 'ts.Write text & "x"')
        self.assertTrue(any("files differ" in d for d in found), found)

    def test_a_wrong_read_is_caught(self):
        found = self.check(ROOT / "data" / "verified" / "c046" / "spec.vcl", 'ForgeReadText = "Read"', 'ForgeReadText = "Missing"')
        self.assertTrue(found)

    def test_a_build_step_that_writes_nothing_is_caught(self):
        found = self.check(ROOT / "vcl" / "examples" / "example4_build.vcl",
                           'ForgeSetValue ncr.Cells(r, cStatus), "Overdue"', "' (removed)", build=True)
        self.assertTrue(found)


class ReservedNames(unittest.TestCase):
    def test_a_variable_named_like_a_vba_function_is_renamed(self):
        vba = compile_text((ROOT / "specs" / "task_09.vcl").read_text(encoding="utf-8"), "task_09.vcl")
        self.assertIn("Dim cDate_v As Long", vba)
        self.assertNotIn("Dim cDate As", vba)

    def test_the_linter_and_the_simulator_refuse_one(self):
        bad = "Option Explicit\nPublic Sub S()\n    Dim cDate As Long\n    cDate = 1\nEnd Sub\n"
        self.assertTrue(any("VBA reserves" in p for _, p in lint(bad)))
        with self.assertRaises(V.VbaUnsupported):
            V.Module(bad)


if __name__ == "__main__":
    unittest.main()
