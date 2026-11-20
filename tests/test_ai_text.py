"""VCL 0.7 part 3e (D112): AI_TEXT("...") marks text the AI wrote, which no TEST can prove right.

It means its quoted text in the interpreter and in VBA; `vclc check` and KRONOS's proof list every one for a person
to check. Also here: the compiler writes text beyond ASCII (° and –, common in KGC's forms) as ChrW(...) pieces.
"""
import datetime as dt
import io
import sys
import unittest
from contextlib import redirect_stdout
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
sys.path.insert(0, str(ROOT / "vcl"))

import vclc  # noqa: E402
from kronos_engine import prove  # noqa: E402
from validator import ai_texts, validate_text  # noqa: E402
from vba_check import check_text  # noqa: E402
from vcl_codegen import compile_text  # noqa: E402

TODAY = dt.date(2026, 11, 20)
PATH = ROOT / "tests" / "data" / "ai_text_wps.vcl"
PRACTICE = PATH.read_text(encoding="utf-8")
WORDS = [(12, "1. Grind the root back to sound metal before side B."),
         (13, "Keep the interpass temperature at or below 250 °C.")]


def errors(text):
    return [p.message for p in validate_text(text) if p.severity == "error"]


class AiText(unittest.TestCase):
    def test_every_ai_text_is_listed_with_its_line(self):
        self.assertEqual(ai_texts(PRACTICE), WORDS)
        self.assertEqual(ai_texts("not a spec"), [])

    def test_vclc_check_shows_them_as_notes_not_problems(self):
        out = io.StringIO()
        with redirect_stdout(out):
            ok = vclc.cmd_check([str(PATH)])
        self.assertTrue(ok)
        self.assertIn("line 13: note: written by the AI, not proven: check it: Keep the interpass", out.getvalue())

    def test_the_proof_carries_them_and_the_spec_is_proven(self):
        proof = prove(PRACTICE, TODAY)
        self.assertTrue(proof.ok, [f.message for f in proof.findings])
        self.assertEqual(proof.ai_texts, WORDS)

    def test_its_text_is_written_in_place_and_it_is_vcl_07(self):
        bad = PRACTICE.replace('AI_TEXT("1. Grind the root back to sound metal before side B.")',
                               'AI_TEXT("1. " & "Grind")')
        self.assertTrue(any("AI_TEXT needs one" in m for m in errors(bad)), errors(bad))
        self.assertTrue(any("is VCL 0.7" in m for m in errors(PRACTICE.replace("VCL 0.7", "VCL 0.6", 1))))

    def test_the_vba_writes_text_beyond_ascii_as_chrw_and_ends_where_the_interpreter_does(self):
        vba = compile_text(PRACTICE, "ai_text_wps")
        self.assertIn('"Keep the interpass temperature at or below 250 " & ChrW(176) & "C."', vba)
        self.assertTrue(vba.isascii())
        outcome = check_text(PRACTICE, "ai_text_wps", TODAY)
        self.assertTrue(outcome.proven, outcome.differences or outcome.skipped)


if __name__ == "__main__":
    unittest.main()
