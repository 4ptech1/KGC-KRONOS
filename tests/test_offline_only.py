"""KRONOS Offline version 6, part 6a (D118): recipes only KRONOS Offline runs.

A DOCUMENT (D115) or a recipe whose TEST reads a second workbook (read=, D114) is proven without VBA when KRONOS Offline
proves it (prove(..., offline=True): the page's kronos_web.check and the wizard's runs); KRONOS.exe (offline=False)
still refuses it at the build. Save offers no .bas for it. A workbook a recipe only reads is loaded with Excel's saved
values, shared formulas and error values too (vcl_template.load(..., reading=True)).
"""
import datetime as dt
import io
import json
import re
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
sys.path.insert(0, str(ROOT / "vcl"))

import kronos_web  # noqa: E402
from kronos_engine import OFFLINE_SAVE, prove  # noqa: E402
from kronos_wizard import Wizard, WizardError  # noqa: E402
from semantics import XlError  # noqa: E402
from vcl_fixtures import Formula  # noqa: E402
from vcl_template import TemplateError, load  # noqa: E402

TODAY = dt.date(2026, 11, 20)
DOCUMENT = (ROOT / "vcl" / "examples" / "example5_document.vcl").read_text(encoding="utf-8")
READ = (ROOT / "tests" / "data" / "wps_from_wps.vcl").read_text(encoding="utf-8")
TASK = (ROOT / "vcl" / "examples" / "example1_rev3.vcl").read_text(encoding="utf-8")


class Proofs(unittest.TestCase):
    def test_kronos_exe_refuses_them_at_the_build_and_kronos_offline_proves_them(self):
        for text, name in ((DOCUMENT, "FillProposal"), (READ, "FillWpsFromOld")):
            with self.subTest(name=name):
                exe = prove(text, TODAY)
                self.assertFalse(exe.ok)
                self.assertEqual([f.stage for f in exe.errors], ["build"])
                offline = prove(text, TODAY, offline=True)
                self.assertTrue(offline.ok, [f.message for f in offline.errors])
                self.assertTrue(offline.offline_only)
                self.assertEqual((offline.module, offline.name), ("", name))

    def test_an_ordinary_recipe_is_proven_with_its_vba_either_way(self):
        proof = prove(TASK, TODAY, offline=True)
        self.assertTrue(proof.ok and proof.module and not proof.offline_only)

    def test_the_page_says_proven_with_no_macro(self):
        doc = json.loads(kronos_web.check(DOCUMENT, TODAY))
        self.assertEqual((doc["proven"], doc["offline_only"], doc["bas"], doc["module"]), (True, True, "", ""))
        task = json.loads(kronos_web.check(TASK, TODAY))
        self.assertTrue(task["proven"] and not task["offline_only"] and task["bas"])


class TheWizard(unittest.TestCase):
    def test_a_recipe_from_a_file_is_proven_and_has_no_macro_to_save(self):
        wizard = Wizard(tempfile.mkdtemp(), TODAY)
        view = wizard.recipe("r1", DOCUMENT, "example5_document.vcl")
        self.assertTrue(view["best"]["proven"])
        self.assertTrue(view["best"]["offline_only"])
        self.assertIsNone(view["starter"])
        self.assertEqual([t["text"] for t in view["best"]["ai_texts"]][:1],
                         ["Offer based on the enquiry data sheet and ASME VIII-1, latest edition."])
        with self.assertRaises(WizardError) as caught:
            wizard.save("r1", {})
        self.assertEqual(str(caught.exception), OFFLINE_SAVE)


def doctored(change):
    """The WPS look-alike with its sheet XML changed: change(xml bytes) -> bytes; plus a link to another workbook."""
    src = zipfile.ZipFile(io.BytesIO((ROOT / "data" / "templates" / "wps_form.xlsx").read_bytes()))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as z:
        for n in src.namelist():
            body = src.read(n)
            z.writestr(n, change(body) if n == "xl/worksheets/sheet1.xml" else body)
        z.writestr("xl/externalLinks/externalLink1.xml", b"<externalLink/>")
    return out.getvalue()


class ReadingAWorkbook(unittest.TestCase):
    def test_saved_values_of_every_formula_and_error_values_are_read(self):
        data = doctored(lambda b: re.sub(rb"<f>(L29&amp;)", rb'<f t="shared" ref="D29" si="0">\1', b, count=1)
                        .replace(b'<c r="C25" t="n"><v>12</v></c>', b'<c r="C25" t="e"><v>#N/A</v></c>'))
        with self.assertRaises(TemplateError):   # a template refuses all three
            load(data, "x.xlsx")
        sheet = load(data, "x.xlsx", reading=True)["WPS"]
        self.assertEqual(sheet.get(29, 4), Formula('=L29&"–"&M29', saved="90–130"))
        self.assertEqual(sheet.get(29, 4).saved, "90–130")
        self.assertEqual(sheet.get(25, 3), XlError("#N/A"))


if __name__ == "__main__":
    unittest.main()
