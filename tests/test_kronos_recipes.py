"""Recipes read from a file (D89): compiler/kronos_recipes.py and kronos_web.read_recipe_file."""
import datetime as dt
import io
import json
import sys
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))

import kronos_recipes as kr  # noqa: E402
import kronos_web  # noqa: E402
from kronos_engine import prove  # noqa: E402

DEMO = ROOT / "docs" / "demo"
DEMO_RECIPES = [(DEMO / n).read_text(encoding="utf-8") for n in
                ("recipe_1_highlight.vcl", "recipe_2_tidy.vcl", "recipe_3_summary.vcl")]
EXAMPLE = (ROOT / "vcl" / "examples" / "example1_rev3.vcl").read_text(encoding="utf-8")
BUILD = (ROOT / "vcl" / "examples" / "example4_build.vcl").read_text(encoding="utf-8")
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def esc(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def docx(paragraphs, table=None):
    """A small Word document: one w:p per paragraph (a list of runs, or a string), and a table if given."""
    body = []
    for p in paragraphs:
        runs = [p] if isinstance(p, str) else p
        body.append("<w:p>" + "".join(r if r.startswith("<") else f'<w:r><w:t xml:space="preserve">{esc(r)}</w:t></w:r>'
                                      for r in runs) + "</w:p>")
    if table:
        body.append("<w:tbl><w:tr><w:tc><w:p><w:r><w:t>" + esc(table) + "</w:t></w:r></w:p></w:tc></w:tr></w:tbl>")
    xml = f'<?xml version="1.0" encoding="UTF-8"?><w:document {NS}><w:body>{"".join(body)}</w:body></w:document>'
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as z:
        z.writestr("[Content_Types].xml", "<Types/>")
        z.writestr("word/document.xml", xml)
    return out.getvalue()


class Files(unittest.TestCase):
    def texts(self, data, name):
        result = kr.read_recipes(data, name)
        self.assertTrue(result["ok"], result["message"])
        return [r["text"] for r in result["recipes"]]

    def test_the_demo_guide_holds_the_three_demo_recipes(self):
        result = kr.read_recipes((DEMO / "KRONOS_Demo_NCR_Register.docx").read_bytes(), "KRONOS_Demo_NCR_Register.docx")
        self.assertEqual(result["message"], "KRONOS found 3 recipes.")
        self.assertEqual([r["text"] for r in result["recipes"]], DEMO_RECIPES)
        self.assertEqual(result["recipes"][0]["title"],
                         "Colour every NCR that is Major and still Open light red, and return how many there are.")

    def test_a_vcl_file_is_its_recipe(self):
        self.assertEqual(self.texts(DEMO_RECIPES[1].encode(), "recipe_2_tidy.vcl"), [DEMO_RECIPES[1]])

    def test_an_ai_reply_saved_as_markdown(self):
        reply = f"Here is the spec.\n\n```vcl\n{EXAMPLE}```\n\nIt deletes cancelled orders.\n"
        self.assertEqual(self.texts(reply.encode(), "reply.md"), [EXAMPLE])
        two = f"First:\n```vcl\n{EXAMPLE}```\nSecond:\n```VCL\n{DEMO_RECIPES[0]}```\n```python\nprint(1)\n```\n"
        self.assertEqual(self.texts(two.encode(), "two.md"), [EXAMPLE, DEMO_RECIPES[0]])

    def test_prose_around_recipes_and_a_build(self):
        text = f"Notes for Monday\n\n{EXAMPLE}\nThen the NCR review:\n{BUILD}\nThat's all.\n"
        self.assertEqual(self.texts(text.encode(), "notes.txt"), [EXAMPLE, BUILD])

    def test_a_mistyped_key_stays_in_the_recipe_so_the_checker_names_it(self):
        broken = EXAMPLE.replace("GOAL ", "GOLA ", 1)
        self.assertEqual(self.texts(f"My notes\n{broken}\nThat's all.\n".encode(), "notes.txt"), [broken])

    def test_windows_files_and_typed_quotes(self):
        utf16 = ("﻿" + EXAMPLE.replace("\n", "\r\n")).encode("utf-16-le")
        self.assertEqual(self.texts(b"\xff\xfe" + utf16[2:], "recipe.txt"), [EXAMPLE])
        curly = EXAMPLE.replace('"Cancelled"', "“Cancelled”").replace("MODIFIES $ws", "MODIFIES $ws")
        self.assertEqual(self.texts(curly.encode("cp1252"), "recipe.txt"), [EXAMPLE])

    def test_a_word_document_with_typed_quotes_tabs_and_an_empty_paragraph(self):
        lines = EXAMPLE.rstrip("\n").split("\n")
        lines = [ln.replace('"', "“", 1).replace('"', "”", 1) for ln in lines]
        paragraphs = ["My recipe:"] + lines[:5] + [""] + lines[5:] + ["Thanks."]
        found = self.texts(docx(paragraphs, table="VCL 0.6"), "mine.docx")
        self.assertEqual(found, [EXAMPLE])
        self.assertTrue(prove(found[0], dt.date(2026, 11, 16)).ok)
        linked = docx([[f'<w:hyperlink><w:r><w:t>{esc(lines[0])}</w:t></w:r></w:hyperlink>'], *lines[1:]])
        self.assertEqual(self.texts(linked, "linked.docx"), [EXAMPLE])

    def test_files_it_turns_away(self):
        for data, name, words in ((b"", "empty.txt", "found no recipe"), (b"just a note", "note.md", "found no recipe"),
                                  (b"%PDF-1.4", "recipe.pdf", ".vcl, .txt, .md and Word .docx"),
                                  (b"\xd0\xcf\x11\xe0", "old.doc", "old Word .doc"),
                                  (b"not a zip", "broken.docx", "doesn't look like a Word document"),
                                  (b"x" * (kr.MAX_BYTES + 1), "big.txt", "up to 2048 KB"),
                                  (b"```vcl\n# QUESTION: Which sheet?\n```\n", "question.md", "found no recipe")):
            with self.subTest(name=name):
                result = kr.read_recipes(data, name)
                self.assertFalse(result["ok"])
                self.assertEqual(result["recipes"], [])
                self.assertIn(words, result["message"])

    def test_every_recipe_found_is_proven(self):
        result = kr.read_recipes((DEMO / "KRONOS_Demo_NCR_Register.docx").read_bytes(), "guide.docx")
        for recipe in result["recipes"]:
            with self.subTest(recipe=recipe["title"][:30]):
                self.assertTrue(prove(recipe["text"], dt.date(2026, 11, 16)).ok)


class Bridge(unittest.TestCase):
    def test_the_page_gets_canonical_json(self):
        text = kronos_web.read_recipe_file(DEMO_RECIPES[0].encode(), "recipe_1_highlight.vcl")
        self.assertEqual(text, kronos_web.canonical(json.loads(text)))
        self.assertEqual(json.loads(text)["recipes"][0]["text"], DEMO_RECIPES[0])


if __name__ == "__main__":
    unittest.main()
