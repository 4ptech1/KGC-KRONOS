"""compiler/gen_context.py: the AI context file's appendices are copies of the repo (D53)."""
import re
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
import gen_context  # noqa: E402

CONTEXT = (ROOT / "docs" / "kgc-kronos-context.md").read_text(encoding="utf-8")


class ContextFile(unittest.TestCase):
    def test_it_is_up_to_date(self):
        self.assertTrue(CONTEXT == gen_context.build(), "docs/kgc-kronos-context.md is out of date: run "
                                                         "python compiler/gen_context.py and commit it")

    def test_it_stays_under_bosss_cap(self):
        self.assertLessEqual(len(CONTEXT.encode("utf-8")), gen_context.LIMIT)
        self.assertEqual(gen_context.LIMIT, 3_000_000)

    def test_the_hand_written_core_comes_first(self):
        core, generated = CONTEXT.split(gen_context.MARKER)
        self.assertIn("## 12. Glossary", core)
        self.assertNotIn("## Appendix", core)
        self.assertEqual(re.findall(r"^## (Appendix [A-Z])", generated, re.M),
                         [f"Appendix {letter}" for letter in "ABCDEFG"])

    def test_every_spec_and_decision_is_in_it(self):
        for path in sorted((ROOT / "specs").glob("*.vcl")):
            with self.subTest(spec=path.name):
                self.assertIn(path.read_text(encoding="utf-8").rstrip(), CONTEXT)
        ids = re.findall(r"^\| ([A-Z]\d+) \|", (ROOT / "DECISIONS.md").read_text(encoding="utf-8"), re.M)
        self.assertTrue(ids and all(f"| {i} |" in CONTEXT for i in ids))

    def test_headings_inside_copies_move_down_but_code_stays(self):
        md = "# Title\n```text\n# not a heading\n```\n## Part"
        self.assertEqual(gen_context.shifted(md), "### Title\n```text\n# not a heading\n```\n#### Part")
        self.assertEqual(gen_context.fenced("a\n```\nb"), "````text\na\n```\nb\n````")


if __name__ == "__main__":
    unittest.main()
