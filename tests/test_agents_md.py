"""AGENTS.md, the front page for OpenAI Codex (D83), and vclc's own wildcard expansion that its Windows commands
rely on."""
import contextlib
import io
import re
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
import vclc  # noqa: E402

AGENTS = ROOT / "AGENTS.md"
CODEX_LIMIT = 32 * 1024   # Codex reads at most this much of AGENTS.md by default (project_doc_max_bytes)
ZIP_ONLY = {"SNAPSHOT.txt", "KRONOS.exe", "my_spec.vcl"}   # only in the zip, or just an example name
EXTENSIONS = (".md", ".txt", ".py", ".bas", ".lark", ".vcl", ".json", ".yml")


class AgentsMd(unittest.TestCase):
    TEXT = AGENTS.read_text(encoding="utf-8")

    def test_codex_reads_all_of_it(self):
        self.assertLess(len(AGENTS.read_bytes()), CODEX_LIMIT)

    def test_it_sends_codex_to_claude_md_first(self):
        first = self.TEXT.split("## ", 1)[0]
        self.assertIn("**`CLAUDE.md`**", first)
        self.assertIn("all of its rules apply to", first)
        self.assertIn("If the two ever disagree, `CLAUDE.md` wins", first)

    def test_the_rules_that_keep_claude_the_only_committer(self):
        for words in ("Never commit, push or open a pull request", "`codex_changes/`", "`codex_changes/NOTE.md`",
                      "**practice only**", "Never touch `data/eval/`", "No real KGC data", "Never weaken",
                      "Excel 2016 only"):
            self.assertIn(words, self.TEXT)

    def test_every_file_it_names_exists(self):
        named = set()
        for token in re.findall(r"`([^`\n]+)`", self.TEXT):
            path = token.strip()
            if " " in path or any(ch in path for ch in "*<>") or path.startswith(("codex_changes", "github.com")):
                continue
            if path in ZIP_ONLY or not ("/" in path or path.endswith(EXTENSIONS)):
                continue
            named.add(path)
        self.assertGreater(len(named), 15)

        def exists(path):   # a path with a folder must be exactly there; a bare name, as in the table, anywhere
            if "/" in path:
                return (ROOT / path.rstrip("/")).exists()
            return (ROOT / path).exists() or any(ROOT.glob(f"*/{path}")) or any(ROOT.glob(f"*/*/{path}"))
        missing = sorted(p for p in named if not exists(p))
        self.assertEqual(missing, [], "AGENTS.md names files that are not in the repo")


class Wildcards(unittest.TestCase):   # Windows cmd and PowerShell pass specs\*.vcl through unexpanded (D83)
    def test_vclc_expands_a_pattern_itself(self):
        pattern = str(ROOT / "specs" / "task_0[12].vcl")
        with contextlib.redirect_stdout(io.StringIO()) as out:
            code = vclc.main(["vclc.py", "check", pattern])
        self.assertEqual(code, 0, out.getvalue())
        self.assertIn("task_01.vcl", out.getvalue())
        self.assertIn("task_02.vcl", out.getvalue())

    def test_a_pattern_that_matches_nothing_is_still_reported(self):
        with contextlib.redirect_stdout(io.StringIO()) as out:
            code = vclc.main(["vclc.py", "check", str(ROOT / "specs" / "no_such_*.vcl")])
        self.assertEqual(code, 1)
        self.assertIn("no_such_*.vcl", out.getvalue())


if __name__ == "__main__":
    unittest.main()
