"""KRONOS against the real Codex CLI, installed and signed out (D80). Skipped unless KRONOS_LIVE_CODEX=1.

The Windows CI job in .github/workflows/build-exe.yml installs the pinned Codex CLI with npm (so it is codex.cmd,
as on Boss's PC) and runs this file. Signed out, KRONOS must stop before `codex exec` ever starts, so nothing is
sent to OpenAI and no account is needed."""
import os
import shutil
import sys
import tempfile
import time
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
sys.path.insert(0, str(ROOT / "app"))
from kronos_ai import AiError, CodexAI  # noqa: E402
from kronos_controller import plain_error  # noqa: E402

LIVE = os.environ.get("KRONOS_LIVE_CODEX") == "1"


@unittest.skipUnless(LIVE, "set KRONOS_LIVE_CODEX=1 with the Codex CLI installed and signed out")
class LiveCodex(unittest.TestCase):
    def codex(self):
        work = Path(tempfile.mkdtemp(prefix="kronos live ")) / "codex"   # a space, as in "Documents\KGC KRONOS"
        return CodexAI(work, (ROOT / "vcl" / "orchestrator_prompt.txt").read_text(encoding="utf-8"))

    def test_it_is_found(self):
        program = self.codex().program()
        print(f"\nCodex found at {program}")
        self.assertEqual(Path(program).name.lower(), "codex.cmd" if os.name == "nt" else "codex")
        self.assertEqual(Path(program), Path(shutil.which("codex")))

    def test_every_flag_is_known(self):
        self.assertEqual(self.codex().check(), [])

    def test_signed_out_is_seen_at_once(self):
        ai = self.codex()
        self.assertEqual(ai.signed_in(), (False, "Not logged in"))
        began = time.monotonic()
        with self.assertRaises(AiError) as caught:
            ai.ask([{"role": "user", "text": "In the Orders sheet, sort the rows by Order Date, newest first."}])
        self.assertLess(time.monotonic() - began, 60)
        self.assertEqual(str(caught.exception), "Codex is not signed in (`codex login status` says: Not logged in)")
        words, details = plain_error(str(caught.exception))
        self.assertTrue(words.startswith("Codex isn't signed in. Open a Command Prompt, type codex login"))
        self.assertFalse((ai.workdir / "last_message.txt").exists())   # `codex exec` never ran


if __name__ == "__main__":
    unittest.main()
