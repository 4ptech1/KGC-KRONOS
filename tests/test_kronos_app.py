"""Item 5 Phase 2b (D76): the KRONOS window opens and takes a practice request through all four steps.

It needs Tk and a display. CI runs the same smoke test as its own step, under a virtual display
(`xvfb-run python app/kronos_app.py --smoke`), so this test only adds a local run where both exist.
"""
import importlib.util
import os
import subprocess
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
HAS_TK = importlib.util.find_spec("tkinter") is not None
HAS_DISPLAY = bool(os.environ.get("DISPLAY")) or os.name == "nt"


@unittest.skipUnless(HAS_TK and HAS_DISPLAY, "needs Tk and a display; CI runs it under xvfb-run")
class Window(unittest.TestCase):
    def test_smoke(self):
        done = subprocess.run([sys.executable, str(ROOT / "app" / "kronos_app.py"), "--smoke"], cwd=ROOT,
                              capture_output=True, text=True, timeout=120)
        self.assertEqual(done.returncode, 0, done.stdout + done.stderr)
        self.assertIn("smoke: OK, all four steps and the saved module", done.stdout)


if __name__ == "__main__":
    unittest.main()
