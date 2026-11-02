"""Item 5 Phase 2b (D76): the KRONOS window opens and takes a practice request through all four steps.

It needs Tk and a display. CI runs the same smoke test as its own step, under a virtual display
(`xvfb-run python app/kronos_app.py --smoke`), so this test only adds a local run where both exist.
"""
import importlib.util
import os
import subprocess
import sys
import tempfile
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
        self.assertIn("smoke: OK, all four steps, the saved module, a practice run, every dialog and a corpus candidate", done.stdout)

    def test_the_version_and_the_log_file(self):   # a windowed exe has no console (D79)
        log = Path(tempfile.mkdtemp()) / "log.txt"
        done = subprocess.run([sys.executable, str(ROOT / "app" / "kronos_app.py"), "--version", "--log", str(log)],
                              cwd=ROOT, capture_output=True, text=True, timeout=60)
        self.assertEqual(done.returncode, 0, done.stderr)
        self.assertEqual(log.read_text(encoding="utf-8").strip(), done.stdout.strip())
        self.assertRegex(done.stdout, r"^KGC KRONOS \d+\.\d+$")

    def test_a_crash_is_kept_in_the_data_folder_and_reported(self):
        data = Path(tempfile.mkdtemp()) / "KGC KRONOS"
        code = ("import sys; sys.path.insert(0, 'app'); import kronos_app as K; "
                "K.messagebox.showerror = lambda title, text, **kw: print('BOX:', text.splitlines()[0]); "
                f"args = K.parse_args(['--data', {str(data)!r}])\n"
                "try:\n    raise ValueError('boom')\nexcept ValueError:\n    print('exit', K.crash(args))")
        done = subprocess.run([sys.executable, "-c", code], cwd=ROOT, capture_output=True, text=True, timeout=60)
        self.assertEqual(done.returncode, 0, done.stderr)
        self.assertIn("BOX: KRONOS stopped on an error: ValueError: boom", done.stdout)
        self.assertIn("exit 1", done.stdout)
        self.assertIn("ValueError: boom", (data / "crash.log").read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
