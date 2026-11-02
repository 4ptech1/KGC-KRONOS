"""The Codex adapter against a real process (D80): a stand-in `codex` program, started the way the real one is.

On Linux and macOS the real Codex replaces its launcher, so the stand-in is one script. On Windows npm starts it as
codex.cmd, which starts node, which starts codex.exe; the stand-in copies that: codex.cmd starts Python, which
starts a second process, so stopping it must stop the whole tree. Nothing here needs Codex, a network or an
account; the Windows CI job runs this file too."""
import os
import subprocess
import sys
import tempfile
import textwrap
import threading
import time
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
from kronos_ai import AiError, CodexAI  # noqa: E402

WINDOWS = os.name == "nt"

STAND_IN = textwrap.dedent('''
    import os, subprocess, sys, time
    from pathlib import Path
    here = Path(__file__).resolve().parent
    mode = (here / "mode.txt").read_text().strip()
    args = sys.argv[1:]
    if args[:1] == ["--child"]:   # what codex.exe is to node on Windows: the process that does the work
        time.sleep(600)
        sys.exit(0)
    if args[:2] == ["login", "status"]:
        sys.stderr.write("Not logged in\\n" if mode == "signed_out" else "Logged in using ChatGPT\\n")
        sys.exit(1 if mode == "signed_out" else 0)
    if args[:1] != ["exec"]:
        sys.exit(2)
    with open(here / "execs.txt", "a") as f:
        f.write("exec\\n")
    message = sys.stdin.buffer.read()
    if mode == "reply":
        out = args[args.index("--output-last-message") + 1]
        Path(out).write_text("got %d bytes, ending %r" % (len(message), message[-12:].decode()), encoding="utf-8")
        sys.exit(0)
    if os.name == "nt":   # a child that keeps running unless the whole tree is stopped
        child = subprocess.Popen([sys.executable, __file__, "--child"])
        (here / "child.txt").write_text(str(child.pid))
    if mode == "offline":   # what codex-cli 0.166.0 prints when it can't reach OpenAI, then it waits for ever
        for n in range(2, 6):
            sys.stderr.write("ERROR: Reconnecting... %d/5\\n" % n)
        sys.stderr.write("warning: Falling back from WebSockets to HTTPS transport. stream disconnected before "
                         "completion: URL error: Proxy connection failed: HTTP CONNECT failed with status 403\\n")
        sys.stderr.write("ERROR: Reconnecting... waiting for network\\n")
        sys.stderr.flush()
    time.sleep(600)
''')


def running(pid):
    """Is a Windows process with this pid still running? (os.kill would stop it on Windows.)"""
    found = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/NH"], capture_output=True, text=True, check=False)
    return str(pid) in found.stdout


class RealProcess(unittest.TestCase):
    def setUp(self):
        self.dir = Path(tempfile.mkdtemp(prefix="kronos codex "))   # a space, as in "Documents\KGC KRONOS"
        (self.dir / "stand_in.py").write_text(STAND_IN, encoding="utf-8")
        if WINDOWS:
            self.program = self.dir / "codex.cmd"
            self.program.write_text(f'@"{sys.executable}" "%~dp0stand_in.py" %*\r\n', encoding="ascii")
        else:
            self.program = self.dir / "codex"
            self.program.write_text(f"#!{sys.executable}\n" + STAND_IN, encoding="utf-8")
            self.program.chmod(0o755)
            (self.dir / "stand_in.py").unlink()
        self.work = self.dir / "work folder"
        self.started = []

    def tearDown(self):
        child = self.dir / "child.txt"
        if WINDOWS and child.exists():   # never leave a stand-in running, even when a test fails
            subprocess.run(["taskkill", "/F", "/T", "/PID", child.read_text().strip()], capture_output=True,
                           check=False)

    def codex(self, mode, timeout=120):
        (self.dir / "mode.txt").write_text(mode, encoding="ascii")
        return CodexAI(self.work, "THE PROMPT " * 20000, timeout=timeout, which=lambda name: str(self.program))

    def execs(self):
        path = self.dir / "execs.txt"
        return len(path.read_text().splitlines()) if path.exists() else 0

    def assert_child_stopped(self):
        if WINDOWS:
            pid = int((self.dir / "child.txt").read_text())
            for _ in range(50):
                if not running(pid):
                    break
                time.sleep(0.2)
            self.assertFalse(running(pid), "the stand-in's child process is still running")

    def test_a_reply(self):   # about 200 KB on stdin: more than any pipe holds, so it must be fed as Codex reads
        ai = self.codex("reply")
        reply = ai.ask([{"role": "user", "text": "Sort the Orders sheet by date."}])
        self.assertRegex(reply, r"^got 2\d{5} bytes, ending '.*by date\.\\n'$")
        self.assertEqual(self.execs(), 1)

    def test_signed_out(self):
        ai = self.codex("signed_out")
        self.assertEqual(ai.signed_in(), (False, "Not logged in"))
        with self.assertRaisesRegex(AiError, "not signed in"):
            ai.ping()
        self.assertEqual(self.execs(), 0)   # `codex exec` never started

    def test_offline_stops_at_once(self):
        ai = self.codex("offline")
        began = time.monotonic()
        with self.assertRaisesRegex(AiError, r"^Codex can't reach OpenAI's servers: .*status 403 \| ERROR: "
                                    r"Reconnecting\.\.\. waiting for network$"):
            ai.ping()
        self.assertLess(time.monotonic() - began, 30)
        self.assert_child_stopped()

    def test_the_time_limit(self):
        ai = self.codex("hang", timeout=3)
        began = time.monotonic()
        with self.assertRaisesRegex(AiError, "^Codex gave no reply within 3 seconds$"):
            ai.ping()
        self.assertLess(time.monotonic() - began, 30)
        self.assert_child_stopped()

    def test_cancel(self):   # the window's Cancel, from another thread
        ai = self.codex("hang")
        threading.Timer(2, ai.cancel).start()
        began = time.monotonic()
        with self.assertRaisesRegex(AiError, "^cancelled$"):
            ai.ping()
        self.assertLess(time.monotonic() - began, 30)
        self.assert_child_stopped()


if __name__ == "__main__":
    unittest.main()
