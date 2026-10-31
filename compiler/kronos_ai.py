"""The AI adapters of the KGC KRONOS engine (item 5 Phase 1, D68, D72).

Each adapter answers one question: ask(conversation) -> the AI's reply text. A conversation is a
list of {"role": "user" | "assistant", "text": ...}, oldest first, and every call sends all of
it, so nothing depends on the AI remembering an earlier call.

  FakeAI   the tests' AI: replies from a script, and keeps every conversation it was sent
  CodexAI  the Codex CLI (`codex exec`), signed in with the person's ChatGPT account

Codex gets the Orchestrator prompt and the conversation as one message on stdin: the prompt is
about 48 KB, too long for a Windows command line and, as far as we know, longer than the 32 KiB
of AGENTS.md that Codex reads (D72). It runs in an empty working folder of KRONOS's data folder
with the read-only sandbox; its last message comes back through --output-last-message. The
command line is a setting, because Codex's flags may change.

Both adapters can be cancelled from another thread (the window's Cancel, D74): cancel() stops the
call in progress, which then raises AiError("cancelled"), and reset() makes the adapter ready again.
"""
import os
import shutil
import subprocess
import threading
from pathlib import Path


class AiError(Exception):
    """The AI gave no reply: not installed, not signed in, too slow, or it failed."""


class FakeAI:
    """Replies from a script, in order. Every conversation it was sent is kept in `seen`. A reply may
    be text, an exception to raise, or a function of the conversation (for the cancel tests)."""

    name = "fake"

    def __init__(self, replies):
        self.replies = list(replies)
        self.seen = []
        self.cancel_event = threading.Event()

    def ask(self, conversation):
        if self.cancel_event.is_set():
            raise AiError("cancelled")
        self.seen.append([dict(m) for m in conversation])
        if len(self.seen) > len(self.replies):
            raise AiError(f"the fake AI has only {len(self.replies)} replies")
        reply = self.replies[len(self.seen) - 1]
        if isinstance(reply, Exception):
            raise reply
        return reply(conversation) if callable(reply) else reply

    def cancel(self):
        self.cancel_event.set()

    def reset(self):
        self.cancel_event.clear()


AGENTS_MD = """# KGC KRONOS working folder

This folder is empty on purpose. Run no commands and create or change no files: KRONOS needs
only the text of your last message. Everything you need is in the message itself.
"""

# {out} is the file Codex writes its last message to; "-" reads the message from stdin.
CODEX_COMMAND = ["codex", "exec", "--sandbox", "read-only", "--skip-git-repo-check",
                 "--output-last-message", "{out}", "-"]
CODEX_TIMEOUT = 300   # seconds: a spec of about 50 lines can take Codex a few minutes
CREATE_NO_WINDOW = 0x08000000   # Windows: the Codex process opens no console window


def render(prompt, conversation):
    """The one message Codex reads: the Orchestrator prompt, then the conversation so far."""
    out = [prompt.rstrip("\n"), "", "CONVERSATION",
           "The user's messages and your earlier replies follow, oldest first. Reply to the last USER",
           "message, following every rule above.", ""]
    for message in conversation:
        out += ["===== USER =====" if message["role"] == "user" else "===== YOUR EARLIER REPLY =====",
                message["text"].rstrip("\n"), ""]
    return "\n".join(out)


class CodexAI:
    """The Codex CLI. `popen` stands in for subprocess.Popen in the tests."""

    name = "codex"

    def __init__(self, workdir, prompt, command=None, timeout=CODEX_TIMEOUT, popen=subprocess.Popen,
                 which=shutil.which, windows=None):
        self.workdir = Path(workdir)
        self.prompt = prompt
        self.command = list(command or CODEX_COMMAND)
        self.timeout = timeout
        self.popen = popen
        self.which = which
        self.windows = os.name == "nt" if windows is None else windows
        self.cancel_event = threading.Event()
        self._proc = None

    def program(self):
        """The full path of the Codex program (on Windows this finds codex.cmd as well as codex.exe)."""
        found = self.which(self.command[0])
        if not found:
            raise AiError(f"the Codex CLI ({self.command[0]}) was not found: install it, run `codex login`, "
                          "and try again (docs/kronos-app.md)")
        return found

    def _run(self, args, stdin=b""):
        extra = {"creationflags": CREATE_NO_WINDOW} if self.windows else {}
        if self.cancel_event.is_set():
            raise AiError("cancelled")
        try:
            proc = self.popen(args, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                              cwd=str(self.workdir), **extra)
        except OSError as exc:
            raise AiError(f"Codex could not start ({type(exc).__name__}: {exc})") from None
        self._proc = proc
        try:
            if self.cancel_event.is_set():   # cancelled while it was starting
                proc.kill()
            out, err = proc.communicate(stdin, timeout=self.timeout)
        except subprocess.TimeoutExpired:
            proc.kill()
            proc.communicate()
            raise AiError(f"Codex gave no reply within {self.timeout} seconds") from None
        finally:
            self._proc = None
        if self.cancel_event.is_set():
            raise AiError("cancelled")
        return subprocess.CompletedProcess(args, proc.returncode, out, err)

    def cancel(self):
        """Stop the call in progress (from another thread): it raises AiError("cancelled")."""
        self.cancel_event.set()
        proc = self._proc
        if proc is not None:
            proc.kill()

    def reset(self):
        self.cancel_event.clear()

    def ask(self, conversation):
        self.workdir.mkdir(parents=True, exist_ok=True)
        (self.workdir / "AGENTS.md").write_text(AGENTS_MD, encoding="utf-8", newline="\n")
        out = self.workdir / "last_message.txt"
        if out.exists():
            out.unlink()   # KRONOS's own file in its own folder: never an old reply
        if "{out}" not in self.command:
            raise AiError("the Codex command line has no {out}: KRONOS can't read the reply without it")
        args = [self.program()] + [str(out) if a == "{out}" else a for a in self.command[1:]]
        done = self._run(args, render(self.prompt, conversation).encode("utf-8"))
        if done.returncode != 0:
            tail = (done.stderr or b"").decode("utf-8", "replace").strip().splitlines()[-5:]
            raise AiError(f"Codex stopped with exit code {done.returncode}: " + (" | ".join(tail) or "no message")
                          + " (is it signed in? run `codex login`)")
        reply = out.read_text(encoding="utf-8").strip() if out.exists() else ""
        if not reply:
            raise AiError("Codex finished without a reply")
        return reply

    def check(self):
        """Problems with the command line: every flag in it must appear in `codex exec --help`."""
        try:
            program = self.program()
        except AiError as exc:
            return [str(exc)]
        self.workdir.mkdir(parents=True, exist_ok=True)
        try:
            done = self._run([program, "exec", "--help"])
        except AiError as exc:
            return [str(exc)]
        text = (done.stdout or b"").decode("utf-8", "replace") + (done.stderr or b"").decode("utf-8", "replace")
        if done.returncode != 0:
            return [f"`codex exec --help` stopped with exit code {done.returncode}"]
        return [f"`codex exec --help` doesn't list {flag}: update the Codex command line in settings.json"
                for flag in self.command[1:] if flag.startswith("--") and flag not in text]
