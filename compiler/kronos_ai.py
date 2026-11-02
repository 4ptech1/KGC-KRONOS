"""The AI adapters of the KGC KRONOS engine (item 5 Phase 1, D68, D72).

Each adapter answers one question: ask(conversation) -> the AI's reply text. A conversation is a
list of {"role": "user" | "assistant", "text": ...}, oldest first, and every call sends all of
it, so nothing depends on the AI remembering an earlier call.

  FakeAI   the tests' AI: replies from a script, and keeps every conversation it was sent
  CodexAI  the Codex CLI (`codex exec`), signed in with the person's ChatGPT account

Codex gets the Orchestrator prompt and the conversation as one message on stdin: the prompt is
about 48 KB, too long for a Windows command line and, as far as we know, longer than the 32 KiB
of AGENTS.md that Codex reads (D72). It runs in an empty working folder of KRONOS's data folder
with the read-only sandbox; its last message comes back through --output-last-message, and
--ephemeral keeps Codex from storing its own copy of the conversation (KRONOS keeps the run). The
command line is a setting, because Codex's flags may change.

Checked against the real Codex CLI 0.166.0 (D80): every flag is in `codex exec --help`, and `-`
reads the message from stdin. Two things it does that KRONOS must handle:
  - Signed out, `codex exec` does not stop: it tries to reach OpenAI anyway. `codex login status`
    answers at once and offline ("Not logged in", exit code 1), so KRONOS asks it first.
  - When it can't reach OpenAI (no network, or a firewall or proxy blocks it), it retries five
    times, prints "Reconnecting... waiting for network" and then waits for ever. KRONOS reads
    Codex's messages as they come and stops it at that line, instead of waiting out the time limit.
On Windows the npm install starts Codex as codex.cmd, which starts node, which starts codex.exe:
stopping Codex means stopping that whole tree (taskkill /T), or codex.exe keeps running.

Both adapters can be cancelled from another thread (the window's Cancel, D74): cancel() stops the
call in progress, which then raises AiError("cancelled"), and reset() makes the adapter ready again.
"""
import os
import shutil
import subprocess
import threading
import time
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

    def ping(self):
        return "OK"

    def cancel(self):
        self.cancel_event.set()

    def reset(self):
        self.cancel_event.clear()


AGENTS_MD = """# KGC KRONOS working folder

This folder is empty on purpose. Run no commands and create or change no files: KRONOS needs
only the text of your last message. Everything you need is in the message itself.
"""

# {out} is the file Codex writes its last message to; "-" reads the message from stdin.
CODEX_COMMAND = ["codex", "exec", "--ephemeral", "--sandbox", "read-only", "--skip-git-repo-check",
                 "--output-last-message", "{out}", "-"]
PING_MESSAGE = "Reply with the single word OK. Do nothing else."
CHECK_TIMEOUT = 30   # seconds: `codex exec --help` answers at once, so a slow one is a problem to report
CODEX_TIMEOUT = 300   # seconds: a spec of about 50 lines can take Codex a few minutes
CREATE_NO_WINDOW = 0x08000000   # Windows: the Codex process opens no console window
NOT_SIGNED_IN = "not logged in"   # what `codex login status` says when nobody is signed in (0.166.0)
OFFLINE = b"waiting for network"   # Codex's line once its retries to reach OpenAI have failed (0.166.0)
POLL = 0.1        # seconds between looks at a running Codex
STOP_WAIT = 10    # seconds to wait for a stopped Codex, and for its pipes to close


def taskkill(pid):
    """Windows: stop a process and every process it started (codex.cmd, node and codex.exe)."""
    program = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32" / "taskkill.exe"
    subprocess.run([str(program), "/F", "/T", "/PID", str(pid)], stdin=subprocess.DEVNULL,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=CREATE_NO_WINDOW,
                   timeout=STOP_WAIT, check=False)


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
    """The Codex CLI. In the tests, `popen` stands in for subprocess.Popen, `clock` for time.monotonic and
    `kill_tree` for taskkill."""

    name = "codex"

    def __init__(self, workdir, prompt, command=None, timeout=CODEX_TIMEOUT, popen=subprocess.Popen,
                 which=shutil.which, windows=None, clock=time.monotonic, kill_tree=taskkill):
        self.workdir = Path(workdir)
        self.prompt = prompt
        self.command = list(command or CODEX_COMMAND)
        self.timeout = timeout
        self.popen = popen
        self.which = which
        self.windows = os.name == "nt" if windows is None else windows
        self.clock = clock
        self.kill_tree = kill_tree
        self.cancel_event = threading.Event()
        self.signed_in_seen = False   # `codex login status` said signed in once: not asked again

    def program(self):
        """The full path of the Codex program (on Windows this finds codex.cmd as well as codex.exe)."""
        found = self.which(self.command[0])
        if not found:
            raise AiError(f"the Codex CLI ({self.command[0]}) was not found: install it, run `codex login`, "
                          "and try again (docs/kronos-app.md)")
        return found

    def _run(self, args, stdin=b"", timeout=None, watch=False):
        """Run Codex to the end, reading what it prints as it comes. It is stopped (the whole process tree on
        Windows) when the person cancels, when the time limit passes, and, with watch, when Codex says it is
        waiting for the network."""
        limit = timeout or self.timeout
        extra = {"creationflags": CREATE_NO_WINDOW} if self.windows else {}
        if self.cancel_event.is_set():
            raise AiError("cancelled")
        try:
            proc = self.popen(args, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                              cwd=str(self.workdir), **extra)
        except OSError as exc:
            raise AiError(f"Codex could not start ({type(exc).__name__}: {exc})") from None
        out, err, offline = [], [], threading.Event()

        def feed():
            try:
                proc.stdin.write(stdin)
            except (OSError, ValueError):
                pass   # Codex stopped before it read everything: its exit code says why
            finally:
                try:
                    proc.stdin.close()
                except (OSError, ValueError):
                    pass

        def read(pipe, lines, watching):
            try:
                for line in iter(pipe.readline, b""):
                    lines.append(line)
                    if watching and OFFLINE in line.lower():
                        offline.set()
            except (OSError, ValueError):
                pass   # the pipe closed because Codex was stopped

        threads = [threading.Thread(target=feed, daemon=True),
                   threading.Thread(target=read, args=(proc.stdout, out, False), daemon=True),
                   threading.Thread(target=read, args=(proc.stderr, err, watch), daemon=True)]
        for thread in threads:
            thread.start()
        deadline, why = self.clock() + limit, None
        while proc.poll() is None:
            why = ("cancelled" if self.cancel_event.is_set() else "offline" if offline.is_set()
                   else "slow" if self.clock() >= deadline else None)
            if why:
                self._stop(proc)
                break
            self.cancel_event.wait(POLL)   # Cancel wakes it at once
        for thread in threads:
            thread.join(STOP_WAIT)
        for pipe, reader in ((proc.stdout, threads[1]), (proc.stderr, threads[2])):
            if not reader.is_alive():   # closing a pipe a reader still waits on could wait too
                try:
                    pipe.close()
                except (OSError, ValueError):
                    pass
        if why == "cancelled" or self.cancel_event.is_set():
            raise AiError("cancelled")
        if why == "slow":
            raise AiError(f"Codex gave no reply within {limit} seconds")
        if why == "offline":
            said = [line.strip() for line in b"".join(err).decode("utf-8", "replace").splitlines() if line.strip()]
            raise AiError("Codex can't reach OpenAI's servers: " + " | ".join(said[-3:]))
        return subprocess.CompletedProcess(args, proc.returncode, b"".join(out), b"".join(err))

    def _stop(self, proc):
        """Stop Codex. On Windows codex.cmd starts node, which starts codex.exe, so the whole tree goes."""
        if self.windows:
            try:
                self.kill_tree(proc.pid)
            except (OSError, subprocess.SubprocessError):
                pass   # proc.kill() below still stops the first process
        try:
            proc.kill()
        except OSError:
            pass
        try:
            proc.wait(STOP_WAIT)
        except subprocess.TimeoutExpired:
            pass

    def cancel(self):
        """Stop the call in progress (from another thread): it raises AiError("cancelled")."""
        self.cancel_event.set()

    def reset(self):
        self.cancel_event.clear()

    def ask(self, conversation):
        return self._exec(render(self.prompt, conversation))

    def ping(self):
        """One cheap call, with the same command line and no Orchestrator prompt: is Codex installed and
        signed in? Returns its reply (the window's Test Codex, D77)."""
        return self._exec(PING_MESSAGE)

    def signed_in(self):
        """`codex login status`, which answers at once and without the network: (True, said) when someone is
        signed in, (False, said) when not, (None, said) when this Codex can't say. AiError when it won't run."""
        self.workdir.mkdir(parents=True, exist_ok=True)
        done = self._run([self.program(), "login", "status"], timeout=CHECK_TIMEOUT)
        text = (done.stdout or b"").decode("utf-8", "replace") + (done.stderr or b"").decode("utf-8", "replace")
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        if done.returncode == 0:
            return True, " ".join(lines[-2:])
        if any(NOT_SIGNED_IN in line.lower() for line in lines):
            return False, next(line for line in lines if NOT_SIGNED_IN in line.lower())
        return None, " | ".join(lines[-3:]) or f"exit code {done.returncode}"

    def _exec(self, message):
        if not self.signed_in_seen:   # signed out, `codex exec` would try anyway, so ask first (D80)
            state, said = self.signed_in()
            if state is False:
                raise AiError(f"Codex is not signed in (`codex login status` says: {said})")
            self.signed_in_seen = state is True
        self.workdir.mkdir(parents=True, exist_ok=True)
        (self.workdir / "AGENTS.md").write_text(AGENTS_MD, encoding="utf-8", newline="\n")
        out = self.workdir / "last_message.txt"
        if out.exists():
            out.unlink()   # KRONOS's own file in its own folder: never an old reply
        if "{out}" not in self.command:
            raise AiError("the Codex command line has no {out}: KRONOS can't read the reply without it")
        args = [self.program()] + [str(out) if a == "{out}" else a for a in self.command[1:]]
        done = self._run(args, message.encode("utf-8"), watch=True)
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
            done = self._run([program, "exec", "--help"], timeout=CHECK_TIMEOUT)
        except AiError as exc:
            return [str(exc)]
        text = (done.stdout or b"").decode("utf-8", "replace") + (done.stderr or b"").decode("utf-8", "replace")
        if done.returncode != 0:
            return [f"`codex exec --help` stopped with exit code {done.returncode}"]
        return [f"`codex exec --help` doesn't list {flag}: update the Codex command line in settings.json"
                for flag in self.command[1:] if flag.startswith("--") and flag not in text]
