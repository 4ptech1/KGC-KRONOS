"""KRONOS Offline (D86, docs/kronos-offline.md): KRONOS in the browser, with nothing installed.

Always on: the launcher and the servers keep the recipe's rules (CRLF, the PowerShell fixes, port 5181, localhost
only), the downloads are pinned, the toolchain manifest covers everything the browser's Python needs, the version
list is in order, the page uses only what Chrome 109 and Firefox 115 have, and compiler/kronos_web.py answers.

Needing the downloads (Pyodide, the lark wheel, Inter): the zip is built and checked; under Pyodide in Node.js,
kronos_web.check gives exactly normal Python's text for the 12 specs and the examples; the built zip's servers serve
it; and with Playwright, with every request outside localhost blocked, the page proves the example, shows the
problems of a broken recipe and cancels. Without the network, Node.js or Playwright these skip; with
KRONOS_OFFLINE_REQUIRED=1 (CI's build-offline.yml) they fail instead.

    KRONOS_OFFLINE_BUILD=<folder>     use a copy build_offline.py already built there (with build-info.json)
    KRONOS_OFFLINE_BROWSER=firefox    the page test's browser (default chromium)
    KRONOS_OFFLINE_BROWSER_PATH=<exe> a browser already on this computer, instead of Playwright's own
    KRONOS_OFFLINE_SHOTS=<folder>     save screenshots of the page there
"""
import json
import os
import re
import shutil
import socket
import subprocess
import sys
import tempfile
import time
import unittest
import urllib.error
import urllib.request
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
sys.path.insert(0, str(ROOT / "packaging"))
import build_offline as B  # noqa: E402
import kronos_web  # noqa: E402
from kronos_engine import bas_text, module_name, prove  # noqa: E402

REQUIRED = os.environ.get("KRONOS_OFFLINE_REQUIRED") == "1"
TODAY = "2026-11-14"
PORT = 5181
URL = f"http://localhost:{PORT}/"
OFFLINE = ROOT / "packaging" / "offline"
LAUNCHER = OFFLINE / B.LAUNCHER
SERVER = OFFLINE / "server"
WEB = ROOT / "web"
EXAMPLE = (ROOT / "vcl" / "examples" / "example1_rev3.vcl").read_text(encoding="utf-8")
BROKEN = EXAMPLE.replace("GOAL ", "GOLA ", 1)   # an unknown key and a missing GOAL: two checker errors
FAILING = EXAMPLE.replace("RESULT >= PRE_COUNTIF", "RESULT < PRE_COUNTIF", 1)   # a TEST that fails
CASES = {p.name: p for p in [*sorted((ROOT / "specs").glob("*.vcl")), *sorted((ROOT / "vcl" / "examples").glob("*.vcl"))]}


def need(reason):
    """Skip here, fail in CI (KRONOS_OFFLINE_REQUIRED=1), as KRONOS_LIVE_CODEX does for the Codex tests."""
    if REQUIRED:
        raise AssertionError(f"KRONOS_OFFLINE_REQUIRED=1, but {reason}")
    raise unittest.SkipTest(reason)


def text_cases():
    out = {name: path.read_text(encoding="utf-8") for name, path in CASES.items()}
    out.update({"broken": BROKEN, "failing test": FAILING, "question": "# QUESTION: Which sheet holds the orders?",
                "no recipe": "Sure! I can help with that.",
                "whole reply": "Here is the spec.\n\n```vcl\n" + EXAMPLE + "```\n"})
    return out


# ---------- the built copy, made once for every test that needs it ----------

_BUILT = {}
_TEMP = []


def temp(prefix):
    folder = Path(tempfile.mkdtemp(prefix=prefix))
    _TEMP.append(folder)
    return folder


def tearDownModule():
    for folder in _TEMP:
        shutil.rmtree(folder, ignore_errors=True)


def built():
    """(folder, build-info): KRONOS_OFFLINE_BUILD's copy, or one built now into a temporary folder."""
    if "info" not in _BUILT:
        given = os.environ.get("KRONOS_OFFLINE_BUILD")
        if given:
            folder = Path(given).resolve()
            info = json.loads((folder / "build-info.json").read_text(encoding="utf-8"))
        else:
            folder = temp("kronos offline ")
            try:
                info = B.build(folder, B.DEFAULT_CACHE, say=lambda text: None)
            except B.Unreachable as exc:
                _BUILT["skip"] = str(exc)
                info = None
        _BUILT.update(folder=folder, info=info)
    if _BUILT["info"] is None:
        need(f"the downloads are not available ({_BUILT['skip']})")
    return _BUILT["folder"], _BUILT["info"]


def unpacked():
    """A fresh folder with the built zip unpacked in it, as Boss's Extract All makes."""
    if "fresh" not in _BUILT:
        folder, info = built()
        fresh = temp("kronos fresh ")
        with zipfile.ZipFile(folder / info["zip"]) as archive:
            archive.extractall(fresh)
        _BUILT["fresh"] = fresh / B.FOLDER
    return _BUILT["fresh"]


# ---------- always on ----------

class Launcher(unittest.TestCase):
    def test_it_keeps_windows_line_endings(self):   # rule 3
        data = LAUNCHER.read_bytes()
        self.assertGreater(data.count(b"\r\n"), 20)
        self.assertEqual(data.count(b"\n"), data.count(b"\r\n"))
        self.assertIn("packaging/offline/*.bat -text", (ROOT / ".gitattributes").read_text(encoding="utf-8"))

    def test_it_tries_node_then_py_then_python_then_powershell(self):
        text = LAUNCHER.read_text(encoding="ascii")
        order = [text.index(words) for words in ('node "server\\serve.mjs"', 'py -3 "server\\serve.py"',
                                                  'python "server\\serve.py"', "powershell -NoProfile")]
        self.assertEqual(order, sorted(order))
        self.assertIn("title KRONOS", text)
        self.assertIn('pushd "%~dp0"', text)
        self.assertTrue(text.rstrip().endswith("pause"))

    def test_it_passes_the_app_folder_to_powershell(self):   # rule 1
        self.assertIn('powershell -NoProfile -ExecutionPolicy Bypass -File "server\\serve.ps1" "%CD%\\app"',
                      LAUNCHER.read_text(encoding="ascii"))


class Servers(unittest.TestCase):
    def read(self, name):
        return (SERVER / name).read_text(encoding="ascii")   # ASCII only: PowerShell 2 reads a .ps1 as ANSI

    def test_all_three_use_port_5181_on_this_computer_only(self):
        for name in B.SERVERS:
            text = self.read(name)
            self.assertRegex(text, r"(PORT|\$port) = 5181\b", name)
            self.assertNotIn("5180", text, name)
            self.assertNotIn("0.0.0.0", text, name)
            self.assertNotIn("BlueForce", text, name)
            self.assertEqual(text.count(B.SERVER_MARK), 1, name)
        self.assertIn("server.listen(PORT, '127.0.0.1'", self.read("serve.mjs"))
        self.assertIn("TCPServer(('127.0.0.1', PORT)", self.read("serve.py"))
        self.assertIn('Prefixes.Add("http://localhost:$port/")', self.read("serve.ps1"))

    def test_the_powershell_fixes_are_there(self):   # rules 1 and 2
        text = self.read("serve.ps1")
        self.assertIn("param([string]$AppFolder)", text)
        self.assertIn("function Open-Browser([string]$url)", text)
        self.assertIn("-notlike 'IE.*'", text)
        for exe in ("chrome.exe", "msedge.exe", "firefox.exe"):
            self.assertIn(exe, text)
        self.assertNotRegex(text, r'Start-Process\s+"http')   # never a bare Start-Process on the address
        self.assertIn("Open Chrome, Edge or Firefox and type:  localhost:$port", text)
        self.assertIn("I can't find the KRONOS app here", text)

    def test_each_server_sends_wasm_as_application_wasm(self):   # Pyodide's .wasm needs it
        for name in B.SERVERS:
            self.assertRegex(self.read(name), r"'\.wasm'\s*[:=]\s*'application/wasm'", name)

    def test_the_python_server_answers_one_request_at_a_time_like_the_powershell_one(self):   # D86
        # BlueForce's serve.py, unchanged in this: so every page test through it also checks that the page loads from
        # a server that answers one request at a time, as serve.ps1 does on the work PC
        self.assertIn("server = socketserver.TCPServer(('127.0.0.1', PORT)", self.read("serve.py"))


class Pins(unittest.TestCase):
    def test_every_download_is_pinned_and_lands_in_app(self):
        pins = B.load_pins()
        targets = []
        for item in pins["downloads"]:
            self.assertTrue(item["url"].startswith("https://"), item["name"])
            self.assertRegex(item["sha256"], r"^[0-9a-f]{64}$")
            parts = [item] if "to" in item else item["extract"]
            self.assertTrue(parts, item["name"])
            for part in parts:
                self.assertRegex(part["sha256"], r"^[0-9a-f]{64}$")
                self.assertRegex(part["to"], r"^app/(pyodide|wheels|fonts)/[A-Za-z0-9._-]+$")
                targets.append(part["to"])
        self.assertEqual(len(targets), len(set(targets)))
        names = {item["name"]: item for item in pins["downloads"]}
        self.assertEqual(names["pyodide"]["version"], "0.24.1")   # the last to cover Chrome 109 (D86)
        self.assertEqual(sorted(Path(p["to"]).name for p in names["pyodide"]["extract"]),
                         ["pyodide-lock.json", "pyodide.asm.js", "pyodide.asm.wasm", "pyodide.js", "python_stdlib.zip"])

    def test_the_lark_wheel_is_the_one_in_requirements_txt(self):
        lark = next(item for item in B.load_pins()["downloads"] if item["name"] == "lark")
        requirements = (ROOT / "requirements.txt").read_text(encoding="utf-8")
        self.assertIn(f"lark=={lark['version']} ", requirements)
        self.assertIn(f"--hash=sha256:{lark['sha256']}", requirements)
        self.assertTrue(lark["url"].endswith(f"/lark-{lark['version']}-py3-none-any.whl"))

    def test_every_licence_is_committed_and_named(self):
        named = [name for item in B.load_pins()["downloads"] for name in item["licences"]]
        on_disk = sorted(p.name for p in B.LICENCES.glob("*.txt"))
        self.assertEqual(sorted(named), on_disk)
        self.assertEqual(on_disk, ["inter-OFL-1.1.txt", "lark-MIT.txt", "pyodide-MPL-2.0.txt", "python-PSF.txt"])

    def test_the_test_tools_are_pinned_with_their_hashes(self):
        for path in sorted((ROOT / "packaging" / "offline_test").glob("*.txt")):
            lines = [ln for ln in path.read_text(encoding="utf-8").splitlines() if ln.strip() and not ln.startswith("#")]
            pins = [ln for ln in lines if not ln.strip().startswith("--hash")]
            hashes = [ln for ln in lines if re.fullmatch(r"\s+--hash=sha256:[0-9a-f]{64}( \\)?", ln)]
            self.assertTrue(any(ln.startswith("playwright==") for ln in pins), path.name)
            self.assertGreaterEqual(len(hashes), len(pins), path.name)
            for pin in pins:
                self.assertRegex(pin, r"^[A-Za-z0-9._-]+==[0-9][^ ]* \\$", path.name)


class Toolchain(unittest.TestCase):
    def test_the_manifest_holds_every_module_kronos_web_imports(self):
        code = ("import sys; from pathlib import Path; sys.path.insert(0, 'compiler'); import kronos_web; "
                "root = Path('.').resolve(); print('\\n'.join(sorted(Path(m.__file__).resolve().relative_to(root)"
                ".as_posix() for m in list(sys.modules.values()) if getattr(m, '__file__', None) and "
                "root in Path(m.__file__).resolve().parents)))")
        loaded = subprocess.run([sys.executable, "-c", code], cwd=ROOT, capture_output=True, text=True, check=True)
        files = set(B.toolchain_files())
        self.assertIn("compiler/kronos_web.py", files)
        for name in loaded.stdout.split():
            self.assertIn(name, files)

    def test_the_manifest_holds_every_compiler_module_and_every_file_kronos_exe_bundles(self):
        files = B.toolchain_files()
        self.assertEqual(files, sorted(set(files)))
        for path in (ROOT / "compiler").glob("*.py"):
            self.assertIn(f"compiler/{path.name}", files)
        for name in ("vcl/semantics.py", "vcl/grammar.lark", "vcl/orchestrator_prompt.txt",
                     "vcl/examples/example1_rev3.vcl", "vba/ForgeLib.bas"):
            self.assertIn(name, files)
        self.assertEqual(set(B.spec_data_files()) - set(files), set())
        for name in files:
            self.assertTrue((ROOT / name).is_file(), name)
            self.assertFalse(name.startswith(("tests/", "data/", "specs/")), name)

    def test_the_version_list_is_in_order(self):
        rows = B.versions()
        self.assertEqual([r[0] for r in rows], list(range(1, len(rows) + 1)))
        for number, date, commit, words in rows:
            time.strptime(date, "%d-%m-%Y")
            self.assertTrue(words, f"version {number} says what changed")
            if number < rows[-1][0]:   # every copy but the newest names the main commit it was built from
                self.assertRegex(commit, r"^`?[0-9a-f]{7,40}`?$", f"version {number}")


class Workflow(unittest.TestCase):
    def test_ci_builds_checks_and_releases_only_by_hand_from_main(self):
        text = (ROOT / ".github" / "workflows" / "build-offline.yml").read_text(encoding="utf-8")
        for words in ("KRONOS_OFFLINE_REQUIRED: \"1\"", "python packaging/build_offline.py", "--require-hashes",
                      "runs-on: ubuntu-22.04", "runs-on: windows-latest", "packaging/offline_test/chromium109.txt",
                      "packaging/offline_test/firefox115.txt", "tests.test_offline.InTheBrowser", "shell: powershell",
                      '"Start KRONOS (Windows).bat"', "gh release create \"kronos-offline-v${VERSION}\"",
                      "github.ref == 'refs/heads/main'", "needs: [build, old-browsers, powershell]"):
            self.assertIn(words, text)
        self.assertEqual(text.count("contents: write"), 1)   # only the release job writes
        self.assertIn("contents: write", text.split("\n  release:\n", 1)[1])


class Page(unittest.TestCase):
    def sources(self):
        return {p.name: p.read_text(encoding="utf-8") for p in WEB.iterdir() if p.is_file()}

    def test_nothing_loads_from_the_internet(self):
        for name, text in self.sources().items():
            for url in re.findall(r"https?://[^\s\"')]+", text):
                self.assertEqual(url, "http://www.w3.org/2000/svg", f"{name} names {url}")

    def test_only_what_chrome_109_and_firefox_115_have(self):
        sources = self.sources()
        css = re.sub(r"/\*.*?\*/", "", sources["styles.css"], flags=re.S)   # the rules, not the comments
        for words in (":has(", "color-mix(", "@container", "@layer", "&"):
            self.assertNotIn(words, css)
        for name in ("app.js", "worker.js"):   # no modules: a classic worker and a classic script
            for words in ("type: 'module'", 'type: "module"', "import.meta"):
                self.assertNotIn(words, sources[name], name)
            self.assertNotRegex(sources[name], r"(?m)^\s*(import|export)\s", name)
            self.assertNotRegex(sources[name], r"(?<![\w.])import\(", name)   # pyodide.pyimport( is Pyodide's
        self.assertIn("new Worker('worker.js')", sources["app.js"])   # a classic worker
        # pyodide.asm.js first, so Pyodide never loads it while its big downloads wait unread (a one-request-at-a-time
        # server would then wait for them forever, D86)
        self.assertIn("importScripts('pyodide/pyodide.js', 'pyodide/pyodide.asm.js')", sources["worker.js"])
        self.assertEqual(sources["index.html"].count(B.PAGE_MARK), 1)

    def test_the_look_is_blueforces(self):
        css = self.sources()["styles.css"]
        for value in ("#0B2A6F", "#D6DEEF", "#FFD8A8", "#EFE6D4", "#E5D9C2", "#DBCDB2", "#F7F2E8", "#F5F5F5",
                      "#595959", "#00870E", "#B26A00", "#C8372D", "height: 96px", "position: sticky", '"Inter"'):
            self.assertIn(value, css)


class Bridge(unittest.TestCase):
    def test_a_proven_recipe(self):
        data = kronos_web.result(EXAMPLE, TODAY)
        proof = prove(EXAMPLE, __import__("datetime").date(2026, 11, 14))
        self.assertEqual((data["reply"], data["proven"], data["kind"], data["name"]),
                         ("spec", True, "TASK", "RemoveCancelledAndSortByDate"))
        self.assertEqual(data["words"], "it passed every check and 3 of 3 practice tests")
        self.assertEqual([t["passed"] for t in data["tests"]], [True, True, True])
        self.assertEqual(data["bas"], bas_text(module_name(proof.name), proof.module))
        self.assertTrue(data["bas"].startswith('Attribute VB_Name = "modKronosRemoveCancelledAndSort"\r\n'))
        self.assertEqual(data["problems"], [])

    def test_a_broken_recipe_and_a_failing_test(self):
        data = kronos_web.result(BROKEN, TODAY)
        self.assertFalse(data["proven"])
        self.assertEqual(data["words"], "the recipe broke 2 rules of the checker")
        self.assertEqual(sorted(f["code"] for f in data["findings"] if f["severity"] == "error"), ["V015", "V324"])
        self.assertIn("spec.vcl:6: error: unknown key GOLA [V015]", data["problems"])
        self.assertEqual(data["bas"], "")
        data = kronos_web.result(FAILING, TODAY)
        self.assertEqual(data["words"], "2 of 3 practice tests passed")
        self.assertTrue(any(line.startswith("FAIL test   spec.vcl  (2/3 tests)") for line in data["problems"]))

    def test_a_question_and_a_reply_with_no_recipe(self):
        data = kronos_web.result("# QUESTION: Which sheet holds the orders?")
        self.assertEqual((data["reply"], data["question"]), ("question", "Which sheet holds the orders?"))
        data = kronos_web.result("Sure! I can help with that.")
        self.assertEqual((data["reply"], data["words"]), ("none", "the reply had no recipe in it"))
        self.assertEqual(data["problems"], ["the reply has no vcl code block: reply with exactly one code block tagged "
                                            "vcl (rule 1) [K003]"])

    def test_the_whole_reply_of_an_ai_works_too(self):
        self.assertEqual(kronos_web.check("Here it is.\n\n```vcl\n" + EXAMPLE + "```\n", TODAY),
                         kronos_web.check(EXAMPLE, TODAY))

    def test_the_answer_is_one_canonical_text_and_the_steps_are_heard(self):
        steps = []
        text = kronos_web.check(EXAMPLE, TODAY, steps.append)
        self.assertEqual(steps, ["checking", "testing", "building"])
        self.assertEqual(text, kronos_web.canonical(json.loads(text)))
        self.assertEqual(text, kronos_web.check(EXAMPLE, TODAY))


# ---------- needing the downloads ----------

class Build(unittest.TestCase):
    def test_the_zip_has_its_name_and_only_what_belongs(self):
        folder, info = built()
        self.assertEqual(info["zip"], f"KRONOS-Offline-v{info['version']}-{info['date']}.zip")
        self.assertEqual(info["version"], B.latest_version())
        with zipfile.ZipFile(folder / info["zip"]) as archive:
            names = archive.namelist()
        self.assertEqual(names, sorted(names))
        self.assertTrue(all(n.startswith("KRONOS-Offline/") for n in names))
        top = sorted({n.split("/")[1] + ("/" if n.count("/") > 1 else "") for n in names if n != "KRONOS-Offline/"})
        self.assertEqual(top, ["Start KRONOS (Windows).bat", "app/", "server/"])
        servers = sorted(n for n in names if n.startswith("KRONOS-Offline/server/") and not n.endswith("/"))
        self.assertEqual(servers, [f"KRONOS-Offline/server/{s}" for s in ("serve.mjs", "serve.ps1", "serve.py")])
        apps = {n[len("KRONOS-Offline/app/"):] for n in names if n.startswith("KRONOS-Offline/app/") and not n.endswith("/")}
        expected = ({p.relative_to(WEB).as_posix() for p in WEB.iterdir() if p.is_file()}
                    | {f"toolchain/{n}" for n in B.toolchain_files()} | {"toolchain/manifest.json"}
                    | {p["to"][len("app/"):] for item in B.load_pins()["downloads"] for p in ([item] if "to" in item
                                                                                             else item["extract"])}
                    | {f"licences/{p.name}" for p in B.LICENCES.glob("*.txt")})
        self.assertEqual(apps, expected)
        digest = (folder / f"{info['zip']}.sha256.txt").read_text(encoding="ascii").split()[0]
        self.assertEqual(digest, B.sha256((folder / info["zip"]).read_bytes()))

    def test_text_files_a_person_opens_are_crlf_and_carry_the_version(self):
        fresh = unpacked()
        _, info = built()
        for path in [fresh / B.LAUNCHER, *(fresh / "server").iterdir(), *(fresh / "app" / "licences").iterdir()]:
            data = path.read_bytes()
            self.assertEqual(data.count(b"\n"), data.count(b"\r\n"), path.name)
            self.assertNotIn(b"\r\r", data, path.name)
        words = f"version {info['version']}, built {info['date']}"
        for name in B.SERVERS:
            self.assertIn(f"KRONOS Offline, {words}", (fresh / "server" / name).read_text(encoding="ascii"))
        page = (fresh / "app" / "index.html").read_text(encoding="utf-8")
        self.assertIn(f"Version {info['version']} · {info['date']}", page)
        self.assertNotIn(B.PAGE_MARK, page)
        self.assertIn("param([string]$AppFolder)", (fresh / "server" / "serve.ps1").read_text(encoding="ascii"))

    def test_the_toolchain_is_the_repos_byte_for_byte(self):
        fresh = unpacked()
        manifest = json.loads((fresh / "app" / "toolchain" / "manifest.json").read_text(encoding="utf-8"))
        self.assertEqual(manifest["files"], B.toolchain_files())
        self.assertEqual(manifest["wheels"], ["lark-1.3.1-py3-none-any.whl"])
        for name in manifest["files"]:
            self.assertEqual((fresh / "app" / "toolchain" / name).read_bytes(), (ROOT / name).read_bytes(), name)

    def test_the_same_inputs_give_the_same_bytes(self):
        _, info = built()
        try:
            again = B.zip_bytes(B.contents(info["version"], info["date"], B.DEFAULT_CACHE), info["date"])
        except B.Unreachable as exc:
            need(f"the downloads are not available ({exc})")
        self.assertEqual(B.sha256(again), info["sha256"])


class UnderPyodide(unittest.TestCase):
    def test_the_browser_engine_gives_exactly_normal_pythons_answers(self):
        node = shutil.which("node")
        if not node:
            need("Node.js is not installed")
        fresh = unpacked()
        texts = text_cases()
        cases = [{"name": name, "text": text, "today": TODAY} for name, text in texts.items()]
        run = subprocess.run([node, str(ROOT / "tests" / "offline_pyodide.js"), str(fresh / "app")],
                             input=json.dumps(cases), capture_output=True, text=True, encoding="utf-8", timeout=600)
        self.assertEqual(run.returncode, 0, run.stderr[-3000:])
        out = json.loads(run.stdout)
        self.assertEqual((out["python"], out["pyodide"]), ("3.11.3", "0.24.1"))
        self.assertEqual(sorted(out["answers"]), sorted(texts))
        proven = 0
        for name, text in texts.items():
            steps = []
            want = kronos_web.check(text, TODAY, steps.append)
            with self.subTest(case=name):
                self.assertEqual(out["answers"][name]["json"], want)   # validator messages, TESTs and the .bas
                self.assertEqual(out["answers"][name]["steps"], steps)
            proven += json.loads(want)["proven"]
        self.assertGreaterEqual(proven, 16)   # the 12 specs, three examples and the whole reply


# ---------- the servers and the page ----------

def port_free():
    with socket.socket() as s:
        return s.connect_ex(("127.0.0.1", PORT)) != 0


def fetch(path):
    try:
        with urllib.request.urlopen(URL + path, timeout=30) as response:
            return response.status, response.headers.get("Content-Type", ""), response.read()
    except urllib.error.HTTPError as exc:
        return exc.code, exc.headers.get("Content-Type", ""), b""


class Served:
    """One of the zip's servers, started from the fresh folder as the launcher starts it, stopped afterwards. If the
    test fails while it runs, what the server printed is shown."""

    def __init__(self, command):
        self.command = command

    def __enter__(self):
        if not port_free():
            raise AssertionError(f"port {PORT} is already in use: stop whatever runs there first")
        # BROWSER: Python's webbrowser runs `true` instead of opening a browser. PYTHONUNBUFFERED: Python prints its
        # start message at once, as it does in the black window (into a pipe it would wait for a full buffer).
        env = dict(os.environ, BROWSER="true", PYTHONUNBUFFERED="1")
        self.process = subprocess.Popen(self.command, cwd=unpacked(), env=env, stdout=subprocess.PIPE,
                                        stderr=subprocess.STDOUT, text=True)
        for _ in range(100):
            if not port_free():
                return self
            if self.process.poll() is not None:
                break
            time.sleep(0.1)
        self.process.kill()
        raise AssertionError(f"{self.command} did not start: {self.process.communicate()[0]}")

    def __exit__(self, kind, value, trace):
        self.process.terminate()
        try:
            self.output = self.process.communicate(timeout=10)[0]
        except subprocess.TimeoutExpired:
            self.process.kill()
            self.output = self.process.communicate()[0]
        if kind is not None:
            print(f"\n--- {self.command[-1]} said:\n{self.output}", file=sys.stderr)
        for _ in range(50):
            if port_free():
                break
            time.sleep(0.1)


def servers():
    """[(name, command)]: the zip's servers this computer can run, each started as the launcher starts it (on Windows,
    PowerShell too, given the app folder as "%CD%\\app"). KRONOS_OFFLINE_SERVERS=serve.ps1,... picks some."""
    found = []
    if shutil.which("node"):
        found.append(("serve.mjs", [shutil.which("node"), os.path.join("server", "serve.mjs")]))
    elif REQUIRED:
        raise AssertionError("KRONOS_OFFLINE_REQUIRED=1, but Node.js is not installed")
    found.append(("serve.py", [sys.executable, os.path.join("server", "serve.py")]))
    if os.name == "nt":   # as the launcher starts it, but with Start-Process a no-op, so that Open-Browser opens no
        # second browser to compete with the test's own (the workflow's launcher step runs it exactly, browser and all)
        app = str(unpacked() / "app")
        found.append(("serve.ps1", ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command",
                                    f"function Start-Process {{ }}; & '.\\server\\serve.ps1' '{app}'"]))
    wanted = os.environ.get("KRONOS_OFFLINE_SERVERS")
    if wanted:
        found = [(name, command) for name, command in found if name in wanted.split(",")]
        if not found:
            raise AssertionError(f"KRONOS_OFFLINE_SERVERS={wanted} names no server this computer can run")
    return found


class Serving(unittest.TestCase):
    def test_each_server_serves_the_app_and_nothing_outside_it(self):
        fresh = unpacked()
        _, info = built()
        for name, command in servers():
            with self.subTest(server=name):
                with Served(command) as served:
                    self.serves_the_app(fresh, info)
                self.assertIn(f"KRONOS Offline, version {info['version']}, built {info['date']}", served.output)
                self.assertIn(f"KRONOS is running at {URL}", served.output)

    def serves_the_app(self, fresh, info):
        status, kind, body = fetch("")
        self.assertEqual((status, kind.split(";")[0]), (200, "text/html"))
        self.assertIn(f"Version {info['version']}".encode(), body)
        status, kind, body = fetch("pyodide/pyodide.asm.wasm")
        self.assertEqual((status, kind), (200, "application/wasm"))
        self.assertEqual(body, (fresh / "app" / "pyodide" / "pyodide.asm.wasm").read_bytes())
        for path in ("pyodide/pyodide.js", "pyodide/python_stdlib.zip", "pyodide/pyodide-lock.json",
                     "toolchain/manifest.json", "toolchain/vcl/grammar.lark", "toolchain/vba/ForgeLib.bas",
                     "toolchain/compiler/kronos_web.py", "wheels/lark-1.3.1-py3-none-any.whl",
                     "fonts/inter-latin-400-normal.woff2", "worker.js"):
            status, kind, body = fetch(path)
            self.assertEqual(status, 200, path)
            self.assertEqual(body, (fresh / "app" / path).read_bytes(), path)
            if path.endswith(".js"):
                self.assertEqual(kind.split(";")[0], "text/javascript", path)
        self.assertEqual(fetch("nothing-here.txt")[0], 404)
        status, _, body = fetch("..%2Fserver%2Fserve.py")   # Windows itself answers 403 to this one, before serve.ps1
        self.assertIn(status, (403, 404))
        self.assertEqual(body, b"")


class InTheBrowser(unittest.TestCase):
    """The page itself, with Playwright, every request outside localhost blocked and recorded."""

    @classmethod
    def setUpClass(cls):
        try:
            from playwright.sync_api import sync_playwright
        except ImportError:
            need("Playwright is not installed (packaging/offline_test/)")
        unpacked()
        cls.browser_name = os.environ.get("KRONOS_OFFLINE_BROWSER", "chromium")
        cls.shots = os.environ.get("KRONOS_OFFLINE_SHOTS")
        cls.playwright = sync_playwright().start()
        try:
            kind = getattr(cls.playwright, cls.browser_name)
            path = os.environ.get("KRONOS_OFFLINE_BROWSER_PATH")
            cls.browser = kind.launch(**({"executable_path": path} if path else {}))
        except Exception as exc:   # noqa: BLE001 (no browser: Playwright's message says why)
            cls.playwright.stop()
            need(f"no {cls.browser_name} for Playwright: {str(exc).splitlines()[0]}")
        print(f"\n{cls.browser_name} {cls.browser.version}", file=sys.stderr)

    @classmethod
    def tearDownClass(cls):
        cls.browser.close()
        cls.playwright.stop()

    def open(self, server):
        self.context = self.browser.new_context(viewport={"width": 1280, "height": 900})
        page = self.context.new_page()
        self.blocked, self.errors = [], []

        def route(r):
            if r.request.url.startswith(URL):
                r.continue_()
            else:
                self.blocked.append(r.request.url)
                r.abort()
        self.context.route("**/*", route)   # the context's route sees the worker's requests too
        page.on("pageerror", lambda error: self.errors.append(str(error)))
        page.on("console", lambda m: self.errors.append(m.text) if m.type == "error" else None)
        page.goto(f"{URL}?today={TODAY}")
        self.server = server
        return page

    def shot(self, page, name):
        if self.shots:
            Path(self.shots).mkdir(parents=True, exist_ok=True)
            page.screenshot(path=str(Path(self.shots) / f"{self.browser_name}-{self.server}-{name}.png"), full_page=True)

    def ready(self, page):
        page.wait_for_function("document.getElementById('python').textContent.indexOf('ready') >= 0", timeout=120000)

    def check(self, page):
        page.click("#check")
        page.wait_for_selector("#banner", timeout=120000)
        return page.text_content("#banner")

    def test_the_page_proves_the_example_shows_problems_and_cancels(self):
        _, info = built()
        for name, command in servers():
            with self.subTest(server=name), Served(command):
                page = self.open(name)
                self.assertEqual(page.text_content("#version"), f"Version {info['version']} · {info['date']}")
                # Cancel while Python is still starting: the panel goes, and a fresh worker starts
                page.fill("#recipe", EXAMPLE)
                page.click("#check")
                self.assertTrue(page.is_visible("#working"))
                page.click("#cancel")
                self.assertFalse(page.is_visible("#working"))
                self.assertIn("Cancelled", page.text_content("#note"))
                self.ready(page)
                self.assertIn("3.11.3 in this browser (Pyodide 0.24.1)", page.text_content("#python"))
                self.shot(page, "start")
                # the practice example: PROVEN, with exactly normal Python's answer
                page.fill("#recipe", "")
                page.click("#load-example")
                page.wait_for_function("document.getElementById('recipe').value.length > 100")
                self.assertEqual(page.input_value("#recipe"), EXAMPLE)
                self.assertEqual(self.check(page), "PROVEN It passed every check and 3 of 3 practice tests.")
                self.assertEqual(page.evaluate("window.kronosLastResult"), kronos_web.check(EXAMPLE, TODAY))
                self.assertEqual(page.locator("table tbody tr").count(), 3)
                self.shot(page, "proven")
                # a broken recipe: NOT PROVEN and its problems
                page.fill("#recipe", BROKEN)
                self.assertEqual(self.check(page), "NOT PROVEN The recipe broke 2 rules of the checker.")
                self.assertIn("unknown key GOLA", page.text_content("#result table"))
                self.assertIn("spec.vcl:6: error: unknown key GOLA [V015]", page.text_content("#problems"))
                self.assertEqual(page.evaluate("window.kronosLastResult"), kronos_web.check(BROKEN, TODAY))
                self.shot(page, "problems")
                # an empty box asks for a recipe instead of checking
                page.fill("#recipe", "  ")
                page.click("#check")
                self.assertIn("Paste a recipe first", page.text_content("#note"))
                self.context.close()
                self.assertEqual(self.blocked, [], "the page asked for something outside this computer")
                self.assertEqual(self.errors, [])


if __name__ == "__main__":
    unittest.main()
