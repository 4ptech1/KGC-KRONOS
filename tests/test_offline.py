"""KRONOS Offline (D86, docs/kronos-offline.md): KRONOS in the browser, with nothing installed.

Always on: the launcher and the servers keep the recipe's rules (CRLF, the PowerShell fixes, port 5181, localhost
only), the downloads are pinned, the toolchain manifest covers everything the browser's Python needs, the version
list is in order, the page uses only what Chrome 109 and Firefox 115 have, and compiler/kronos_web.py answers.

Needing the downloads (Pyodide, the lark wheel, Inter): the zip is built and checked; under Pyodide in Node.js,
kronos_web.check gives exactly normal Python's text for the 12 specs and the examples; the built zip's servers serve
it, also while a connection that asks for nothing is open; and with Playwright, with every request outside localhost
blocked, the page proves the example, shows the problems of a broken recipe and cancels, through each server and
through a stand-in that answers one request at a time as the PowerShell server does. Without the network, Node.js or
Playwright these skip; with KRONOS_OFFLINE_REQUIRED=1 (CI's build-offline.yml) they fail instead.

    KRONOS_OFFLINE_BUILD=<folder>     use a copy build_offline.py already built there (with build-info.json)
    KRONOS_OFFLINE_BROWSER=firefox    the page test's browser (default chromium)
    KRONOS_OFFLINE_BROWSER_PATH=<exe> a browser already on this computer, instead of Playwright's own
    KRONOS_OFFLINE_SHOTS=<folder>     save screenshots of the page there
"""
import http.server
import json
import os
import queue
import re
import shutil
import socket
import subprocess
import sys
import tempfile
import threading
import time
import unittest
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
sys.path.insert(0, str(ROOT / "packaging"))
import build_offline as B  # noqa: E402
import kronos_web  # noqa: E402
try:
    import openpyxl  # noqa: E402
except ImportError:   # CI installs requirements.txt
    openpyxl = None
from kronos_engine import bas_text, module_name, prove  # noqa: E402

REQUIRED = os.environ.get("KRONOS_OFFLINE_REQUIRED") == "1"
TODAY = "2026-11-14"
PORT = 5181
URL = f"http://localhost:{PORT}/"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
OFFLINE = ROOT / "packaging" / "offline"
LAUNCHER = OFFLINE / B.LAUNCHER
SERVER = OFFLINE / "server"
WEB = ROOT / "web"
EXAMPLE = (ROOT / "vcl" / "examples" / "example1_rev3.vcl").read_text(encoding="utf-8")
BROKEN = EXAMPLE.replace("GOAL ", "GOLA ", 1)   # an unknown key and a missing GOAL: two checker errors
FAILING = EXAMPLE.replace("RESULT >= PRE_COUNTIF", "RESULT < PRE_COUNTIF", 1)   # a TEST that fails
REQUEST = "Delete every cancelled order from the Orders sheet, then sort the rest by order date, newest first."
QUESTION_REPLY = "```vcl\n# QUESTION: Which sheet holds the orders?\n```"


def reply(spec):
    return f"Here is the spec.\n\n```vcl\n{spec}```\n"
CASES = {p.name: p for p in [*sorted((ROOT / "specs").glob("*.vcl")), *sorted((ROOT / "vcl" / "examples").glob("*.vcl"))]}


def need(reason):
    """Skip here, fail in CI (KRONOS_OFFLINE_REQUIRED=1), as KRONOS_LIVE_CODEX does for the Codex tests."""
    if REQUIRED:
        raise AssertionError(f"KRONOS_OFFLINE_REQUIRED=1, but {reason}")
    raise unittest.SkipTest(reason)


def orders_workbook(rows=24, filtered=False):
    """A synthetic orders workbook with a few styles, as bytes: what the upload tests give KRONOS."""
    import datetime
    import io
    from openpyxl.styles import Font
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Orders"
    ws.append(["Order No", "Order Date", "Customer", "Status"])
    statuses = ["Open", "Cancelled", "Closed", " cancelled "]
    for i in range(1, rows + 1):
        ws.append([i, datetime.date(2026, 1, 1) + datetime.timedelta(days=i * 3), f"Cust {i}", statuses[i % 4]])
    ws["C5"].font = Font(bold=True)
    ws.column_dimensions["C"].width = 33
    wb.create_sheet("Notes")["A1"] = "keep me"
    if filtered:   # Excel would adjust the filter when rows go; KRONOS says so (merged cells refuse a row op, D113, D119)
        ws.auto_filter.ref = "A1:D25"
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def zip_parts(data):
    """{part: bytes} of a zip, unpacked: two zlibs can pack the same file into different bytes."""
    import io
    import zipfile
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        return {n: z.read(n) for n in z.namelist()}


def workbook_state(data):
    """What a workbook holds that matters to the tests: values by cell, bold cells, column widths, per sheet."""
    import io
    wb = openpyxl.load_workbook(io.BytesIO(data))
    return {ws.title: ([(c.coordinate, str(c.value), bool(c.font.bold)) for row in ws.iter_rows() for c in row if c.value is not None],
                       sorted((k, round(d.width, 2)) for k, d in ws.column_dimensions.items() if d.width))
            for ws in wb.worksheets}


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

    def test_the_python_server_gives_every_connection_its_own_thread(self):   # D86: a spare connection held it up
        text = self.read("serve.py")
        self.assertIn("server = socketserver.ThreadingTCPServer(('127.0.0.1', PORT)", text)
        self.assertIn("socketserver.ThreadingTCPServer.daemon_threads = True", text)


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

    def test_the_workbook_wheels_are_the_ones_in_requirements_txt(self):   # D87
        requirements = (ROOT / "requirements.txt").read_text(encoding="utf-8")
        names = {item["name"]: item for item in B.load_pins()["downloads"]}
        for name, wheel in (("openpyxl", "openpyxl-{v}-py2.py3-none-any.whl"), ("et_xmlfile", "et_xmlfile-{v}-py3-none-any.whl")):
            item = names[name]
            self.assertIn(f"{name.replace('_', '-')}=={item['version']} ", requirements)
            self.assertIn(f"--hash=sha256:{item['sha256']}", requirements)
            self.assertTrue(item["url"].endswith("/" + wheel.format(v=item["version"])))
            self.assertEqual(item["to"], "app/wheels/" + wheel.format(v=item["version"]))

    def test_every_licence_is_committed_and_named(self):
        named = [name for item in B.load_pins()["downloads"] for name in item["licences"]]
        on_disk = sorted(p.name for p in B.LICENCES.glob("*.txt"))
        self.assertEqual(sorted(named), on_disk)
        self.assertEqual(on_disk, ["et_xmlfile-MIT.txt", "et_xmlfile-PSF.txt", "inter-OFL-1.1.txt", "lark-MIT.txt",
                                   "openpyxl-MIT.txt", "pyodide-MPL-2.0.txt", "python-PSF.txt"])

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
            if name.startswith("data/templates/"):   # the look-alike templates are the one part of data/ (D118)
                self.assertRegex(name, r"^data/templates/[a-z0-9_]+\.(xlsx|docx)$")
                continue
            self.assertFalse(name.startswith(("tests/", "data/", "specs/")), name)
        self.assertTrue(any(n.endswith(".docx") for n in files) and any(n.endswith(".xlsx") for n in files))

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
                      "--latest=false",   # the Latest badge stays on KRONOS.exe's release
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
        # the "this browser can't run KRONOS" line is hidden at once where KRONOS can run, and stays in Internet Explorer
        head = sources["index.html"].split("</head>")[0]
        self.assertIn("if (window.WebAssembly && window.fetch && window.Worker) { document.documentElement.className = "
                      "'can-run'; }", head)
        self.assertIn(".can-run .old-browser { display: none; }", css)
        self.assertIn('id="old-browser">This browser can\'t run KRONOS. Open Chrome, Edge or Firefox and type '
                      "localhost:5181.</p>", sources["index.html"])

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
        self.assertEqual(manifest["wheels"], ["et_xmlfile-2.0.0-py3-none-any.whl", "lark-1.3.1-py3-none-any.whl",
                                              "openpyxl-3.1.5-py2.py3-none-any.whl"])
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


    @unittest.skipIf(openpyxl is None, "openpyxl is not installed")
    def test_the_upload_gives_exactly_normal_pythons_answers(self):   # D87
        import base64
        node = shutil.which("node")
        if not node:
            need("Node.js is not installed")
        fresh = unpacked()
        data = orders_workbook()
        b64 = base64.b64encode(data).decode()
        cases = [{"name": "describe", "kind": "describe", "data": b64, "file": "orders.xlsx"},
                 {"name": "not a workbook", "kind": "describe", "data": base64.b64encode(b"nope").decode(), "file": "x.xlsx"},
                 {"name": "arguments", "kind": "arguments", "text": EXAMPLE},
                 {"name": "apply", "kind": "apply", "text": EXAMPLE, "data": b64, "file": "orders.xlsx",
                  "args": {"ws": "Orders"}, "today": TODAY},
                 {"name": "no such sheet", "kind": "apply", "text": EXAMPLE, "data": b64, "file": "orders.xlsx",
                  "args": {"ws": "Nope"}, "today": TODAY}]
        run = subprocess.run([node, str(ROOT / "tests" / "offline_pyodide.js"), str(fresh / "app")],
                             input=json.dumps(cases), capture_output=True, text=True, encoding="utf-8", timeout=600)
        self.assertEqual(run.returncode, 0, run.stderr[-3000:])
        answers = json.loads(run.stdout)["answers"]
        self.assertEqual(answers["describe"]["json"], kronos_web.describe_book(data, "orders.xlsx"))
        self.assertEqual(answers["not a workbook"]["json"], kronos_web.describe_book(b"nope", "x.xlsx"))
        self.assertEqual(answers["arguments"]["json"], kronos_web.arguments(EXAMPLE))
        steps = []
        want, copy = kronos_web.apply_changes(EXAMPLE, data, "orders.xlsx", json.dumps({"ws": "Orders"}), TODAY, steps.append)
        self.assertEqual(answers["apply"]["json"], want)
        self.assertEqual(answers["apply"]["steps"], steps)
        self.assertEqual(steps[-3:], ["reading", "running", "writing"])
        self.assertEqual(workbook_state(base64.b64decode(answers["apply"]["data"])), workbook_state(copy))
        refused, nothing = kronos_web.apply_changes(EXAMPLE, data, "orders.xlsx", json.dumps({"ws": "Nope"}), TODAY)
        self.assertEqual(answers["no such sheet"]["json"], refused)
        self.assertEqual((answers["no such sheet"]["data"], nothing), ("", b""))

    @unittest.skipIf(openpyxl is None, "openpyxl is not installed")
    def test_the_wizard_and_recipe_files_give_exactly_normal_pythons_answers(self):   # D89
        import base64
        node = shutil.which("node")
        if not node:
            need("Node.js is not installed")
        fresh = unpacked()
        demo = ROOT / "docs" / "demo"
        guide = (demo / "KRONOS_Demo_NCR_Register.docx").read_bytes()
        book = (demo / "NCR_Register_Demo.xlsx").read_bytes()
        texts = json.loads(kronos_web.read_recipe_file(guide, "guide.docx"))["recipes"]
        args = json.dumps([{"ws": "NCR Register"}, {"ws": "NCR Register"}, {"book": "", "ws": "NCR Register"}])
        b64 = lambda data: {"$b64": base64.b64encode(data).decode()}   # noqa: E731
        calls = [("wizard_reset", ["/kronos-data", TODAY], False),
                 ("wizard_call", ["ask", "r1", json.dumps([REQUEST, None])], True),
                 ("wizard_call", ["message", "r1", json.dumps([True])], False),
                 ("wizard_call", ["paste", "r1", json.dumps([reply(BROKEN)])], True),
                 ("wizard_call", ["message", "r1", json.dumps([False])], False),
                 ("wizard_call", ["paste", "r1", json.dumps([reply(EXAMPLE)])], True),
                 ("wizard_call", ["save", "r1", json.dumps([{"ws": "Orders"}])], False),
                 ("wizard_call", ["recipe", "f1", json.dumps([texts[2]["text"], "guide.docx"])], True),
                 ("wizard_call", ["save", "f1", json.dumps([{"ws": "NCR Register"}])], False),
                 ("words", [], False),
                 ("read_recipe_file", [b64(guide), "guide.docx"], False),
                 ("apply_changes_many", [json.dumps([t["text"] for t in texts]), b64(book), "NCR_Register_Demo.xlsx", args, TODAY],
                  True)]
        cases = [{"name": f"{n}: {fn}", "kind": "call", "fn": fn, "args": a, "step": st} for n, (fn, a, st) in enumerate(calls)]
        run = subprocess.run([node, str(ROOT / "tests" / "offline_pyodide.js"), str(fresh / "app")],
                             input=json.dumps(cases), capture_output=True, text=True, encoding="utf-8", timeout=900)
        self.assertEqual(run.returncode, 0, run.stderr[-3000:])
        answers = json.loads(run.stdout)["answers"]
        folder = temp("kronos python bridge ")
        for n, (fn, a, st) in enumerate(calls):
            a = [str(folder) if x == "/kronos-data" else (base64.b64decode(x["$b64"]) if isinstance(x, dict) else x) for x in a]
            steps = []
            got = getattr(kronos_web, fn)(*a, *([steps.append] if st else []))
            text = got if isinstance(got, str) else got[0]
            with self.subTest(call=f"{n}: {fn} {a[0] if a else ''}"):
                self.assertEqual(answers[f"{n}: {fn}"]["json"], text)
                self.assertEqual(answers[f"{n}: {fn}"]["steps"], steps)
                if not isinstance(got, str):
                    self.assertEqual(workbook_state(base64.b64decode(answers[f"{n}: {fn}"]["data"])), workbook_state(got[1]))


# ---------- the servers and the page ----------

def port_free():
    with socket.socket() as s:
        return s.connect_ex(("127.0.0.1", PORT)) != 0


def fetch(path, timeout=30):
    try:
        with urllib.request.urlopen(URL + path, timeout=timeout) as response:
            return response.status, response.headers.get("Content-Type", ""), response.read()
    except urllib.error.HTTPError as exc:
        return exc.code, exc.headers.get("Content-Type", ""), b""


class Served:
    """One of the zip's servers, started from the fresh folder as the launcher starts it, stopped afterwards. If the
    test fails while it runs, what the server printed is shown.

    What it prints goes to a file, never to a pipe. Nobody reads a pipe while the server runs, and a full one (4 KB
    on Windows) stops a server in the middle of a line: serve.py prints a traceback for every download the browser
    gives up, and on its third one it stopped serving, which failed the page test on Windows once (D86)."""

    def __init__(self, command):
        self.command = command

    def said(self):
        self.log.seek(0)
        return self.log.read().decode("utf-8", "replace").replace("\r\n", "\n")

    def __enter__(self):
        if not port_free():
            raise AssertionError(f"port {PORT} is already in use: stop whatever runs there first")
        # BROWSER: Python's webbrowser runs `true` instead of opening a browser. PYTHONUNBUFFERED: Python writes its
        # start message at once, as it does in the black window (into a file it would wait for a full buffer).
        env = dict(os.environ, BROWSER="true", PYTHONUNBUFFERED="1")
        self.log = tempfile.TemporaryFile()
        self.process = subprocess.Popen(self.command, cwd=unpacked(), env=env, stdout=self.log,
                                        stderr=subprocess.STDOUT)
        for _ in range(100):
            if not port_free():
                return self
            if self.process.poll() is not None:
                break
            time.sleep(0.1)
        self.process.kill()
        self.process.wait()
        raise AssertionError(f"{self.command} did not start: {self.said()}")

    def __exit__(self, kind, value, trace):
        self.process.terminate()
        try:
            self.process.wait(timeout=10)
        except subprocess.TimeoutExpired:
            self.process.kill()
            self.process.wait()
        self.output = self.said()
        self.log.close()
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


class OneAnswerAtATime:
    """A stand-in for serve.ps1 that runs on any computer, for the page tests. Like the real one, every request is
    taken in as it comes (Windows does that for serve.ps1), and one loop answers them in the order they came, each
    to its end before the next, with the same headers and the whole file every time. A page that waits for one file
    while it leaves an earlier one unread stops for ever on such a server (D86): this is how Chromium 109 and
    Firefox 115 are checked against that on Linux, where PowerShell's own server can't run."""

    TYPES = {".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8",
             ".css": "text/css; charset=utf-8", ".svg": "image/svg+xml", ".json": "application/json",
             ".wasm": "application/wasm", ".woff2": "font/woff2"}

    def __enter__(self):
        if not port_free():
            raise AssertionError(f"port {PORT} is already in use: stop whatever runs there first")
        self.root, self.waiting = (unpacked() / "app").resolve(), queue.Queue()
        waiting = self.waiting

        class Handler(http.server.BaseHTTPRequestHandler):
            protocol_version = "HTTP/1.1"   # a connection stays open between requests, as with the real one
            disable_nagle_algorithm = True   # send each piece at once, so that a small answer doesn't wait 40 ms

            def do_GET(self):   # hand the request to the loop, and wait for its answer
                answered = threading.Event()
                waiting.put((self, answered))
                answered.wait()

            def log_message(self, *args):
                pass

        class Server(http.server.ThreadingHTTPServer):
            def handle_error(self, request, client_address):   # a browser that stopped listening is not news
                pass

        self.server = Server(("127.0.0.1", PORT), Handler)
        threading.Thread(target=self.answer, daemon=True).start()
        threading.Thread(target=self.server.serve_forever, kwargs={"poll_interval": 0.05}, daemon=True).start()
        return self

    def answer(self):
        """serve.ps1's loop: the next request that came, its whole file, then the one after it."""
        while True:
            request, answered = self.waiting.get()
            if request is None:
                return
            try:
                relative = urllib.parse.unquote(urllib.parse.urlsplit(request.path).path).lstrip("/") or "index.html"
                file = (self.root / relative).resolve()
                found = self.root in file.parents and file.is_file()
                body = file.read_bytes() if found else b""
                request.send_response(200 if found else 404)
                if found:
                    request.send_header("Content-Type", self.TYPES.get(file.suffix.lower(), "application/octet-stream"))
                    request.send_header("Cache-Control", "no-cache")
                request.send_header("Content-Length", str(len(body)))
                request.end_headers()
                request.wfile.write(body)
            except Exception:   # noqa: BLE001 (the browser hung up, as after Cancel: on to the next request)
                request.close_connection = True
            finally:
                answered.set()

    def __exit__(self, kind, value, trace):
        self.server.shutdown()
        self.server.server_close()
        self.waiting.put((None, None))
        self.output = ""


def page_servers():
    """[(name, a way to start it)]: the zip's servers, and the stand-in for the PowerShell one."""
    real = [(name, lambda command=command: Served(command)) for name, command in servers()]
    return real + [("one-answer-at-a-time", OneAnswerAtATime)]


class Serving(unittest.TestCase):
    def test_a_connection_that_asks_for_nothing_holds_up_nobody(self):
        """A browser keeps spare connections open without asking for anything. A server that reads one connection at
        a time waits on such a connection and answers nobody else: BlueForce's serve.py does (D86). Not seen with a
        real browser in any run; this shows it without one."""
        fresh = unpacked()
        for name, command in servers():
            with self.subTest(server=name), Served(command):
                with socket.create_connection(("127.0.0.1", PORT), timeout=10):
                    time.sleep(0.5)   # the spare connection is the first one the server gets
                    began = time.monotonic()
                    try:
                        status, _, body = fetch("worker.js", timeout=20)
                    except OSError as exc:   # a timeout, when the server waits on the spare connection
                        self.fail(f"{name} gave no answer while a connection that asks for nothing was open: {exc}")
                    self.assertEqual(status, 200)
                    self.assertEqual(body, (fresh / "app" / "worker.js").read_bytes())
                    self.assertLess(time.monotonic() - began, 15)

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

    def go_ready(self, server):
        page = self.open(server)
        self.ready(page)
        page.wait_for_selector("#tips li", timeout=60000)   # KRONOS's words came from Python
        return page

    def review_banners(self, page):
        page.wait_for_selector("#step-review:not([hidden]) .review .banner", timeout=180000)
        return [b.text_content() for b in page.locator("#review-list .review .banner").all()]

    def finish(self):
        self.context.close()
        self.assertEqual(self.blocked, [], "the page asked for something outside this computer")
        self.assertEqual(self.errors, [])

    def test_ask_paste_repair_review_and_save_as_kronos_exe_does(self):   # D89
        import kronos_wizard
        folder = temp("kronos python wizard ")
        python = kronos_wizard.Wizard(folder, TODAY)
        python.ask("x", REQUEST)
        want_first = python.message("x")
        python.paste("x", reply(BROKEN))
        want_repair = python.message("x", instructions=False)
        python.paste("x", reply(EXAMPLE))
        want_saved = python.save("x", {"ws": "Orders"})
        for name, start in page_servers():
            with self.subTest(server=name), start():
                page = self.open(name)
                self.assertEqual(page.text_content("#version"), f"Version {built()[1]['version']} · {built()[1]['date']}")
                self.assertEqual(page.locator("#old-browser").count(), 0)
                # Cancel while Python is still starting: the panel goes and a fresh worker starts
                page.click("#load-example")
                page.wait_for_selector("#working", state="visible", timeout=30000)
                page.click("#cancel")
                self.assertFalse(page.is_visible("#working"))
                self.assertIn("Cancelled", page.text_content("#file-note"))
                self.ready(page)
                page.wait_for_selector("#tips li", timeout=60000)
                self.assertIn("3.11.3 in this browser (Pyodide 0.24.1)", page.text_content("#python"))
                self.assertIn("Name the sheets", page.text_content("#tips"))
                # 1 Ask: the message with instructions is exactly what KRONOS.exe sends Codex
                page.fill("#request", REQUEST)
                page.click("#write-message")
                page.wait_for_selector("#ai-panel-ask .copy-instructions", timeout=120000)
                page.click("#ai-panel-ask .copy-instructions")
                self.assertEqual(page.evaluate("window.kronosLastCopied"), want_first)
                self.shot(page, "ask")
                # a broken reply: KRONOS sends the toolchain's problems back, as KRONOS.exe would
                page.fill("#ai-panel-ask-reply", reply(BROKEN))
                page.click("#ai-panel-ask .check-reply")
                self.assertEqual(self.review_banners(page), ["WAITING FOR THE AI Send the problems to the AI and paste its new reply."])
                page.click("#review-list .copy-orchestrator")
                self.assertEqual(page.evaluate("window.kronosLastCopied"), want_repair)
                self.assertIn("unknown key GOLA", page.text_content("#review-list"))
                page.fill("#review-list .reply", reply(EXAMPLE))
                page.click("#review-list .check-reply")
                page.wait_for_function("document.querySelector('#review-list .banner').textContent.indexOf('PROVEN') === 0",
                                       timeout=180000)
                self.assertEqual(self.review_banners(page), ["PROVEN It passed every check and 3 of 3 practice tests."])
                self.assertTrue(page.is_disabled("#go-save"))
                self.shot(page, "review")
                page.check("#review-list input[type=checkbox]")
                page.click("#go-save")
                # 4 Save: the .bas and the practice run, byte for byte KRONOS.exe's
                page.wait_for_selector("#step-save:not([hidden]) .download-bas")
                self.assertEqual(page.input_value("#save-list input"), "Orders")
                with page.expect_download() as download:
                    page.click(".download-bas")
                self.assertEqual(download.value.suggested_filename, want_saved["bas_name"])
                self.assertEqual(Path(download.value.path()).read_bytes(), want_saved["bas"].encode("ascii"))
                with page.expect_download() as practice:
                    page.click(".download-practice")
                self.assertEqual(Path(practice.value.path()).read_bytes(), want_saved["practice"].encode("ascii"))
                self.assertIn("Alt+F11", page.text_content("#save-list"))
                self.shot(page, "save")
                self.finish()

    def test_the_question_step(self):   # D89
        for name, start in page_servers():
            with self.subTest(server=name), start():
                page = self.go_ready(name)
                page.fill("#request", REQUEST)
                page.click("#write-message")
                page.wait_for_selector("#ai-panel-ask .check-reply", timeout=120000)
                page.fill("#ai-panel-ask-reply", QUESTION_REPLY)
                page.click("#ai-panel-ask .check-reply")
                page.wait_for_selector("#step-question:not([hidden])", timeout=120000)
                self.assertEqual(page.text_content("#question-text"), "Which sheet holds the orders?")
                page.fill("#answer", "The sheet is Orders.")
                page.click("#write-answer")
                page.wait_for_selector("#ai-panel-question .copy-orchestrator", timeout=120000)
                page.click("#ai-panel-question .copy-orchestrator")
                self.assertEqual(page.evaluate("window.kronosLastCopied"), "The sheet is Orders.\n")
                self.shot(page, "question")
                page.fill("#ai-panel-question-reply", reply(EXAMPLE))
                page.click("#ai-panel-question .check-reply")
                self.assertEqual(self.review_banners(page), ["PROVEN It passed every check and 3 of 3 practice tests."])
                self.finish()

    @unittest.skipIf(openpyxl is None, "openpyxl is not installed")
    def test_a_recipe_file_with_three_recipes_on_one_copy_of_the_workbook(self):   # D89
        import kronos_book
        demo = ROOT / "docs" / "demo"
        texts = [(demo / n).read_text(encoding="utf-8") for n in ("recipe_1_highlight.vcl", "recipe_2_tidy.vcl", "recipe_3_summary.vcl")]
        book = (demo / "NCR_Register_Demo.xlsx").read_bytes()
        sheet = "NCR Register"
        want, want_copy = kronos_book.apply_many(texts, book, "NCR_Register_Demo.xlsx",
                                                  [{"ws": sheet}, {"ws": sheet}, {"book": "", "ws": sheet}], TODAY)
        for name, start in page_servers():
            with self.subTest(server=name), start():
                page = self.go_ready(name)
                # an .xls and a file that is not a workbook are turned away in words
                page.set_input_files("#book-file", files={"name": "old.xls", "mimeType": "application/vnd.ms-excel", "buffer": b"x"})
                self.assertIn("old .xls workbook", page.text_content("#book-note"))
                page.set_input_files("#book-file", files={"name": "notes.xlsx", "mimeType": XLSX, "buffer": b"just text"})
                page.wait_for_function("document.getElementById('book-note').textContent.indexOf('look like') >= 0", timeout=120000)
                page.set_input_files("#book-file", files={"name": "NCR_Register_Demo.xlsx", "mimeType": XLSX, "buffer": book})
                page.wait_for_selector("#book-info:not([hidden])", timeout=120000)
                self.assertTrue(page.is_visible("#include-line"))
                # the demo's Word guide holds three recipes: all three are chosen and checked
                page.set_input_files("#recipe-file", str(demo / "KRONOS_Demo_NCR_Register.docx"))
                page.wait_for_selector("#choose-all", timeout=120000)
                self.assertIn("KRONOS found 3 recipes.", page.text_content("#recipe-list"))
                self.shot(page, "recipes")
                page.click("#check-chosen")
                self.assertIn("Tick at least one recipe", page.text_content("#file-note"))
                page.click("#choose-all")
                page.click("#check-chosen")
                self.assertEqual(self.review_banners(page), ["PROVEN It passed every check and 3 of 3 practice tests."] * 3)
                for box in page.locator("#review-list input[type=checkbox]").all():
                    box.check()
                page.click("#go-save")
                page.wait_for_selector("#apply-card:not([hidden]) select", timeout=120000)
                self.assertEqual(page.locator("#save-list .save").count(), 3)
                for select in page.locator("#apply-fields select").all():
                    select.select_option(sheet)
                page.click("#apply")
                page.wait_for_selector("#applied-banner", timeout=180000)
                self.assertEqual(json.loads(page.evaluate("window.kronosLastApplied")), want)   # exactly normal Python's answer
                self.assertEqual(page.locator("#applied h2").count(), 4)   # one part per recipe, then all together
                self.shot(page, "copy")
                with page.expect_download() as download:
                    page.click("#download")
                self.assertEqual(workbook_state(Path(download.value.path()).read_bytes()), workbook_state(want_copy))
                self.finish()

    @unittest.skipIf(openpyxl is None, "openpyxl is not installed")
    def test_a_recipe_that_reads_a_second_workbook(self):   # D120
        import kronos_book
        recipe = (ROOT / "tests" / "data" / "wps_from_wps.vcl").read_bytes()
        new = (ROOT / "data" / "templates" / "wps_form.xlsx").read_bytes()
        old = (ROOT / "data" / "templates" / "wps_form_b.xlsx").read_bytes()
        want, want_copy = kronos_book.apply_many([recipe.decode("utf-8")], new, "WPS-201.xlsx", [{"old": "", "ws": "WPS"}],
                                                  TODAY, read=(old, "WPS-105.xlsx"))
        for name, start in page_servers():
            with self.subTest(server=name), start():
                page = self.go_ready(name)
                page.set_input_files("#book-file", files={"name": "WPS-201.xlsx", "mimeType": XLSX, "buffer": new})
                page.wait_for_selector("#book-info:not([hidden])", timeout=120000)
                page.set_input_files("#recipe-file", files={"name": "wps_from_wps.vcl", "mimeType": "text/plain",
                                                            "buffer": recipe})
                self.assertEqual(self.review_banners(page), ["PROVEN It passed every check and 2 of 2 practice tests."])
                page.check("#review-list input[type=checkbox]")
                page.click("#go-save")
                page.wait_for_selector("#apply-card:not([hidden]) .choose-read", timeout=120000)
                self.assertIn("no macro to download", page.text_content("#save-list"))
                self.assertIn("Workbook to read from (old)", page.text_content("#apply-fields"))
                page.click("#apply")
                self.assertEqual(page.text_content("#apply-note"), "Choose the workbook to read from first.")
                page.set_input_files("#read-file", files={"name": "WPS-105.xlsx", "mimeType": XLSX, "buffer": old})
                page.wait_for_function("document.querySelector('#apply-fields .read-name') && document.querySelector("
                                       "'#apply-fields .read-name').textContent.indexOf('WPS-105.xlsx') === 0", timeout=60000)
                page.click("#apply")
                page.wait_for_selector("#applied-banner", timeout=180000)
                self.assertEqual(json.loads(page.evaluate("window.kronosLastApplied")), want)   # exactly normal Python's answer
                self.assertIn("It read WPS-105.xlsx, which is not changed.", page.text_content("#applied"))
                self.shot(page, "two-workbooks")
                with page.expect_download() as download:
                    page.click("#download")
                self.assertEqual(zip_parts(Path(download.value.path()).read_bytes()), zip_parts(want_copy))   # (its zlib differs)
                self.finish()

    @unittest.skipIf(openpyxl is None, "openpyxl is not installed")
    def test_a_word_template_filled_from_a_workbook(self):   # D121
        import kronos_book
        recipe = (ROOT / "tests" / "data" / "doc_estimate.vcl").read_bytes()
        template = (ROOT / "data" / "templates" / "estimate_template.docx").read_bytes()
        book = (ROOT / "data" / "templates" / "estimation.xlsx").read_bytes()
        want, want_doc = kronos_book.fill_document(recipe.decode("utf-8"), template, "Estimate.docx",
                                                   {"estNo": "EST-9001", "onDate": "20 Nov 2026"}, TODAY,
                                                   book=(book, "E-9100.xlsx"))
        for name, start in page_servers():
            with self.subTest(server=name), start():
                page = self.go_ready(name)
                page.set_input_files("#book-file", files={"name": "E-9100.xlsx", "mimeType": XLSX, "buffer": book})
                page.wait_for_selector("#book-info:not([hidden])", timeout=120000)
                page.set_input_files("#recipe-file", files={"name": "doc_estimate.vcl", "mimeType": "text/plain",
                                                            "buffer": recipe})
                self.assertEqual(self.review_banners(page), ["PROVEN It passed every check and 2 of 2 practice tests."])
                page.check("#review-list input[type=checkbox]")
                page.click("#go-save")
                page.wait_for_selector("#doc-list .doc .fill", timeout=120000)
                self.assertTrue(page.is_hidden("#apply-card"))   # a DOCUMENT changes no workbook
                self.assertIn("Ex-works, packing included", page.text_content("#doc-list .ai-texts"))
                self.assertIn("E-9100.xlsx, as Excel last saved it", page.text_content("#doc-list"))
                page.click("#doc-list .fill")
                self.assertEqual(page.text_content("#doc-list .note"), "Choose the Word template first.")
                with page.expect_file_chooser() as chooser:
                    page.click("#doc-list .choose-template")
                chooser.value.set_files(files=[{"name": "Estimate.docx", "mimeType": "application/vnd.openxmlformats-"
                                                "officedocument.wordprocessingml.document", "buffer": template}])
                page.wait_for_function("document.querySelector('#doc-list .template-name') && document.querySelector("
                                       "'#doc-list .template-name').textContent.indexOf('Estimate.docx') === 0", timeout=60000)
                page.click("#doc-list .fill")
                page.wait_for_function("document.querySelector('#doc-list .banner') || !document.querySelector('#doc-list "
                                       ".note').hidden", timeout=180000)
                self.assertEqual(page.text_content("#doc-list .note"), "")
                self.assertEqual(json.loads(page.evaluate("window.kronosLastFilled")), want)   # exactly normal Python's answer
                self.assertIn("WORD FILE READY", page.text_content("#doc-list .banner"))
                self.shot(page, "word")
                with page.expect_download() as download:
                    page.click("#doc-list .download-doc")
                self.assertTrue(download.value.suggested_filename.startswith("Estimate (KRONOS "))
                self.assertEqual(zip_parts(Path(download.value.path()).read_bytes()), zip_parts(want_doc))
                self.finish()

    def test_one_recipe_files_are_checked_straight_away(self):   # D89
        for name, start in page_servers():
            with self.subTest(server=name), start():
                page = self.go_ready(name)
                page.set_input_files("#recipe-file", files={"name": "reply.md", "mimeType": "text/markdown",
                                                            "buffer": reply(EXAMPLE).encode()})
                self.assertEqual(self.review_banners(page), ["PROVEN It passed every check and 3 of 3 practice tests."])
                page.click("#step-review .back")
                page.set_input_files("#recipe-file", files={"name": "broken.txt", "mimeType": "text/plain", "buffer": BROKEN.encode()})
                page.wait_for_function("document.querySelector('#review-list .banner') && "
                                       "document.querySelector('#review-list .banner').textContent.indexOf('NOT PROVEN') === 0",
                                       timeout=180000)
                self.assertTrue(page.is_visible("#review-list .send-problems"))
                page.click("#step-review .back")
                page.set_input_files("#recipe-file", files={"name": "notes.txt", "mimeType": "text/plain", "buffer": b"no recipe"})
                page.wait_for_function("document.getElementById('file-note').textContent.indexOf('found no recipe') >= 0",
                                       timeout=120000)
                self.finish()

    @unittest.skipIf(openpyxl is None, "openpyxl is not installed")
    def test_what_a_copy_cannot_keep_must_be_ticked_before_the_download(self):   # D87
        data = orders_workbook(filtered=True)
        for name, start in page_servers():
            with self.subTest(server=name), start():
                page = self.go_ready(name)
                page.set_input_files("#book-file", files={"name": "orders.xlsx", "mimeType": XLSX, "buffer": data})
                page.wait_for_selector("#book-info:not([hidden])", timeout=120000)
                page.click("#load-example")
                self.review_banners(page)
                page.check("#review-list input[type=checkbox]")
                page.click("#go-save")
                page.wait_for_selector("#apply-card:not([hidden]) select", timeout=120000)
                self.assertEqual(page.input_value("#apply-fields select"), "Orders")
                page.click("#apply")
                page.wait_for_selector("#applied-banner", timeout=180000)
                self.assertIn("the filter", page.text_content("#warnings"))
                self.assertTrue(page.is_disabled("#download"))
                page.check("#understand")
                self.assertTrue(page.is_enabled("#download"))
                self.shot(page, "warning")
                self.finish()

if __name__ == "__main__":
    unittest.main()
