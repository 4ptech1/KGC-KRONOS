"""Item 5 Phase 3 (D79): KRONOS.exe's packaging, checked on any platform.

The exe itself is built and self-tested on Windows by .github/workflows/build-exe.yml. These tests protect what
can be checked anywhere: the program finds its data files through compiler/resources.py (so it works from the
unpacking folder), the PyInstaller spec bundles every file the code reads, and the build tools are pinned.
"""
import re
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SPEC = ROOT / "packaging" / "kronos.spec"
GOOD = (ROOT / "vcl" / "examples" / "example1_rev3.vcl").read_text(encoding="utf-8")


def spec_files():
    """The data files the spec bundles, as paths under the repo: ('vcl/grammar.lark', ...)."""
    found = re.findall(r'\(str\(root((?: / "[^"]+")+)\), "[^"]+"\)', SPEC.read_text(encoding="utf-8"))
    return {"/".join(re.findall(r'"([^"]+)"', part)) for part in found}


def runtime_files():
    """The data files the code reads through resources.ROOT: every `ROOT / "a" / "b"` in the modules that import
    it (the other tools in compiler/ have a ROOT of their own and are not part of the exe)."""
    found = set()
    for path in [*(ROOT / "compiler").glob("*.py"), *(ROOT / "app").glob("*.py")]:
        text = path.read_text(encoding="utf-8")
        if path.name == "resources.py" or "from resources import" not in text:
            continue
        for part in re.findall(r'\bROOT((?: / "[^"]+")+)', text):
            found.add("/".join(re.findall(r'"([^"]+)"', part)))
    return found


class Spec(unittest.TestCase):
    def test_every_file_the_code_reads_is_bundled(self):
        self.assertEqual(runtime_files(), spec_files())

    def test_every_bundled_file_exists(self):
        for name in spec_files():
            self.assertTrue((ROOT / name).is_file(), name)

    def test_the_program_is_found_through_the_paths_the_code_adds_at_run_time(self):
        text = SPEC.read_text(encoding="utf-8")
        for folder in ("app", "compiler", "vcl"):   # the build can't see their sys.path.insert calls
            self.assertIn(f'root / "{folder}"', text.split("datas=")[0])


class BuildTools(unittest.TestCase):
    def test_every_build_tool_is_pinned_with_its_hash(self):
        lines = [ln for ln in (ROOT / "requirements-build.txt").read_text(encoding="utf-8").splitlines()
                 if ln.strip() and not ln.startswith("#")]
        pins = [ln for ln in lines if not ln.strip().startswith("--hash")]
        hashes = [ln for ln in lines if ln.strip().startswith("--hash=sha256:")]
        self.assertGreaterEqual(len(pins), 7)
        self.assertEqual(len(pins), len(hashes))
        for pin in pins:
            self.assertRegex(pin, r"^[A-Za-z0-9._-]+==[0-9][^ ]* \\$")
        self.assertTrue(any(ln.startswith("pyinstaller==") for ln in pins))

    def test_the_workflow_builds_on_windows_and_self_tests_the_exe(self):
        text = (ROOT / ".github" / "workflows" / "build-exe.yml").read_text(encoding="utf-8")
        for words in ("windows-latest", "--require-hashes", "packaging/kronos.spec", "--self-test", "--log"):
            self.assertIn(words, text)


class Frozen(unittest.TestCase):
    def test_the_data_files_come_from_the_unpacking_folder(self):
        """With sys._MEIPASS set (as in KRONOS.exe), the grammar, ForgeLib and the prompt are read from there."""
        unpacked = Path(tempfile.mkdtemp())
        for name in spec_files():
            (unpacked / name).parent.mkdir(parents=True, exist_ok=True)
            text = (ROOT / name).read_text(encoding="utf-8")
            (unpacked / name).write_text("' UNPACKED COPY\n" + text if name.endswith(".bas") else text,
                                         encoding="utf-8")
        code = (f"import sys; sys._MEIPASS = {str(unpacked)!r}; sys.path.insert(0, {str(ROOT / 'compiler')!r}); "
                "import resources, kronos_engine, vcl_codegen, vcl_parse; "
                "print(resources.FROZEN, resources.ROOT == kronos_engine.ROOT); "
                "print(vcl_parse.GRAMMAR.parent.parent, vcl_codegen.FORGELIB.parent.parent, "
                "kronos_engine.PROMPT.parent.parent); "
                f"p = kronos_engine.prove(open({str(ROOT / 'vcl' / 'examples' / 'example1_rev3.vcl')!r}).read()); "
                "print(p.ok, p.module.count('UNPACKED COPY'))")
        done = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, timeout=120)
        self.assertEqual(done.returncode, 0, done.stderr)
        frozen, folders, proven = done.stdout.strip().splitlines()
        self.assertEqual(frozen, "True True")
        self.assertEqual(set(folders.split(" ")), {str(unpacked)})
        self.assertEqual(proven, "True 1")   # a recipe proved, and the module built from the unpacked ForgeLib

    def test_from_a_checkout_the_root_is_the_repo(self):
        code = (f"import sys; sys.path.insert(0, {str(ROOT / 'compiler')!r}); import resources; "
                "print(resources.FROZEN, resources.ROOT)")
        done = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, timeout=60)
        self.assertEqual(done.stdout.strip(), f"False {ROOT}")


if __name__ == "__main__":
    unittest.main()
