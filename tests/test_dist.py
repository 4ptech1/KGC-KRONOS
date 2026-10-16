"""dist/forge_module.bas is committed so Boss can copy it any time: it must match the specs."""
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
from vcl_codegen import assemble, compile_text, module_title  # noqa: E402


class CommittedModule(unittest.TestCase):
    def test_forge_module_matches_the_specs(self):
        specs = sorted((ROOT / "specs").glob("*.vcl"))
        procs = [compile_text(p.read_text(encoding="utf-8"), p.name) for p in specs]
        fresh = assemble(procs, module_title(len(procs)))
        committed = (ROOT / "dist" / "forge_module.bas").read_text(encoding="utf-8")
        self.assertTrue(fresh == committed, "dist/forge_module.bas is out of date: run "
                                            "python compiler/vclc.py all specs/*.vcl -o dist and commit it")


if __name__ == "__main__":
    unittest.main()
