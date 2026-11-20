"""samples/kgc/ (D105, D106): Boss's real KGC samples, the one exception to the synthetic-only rule.

The folder holds exactly the files its README lists, unchanged (SHA-256), and nothing that tests, specs, FIXTUREs,
the corpus or KRONOS's code uses ever reads from it.
"""
import hashlib
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SAMPLES = ROOT / "samples" / "kgc"
README = SAMPLES / "README.md"


def listed():
    rows = re.findall(r"^\| `([^`]+)` \|.*\| `([0-9a-f]{64})` \|$", README.read_text(encoding="utf-8"), re.M)
    return dict(rows)


class TheSamplesFolder(unittest.TestCase):
    def test_it_holds_exactly_the_listed_files_unchanged(self):
        files = listed()
        self.assertEqual(len(files), 11)
        on_disk = {p.name for p in SAMPLES.iterdir()} - {"README.md"}
        self.assertEqual(on_disk, set(files))
        for name, digest in files.items():
            with self.subTest(file=name):
                self.assertEqual(hashlib.sha256((SAMPLES / name).read_bytes()).hexdigest(), digest)

    def test_nothing_that_is_tested_or_trained_on_reads_them(self):
        # docs and decision records may name the folder; code, specs, tests and data may not use it
        allowed = {Path(__file__).resolve()}
        places = [ROOT / d for d in ("compiler", "vcl", "specs", "data", "app", "web", "vba", "packaging", "tests")]
        for place in places:
            for path in place.rglob("*"):
                if path.is_file() and path.resolve() not in allowed and path.suffix in (
                        ".py", ".vcl", ".txt", ".json", ".lark", ".js", ".html", ".bas", ".md", ".spec"):
                    with self.subTest(file=str(path.relative_to(ROOT))):
                        self.assertNotIn("samples/kgc", path.read_text(encoding="utf-8", errors="replace"))


if __name__ == "__main__":
    unittest.main()
