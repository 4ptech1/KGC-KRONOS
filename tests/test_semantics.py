import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "vcl"))
sys.path.insert(0, str(ROOT / "tests"))
import semantics  # noqa: E402
from conformance_cases import CASES  # noqa: E402


class Conformance(unittest.TestCase):
    def test_cases(self):
        for fn, args, expected, why in CASES:
            with self.subTest(fn=fn, why=why):
                self.assertEqual(getattr(semantics, fn)(*args), expected)

    def test_equals_refuses_unguarded_mixed_types(self):
        with self.assertRaises(TypeError):
            semantics.equals(25, semantics.XlError("#N/A"))


if __name__ == "__main__":
    unittest.main()
