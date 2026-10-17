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

    def test_mod_by_zero_is_an_error(self):
        with self.assertRaises(ZeroDivisionError):
            semantics.mod(7, 0)

    def test_parse_date_needs_a_known_format(self):
        with self.assertRaises(ValueError):
            semantics.parse_date("15/10/2026", "d/m/y")

    def test_every_vcl_02_function_has_cases(self):
        names = {fn for fn, *_ in CASES}
        for fn in ("same", "clean_text", "is_error", "parse_date", "iif", "mod"):
            self.assertIn(fn, names)


if __name__ == "__main__":
    unittest.main()
