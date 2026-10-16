import random
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
from vcl_parse import parse, parser  # noqa: E402


def op_tree(line):
    return parser().parse(line, start="op_value").children[0]


class RealSpecs(unittest.TestCase):
    def test_all_specs_parse(self):
        for p in sorted((ROOT / "specs").glob("*.vcl")):
            spec, problems = parse(p.read_text())
            self.assertEqual(problems, [], p.name)
            self.assertEqual(spec["kind"], "TASK")

    def test_example1_rev3(self):
        spec, problems = parse((ROOT / "vcl" / "examples" / "example1_rev3.vcl").read_text())
        self.assertEqual(problems, [])
        loop = [n for n in spec["ops"] if n["kind"] == "loop_rows"][0]
        self.assertEqual(loop["children"][0]["kind"], "if_block")
        self.assertEqual(len(loop["children"][0]["children"]), 2)


class Disambiguation(unittest.TestCase):
    def test_option_vs_comparison(self):
        call = op_tree("SET_VALUE(CELL($ws, r, 1), got = 0)")
        args = call.children[1].children
        self.assertEqual(args[1].data, "comparison")          # got = 0 is a comparison
        call2 = op_tree("DICT_ADD(d, k, 1, mode=count)")
        self.assertEqual(call2.children[1].children[3].data, "option")

    def test_reserved_option_word_needs_parentheses_to_compare(self):
        # "name" is a reserved option key; a comparison on a variable called name must be wrapped
        call = op_tree("SET_VALUE(CELL($ws, 1, 1), (name = 1))")
        self.assertEqual(call.children[1].children[1].data, "comparison")

    def test_precedence_matches_vba(self):
        t = op_tree('LET(x, "a" & 1 + 2)')
        concat = t.children[1].children[1]
        self.assertEqual(concat.data, "concat")             # & binds looser than +
        t2 = op_tree("IF NOT a = b THEN LET(x, 1)")
        self.assertEqual(t2.children[0].data, "not_op")     # NOT (a = b)

    def test_comparison_is_not_associative(self):
        _, problems = parse("VCL 0.1\nTASK\nPROC     P\nKIND     Sub\nMODULE   m\nOP       IF a < b < c THEN LET(x, 1)\nERR      standard")
        self.assertTrue(any("syntax error" in m for _, m in problems))

    def test_else_attaches_to_if(self):
        text = "\n".join(["VCL 0.1", "TASK", "PROC     P", "KIND     Sub", "MODULE   m",
                          "OP       IF a > 1", "OP         LET(x, 1)", "OP       ELSE",
                          "OP         LET(x, 2)", "ERR      standard"])
        spec, problems = parse(text)
        self.assertEqual(problems, [])
        node = spec["ops"][0]
        self.assertEqual((len(node["children"]), len(node["else"])), (1, 1))


class Reporting(unittest.TestCase):
    def test_bad_lines_are_reported_not_raised(self):
        cases = ["VAR      x as String", "OP       lowercase_op(1)", "FIXTURE  f.xlsx cols=A:X rows=1",
                 "OP       LET(x, sh.Name)", "PRE      XLOOKUP($nope, bogus) >", "OP       IF a > 1 THEN"]
        for c in cases:
            text = "VCL 0.1\nTASK\nPROC     P\nKIND     Sub\nMODULE   m\n" + c + "\nERR      standard"
            spec, problems = parse(text)   # must not raise
            self.assertIsInstance(problems, list)

    def test_question_reply(self):
        spec, problems = parse("# QUESTION: Which folder should I clean?")
        self.assertEqual((spec["kind"], problems), ("QUESTION", []))

    def test_fuzz_never_raises(self):
        rng = random.Random(7)
        base = (ROOT / "specs" / "task_01.vcl").read_text()
        for _ in range(400):
            chars = list(base)
            for _ in range(rng.randint(1, 8)):
                i = rng.randrange(len(chars))
                chars[i] = rng.choice(' ()=,"$[]!#;:\nAZaz09-+')
            parse("".join(chars))


if __name__ == "__main__":
    unittest.main()
