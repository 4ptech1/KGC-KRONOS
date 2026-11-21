"""The grammar that locks the model's output (train/grammar_gbnf.py, train/gbnf_check.py, train/vcl.gbnf; D127).

The converter is tested through an independent reader (gbnf_check), so no model and no runner is needed: the grammar
must accept every canonical recipe this repo has, reject broken ones, and mean for each line what vcl/grammar.lark means.
"""
import glob
import random
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "train"))
sys.path.insert(0, str(ROOT / "compiler"))

import gbnf_check as gc  # noqa: E402
import grammar_gbnf as gg  # noqa: E402
import vcl_parse  # noqa: E402

GRAMMAR_FILE = ROOT / "train" / "vcl.gbnf"


def spec_files():
    """The recipes the model is meant to write: VCL 0.6 and 0.7, canonical (the corpus, the held-out items, the examples)."""
    out = sorted(glob.glob(str(ROOT / "data/verified/*/spec.vcl")) + glob.glob(str(ROOT / "data/eval/*/spec.vcl"))
                 + glob.glob(str(ROOT / "vcl/examples/example[45]*.vcl")))
    return [Path(p) for p in out]


def task(*ops):
    """A small canonical TASK whose OP lines are (depth, text) pairs."""
    lines = "".join(f"OP{' ' * (7 + 2 * d)}{text}\n" for d, text in ops)
    return ("VCL 0.6\nTASK\nGOAL     Synthetic task used to test the grammar's nesting rules in detail\n"
            "PROC     T\nKIND     Sub\nMODULE   modT\n" + lines +
            "ERR      standard\nFIXTURE  a_01.xlsx sheet=S cols=A:X rows=1\n"
            "TEST     fixture=a_01.xlsx call=T() assert=ERROR_COUNT = 0\n")


class Checker(unittest.TestCase):
    """gbnf_check on small grammars: it must be right, or the tests that rely on it mean nothing."""

    def g(self, text):
        return gc.Grammar(text)

    def test_literals_classes_and_repeats(self):
        g = self.g('root ::= "ab" [0-9]+ ("x" | "y")? [^"\\n]{2,3}\n')
        self.assertTrue(g.accepts("ab12xzz"))
        self.assertTrue(g.accepts("ab7yq q"))
        self.assertTrue(g.accepts("ab7zzz"))
        for bad in ("ab", "ab1x", "ab1zzzz", "ab1z\n", "b1zz", 'ab1"zz'):
            self.assertFalse(g.accepts(bad), bad)

    def test_escapes_and_ranges(self):
        g = self.g(r'root ::= "\"" [\]\-a-c]+ "\\" "\n"' + "\n")
        self.assertTrue(g.accepts('"]-ab\\\n'))
        self.assertFalse(g.accepts('"d\\\n'))

    def test_recursion_and_ambiguity(self):
        g = self.g('root ::= e\ne ::= t ("+" t)*\nt ::= "(" e ")" | [a-z]\n')
        self.assertTrue(g.accepts("(a+(b+c))+d"))
        self.assertFalse(g.accepts("(a+b"))
        g = self.g('root ::= a a\na ::= "x"* \n')
        self.assertTrue(g.accepts("xxxx"))

    def test_it_finds_what_llama_cpp_refuses(self):
        self.assertEqual(self.g('root ::= a\n').problems(), ["root uses a, which is not defined"])
        self.assertEqual(self.g('x ::= "a"\n').problems(), ["there is no root rule"])
        self.assertIn("left recursion in e", self.g('root ::= e\ne ::= e "+" "a" | "a"\n').problems())
        self.assertIn("left recursion in a", self.g('root ::= a\na ::= b a | "x"\nb ::= "y"?\n').problems())

    def test_it_refuses_what_it_cannot_read(self):
        for bad in ('root ::= "a\n', 'root ::= (a\n', 'root = "a"\n', 'root ::= "a"\nroot ::= "b"\n', 'root ::= [a\n'):
            with self.assertRaises(gc.GbnfError, msg=bad):
                self.g(bad)


class Converter(unittest.TestCase):
    def regex(self, pattern):
        return gg._Regex(pattern).parse()

    def test_regexes(self):
        self.assertEqual(self.regex(r"[A-Za-z_][A-Za-z0-9_]*"), "[A-Za-z_] [A-Za-z0-9_]*")
        self.assertEqual(self.regex(r"\d+(\.\d+)?"), '[0-9]+ ("." [0-9]+)?')
        self.assertEqual(self.regex(r"#[0-9A-Fa-f]{6}"), '"#" [0-9A-Fa-f]{6}')
        self.assertEqual(self.regex(r'"(?:[^"]|"")*"'), '"\\"" ([^"\\n\\r] | "\\"\\"")* "\\""')
        self.assertEqual(self.regex(r"/".replace("/", "x") + r"\$?[A-Z]{1,3}"), '"x" "$"? [A-Z]{1,3}')

    def test_a_lookahead_is_dropped_and_a_lazy_quantifier_is_greedy(self):
        self.assertEqual(self.regex(r"(dir|from)(?==)"), '("dir" | "from")')
        self.assertEqual(self.regex(r"[^,\r\n]+?(?=,[A-Z]{1,3}:| rows=|$)"), "[^,\\r\\n]+")

    def test_a_negated_class_never_runs_over_a_line_end(self):
        self.assertIn("\\n", self.regex(r"[^x]"))

    def test_what_it_cannot_convert_it_refuses(self):
        for pattern in (r"a.b", r"^a", r"\w+", r"a$"):
            with self.assertRaises(gg.GrammarError, msg=pattern):
                self.regex(pattern)

    def test_the_lark_reader_keeps_comments_out_of_the_next_lines(self):
        rules, terminals, ignore = gg.read_lark(
            '// a comment\n'
            'start: item ("," item)*   // trailing\n'
            '     | "x"                // another // one\n'
            '     | "//" -> slashes\n'
            'item: NAME\n'
            'NAME: /[a-z]+/\n'
            '%ignore /[ \\t]+/\n')
        self.assertEqual(set(rules), {"start", "item"})
        self.assertEqual(set(terminals), {"NAME"})
        self.assertEqual(ignore, "/[ \\t]+/")
        self.assertEqual(len(rules["start"][1]), 3)
        self.assertEqual(rules["start"][1][2], ("str", "//"))

    def test_spaces_go_between_tokens_and_inside_optional_groups(self):
        rules, terminals, _ = gg.read_lark('r: A ("," A)* [B] "c"\nA: "a"\nB: "b"\n')
        conv = gg.Converter(rules, terminals)
        self.assertEqual(conv.rule_line("r", rules["r"]), 'r ::= t-a (sp "," sp t-a)* (sp t-b)? sp "c"')

    def test_a_sequence_that_can_start_empty_is_refused(self):
        rules, terminals, _ = gg.read_lark('r: [A] B\nA: "a"\nB: "b"\n')
        with self.assertRaises(gg.GrammarError):
            gg.Converter(rules, terminals).rule_line("r", rules["r"])

    def test_every_key_the_parser_reads_has_a_count(self):
        gg.check_tables(gg.read_lark(vcl_parse.GRAMMAR.read_text(encoding="utf-8"))[0])   # raises if not


class TheGrammar(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.text = GRAMMAR_FILE.read_text(encoding="utf-8")
        cls.g = gc.Grammar(cls.text)

    def accepts(self, text):
        return self.g.accepts(text)

    def test_the_committed_file_is_what_the_sources_give(self):
        self.assertEqual(self.text, gg.generate(), "run python train/grammar_gbnf.py")

    def test_llama_cpp_could_read_it(self):
        self.assertEqual(self.g.problems(), [])
        self.assertTrue(all("\n" not in line for line in self.text.split("\n")))   # one rule per line

    def test_it_accepts_every_recipe_we_have(self):
        files = spec_files()
        self.assertGreater(len(files), 150)
        for path in files:
            with self.subTest(file=str(path.relative_to(ROOT))):
                self.assertTrue(self.accepts(path.read_text(encoding="utf-8")))

    def test_it_accepts_each_kind_of_file(self):
        kinds = {p.read_text(encoding="utf-8").split("\n")[1] for p in spec_files()}
        self.assertEqual(kinds, {"TASK", "BUILD", "DOCUMENT"})

    def test_it_rejects_broken_recipes(self):
        base = (ROOT / "data/verified/c001/spec.vcl").read_text(encoding="utf-8")
        lines = base.split("\n")
        i = next(k for k, line in enumerate(lines) if line.startswith("OP") and ")" in line)

        def edit(old, new):
            return "\n".join(lines[:i] + [lines[i].replace(old, new, 1)] + lines[i + 1:])

        broken = {
            "a missing )": edit(")", ""),
            "an unterminated string": edit('"', ""),
            "keys out of order": "\n".join(lines[:2] + [x for x in lines[2:] if x.startswith("PROC")]
                                           + [x for x in lines[2:] if not x.startswith("PROC")]),
            "no ERR": "\n".join(x for x in lines if not x.startswith("ERR")),
            "no TEST": "\n".join(x for x in lines if not x.startswith("TEST")),
            "no GOAL": "\n".join(x for x in lines if not x.startswith("GOAL")),
            "an OP two levels too deep": edit("OP       ", "OP           "),
            "no final newline": base.rstrip("\n"),
            "a code fence": "```vcl\n" + base + "```\n",
            "a blank line": "\n".join(lines[:5] + [""] + lines[5:]),
            "an old version line": base.replace("VCL 0.6", "VCL 0.5", 1),
            "talk before the recipe": "Here is the recipe:\n" + base,
            "nothing": "",
        }
        for what, text in broken.items():
            with self.subTest(what):
                self.assertFalse(self.accepts(text))

    def test_the_goal_has_the_length_the_validator_wants(self):
        base = (ROOT / "data/verified/c001/spec.vcl").read_text(encoding="utf-8")
        goal = next(x for x in base.split("\n") if x.startswith("GOAL"))
        for n, ok in ((9, False), (10, True), (200, True), (201, False)):
            with self.subTest(length=n):
                self.assertEqual(self.accepts(base.replace(goal, "GOAL     " + "x" * n, 1)), ok)

    def test_steps_nest_one_level_at_a_time(self):
        loop = "LOOP_ROWS var=r from=2 to=9"
        self.assertTrue(self.accepts(task((0, loop), (1, "LET(x, 1)"))))
        self.assertTrue(self.accepts(task((0, loop), (1, loop), (2, "LET(x, 1)"), (1, "LET(y, 2)"), (0, "LET(z, 3)"))))
        self.assertFalse(self.accepts(task((0, loop), (2, "LET(x, 1)"))), "skips a level")
        self.assertFalse(self.accepts(task((0, loop))), "an opener needs a body")
        self.assertFalse(self.accepts(task((1, "LET(x, 1)"))), "the first OP is at depth 0")

    def test_if_and_else_pair_up(self):
        self.assertTrue(self.accepts(task((0, "IF x > 1"), (1, "LET(a, 1)"), (0, "ELSE"), (1, "LET(a, 2)"))))
        self.assertTrue(self.accepts(task((0, "IF x > 1 THEN LET(a, 1)"), (0, "LET(b, 2)"))))
        self.assertFalse(self.accepts(task((0, "ELSE"), (1, "LET(a, 2)"))), "ELSE with no IF")
        self.assertFalse(self.accepts(task((0, "IF x > 1"), (1, "LET(a, 1)"), (0, "ELSE"), (1, "LET(a, 2)"),
                                           (0, "ELSE"), (1, "LET(a, 3)"))), "two ELSEs")

    def test_cases_belong_to_a_select_case(self):
        ok = task((0, "SELECT_CASE x"), (1, "CASE 1"), (2, "LET(a, 1)"), (1, "CASE 2, 3"), (1, "CASE_ELSE"),
                  (2, "LET(a, 0)"))
        self.assertTrue(self.accepts(ok))
        self.assertFalse(self.accepts(task((0, "CASE 1"), (1, "LET(a, 1)"))), "CASE outside SELECT_CASE")
        self.assertFalse(self.accepts(task((0, "SELECT_CASE x"), (1, "LET(a, 1)"))), "a step directly under SELECT_CASE")
        # CASE after CASE_ELSE is not refused here: with an optional space between tokens, CASE_ELSE can also be read as
        # CASE and the name _ELSE. The validator refuses that order (V115), as it refuses any word joined to the next.

    def test_nesting_stops_at_the_bound(self):
        deepest = gg.MAX_DEPTH - 1
        ops = [(d, "LOOP_ROWS var=r from=2 to=9") for d in range(deepest)] + [(deepest, "LET(x, 1)")]
        self.assertTrue(self.accepts(task(*ops)))
        ops = [(d, "LOOP_ROWS var=r from=2 to=9") for d in range(deepest + 1)] + [(deepest + 1, "LET(x, 1)")]
        self.assertFalse(self.accepts(task(*ops)))

    def test_a_build_repeats_its_own_version(self):
        text = (ROOT / "vcl/examples/example4_build.vcl").read_text(encoding="utf-8")
        self.assertTrue(self.accepts(text))
        i = text.index("---\nVCL 0.7")
        self.assertFalse(self.accepts(text[:i] + text[i:].replace("VCL 0.7", "VCL 0.6", 1)))

    def test_a_question_is_not_a_recipe(self):
        self.assertFalse(self.accepts("# QUESTION: which sheet?\n"))
        with_question = gc.Grammar(gg.generate(questions=True))
        self.assertTrue(with_question.accepts("# QUESTION: which sheet?\n"))

    def test_it_means_for_a_line_what_grammar_lark_means(self):
        """Mutate real lines; the grammar and the real parser agree on nearly all, and the grammar never refuses a
        canonical line. They differ only in spacing (the grammar allows one space, canonical form) and in two words
        written without a space between them, which the validator catches."""
        random.seed(11)
        values = []
        for path in spec_files()[::3]:
            for line in path.read_text(encoding="utf-8").split("\n"):
                key = line.split(" ", 1)[0]
                if key in vcl_parse.START and " " in line:
                    values.append((key, line.split(None, 1)[1].strip()))
        agree = total = 0
        for key, value in values:
            start = gg.rule_name(vcl_parse.START[key])
            self.assertTrue(self.g.accepts(value, start=start), f"{key} {value}")
            for _ in range(2):
                i = random.randrange(len(value))
                mutant = value[:i] + value[i + 1:] if random.random() < 0.5 else value[:i] + random.choice('()",=[] &') + value[i:]
                try:
                    vcl_parse.parser().parse(mutant, start=vcl_parse.START[key])
                    lark = True
                except Exception:
                    lark = False
                total += 1
                agree += lark == self.g.accepts(mutant, start=start)
        self.assertGreater(total, 1000)
        self.assertGreater(agree / total, 0.95)


if __name__ == "__main__":
    unittest.main()
