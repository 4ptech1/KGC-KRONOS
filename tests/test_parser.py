"""Unit tests for compiler/parser.py. Run: python -m unittest discover tests"""
import os
import random
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "compiler"))

from parser import parse_spec, tokenize  # noqa: E402


def read(path):
    with open(os.path.join(ROOT, path), encoding="utf-8") as f:
        return f.read()


MINIMAL_BUILD = """VCL 0.1
BUILD
# ASSUME: Data headers (row 1) include "ID" and "Status"
NAME     StatusCounter
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modFormat
SIG      modMain.Run()
SIG      modMain.RunHeadless(ws As Worksheet) As Long
SIG      modFormat.BoldHeader(ws As Worksheet)
FLOW     modMain.Run -> modMain.RunHeadless -> modFormat.BoldHeader
FIXTURE  demo_01.xlsx sheet=Data cols=A:ID,B:Status rows=10 data=A:seq(1,1);B:values("Open"x6,"Closed"x4)
TEST     fixture=demo_01.xlsx call=modMain.RunHeadless(Data) runs=2 assert=RESULT = PRE_COUNTIF(Data![Status], "Open") AND ERROR_COUNT = 0
TEST     fixture=demo_01.xlsx call=modMain.RunHeadless(Data) assert=COUNTA(Data![Status]) = PRE_COUNTA(Data![Status])
---
VCL 0.1
TASK
PROC     Run
KIND     Sub
MODULE   modMain
OP       CALL(modMain.RunHeadless, ACTIVE_SHEET)
ERR      standard
---
VCL 0.1
TASK
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     ws As Worksheet
RET      Long
VAR      cStatus As Long
VAR      found As Long
PRE      COL($ws, "Status") > 0
OP       LET(cStatus, COL($ws, "Status"))
OP       LET(found, 0)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cStatus) var=r
OP         IF UPPER(TRIM(CELL($ws, r, cStatus))) = "OPEN" THEN LET(found, found + 1)
OP       CALL(modFormat.BoldHeader, $ws)
OP       RETURN(found)
ERR      standard
---
VCL 0.1
TASK
PROC     BoldHeader
KIND     Sub
MODULE   modFormat
ARGS     ws As Worksheet
OP       FORMAT(ROW_RANGE($ws, 1, 1, LASTCOL($ws, 1)), bold=yes)
ERR      standard
"""


class ParseTask(unittest.TestCase):
    def setUp(self):
        self.spec = parse_spec(read("specs/task_01.vcl"))

    def test_header_and_assumptions(self):
        self.assertEqual(self.spec.kind, "TASK")
        self.assertEqual(self.spec.problems, [])
        self.assertEqual(len(self.spec.assumes), 5)
        self.assertEqual(self.spec.assumes[0][0], 3)

    def test_keys_args_and_vars(self):
        self.assertEqual(self.spec.value("PROC"), "BuildShortageList")
        self.assertEqual([(d.name, d.type) for d in self.spec.args],
                         [("bom", "Worksheet"), ("grn", "Worksheet"), ("out", "Worksheet")])
        self.assertEqual(len(self.spec.vars), 8)
        self.assertEqual(self.spec.vars[0].name, "received")

    def test_op_lines_form_a_tree(self):
        tops = self.spec.ops
        self.assertEqual(len(tops), 14)
        loop = next(n for n in tops if n.text.startswith("LOOP_ROWS") and n.line == 35)
        block_if = loop.children[0]
        self.assertTrue(block_if.text.startswith("IF NOT IS_EMPTY"))
        texts = [c.text.split("(")[0].split()[0] for c in block_if.children]
        self.assertEqual(texts, ["LET", "LET", "IF", "IF", "ELSE"])
        self.assertEqual(len(block_if.children[4].children), 1)  # ELSE holds the nested IF
        self.assertEqual(sum(1 for _ in self.spec.all_ops()), 33)

    def test_fixtures(self):
        bom, grn, short = self.spec.fixtures
        self.assertEqual((bom.file, bom.sheet, bom.rows), ("bom_grn_01.xlsx", "BOM", 40))
        self.assertEqual(bom.headers, [("A", "Part No"), ("B", "Description"), ("C", "Required Qty")])
        self.assertEqual([(r.col, r.kind, r.args, r.blank) for r in bom.rules],
                         [("A", "seq", ["1001", "1"], 0), ("B", "text", ["12"], 0), ("C", "int", ["1", "50"], 3)])
        self.assertEqual((short.rows, short.rules), (0, []))

    def test_tests(self):
        first = self.spec.tests[0]
        self.assertEqual(first.fixture, "bom_grn_01.xlsx")
        self.assertEqual(first.procedure, "BuildShortageList")
        self.assertEqual(first.args, ["BOM", "GRN", "Shortages"])
        self.assertEqual(first.runs, 2)
        self.assertTrue(first.assertion.startswith("RESULT >= 5 AND"))
        self.assertEqual(self.spec.tests[1].runs, 1)

    def test_values_rule(self):
        spec = parse_spec(read("specs/task_03.vcl"))
        rule = spec.fixtures[0].rules[1]
        self.assertEqual(rule.args, [("W-01", 20), ("W-02", 20), ("W-03", 15), ("", 5)])


class ParseOtherKinds(unittest.TestCase):
    def test_question_reply(self):
        spec = parse_spec("# QUESTION: Which folder should I clean?\n")
        self.assertEqual(spec.kind, "QUESTION")
        self.assertEqual(spec.question, "Which folder should I clean?")
        self.assertEqual(spec.problems, [])

    def test_build_sections(self):
        spec = parse_spec(MINIMAL_BUILD)
        self.assertEqual(spec.kind, "BUILD")
        self.assertEqual(spec.problems, [])
        self.assertEqual([f"{s.module}.{s.proc}" for s in spec.sigs],
                         ["modMain.Run", "modMain.RunHeadless", "modFormat.BoldHeader"])
        self.assertEqual(spec.sigs[1].ret, "Long")
        self.assertEqual([t.value("PROC") for t in spec.tasks], ["Run", "RunHeadless", "BoldHeader"])
        self.assertEqual(spec.tasks[1].line, 26)  # line numbers stay those of the whole file
        self.assertEqual(spec.flows[0][1], ["modMain.Run", "modMain.RunHeadless", "modFormat.BoldHeader"])


class ParseBadInput(unittest.TestCase):
    def test_problems_carry_line_numbers(self):
        text = read("specs/task_07.vcl").replace("VAR      status As String", "VAR      status as String")
        spec = parse_spec(text)
        self.assertEqual([(p.line, p.severity) for p in spec.problems], [(17, "error")])
        self.assertIn("expected '<name> As <Type>'", spec.problems[0].message)

    def test_bad_nesting_is_reported(self):
        text = read("specs/task_02.vcl").replace("OP         IF NOT IS_EMPTY", "OP           IF NOT IS_EMPTY")
        spec = parse_spec(text)
        self.assertTrue(any("bad nesting" in p.message and p.line == 24 for p in spec.problems))

    def test_never_raises(self):
        samples = [None, "", "\n\n", "VCL 0.1", "VCL 0.1\nTASK", "VCL 0.1\nTASK\nOP", "VCL 0.1\nBUILD\n---\n---",
                   "﻿VCL 0.1\nTASK\nPROC X", "VCL 0.1\nTASK\nFIXTURE a.xlsx sheet=S cols=A rows=x data=A:",
                   "VCL 0.1\nTASK\nTEST fixture= call=( assert=", "VCL 0.1\nTASK\n  OP x\n\tOP y",
                   "VCL 0.1\nTASK\nOP       IF\nOP             X\nOP   Y"]
        for text in samples:
            with self.subTest(text=text):
                parse_spec(text)

    def test_random_damage_never_raises(self):
        rng = random.Random(7)
        texts = [read(f"specs/task_{i:02d}.vcl") for i in range(1, 11)]
        pool = ' ()[],;=<>"$#!:+-*/&.\t\nabcXYZ019_'
        for _ in range(300):
            chars = list(rng.choice(texts))
            for _ in range(rng.randint(1, 8)):
                at = rng.randrange(len(chars))
                if rng.random() < 0.5:
                    del chars[at]
                else:
                    chars.insert(at, rng.choice(pool))
            parse_spec("".join(chars))


class Tokens(unittest.TestCase):
    def test_header_ranges(self):
        tokens, err = tokenize('VALUE(Orders![Order Date]2) >= 3')
        self.assertEqual(err, "")
        rng = tokens[2]
        self.assertEqual((rng.kind, rng.sheet, rng.header, rng.row), ("range", "Orders", "Order Date", "2"))

    def test_letter_ranges_and_colours(self):
        tokens, _ = tokenize("FILL_COUNT(Data!A:A, #FFCCCC)")
        self.assertEqual([t.kind for t in tokens], ["name", "op", "range", "op", "hex", "op"])

    def test_unclosed_quote(self):
        self.assertEqual(tokenize('LET(x, "abc)')[1], "unclosed quote")


if __name__ == "__main__":
    unittest.main()
