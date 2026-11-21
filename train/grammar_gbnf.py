"""VCL's grammar as a GBNF grammar for the offline runner (D127; plan `docs/ai-plan.md`, Stage C).

    python train/grammar_gbnf.py            write train/vcl.gbnf
    python train/grammar_gbnf.py --check    exit 1 if train/vcl.gbnf is not what the sources give today
    python train/grammar_gbnf.py -o FILE    write somewhere else

From March 2027 the model's reply is generated with this grammar (llama.cpp `--grammar-file`), so it can only be text
that has the shape of a VCL recipe: no chat, no code fence, balanced brackets, keys in order, steps indented the way
the parser wants. The validator, the type checker and the TESTs still decide whether the recipe is right; the grammar
only keeps the reply from being unusable.

Two sources, never copied by hand:
- `vcl/grammar.lark` gives every line's value (what follows a key, or an OP's statement). It is read by the small lark
  reader below and converted rule by rule: the EBNF is kept as it is (no left recursion appears, which GBNF forbids),
  and a lark terminal that is a regex is translated, with its look-ahead dropped (GBNF has none: the `=` that follows an
  option key is in the rule anyway).
- `compiler/vcl_parse.py` gives the file's shape: the versions, the key order of a TASK, a BUILD and a DOCUMENT. The
  parser enforces these in Python, not in the grammar. How many of each key (the validator's counts) is the table
  here; the converter stops if a key is in the parser and not in the table.

What the grammar is, exactly:
- the model's *canonical* form: key names padded to column 9, an OP's value indented 2 spaces per level (the lines
  `vclc fmt` writes), `# ASSUME:` lines first, one final newline. Inside a line a space is optional between tokens
  (`sp`), at most one, so the grammar is looser than the formatter and KRONOS runs `vclc fmt` on the result anyway;
- slightly stricter than the parser where the validator is stricter: a TASK has GOAL, PROC, KIND, MODULE, an OP and
  ERR; a top-level TASK or DOCUMENT has FIXTURE and TEST lines; a block step has a body; CASE only inside SELECT_CASE;
  a BUILD repeats its own version line in every TASK;
- bounded: steps nest at most MAX_DEPTH levels (the corpus uses 6).
It does not list op names or option names (the validator's catalogue does): the grammar says syntax only, as
`vcl/grammar.lark` does.
"""
import argparse
import ast
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(ROOT / "compiler"))
import vcl_parse  # noqa: E402

OUT = HERE / "vcl.gbnf"
VERSIONS = ("0.6", "0.7")   # what the model writes; the parser reads 0.1 to 0.7
MAX_DEPTH = 10              # levels of OP nesting (0 to 9)
KEY_WIDTH = 9               # `vclc fmt`: key names padded to column 9

# How many of each key (the validator's rules, compiler/validator.py check_keys and BUILD/DOCUMENT checks).
# "1" exactly one, "?" at most one, "*" any, "+" at least one; the TASK inside a BUILD has no FIXTURE or TEST.
TASK_COUNT = {"GOAL": "1", "PROC": "1", "KIND": "1", "MODULE": "1", "ARGS": "?", "RET": "?", "VAR": "*", "PRE": "?",
              "MODIFIES": "*", "OP": "1", "PERF": "?", "ERR": "1", "ALLOW": "?", "FIXTURE": "+", "TEST": "+"}
BUILD_COUNT = {"GOAL": "1", "NAME": "1", "TARGET": "1", "ENTRY": "1", "MODULE": "+", "FORM": "*", "SIG": "+",
               "FLOW": "+", "FIXTURE": "+", "TEST": "+"}
DOC_COUNT = {"GOAL": "1", "NAME": "1", "TEMPLATE": "1", "ARGS": "?", "FIELD": "*", "REPEAT": "*", "KEEP": "*",
             "FIXTURE": "+", "TEST": "+"}
# the statements an OP line can hold, by what they do to the nesting (checked against the grammar's `statement`)
LEAF = ["if_inline", "exit_if", "skip_if", "exit_loop_if", "note", "call"]
LOOP = ["loop_rows", "loop_cols", "for_each"]
OPENING = {"if_block", "else_line", "select_case", "case_line", "case_else"}


class GrammarError(Exception):
    pass


# ---------------------------------------------------------------- reading grammar.lark

TOKEN = re.compile(r"""
    (?P<ws>[ \t]+)
  | (?P<str>"(?:[^"\\]|\\.)*")
  | (?P<re>/(?:[^/\\\n]|\\.)+/)
  | (?P<alias>->)
  | (?P<name>[A-Za-z_][A-Za-z0-9_]*)
  | (?P<punct>[()\[\]*+?|])
""", re.X)
DEFINITION = re.compile(r"^(\??)([A-Za-z_][A-Za-z0-9_]*)(?:\.\d+)?:(.*)$")


def strip_comment(line):
    """The line without its // comment (a // inside a string or a regex is not one)."""
    i = 0
    while i < len(line):
        if line.startswith("//", i):
            return line[:i]
        m = TOKEN.match(line, i)
        i = m.end() if m else i + 1   # the `name:` of a definition is not a token: step over it
    return line


def tokenize(text):
    out, i = [], 0
    while i < len(text):
        m = TOKEN.match(text, i)
        if not m:
            raise GrammarError(f"cannot read {text[i:i + 20]!r}")
        i = m.end()
        if m.lastgroup != "ws":
            out.append((m.lastgroup, m.group()))
    return out


class _Body:
    """The body of one lark rule or terminal as a tree: seq, alt, opt, star, plus, ref, str, re."""

    def __init__(self, tokens):
        self.t = tokens
        self.i = 0

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def alts(self):
        branches = [self.seq()]
        while self.peek()[1] == "|":
            self.i += 1
            branches.append(self.seq())
        return branches[0] if len(branches) == 1 else ("alt", branches)

    def seq(self):
        items = []
        while True:
            kind, value = self.peek()
            if kind is None or value in ("|", ")", "]"):
                break
            if kind == "alias":
                self.i += 2   # `-> name`: a tree name, no syntax
                continue
            items.append(self.item())
        return items[0] if len(items) == 1 else ("seq", items)

    def item(self):
        node = self.atom()
        while self.peek()[1] in ("*", "+", "?"):
            node = ({"*": "star", "+": "plus", "?": "opt"}[self.peek()[1]], node)
            self.i += 1
        return node

    def atom(self):
        kind, value = self.peek()
        self.i += 1
        if kind == "str":
            return ("str", ast.literal_eval(value))
        if kind == "re":
            return ("re", value[1:-1])
        if kind == "name":
            return ("ref", value)
        if value == "(":
            node = self.alts()
            self.i += 1
            return node
        if value == "[":   # [x] is x? (the parser runs with maybe_placeholders=False)
            node = self.alts()
            self.i += 1
            return ("opt", node)
        raise GrammarError(f"unexpected {value!r}")


def read_lark(text):
    """({name: tree} for the rules, {name: tree} for the terminals, the %ignore pattern)."""
    chunks, ignore = [], None
    for physical in text.split("\n"):
        if physical.startswith("%ignore"):
            ignore = physical.split(None, 1)[1].split("//")[0].strip()
            continue
        raw = strip_comment(physical).rstrip()
        stripped = raw.strip()
        if not stripped:
            continue
        m = DEFINITION.match(raw)
        if m and not raw.startswith((" ", "\t", "|")):
            chunks.append([m.group(2), m.group(3)])
        elif chunks:
            chunks[-1][1] += " " + raw
        else:
            raise GrammarError(f"text before the first rule: {raw[:30]!r}")
    rules, terminals = {}, {}
    for name, body in chunks:
        tree = _Body(tokenize(body))
        node = tree.alts()
        if tree.i != len(tree.t):
            raise GrammarError(f"{name}: cannot read all of its body")
        (terminals if re.fullmatch(r"[A-Z][A-Z0-9_]*", name) else rules)[name] = node
    return rules, terminals, ignore


# ---------------------------------------------------------------- a lark regex as GBNF

def _lit(text):
    out = []
    for c in text:
        out.append({'"': '\\"', "\\": "\\\\", "\n": "\\n", "\r": "\\r", "\t": "\\t"}.get(c, c))
    return '"' + "".join(out) + '"'


def _class_char(c):
    return {"]": "\\]", "[": "\\[", "\\": "\\\\", "^": "\\^", "-": "\\-", "\n": "\\n", "\r": "\\r", "\t": "\\t"}.get(c, c)


class _Regex:
    """The regex subset grammar.lark uses -> GBNF: literals, classes, groups, alternation, * + ? {m} {m,n}, lazy
    quantifiers read as greedy, and a look-ahead (?=...) dropped."""

    def __init__(self, pattern):
        self.p = pattern
        self.i = 0

    def parse(self):
        out = self.alt()
        if self.i != len(self.p):
            raise GrammarError(f"regex {self.p!r}: stopped at {self.p[self.i:]!r}")
        return out

    def alt(self):
        branches = [self.seq()]
        while self.i < len(self.p) and self.p[self.i] == "|":
            self.i += 1
            branches.append(self.seq())
        return " | ".join(b for b in branches if b != "") if len(branches) > 1 else branches[0]

    def seq(self):
        atoms = []   # (gbnf text, is a plain character, the character)
        while self.i < len(self.p) and self.p[self.i] not in "|)":
            atom = self.atom()
            if atom is None:   # a dropped look-ahead
                continue
            quant = self.quantifier()
            if quant:
                atoms.append((atom[0] + quant, False, ""))
            else:
                atoms.append(atom)
        out, run = [], ""
        for text, plain, char in atoms:
            if plain:
                run += char
                continue
            if run:
                out.append(_lit(run))
                run = ""
            out.append(text)
        if run:
            out.append(_lit(run))
        return " ".join(out)

    def quantifier(self):
        if self.i >= len(self.p):
            return ""
        c = self.p[self.i]
        if c in "*+?":
            self.i += 1
        elif c == "{":
            end = self.p.index("}", self.i)
            c = self.p[self.i:end + 1]
            self.i = end + 1
        else:
            return ""
        if self.i < len(self.p) and self.p[self.i] == "?":   # lazy: greedy is a superset
            self.i += 1
        return c

    def atom(self):
        c = self.p[self.i]
        if c == "(":
            if self.p.startswith("(?=", self.i):
                depth, j = 0, self.i
                while True:
                    depth += {"(": 1, ")": -1}.get(self.p[j], 0) if self.p[j - 1] != "\\" else 0
                    j += 1
                    if depth == 0:
                        break
                self.i = j
                return None
            self.i += 3 if self.p.startswith("(?:", self.i) else 1
            inner = self.alt()
            if self.i >= len(self.p) or self.p[self.i] != ")":
                raise GrammarError(f"regex {self.p!r}: missing )")
            self.i += 1
            return ("(" + inner + ")", False, "")
        if c == "[":
            return self.char_class()
        if c == "\\":
            return self.escape()
        if c in ".$^":
            raise GrammarError(f"regex {self.p!r}: {c!r} is not supported")
        self.i += 1
        return (_lit(c), True, c)

    def escape(self):
        c = self.p[self.i + 1]
        self.i += 2
        if c == "d":
            return ("[0-9]", False, "")
        if c in "rnt":
            ch = {"r": "\r", "n": "\n", "t": "\t"}[c]
            return (_lit(ch), True, ch)
        if c.isalnum():
            raise GrammarError(f"regex {self.p!r}: \\{c} is not supported")
        return (_lit(c), True, c)

    def char_class(self):
        self.i += 1
        negate = self.p[self.i] == "^"
        self.i += negate
        parts = []
        while self.p[self.i] != "]":
            if self.p[self.i] == "\\":
                e = self.p[self.i + 1]
                self.i += 2
                if e == "d":
                    parts.append("0-9")
                    continue
                lo = {"r": "\r", "n": "\n", "t": "\t"}.get(e, e)
            else:
                lo, self.i = self.p[self.i], self.i + 1
            if self.p[self.i] == "-" and self.p[self.i + 1] != "]":
                hi = self.p[self.i + 1]
                self.i += 2
                parts.append(_class_char(lo) + "-" + _class_char(hi))
            else:
                parts.append(_class_char(lo))
        self.i += 1
        if negate:   # a value is one line: no negated class may run over a line end
            parts += ["\\n", "\\r"] if "\\n" not in parts else []
        return ("[" + ("^" if negate else "") + "".join(parts) + "]", False, "")


# ---------------------------------------------------------------- lark trees as GBNF

def rule_name(name):
    return name.replace("_", "-")


def terminal_name(name):
    return "t-" + name.lower().replace("_", "-")


class Converter:
    SP = "sp"   # at most one space between tokens

    def __init__(self, rules, terminals):
        self.rules, self.terminals = rules, terminals
        self.nullable = self._nullable_rules()

    # nullability, to be sure no sequence starts with something that can be empty (a space would dangle)
    def _nullable_rules(self):
        null = set()
        changed = True
        while changed:
            changed = False
            for name, node in self.rules.items():
                if name not in null and self.is_nullable(node, null):
                    null.add(name)
                    changed = True
        return null

    def is_nullable(self, node, null=None):
        null = self.nullable if null is None else null
        kind = node[0]
        if kind in ("str", "re"):
            return False
        if kind == "ref":
            return node[1] in null
        if kind == "seq":
            return all(self.is_nullable(c, null) for c in node[1])
        if kind == "alt":
            return any(self.is_nullable(c, null) for c in node[1])
        return kind in ("opt", "star") or self.is_nullable(node[1], null)

    def atom(self, node):
        kind = node[0]
        if kind == "str":
            return _lit(node[1])
        if kind == "ref":
            name = node[1]
            if name in self.terminals:
                return terminal_name(name)
            if name in self.rules:
                return rule_name(name)
            raise GrammarError(f"{name} is not defined")
        raise AssertionError(kind)

    def emit(self, node, lead):
        """GBNF for node. With lead, `sp` goes before it, but only when it matches something, so a group that can be
        empty has its `sp` inside."""
        kind = node[0]
        if kind in ("str", "ref"):
            return (self.SP + " " if lead else "") + self.atom(node)
        if kind == "seq":
            items = node[1]
            if not lead and self.is_nullable(items[0]):
                raise GrammarError("a sequence starts with something that can be empty")
            return " ".join(self.emit(item, lead or i > 0) for i, item in enumerate(items))
        if kind == "alt":
            return "(" + " | ".join(self.emit(c, lead) for c in node[1]) + ")"
        if kind == "opt":
            return "(" + self.emit(node[1], lead) + ")?"
        if kind == "star":
            if lead:
                return "(" + self.emit(node[1], True) + ")*"
            return "(" + self.emit(node[1], False) + " (" + self.emit(node[1], True) + ")*)?"
        if kind == "plus":
            if lead:
                return "(" + self.emit(node[1], True) + ")+"
            return "(" + self.emit(node[1], False) + ") (" + self.emit(node[1], True) + ")*"
        raise GrammarError(f"cannot convert {kind}")

    def rule_line(self, name, node):
        if node[0] == "alt":
            return f"{rule_name(name)} ::= " + " | ".join(self.emit(c, False) for c in node[1])
        return f"{rule_name(name)} ::= {self.emit(node, False)}"

    def terminal_line(self, name, node):
        return f"{terminal_name(name)} ::= {self.terminal_body(node)}"

    def terminal_body(self, node):
        kind = node[0]
        if kind == "str":
            return _lit(node[1])
        if kind == "re":
            return _Regex(node[1]).parse()
        if kind == "alt":
            return " | ".join(self.terminal_body(c) for c in node[1])
        raise GrammarError(f"cannot convert a terminal made of {kind}")


# ---------------------------------------------------------------- the file's shape

def pad(key, extra=0):
    return _lit(" " * (KEY_WIDTH - len(key) + extra))


def key_rules(keys_in_order):
    """The `<key>-line` rules of every key the parser reads (OP and GOAL are written separately)."""
    out, seen = [], set()
    for key in keys_in_order:
        if key in ("GOAL", "OP") or key in seen:
            continue
        seen.add(key)
        out.append(f'{key.lower()}-line ::= "{key}" {pad(key)} {rule_name(vcl_parse.START[key])} nl')
    return out


def sequence(order, counts, skip=()):
    parts = []
    for key in order:
        if key in skip:
            continue
        ref = "op-block" if key == "OP" else f"{key.lower()}-line"
        parts.append(ref + ("" if counts[key] == "1" else counts[key]))
    return " ".join(parts)


def check_tables(rules):
    for name, order, counts in (("TASK", vcl_parse.TASK_ORDER, TASK_COUNT), ("BUILD", vcl_parse.BUILD_ORDER, BUILD_COUNT),
                                ("DOCUMENT", vcl_parse.DOC_ORDER, DOC_COUNT)):
        if set(order) != set(counts):
            raise GrammarError(f"the {name} keys in vcl_parse ({sorted(order)}) and in this table ({sorted(counts)}) differ")
    missing = [k for k in set(vcl_parse.TASK_ORDER + vcl_parse.BUILD_ORDER + vcl_parse.DOC_ORDER)
               if k not in ("GOAL", "OP") and k not in vcl_parse.START]
    if missing:
        raise GrammarError(f"no start rule in vcl_parse.START for {sorted(missing)}")
    statements = set(_names(rules["statement"]))
    if statements != set(LEAF + LOOP) | OPENING:
        raise GrammarError(f"the grammar's statement kinds changed: {sorted(statements ^ (set(LEAF + LOOP) | OPENING))}")


def _names(node):
    if node[0] == "ref":
        yield node[1]
    elif node[0] in ("seq", "alt"):
        for c in node[1]:
            yield from _names(c)
    elif node[0] in ("opt", "star", "plus"):
        yield from _names(node[1])


def op_rules(max_depth):
    """The OP steps, one set of rules per nesting depth, so a block opener is followed by a deeper, non-empty body.
    CASE and CASE_ELSE sit one level under their SELECT_CASE, and their bodies one level under them."""
    def line(d, what):
        return f'"OP" {_lit(" " * (KEY_WIDTH - 2 + 2 * d))} {what} nl'

    leaf = "(" + " | ".join(rule_name(n) for n in LEAF) + ")"
    loops = "(" + " | ".join(rule_name(n) for n in LOOP) + ")"
    out = ["op-block ::= stmt-0+"]
    for d in range(max_depth):
        out.append(f"leaf-{d} ::= {line(d, leaf)}")
        kinds = [f"leaf-{d}"]
        if d + 1 < max_depth:   # a block opener, with its body one level down
            kinds += [f"loop-{d}", f"if-{d}"]
            out.append(f"block-{d + 1} ::= stmt-{d + 1}+")
            out.append(f"loop-{d} ::= {line(d, loops)} block-{d + 1}")
            out.append(f"if-{d} ::= {line(d, 'if-block')} block-{d + 1} ({line(d, 'else-line')} block-{d + 1})?")
        if d + 2 < max_depth:   # SELECT_CASE: its cases, and their bodies two levels down
            kinds.append(f"select-{d}")
            out.append(f"select-{d} ::= {line(d, 'select-case')} (case-{d + 1}+ case-else-{d + 1}? | case-else-{d + 1})")
            out.append(f"case-{d + 1} ::= {line(d + 1, 'case-line')} stmt-{d + 2}*")
            out.append(f"case-else-{d + 1} ::= {line(d + 1, 'case-else')} block-{d + 2}")
        out.append(f"stmt-{d} ::= " + " | ".join(kinds))
    return out


def generate(versions=VERSIONS, max_depth=MAX_DEPTH, questions=False):
    """The grammar as GBNF text."""
    rules, terminals, ignore = read_lark(vcl_parse.GRAMMAR.read_text(encoding="utf-8"))
    if ignore != r"/[ \t]+/":
        raise GrammarError(f"grammar.lark ignores {ignore}, not spaces and tabs: the converter's `sp` is wrong")
    check_tables(rules)
    conv = Converter(rules, terminals)
    out = ["# VCL as a GBNF grammar for llama.cpp (--grammar-file). Written by train/grammar_gbnf.py from vcl/grammar.lark",
           "# and compiler/vcl_parse.py: do not edit; run `python train/grammar_gbnf.py`. See that file for what it is.", ""]

    # the file
    ver = " | ".join(_lit(f"VCL {v}") for v in versions)
    files = ["task-file", "build-file"] + (["doc-file"] if "0.7" in versions else []) + (["question"] if questions else [])
    out += ["root ::= " + " | ".join(files), "nl ::= \"\\n\"", "sp ::= \" \"?",
            'assume-line ::= "# ASSUME: " [^\\n]+ nl',
            f'goal-line ::= "GOAL" {pad("GOAL")} [ -~]{{10,200}} nl']
    if questions:
        out.append('question ::= "# QUESTION: " [^\\n]+ nl')
    out += key_rules(vcl_parse.TASK_ORDER + vcl_parse.BUILD_ORDER + vcl_parse.DOC_ORDER)
    task_head = sequence(vcl_parse.TASK_ORDER, TASK_COUNT, skip=("FIXTURE", "TEST"))
    out += [f"task-head ::= {task_head}",
            f"task-file ::= ({ver}) nl \"TASK\" nl assume-line* task-head fixture-line+ test-line+"]
    build_head = sequence(vcl_parse.BUILD_ORDER, BUILD_COUNT)
    builds = []
    for v in versions:
        if v < "0.5":
            continue
        tag = v.replace(".", "")
        builds.append(f"build-{tag}")
        out.append(f'build-{tag} ::= "VCL {v}" nl "BUILD" nl assume-line* {build_head} '
                   f'("---" nl "VCL {v}" nl "TASK" nl assume-line* task-head)+')
    out.append("build-file ::= " + " | ".join(builds))
    if "0.7" in versions:
        out.append(f'doc-file ::= "VCL 0.7" nl "DOCUMENT" nl assume-line* {sequence(vcl_parse.DOC_ORDER, DOC_COUNT)}')
    out += op_rules(max_depth)

    # the lines
    out.append("")
    out.append("# the value of each line: vcl/grammar.lark")
    for name, node in rules.items():
        out.append(conv.rule_line(name, node))
    out.append("")
    out.append("# the terminals")
    for name, node in terminals.items():
        out.append(conv.terminal_line(name, node))
    names = [line.split(" ::= ")[0] for line in out if " ::= " in line]
    dup = sorted({n for n in names if names.count(n) > 1})
    if dup:
        raise GrammarError(f"rule names used twice: {dup}")
    return "\n".join(out) + "\n"


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("-o", "--out", default=str(OUT))
    ap.add_argument("--check", action="store_true", help="exit 1 when the file is not what the sources give")
    args = ap.parse_args(argv)
    text = generate()
    path = Path(args.out)
    if args.check:
        if not path.is_file() or path.read_text(encoding="utf-8") != text:
            print(f"{path} is stale: run python train/grammar_gbnf.py")
            return 1
        print(f"{path} is up to date")
        return 0
    path.write_text(text, encoding="utf-8")
    print(f"wrote {path} ({len(text.splitlines())} lines)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
