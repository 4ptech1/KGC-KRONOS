"""A small recognizer for the GBNF grammars `grammar_gbnf.py` writes (D127).

    python train/gbnf_check.py GRAMMAR.gbnf FILE...      which files the grammar accepts

GBNF is the grammar format llama.cpp (the offline runner planned for March 2027) uses to lock a model's output. This
module reads the grammar *text* and answers "does it accept this string?", so the converter's output is tested without
loading any model or any runner: it is an independent reader (it knows nothing of the converter or of lark).

It understands exactly the subset the converter writes, one rule per line: `name ::= alternatives`, "literals" with
escapes, [classes] (ranges, ^ negation, escapes), (groups), the postfix * + ? {m} {m,n} {m,}, and # comments. Anything
else is an error, so a grammar this module reads is one llama.cpp can read too. It also finds what llama.cpp would
refuse: a rule that is used and not defined, no `root`, left recursion.
"""
import re
import sys

NAME = re.compile(r"[a-zA-Z0-9-]+")
sys.setrecursionlimit(20000)


class GbnfError(Exception):
    pass


def _escape(text, i):
    """The character an escape at text[i] (the backslash) stands for, and the index after it."""
    c = text[i + 1]
    simple = {"n": "\n", "r": "\r", "t": "\t", "\\": "\\", '"': '"', "]": "]", "[": "[", "^": "^", "-": "-"}
    if c in simple:
        return simple[c], i + 2
    if c in "xuU":
        width = {"x": 2, "u": 4, "U": 8}[c]
        return chr(int(text[i + 2:i + 2 + width], 16)), i + 2 + width
    raise GbnfError(f"unknown escape \\{c}")


class _Parser:
    def __init__(self, text):
        self.s = text
        self.i = 0

    def space(self):
        while self.i < len(self.s) and self.s[self.i] in " \t":
            self.i += 1

    def peek(self):
        self.space()
        return self.s[self.i] if self.i < len(self.s) else ""

    def alt(self):
        branches = [self.seq()]
        while self.peek() == "|":
            self.i += 1
            branches.append(self.seq())
        return branches[0] if len(branches) == 1 else ("alt", branches)

    def seq(self):
        items = []
        while self.peek() not in ("", "|", ")"):
            items.append(self.item())
        return ("seq", items)

    def item(self):
        node = self.primary()
        while True:
            c = self.peek()
            if c == "*":
                node, self.i = ("rep", node, 0, None), self.i + 1
            elif c == "+":
                node, self.i = ("rep", node, 1, None), self.i + 1
            elif c == "?":
                node, self.i = ("rep", node, 0, 1), self.i + 1
            elif c == "{":
                m = re.compile(r"\{(\d+)(?:(,)(\d*))?\}").match(self.s, self.i)
                if not m:
                    raise GbnfError(f"bad repetition at {self.s[self.i:self.i + 12]!r}")
                lo = int(m.group(1))
                hi = lo if not m.group(2) else (int(m.group(3)) if m.group(3) else None)
                node, self.i = ("rep", node, lo, hi), m.end()
            else:
                return node

    def primary(self):
        c = self.peek()
        if c == '"':
            self.i += 1
            out = []
            while self.s[self.i] != '"':
                if self.s[self.i] == "\\":
                    ch, self.i = _escape(self.s, self.i)
                else:
                    ch, self.i = self.s[self.i], self.i + 1
                out.append(ch)
            self.i += 1
            return ("lit", "".join(out))
        if c == "[":
            return self.char_class()
        if c == "(":
            self.i += 1
            node = self.alt()
            if self.peek() != ")":
                raise GbnfError("missing )")
            self.i += 1
            return node
        m = NAME.match(self.s, self.i)
        if not m:
            raise GbnfError(f"cannot read {self.s[self.i:self.i + 12]!r}")
        self.i = m.end()
        return ("ref", m.group(0))

    def char_class(self):
        self.i += 1
        negate = self.s[self.i] == "^"
        self.i += negate
        ranges = []
        while self.s[self.i] != "]":
            lo, self.i = _escape(self.s, self.i) if self.s[self.i] == "\\" else (self.s[self.i], self.i + 1)
            if self.s[self.i] == "-" and self.s[self.i + 1] != "]":
                self.i += 1
                hi, self.i = _escape(self.s, self.i) if self.s[self.i] == "\\" else (self.s[self.i], self.i + 1)
                ranges.append((lo, hi))
            else:
                ranges.append((lo, lo))
        self.i += 1
        return ("cls", negate, tuple(ranges))


class Grammar:
    def __init__(self, text):
        self.rules = {}
        self.order = []
        for number, raw in enumerate(text.split("\n"), 1):
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            name, arrow, body = line.partition("::=")
            name = name.strip()
            if not arrow or not NAME.fullmatch(name):
                raise GbnfError(f"line {number}: not a rule: {line[:40]!r}")
            if name in self.rules:
                raise GbnfError(f"line {number}: rule {name} is defined twice")
            parser = _Parser(body)
            try:
                node = parser.alt()
                if parser.peek() != "":
                    raise GbnfError(f"unexpected {parser.s[parser.i:parser.i + 8]!r}")
            except (IndexError, GbnfError) as err:
                raise GbnfError(f"line {number} ({name}): {err}") from None
            self.rules[name] = node
            self.order.append(name)

    # ---- what llama.cpp would refuse

    def problems(self):
        out = []
        if "root" not in self.rules:
            out.append("there is no root rule")

        def refs(node):
            if node[0] == "ref":
                yield node[1]
            elif node[0] in ("seq", "alt"):
                for child in node[1]:
                    yield from refs(child)
            elif node[0] == "rep":
                yield from refs(node[1])

        for name, node in self.rules.items():
            for ref in sorted(set(refs(node))):
                if ref not in self.rules:
                    out.append(f"{name} uses {ref}, which is not defined")
        if out:
            return out
        nullable = self._nullable()
        left = {name: self._starts(node, nullable) for name, node in self.rules.items()}
        for name in self.rules:
            seen, stack = set(), list(left[name])
            while stack:
                n = stack.pop()
                if n == name:
                    out.append(f"left recursion in {name}")
                    break
                if n not in seen:
                    seen.add(n)
                    stack += left[n]
        return out

    def _nullable(self):
        nullable = set()
        changed = True
        while changed:
            changed = False
            for name, node in self.rules.items():
                if name not in nullable and self._is_nullable(node, nullable):
                    nullable.add(name)
                    changed = True
        return nullable

    def _is_nullable(self, node, nullable):
        kind = node[0]
        if kind == "lit":
            return node[1] == ""
        if kind == "cls":
            return False
        if kind == "ref":
            return node[1] in nullable
        if kind == "seq":
            return all(self._is_nullable(c, nullable) for c in node[1])
        if kind == "alt":
            return any(self._is_nullable(c, nullable) for c in node[1])
        return node[2] == 0 or self._is_nullable(node[1], nullable)

    def _starts(self, node, nullable):
        """The rules that can be the first thing node reads."""
        kind = node[0]
        if kind == "ref":
            return {node[1]}
        if kind == "seq":
            out = set()
            for child in node[1]:
                out |= self._starts(child, nullable)
                if not self._is_nullable(child, nullable):
                    break
            return out
        if kind == "alt":
            return set().union(*(self._starts(c, nullable) for c in node[1]))
        if kind == "rep":
            return self._starts(node[1], nullable)
        return set()

    # ---- does it accept this text

    def accepts(self, text, start="root"):
        return len(text) in self.ends(("ref", start), text, 0, {})

    def ends(self, node, s, i, memo):
        """Every index a match of node starting at s[i] can end at."""
        kind = node[0]
        if kind == "lit":
            lit = node[1]
            return {i + len(lit)} if s.startswith(lit, i) else set()
        if kind == "cls":
            if i >= len(s):
                return set()
            c = s[i]
            hit = any(lo <= c <= hi for lo, hi in node[2])
            return {i + 1} if hit != node[1] else set()
        if kind == "ref":
            key = (node[1], i)
            if key not in memo:
                memo[key] = frozenset(self.ends(self.rules[node[1]], s, i, memo))
            return memo[key]
        if kind == "seq":
            here = {i}
            for child in node[1]:
                nxt = set()
                for p in here:
                    nxt |= self.ends(child, s, p, memo)
                here = nxt
                if not here:
                    break
            return here
        if kind == "alt":
            out = set()
            for child in node[1]:
                out |= self.ends(child, s, i, memo)
            return out
        _, child, lo, hi = node
        out = {i} if lo == 0 else set()
        level, count, seen = {i}, 0, {i}
        while level and (hi is None or count < hi):
            nxt = set()
            for p in level:
                nxt |= self.ends(child, s, p, memo)
            count += 1
            if count >= lo:
                out |= nxt
                nxt -= seen
                seen |= nxt
            level = nxt
        return out


def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 2
    grammar = Grammar(open(argv[0], encoding="utf-8").read())
    found = grammar.problems()
    for line in found:
        print("PROBLEM:", line)
    bad = bool(found)
    for path in argv[1:]:
        ok = grammar.accepts(open(path, encoding="utf-8").read())
        print(("accepted " if ok else "REJECTED ") + path)
        bad |= not ok
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
