"""Parse VCL 0.1, 0.2 and 0.3 specs into Python objects.

    from parser import parse_spec
    spec = parse_spec(text)

parse_spec() never raises on bad input. Anything it can't read becomes a Problem with a
line number, and parsing carries on with the next line. This module only reads the format;
the rules (names, ops, tests, headers) are checked by validator.py.
"""
import datetime
import re
from dataclasses import dataclass, field

VERSION = "VCL 0.1"
VERSIONS = ("VCL 0.1", "VCL 0.2", "VCL 0.3")  # each version keeps every spec of the one before valid
TASK_KEYS = ["PROC", "KIND", "MODULE", "ARGS", "RET", "VAR", "PRE", "MODIFIES", "OP", "PERF", "ERR",
             "ALLOW", "FIXTURE", "TEST"]
BUILD_KEYS = ["NAME", "TARGET", "ENTRY", "MODULE", "FORM", "SIG", "FLOW", "FIXTURE", "TEST"]


@dataclass
class Problem:
    line: int
    severity: str  # "error" or "warning"
    message: str


@dataclass
class KeyLine:
    line: int
    key: str
    value: str


@dataclass
class Decl:
    """One '<name> As <Type>' from ARGS, VAR or a SIG."""
    line: int
    name: str
    type: str


@dataclass
class OpNode:
    """One OP line. Steps nested under it (2 more spaces) are its children."""
    line: int
    text: str
    children: list = field(default_factory=list)

    def walk(self):
        yield self
        for child in self.children:
            yield from child.walk()


@dataclass
class Rule:
    """One FIXTURE data rule with its modifiers, e.g. C:int(1,50)+boundary(0)+blank(3)."""
    col: str
    kind: str   # values, seq, int, num, dates, reldates or text
    args: list  # values: [(text, count)]; the others: their numbers or dates as text
    blank: int = 0
    boundary: list = field(default_factory=list)  # [(kind, value)]: kind is text, number or date
    dupes: int = 0
    mods: list = field(default_factory=list)      # modifier names in the order written


@dataclass
class Fixture:
    line: int
    file: str
    sheet: str
    headers: list  # [(column letter, header)]
    rows: int
    rules: list    # [Rule]
    stale: int = 0


@dataclass
class Test:
    line: int
    fixture: str
    procedure: str
    args: list
    runs: int
    assertion: str


@dataclass
class Sig:
    """One BUILD SIG line: <module>.<Proc>(<params>) [As <Type>]."""
    line: int
    module: str
    proc: str
    params: list  # [Decl]
    ret: str      # "" for a Sub


@dataclass
class Form:
    line: int
    name: str
    controls: list  # [(name, control type)]


@dataclass
class Spec:
    kind: str = None    # "TASK", "BUILD", "QUESTION", or None if the header is unreadable
    version: str = VERSION  # the version line: "VCL 0.1", "VCL 0.2" or "VCL 0.3"
    line: int = 1       # line of the TASK / BUILD header (or of the # QUESTION: line)
    assumes: list = field(default_factory=list)   # [(line, text)]
    question: str = ""
    keys: list = field(default_factory=list)      # [KeyLine], in file order
    args: list = field(default_factory=list)      # [Decl]
    vars: list = field(default_factory=list)      # [Decl]
    ops: list = field(default_factory=list)       # top-level OpNodes
    fixtures: list = field(default_factory=list)  # [Fixture]
    tests: list = field(default_factory=list)     # [Test]
    sigs: list = field(default_factory=list)      # BUILD: [Sig]
    forms: list = field(default_factory=list)     # BUILD: [Form]
    flows: list = field(default_factory=list)     # BUILD: [(line, [names])]
    tasks: list = field(default_factory=list)     # BUILD: one TASK Spec per section after ---
    problems: list = field(default_factory=list)  # [Problem]

    def lines(self, key):
        return [k for k in self.keys if k.key == key]

    def value(self, key):
        found = self.lines(key)
        return found[0].value if found else ""

    def all_ops(self):
        for node in self.ops:
            yield from node.walk()

    def error(self, line, message):
        self.problems.append(Problem(line, "error", message))


# ---------------------------------------------------------------- expression tokens

@dataclass
class Token:
    kind: str     # string, range, hex, number, arg, name or op
    text: str
    spaced: bool  # True if whitespace (or the start of the text) comes right before it
    sheet: str = ""
    header: str = None  # for <Sheet>![Header] ranges
    row: str = ""       # for <Sheet>![Header]<n> ranges


TOKEN_RE = re.compile(r"""
    (?P<space>\s+)
  | (?P<string>"(?:[^"]|"")*")
  | (?P<range>[A-Za-z_]\w*!(?:\[[^\]]*\]\d*|\$?[A-Z]{1,3}\$?\d*(?::\$?[A-Z]{1,3}\$?\d*)?))
  | (?P<hex>\#\w*)
  | (?P<number>\d+(?:\.\d+)?)
  | (?P<arg>\$\w+)
  | (?P<name>[A-Za-z_]\w*(?:\.\w+)*)
  | (?P<op><>|<=|>=|[=<>+\-*/&(),;])
""", re.X)


def tokenize(text):
    """Split an OP, PRE or assertion text into tokens. Returns (tokens, error message)."""
    tokens, pos, spaced = [], 0, True
    while pos < len(text):
        m = TOKEN_RE.match(text, pos)
        if not m:
            if text[pos] == '"':
                return tokens, "unclosed quote"
            return tokens, f"unexpected character {text[pos]!r}"
        if m.lastgroup == "space":
            spaced = True
        else:
            tok = Token(m.lastgroup, m.group(), spaced)
            if tok.kind == "range":
                tok.sheet, ref = tok.text.split("!", 1)
                if ref.startswith("["):
                    tok.header = ref[1:ref.index("]")]
                    tok.row = ref[ref.index("]") + 1:]
            tokens.append(tok)
            spaced = False
        pos = m.end()
    return tokens, ""


def split_top(text, sep):
    """Split text on sep, but not inside "quotes" or (brackets)."""
    parts, depth, quoted, start = [], 0, False, 0
    for i, ch in enumerate(text):
        if ch == '"':
            quoted = not quoted
        elif not quoted and ch == "(":
            depth += 1
        elif not quoted and ch == ")":
            depth -= 1
        elif not quoted and depth == 0 and ch == sep:
            parts.append(text[start:i])
            start = i + 1
    parts.append(text[start:])
    return parts


# ---------------------------------------------------------------- spec structure

FIXTURE_RE = re.compile(r"(?P<file>\S+)\s+sheet=(?P<sheet>\S+)\s+cols=(?P<cols>.+?)\s+rows=(?P<rows>\S+)"
                        r"(?:\s+stale=(?P<stale>\S+))?(?:\s+data=(?P<data>.*))?")
CALL_HEAD_RE = re.compile(r"(\w+)\(")
TEST_RE = re.compile(r"fixture=(?P<fixture>\S+)\s+call=(?P<proc>[A-Za-z_][\w.]*)\((?P<args>[^()]*)\)"
                     r"(?:\s+runs=(?P<runs>\S+))?\s+assert=(?P<assertion>.+)")
SIG_RE = re.compile(r"(?P<module>[A-Za-z_]\w*)\.(?P<proc>[A-Za-z_]\w*)\((?P<params>[^()]*)\)"
                    r"(?:\s+As\s+(?P<ret>\w+))?")
DECL_RE = re.compile(r"\s*([A-Za-z_]\w*)\s+As\s+([A-Za-z_]\w*)\s*")
NUMBER_RE = re.compile(r"-?\d+(?:\.\d+)?")


def parse_spec(text):
    """Parse one spec (TASK, BUILD or a # QUESTION: reply) into a Spec. Never raises."""
    spec = Spec()
    try:
        if text.startswith("﻿"):
            text = text[1:]
        _parse(spec, [(n, line.rstrip()) for n, line in enumerate(text.splitlines(), 1)])
    except Exception as exc:  # safety net: a bug here must not crash the validator
        spec.error(1, f"the parser could not read this spec ({type(exc).__name__}: {exc})")
    return spec


def _parse(spec, lines):
    body = [(n, line) for n, line in lines if line.strip()]
    if not body:
        spec.error(1, "the file is empty")
        return
    questions = [(n, line) for n, line in body if line.startswith("# QUESTION:")]
    if questions:
        spec.kind = "QUESTION"
        spec.line = questions[0][0]
        spec.question = questions[0][1][len("# QUESTION:"):].strip()
        if not spec.question:
            spec.error(spec.line, "# QUESTION: needs a question after it")
        others = [n for n, _ in body if n != spec.line]
        if others:
            spec.error(others[0], "a QUESTION reply contains only its '# QUESTION: <question>' line")
        return
    if body[0][1] in VERSIONS:
        spec.version = body[0][1]
    else:
        spec.error(body[0][0], f"the first line must be '{VERSIONS[0]}', '{VERSIONS[1]}' or '{VERSIONS[2]}'")
    if len(body) < 2 or body[1][1] not in ("TASK", "BUILD"):
        spec.error(body[1][0] if len(body) > 1 else body[0][0], "the second line must be TASK or BUILD")
        return
    spec.kind, spec.line = body[1][1], body[1][0]
    if spec.kind == "TASK":
        _parse_keys(spec, body[2:], TASK_KEYS)
    else:
        _parse_build(spec, body[2:])


def _parse_build(spec, body):
    sections, separators = [[]], []
    for n, line in body:
        if line.strip() == "---":
            sections.append([])
            separators.append(n)
        else:
            sections[-1].append((n, line))
    _parse_keys(spec, sections[0], BUILD_KEYS)
    if not separators:
        spec.error(spec.line, "a BUILD needs one TASK spec per SIG, each after a '---' line")
    for sep_line, section in zip(separators, sections[1:]):
        task = Spec(kind="TASK", line=sep_line, version=spec.version)
        if len(section) < 2 or section[0][1] != spec.version or section[1][1] != "TASK":
            spec.error(section[0][0] if section else sep_line,
                       f"each section after '---' starts with '{spec.version}' and 'TASK'")
            continue
        task.line = section[1][0]
        _parse_keys(task, section[2:], TASK_KEYS)
        spec.tasks.append(task)


def _parse_keys(spec, body, allowed):
    order_seen, last_key, seen_key = -1, None, False
    op_col, op_level, op_stack = None, -1, []
    for n, line in body:
        if line.startswith("#"):
            if line.startswith("# ASSUME:"):
                if seen_key:
                    spec.error(n, "# ASSUME: lines go at the top, before the first key")
                spec.assumes.append((n, line[len("# ASSUME:"):].strip()))
            elif line.startswith("# QUESTION:"):
                spec.error(n, "a QUESTION reply contains only its '# QUESTION: <question>' line")
            else:
                spec.error(n, "unknown comment: the only comment lines are '# ASSUME: <assumption>'")
            continue
        if line[0].isspace():
            spec.error(n, "line starts with spaces: every line starts with a key")
            continue
        key = line.split()[0]
        if key not in allowed:
            spec.error(n, f"unknown key {key}")
            continue
        seen_key = True
        index = allowed.index(key)
        if index < order_seen:
            spec.error(n, f"{key} is out of order: it must come before {last_key}")
        else:
            order_seen, last_key = index, key
        raw = line[len(key):]
        value = raw.strip()
        spec.keys.append(KeyLine(n, key, value))

        if key == "OP":
            step = raw.lstrip(" ")
            if not step.strip():
                spec.error(n, "OP needs a step")
                continue
            if step[0] == "\t":
                spec.error(n, "use spaces, not tabs, to nest OP steps")
                step = step.lstrip()
            col = len(line) - len(step)
            if op_col is None:
                op_col = col
            extra = col - op_col
            level = extra // 2
            if extra < 0 or extra % 2:
                spec.error(n, "bad nesting: nest a step exactly 2 spaces deeper than the step it belongs to")
                level = max(0, min(level, op_level + 1))
            elif level > op_level + 1:
                spec.error(n, "bad nesting: a step can only be one level deeper than the step above it")
                level = op_level + 1
            node = OpNode(n, step.strip())
            while op_stack and op_stack[-1][0] >= level:
                op_stack.pop()
            (op_stack[-1][1].children if op_stack else spec.ops).append(node)
            op_stack.append((level, node))
            op_level = level
        elif key == "ARGS":
            for part in value.split(","):
                _decl(spec, n, part, "ARGS", spec.args)
        elif key == "VAR":
            if "," in value:
                spec.error(n, "VAR declares one variable per line")
            _decl(spec, n, value, "VAR", spec.vars)
        elif key == "FIXTURE":
            _fixture(spec, n, value)
        elif key == "TEST":
            _test(spec, n, value)
        elif key == "SIG":
            _sig(spec, n, value)
        elif key == "FORM":
            _form(spec, n, value)
        elif key == "FLOW":
            spec.flows.append((n, [name.strip() for name in value.split("->")]))


def _decl(spec, n, text, where, into):
    m = DECL_RE.fullmatch(text)
    if not m:
        spec.error(n, f"{where}: expected '<name> As <Type>', got '{text.strip()}'")
        return
    into.append(Decl(n, m.group(1), m.group(2)))


def _fixture(spec, n, value):
    m = FIXTURE_RE.fullmatch(value)
    if not m:
        missing = [part for part in ("sheet=", "cols=", "rows=") if part not in value]
        hint = f" (missing {', '.join(missing)})" if missing else ""
        spec.error(n, "FIXTURE must look like: <name>.xlsx sheet=<name> cols=<A:Header,...> rows=<n> "
                      f"data=<col>:<rule>;...{hint}")
        return
    headers = []
    for item in m["cols"].split(","):
        col = re.fullmatch(r"\s*([A-Z]{1,3}):(.*\S)\s*", item)
        if not col:
            spec.error(n, f"FIXTURE cols: expected <Letter>:<Header>, got '{item.strip()}'")
            continue
        headers.append((col.group(1), col.group(2)))
    rows = 0
    if m["rows"].isdigit():
        rows = int(m["rows"])
    else:
        spec.error(n, f"FIXTURE rows= must be a whole number, got '{m['rows']}'")
    stale = 0
    if m["stale"] is not None:
        if m["stale"].isdigit() and int(m["stale"]) > 0:
            stale = int(m["stale"])
        else:
            spec.error(n, f"FIXTURE stale= must be a whole number of 1 or more, got '{m['stale']}'")
    rules = []
    if m["data"] is not None:
        for item in split_top(m["data"], ";"):
            rule = _rule(spec, n, item.strip())
            if rule:
                rules.append(rule)
    spec.fixtures.append(Fixture(n, m["file"], m["sheet"], headers, rows, rules, stale))


def _calls(text):
    """'kind(args)+mod(args)+...' -> [(name, args)], or None if the text doesn't read that way."""
    out, pos = [], 0
    while pos < len(text):
        m = CALL_HEAD_RE.match(text, pos)
        if not m:
            return None
        depth, quoted, i = 1, False, m.end()
        while i < len(text) and depth:
            if text[i] == '"':
                quoted = not quoted
            elif not quoted and text[i] in "()":
                depth += 1 if text[i] == "(" else -1
            i += 1
        if depth:
            return None
        out.append((m.group(1), text[m.end():i - 1]))
        pos = i
        if pos < len(text):
            if text[pos] != "+" or pos + 1 == len(text):
                return None
            pos += 1
    return out


def _boundary_values(text):
    """The values of +boundary(...): [(kind, value)], or None if one can't be read."""
    out = []
    for part in (p.strip() for p in split_top(text, ",")):
        if re.fullmatch(r'"(?:[^"]|"")*"', part):
            out.append(("text", part[1:-1].replace('""', '"')))
        elif re.fullmatch(r"\d{4}-\d{2}-\d{2}", part):
            try:
                datetime.date.fromisoformat(part)
            except ValueError:
                return None
            out.append(("date", part))
        elif NUMBER_RE.fullmatch(part):
            out.append(("number", part))
        else:
            return None
    return out


def _rule(spec, n, item):
    m = re.fullmatch(r"(?P<col>[A-Z]{1,3}):(?P<rest>.+)", item)
    calls = _calls(m["rest"]) if m else None
    if not calls:
        spec.error(n, f"FIXTURE data: expected <col>:<rule>[+<modifier>...], got '{item}'")
        return None
    (kind, args), modifiers = calls[0], calls[1:]
    rule = Rule(m["col"], kind, [])
    for name, margs in modifiers:
        rule.mods.append(name)
        if name in ("blank", "dupes"):
            if not margs.strip().isdigit():
                spec.error(n, f"FIXTURE data: +{name}() takes one whole number, got '{margs}'")
                return None
            setattr(rule, name, int(margs))
        elif name == "boundary":
            values = _boundary_values(margs)
            if not values:
                spec.error(n, f'FIXTURE data: +boundary() takes "text", numbers or yyyy-mm-dd dates, got \'{margs}\'')
                return None
            rule.boundary = values
        else:
            spec.error(n, f"FIXTURE data: unknown modifier +{name}() (use +boundary, +blank or +dupes)")
            return None
    parts = [a.strip() for a in args.split(",")]
    if kind == "values":
        if not re.fullmatch(r'"(?:[^"]|"")*"x\d+(?:,"(?:[^"]|"")*"x\d+)*', args):
            spec.error(n, f'FIXTURE data: values() takes "<text>"x<n>, ... , got \'{args}\'')
            return None
        rule.args = [(text.replace('""', '"'), int(count))
                     for text, count in re.findall(r'"((?:[^"]|"")*)"x(\d+)', args)]
    elif kind in ("seq", "int", "num", "reldates"):
        pattern = r"-?\d+" if kind in ("int", "reldates") else NUMBER_RE.pattern
        if len(parts) != 2 or not all(re.fullmatch(pattern, p) for p in parts):
            what = "two whole numbers" if kind in ("int", "reldates") else "two numbers"
            spec.error(n, f"FIXTURE data: {kind}() takes {what}, got '{args}'")
            return None
        rule.args = parts
    elif kind == "dates":
        try:
            if len(parts) != 2:
                raise ValueError
            for p in parts:
                if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", p):
                    raise ValueError
                datetime.date.fromisoformat(p)
        except ValueError:
            spec.error(n, f"FIXTURE data: dates() takes two yyyy-mm-dd dates, got '{args}'")
            return None
        rule.args = parts
    elif kind == "text":
        if len(parts) != 1 or not parts[0].isdigit():
            spec.error(n, f"FIXTURE data: text() takes one length, got '{args}'")
            return None
        rule.args = parts
    else:
        spec.error(n, f"FIXTURE data: unknown rule {kind}() (use values, seq, int, num, dates, reldates or text)")
        return None
    return rule


def _test(spec, n, value):
    m = TEST_RE.fullmatch(value)
    if not m:
        missing = [part for part in ("fixture=", "call=", "assert=") if part not in value]
        hint = f" (missing {', '.join(missing)})" if missing else ""
        spec.error(n, "TEST must look like: fixture=<name>.xlsx call=<Proc>(<args>) [runs=<n>] "
                      f"assert=<assertion>{hint}")
        return
    runs = 1
    if m["runs"] is not None:
        if m["runs"].isdigit() and int(m["runs"]) > 0:
            runs = int(m["runs"])
        else:
            spec.error(n, f"runs= must be a whole number of 1 or more, got '{m['runs']}'")
    args = [a.strip() for a in split_top(m["args"], ",")] if m["args"].strip() else []
    spec.tests.append(Test(n, m["fixture"], m["proc"], args, runs, m["assertion"].strip()))


def _sig(spec, n, value):
    m = SIG_RE.fullmatch(value)
    if not m:
        spec.error(n, f"SIG must look like: <module>.<Proc>(<name> As <Type>, ...) [As <Type>], got '{value}'")
        return
    params = []
    if m["params"].strip():
        for part in m["params"].split(","):
            _decl(spec, n, part, "SIG", params)
    spec.sigs.append(Sig(n, m["module"], m["proc"], params, m["ret"] or ""))


def _form(spec, n, value):
    m = re.fullmatch(r"(?P<name>\w+)\s+CONTROLS\s+(?P<controls>.+)", value)
    if not m:
        spec.error(n, f"FORM must look like: <frmName> CONTROLS <name>:<ControlType> ..., got '{value}'")
        return
    controls = []
    for item in m["controls"].split():
        if not re.fullmatch(r"\w+:\w+", item):
            spec.error(n, f"FORM control must look like <name>:<ControlType>, got '{item}'")
            continue
        controls.append(tuple(item.split(":")))
    spec.forms.append(Form(n, m["name"], controls))
