"""Static checks on a generated VBA module: catches what Excel would reject at compile time.

lint(text) -> list of (line_no, message). Checks: ASCII only, Option Explicit, balanced blocks
(Sub/Function, If, For/For Each with matching Next variable, Select Case, Do/Loop, With),
declared variables under Option Explicit, unique line labels, GoTo targets, Forge* helpers that
exist, banned constructs, and quote balance.
"""
import re
import sys

BANNED = [(r"\.Select\b", "Select is banned"), (r"\.Activate\b", "Activate is banned"),
          (r"\bSelection\b", "Selection is banned"), (r"\bSendKeys\b", "SendKeys is banned"),
          (r"\bAs\s+Integer\b", "use Long, not Integer"), (r"^\s*End\s*$", "bare End statement")]
KNOWN_GLOBALS = {"Application", "ThisWorkbook", "ActiveWorkbook", "ActiveSheet", "Err", "Debug", "Date",
                 "Now", "True", "False", "Nothing", "Empty", "Null"}


def strip_strings_and_comments(line):
    out, i, in_str = [], 0, False
    while i < len(line):
        ch = line[i]
        if in_str:
            if ch == '"':
                if i + 1 < len(line) and line[i + 1] == '"':
                    i += 2
                    continue
                in_str = False
            i += 1
            continue
        if ch == '"':
            in_str = True
            out.append('""')
            i += 1
            continue
        if ch == "'":
            break
        out.append(ch)
        i += 1
    return "".join(out), in_str


def lint(text):
    problems = []
    lines = text.splitlines()
    defined_helpers = set(re.findall(r"^Public (?:Function|Sub) (Forge\w+)", text, re.M))
    defined_helpers |= set(re.findall(r"^Private (?:Function|Sub) (Forge\w+)", text, re.M))
    module_vars = set(re.findall(r"^(?:Private|Public|Dim) (\w+) As ", text, re.M))
    option_explicit_seen = False
    proc = None
    stack = []
    declared = set()
    labels = set()
    gotos = []
    continuation = ""
    for no, raw in enumerate(lines, start=1):
        if any(ord(ch) > 127 for ch in raw):
            problems.append((no, "non-ASCII character"))
        code, open_str = strip_strings_and_comments(raw)
        if open_str:
            problems.append((no, "unbalanced quotes"))
        if continuation:
            code = continuation + " " + code.strip()
        if code.rstrip().endswith(" _"):
            continuation = code.rstrip()[:-2]
            continue
        continuation = ""
        stmt = code.strip()
        if not stmt:
            continue
        if stmt == "Option Explicit":
            if proc is not None:
                problems.append((no, "Option Explicit inside a procedure"))
            option_explicit_seen = True
            continue
        m = re.match(r"^(\d+)\s+(.*)$", stmt)
        if m:
            if proc is None:
                problems.append((no, "line label outside a procedure"))
            elif m.group(1) in labels:
                problems.append((no, f"duplicate line label {m.group(1)}"))
            labels.add(m.group(1))
            stmt = m.group(2)
        m = re.match(r"^(\w+):$", stmt)
        if m:
            labels.add(m.group(1))
            continue
        for pat, msg in BANNED:
            if re.search(pat, stmt):
                problems.append((no, msg))
        m = re.match(r"^(?:Public |Private )?(Function|Sub) (\w+)\((.*?)\)", stmt)
        if m:
            if proc is not None:
                problems.append((no, f"{m.group(2)} starts before {proc[1]} ends"))
            if not option_explicit_seen:
                problems.append((no, "procedure before Option Explicit"))
            proc = (m.group(1), m.group(2))
            stack, labels, gotos = [], set(), []
            declared = {m.group(2)} | module_vars
            for p in filter(None, (x.strip() for x in m.group(3).split(","))):
                pm = re.match(r"(?:ByVal |ByRef |Optional )*(\w+)", p)
                declared.add(pm.group(1))
            continue
        m = re.match(r"^End (Function|Sub)$", stmt)
        if m:
            if proc is None or proc[0] != m.group(1):
                problems.append((no, f"End {m.group(1)} without a matching start"))
            if stack:
                problems.append((no, f"unclosed {stack[-1][0]} from line {stack[-1][1]}"))
            for target, gline in gotos:
                if target not in labels:
                    problems.append((gline, f"GoTo {target}: no such label"))
            proc = None
            continue
        if proc is None:
            if re.match(r"^(Private|Public|Dim) \w+ As ", stmt):
                continue
            problems.append((no, f"statement outside a procedure: {stmt[:40]}"))
            continue
        for dm in re.finditer(r"\bDim (\w+) As ", stmt):
            declared.add(dm.group(1))
        for g in re.findall(r"\b(?:GoTo|Resume) (\w+)", stmt):
            if not g.isdigit() and g != "Next":
                gotos.append((g, no))
        for h in re.findall(r"\b(Forge\w+)\b", stmt):
            if h not in defined_helpers and h not in declared and not h.startswith("forge"):
                problems.append((no, f"unknown helper {h}"))
        # assignment targets must be declared
        am = re.match(r"^(?:Set )?(\w+)\s*=", stmt)
        if am and am.group(1) not in declared and am.group(1) not in KNOWN_GLOBALS:
            problems.append((no, f"variable {am.group(1)} is not declared"))
        # block structure
        single_if = re.match(r"^If .* Then\s+\S", stmt)
        if re.match(r"^If .* Then$", stmt):
            stack.append(("If", no))
        elif single_if:
            pass
        elif re.match(r"^ElseIf .* Then$", stmt) or stmt == "Else":
            if not stack or stack[-1][0] != "If":
                problems.append((no, f"{stmt.split()[0]} without If"))
        elif stmt == "End If":
            if not stack or stack[-1][0] != "If":
                problems.append((no, "End If without If"))
            else:
                stack.pop()
        elif re.match(r"^For Each (\w+) In ", stmt):
            v = re.match(r"^For Each (\w+) In ", stmt).group(1)
            if v not in declared:
                problems.append((no, f"loop variable {v} is not declared"))
            stack.append(("For", no, v))
        elif re.match(r"^For (\w+) = ", stmt):
            v = re.match(r"^For (\w+) = ", stmt).group(1)
            if v not in declared:
                problems.append((no, f"loop variable {v} is not declared"))
            stack.append(("For", no, v))
        elif re.match(r"^Next\b", stmt):
            v = stmt[4:].strip()
            if not stack or stack[-1][0] != "For":
                problems.append((no, "Next without For"))
            else:
                top = stack.pop()
                if v and v != top[2]:
                    problems.append((no, f"Next {v} closes For {top[2]}"))
        elif re.match(r"^Select Case ", stmt):
            stack.append(("Select", no))
        elif stmt == "End Select":
            if not stack or stack[-1][0] != "Select":
                problems.append((no, "End Select without Select Case"))
            else:
                stack.pop()
        elif re.match(r"^Do\b", stmt):
            stack.append(("Do", no))
        elif re.match(r"^Loop\b", stmt):
            if not stack or stack[-1][0] != "Do":
                problems.append((no, "Loop without Do"))
            else:
                stack.pop()
        elif re.match(r"^With ", stmt):
            stack.append(("With", no))
        elif stmt == "End With":
            if not stack or stack[-1][0] != "With":
                problems.append((no, "End With without With"))
            else:
                stack.pop()
        if len(raw) > 1023:
            problems.append((no, "line longer than 1023 characters"))
    if proc is not None:
        problems.append((len(lines), f"{proc[1]} never ends"))
    if not option_explicit_seen:
        problems.append((1, "Option Explicit is missing"))
    return problems


if __name__ == "__main__":
    total = 0
    for p in sys.argv[1:]:
        found = lint(open(p, encoding="utf-8").read())
        total += len(found)
        print(f"{'OK  ' if not found else 'FAIL'} {p}")
        for ln, msg in found:
            print(f"     line {ln}: {msg}")
    sys.exit(1 if total else 0)
