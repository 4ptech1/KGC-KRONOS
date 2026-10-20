"""Static checks on a generated VBA module: catches what Excel would reject at compile time.

lint(text) -> list of (line_no, message). Checks: ASCII only, Option Explicit, balanced blocks
(Sub/Function, If, For/For Each with matching Next variable, Select Case, Do/Loop, With),
declared variables under Option Explicit, unique line labels, GoTo targets, Forge* helpers that
exist, banned constructs, and quote balance. Module-level declarations come before the first
procedure, as Excel requires.

File and system safety (CLAUDE.md, docs/vep-0.3-file-ops.md): nothing deletes, moves or renames a
file or folder, runs a program, reads the registry or declares a Windows API. CreateObject makes
only the classes in CREATE_OBJECT, each in its own procedures, and ONLY_IN keeps the file writes
inside ForgeLib's helpers and the fixture module's sandbox builder (D30), so generated code
reaches the disk only through them. Procedure names are unique, as Excel requires, so a
procedure can't borrow a helper's name.
"""
import re
import sys

BANNED = [(r"\.Select\b", "Select is banned"), (r"\.Activate\b", "Activate is banned"),
          (r"\bSelection\b", "Selection is banned"), (r"\bSendKeys\b", "SendKeys is banned"),
          (r"\bAs\s+Integer\b", "use Long, not Integer"), (r"^\s*End\s*$", "bare End statement"),
          (r"(?i)\b(Kill|RmDir|MkDir|ChDir|ChDrive|Shell|GetObject|Declare)\b", "file, process and API statements are banned"),
          (r"(?i)\.(DeleteFile|DeleteFolder|MoveFile|MoveFolder|CopyFolder|CreateTextFile|OpenTextFile)\b",
           "this file system method is banned"),
          (r"(?i)\b(SaveSetting|GetSetting|GetAllSettings|DeleteSetting)\b", "registry access is banned"),
          (r"(?i)\.(SaveAs|SaveCopyAs)\b", "SaveAs is banned"),
          (r"(?i)^Name\s+.+\s+As\s+", "the Name statement (rename) is banned")]
# CreateObject: the only classes, and the procedures that make them.
CREATE_OBJECT = {"Scripting.Dictionary": {"ForgeNewDict"}, "Scripting.FileSystemObject": {"ForgeFso", "SandboxFso"}}
# (what, pattern, the only procedures where it may appear): every file write goes through these.
ONLY_IN = [("CopyFile", r"\bCopyFile\b", {"ForgeCopyFile"}),
           ("CreateFolder", r"\bCreateFolder\b", {"ForgeMakeFolder", "SandboxFolder"}),
           ("ForgeFso", r"\bForgeFso\b", {"ForgeFso", "ForgeFolderExists", "ForgeBelowRoot", "ForgeMakeFolder",
                                         "ForgeCopyFile", "ForgeListFiles"}),
           ("SandboxFso", r"\bSandboxFso\b", {"SandboxFso", "SandboxNewRoot", "SandboxFolder", "SandboxFile",
                                             "PilotFiles"}),
           ("Writing files with Open", r"(?i)^Open\s+.*\bFor\s+(Output|Append|Binary|Random)\b",
            {"SandboxFile", "SandboxLock"})]
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


def public_helpers(text):
    """The Forge* procedures a module makes public: what another module of the project may call."""
    return set(re.findall(r"^Public (?:Function|Sub) (Forge\w+)", text, re.M))


def lint(text, helpers=()):
    """helpers: Forge* procedures defined Public in another module of the same VBA project."""
    problems = []
    lines = text.splitlines()
    defined_helpers = public_helpers(text) | set(helpers)
    defined_helpers |= set(re.findall(r"^Private (?:Function|Sub) (Forge\w+)", text, re.M))
    module_vars = set(re.findall(r"^(?:Private|Public|Dim) (\w+) As ", text, re.M))
    option_explicit_seen = False
    proc = None
    seen_proc = False
    proc_lines = {}
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
        header = re.match(r"^(?:Public |Private )?(?:Function|Sub) (\w+)\(", stmt)
        here = header.group(1) if header else (proc[1] if proc else None)
        for cls in re.findall(r"\bCreateObject\((\"[^\"]*\"|[^)]*)\)", raw.split("'")[0]):
            cls = cls.strip('"')
            if here not in CREATE_OBJECT.get(cls, ()):
                problems.append((no, f"CreateObject({cls}) is only allowed for "
                                     + "; ".join(f"{k} in {', '.join(sorted(v))}" for k, v in CREATE_OBJECT.items())))
        for what, pat, allowed in ONLY_IN:
            if re.search(pat, stmt) and here not in allowed:
                problems.append((no, f"{what} is only allowed in {', '.join(sorted(allowed))}"))
        m = re.match(r"^(?:Public |Private )?(Function|Sub) (\w+)\((.*?)\)", stmt)
        if m:
            if proc is not None:
                problems.append((no, f"{m.group(2)} starts before {proc[1]} ends"))
            if not option_explicit_seen:
                problems.append((no, "procedure before Option Explicit"))
            proc = (m.group(1), m.group(2))
            seen_proc = True
            if m.group(2).lower() in proc_lines:
                problems.append((no, f"{m.group(2)} is defined twice (line {proc_lines[m.group(2).lower()]}): "
                                     "Excel stops with 'Ambiguous name detected'"))
            proc_lines.setdefault(m.group(2).lower(), no)
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
            if re.match(r"^(Private|Public|Dim) \w+ As |^(Private |Public )?Const \w+ As \w+ = ", stmt):
                if seen_proc:
                    problems.append((no, "module-level declaration after a procedure: Excel only allows comments there"))
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
