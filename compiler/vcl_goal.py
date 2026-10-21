"""Move a spec to VCL 0.6 by adding its GOAL lines (VEP 0.6, D61).

    from vcl_goal import add_goals
    new = add_goals(text, "Flag every weld with a pending NDE status.")        # a TASK
    new = add_goals(text, {"": "The BUILD's goal.", "Run": "...", "Mark": "..."})   # a BUILD

A GOAL line goes right before PROC in a TASK (after its # ASSUME: lines) and right before NAME in a
BUILD; every TASK inside a BUILD has its own, found by its PROC name. The version line becomes
VCL 0.6 in every section. Nothing else changes, so a spec that was valid gives the same results.
"""
import re

KEY_WIDTH = 9


def add_goals(text, goals, version="VCL 0.6"):
    """The spec with GOAL lines and `version`. goals: a str for a TASK, or {"": the BUILD's goal,
    "<Proc>": that TASK's goal} for a BUILD. Raises ValueError when a goal is missing."""
    lines = text.split("\n")
    kind = next((l.strip() for l in lines[1:3] if l.strip() in ("TASK", "BUILD")), "")
    if isinstance(goals, str):
        goals = {"": goals}
    out, section_kind, pending, used = [], None, None, set()
    for i, line in enumerate(lines):
        if re.fullmatch(r"VCL 0\.\d", line.strip()) and (i == 0 or lines[i - 1].strip() == "---"):
            out.append(version)
            section_kind = None
            continue
        if line.strip() in ("TASK", "BUILD") and section_kind is None:
            section_kind = line.strip()
            pending = "NAME" if section_kind == "BUILD" else "PROC"
            out.append(line)
            continue
        if pending and line.split(" ")[0] == pending:
            if pending == "PROC":
                name = line.split()[1] if len(line.split()) > 1 else ""
                key = name if (kind == "BUILD" or name in goals) else ""
            else:
                key = ""
            if key not in goals:
                raise ValueError(f"no goal for {key or 'the spec'}")
            out.append(f"{'GOAL':<{KEY_WIDTH}}{goals[key]}")
            used.add(key)
            pending = None
        out.append(line)
    if kind == "":
        raise ValueError("not a TASK or a BUILD spec")
    unused = sorted(set(goals) - used)
    if unused:
        raise ValueError(f"a goal for {unused[0] or 'the spec'}, which the spec doesn't have")
    return "\n".join(out)
