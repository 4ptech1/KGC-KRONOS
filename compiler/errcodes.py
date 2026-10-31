"""Stable codes for the messages the checks give (VEP 0.6, D60).

    from errcodes import coded, code_of
    self.error(coded(f"unknown key {key}", "V012"))

A message is a plain str that also carries its code in `.code`, so every place that compares,
prints or searches messages keeps working. A code, once published, never changes meaning and is
never reused: V the validator's rules, G the grammar, T the type checker, X the fixture builder,
R what a TEST run reports, K the KRONOS engine's reading of an AI reply (D72).
`python compiler/gen_error_codes.py` gives new messages their codes and writes
`docs/error-codes.md`; `--check` (and a unit test) fails on an uncoded or reused one.
"""
import os


class Coded(str):
    """A message with its code. `str(message)` keeps the code, f-strings and + give a plain str."""

    code = ""

    def __str__(self):
        return self


def coded(message, code=""):
    """The message as a Coded; `code` is the literal code gen_error_codes.py assigns."""
    out = Coded(message)
    out.code = code
    return out


def code_of(message):
    """The code a message carries, or "" for a plain str."""
    return getattr(message, "code", "")


def tag(message):
    """` [V012]` after a message on the command line, or nothing for an uncoded one."""
    code = code_of(message)
    return f" [{code}]" if code else ""


def strict():
    """True in the unit tests (KGC_STRICT_CODES=1): a Problem without a code is a bug."""
    return os.environ.get("KGC_STRICT_CODES") == "1"
