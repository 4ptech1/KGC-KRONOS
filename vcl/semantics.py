"""Reference semantics for VCL 0.1 core value functions (milestone M3 seed).

These functions ARE the meaning of VCL. ForgeLib (VBA) must behave identically; the
conformance table in tests/test_semantics.py is shared by both and is the contract.

Cell values in the model: None (empty), str, int/float (numbers), bool, datetime.date, XlError.
"""
import datetime as _dt
import re as _re

WHITESPACE = " \u00a0\t\r\n"        # space, non-breaking space, tab, CR, LF
_PLAIN_NUMBER = _re.compile(r"^[+-]?(\d+(\.\d+)?|\.\d+)$")


class XlError:
    """An Excel error value such as #N/A."""
    def __init__(self, text):
        self.text = text

    def __eq__(self, other):
        return isinstance(other, XlError) and other.text == self.text

    def __repr__(self):
        return self.text


def _is_num_type(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def is_empty(v):
    """S1: empty cell, empty text, or text made only of spaces, non-breaking spaces, tabs or line breaks."""
    if v is None:
        return True
    if isinstance(v, str):
        return all(ch in WHITESPACE for ch in v)
    return False


def trim(v):
    """S2: text of v with whitespace (incl. non-breaking spaces) removed at both ends. VBA Trim only removes spaces."""
    return text_of(v).strip(WHITESPACE)


def is_number(v):
    """S3: true for numbers and for text that is a plain number after TRIM ("25", " -3.5 ").
    False for blanks, booleans, dates, errors, and text like "1,250", "1e3", "25%", "(25)"."""
    if _is_num_type(v):
        return v == v  # NaN is not a number
    if isinstance(v, str):
        return bool(_PLAIN_NUMBER.match(v.strip(WHITESPACE)))
    return False


def to_number(v):
    """The number of a value for which is_number is true."""
    if _is_num_type(v):
        return v
    s = v.strip(WHITESPACE)
    return int(s) if _re.match(r"^[+-]?\d+$", s) else float(s)


def is_date(v):
    """S4: true only for real date values. Text that looks like a date ("15/10/2026") is false."""
    return isinstance(v, (_dt.date, _dt.datetime))


def is_error(v):
    """S5 (VCL 0.2 candidate): true for Excel error values (#N/A, #DIV/0!, ...)."""
    return isinstance(v, XlError)


def text_of(v):
    """S6: the canonical text of any value, independent of the PC's regional settings.
    empty -> ""; whole numbers without ".0"; "." as decimal point; dates -> yyyy-mm-dd;
    booleans -> TRUE/FALSE; errors -> their Excel text (#N/A)."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if _is_num_type(v):
        if isinstance(v, float) and v.is_integer() and abs(v) < 1e15:
            return str(int(v))
        return repr(v) if isinstance(v, float) else str(v)
    if isinstance(v, _dt.datetime):
        return v.strftime("%Y-%m-%d") if (v.hour, v.minute, v.second) == (0, 0, 0) else v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, _dt.date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, XlError):
        return v.text
    return str(v)


def upper(v):
    return text_of(v).upper()


def lower(v):
    return text_of(v).lower()


def length(v):
    return len(text_of(v))


def left(v, n):
    """S8: n below 0 counts as 0; n beyond the length returns the whole text."""
    return text_of(v)[:max(0, int(n))]


def right(v, n):
    n = max(0, int(n))
    s = text_of(v)
    return s[len(s) - n:] if n else ""


def equals(a, b):
    """S7: '=' between a cell and text compares the cell's canonical text, exactly (case and
    spaces count). So a cell holding the number 25 equals "25". Plain VBA would say False."""
    if isinstance(a, str) or isinstance(b, str):
        return text_of(a) == text_of(b)
    if is_date(a) and is_date(b):
        return a == b
    if _is_num_type(a) and _is_num_type(b):
        return a == b
    if isinstance(a, bool) and isinstance(b, bool):
        return a == b
    raise TypeError("compare only after guarding with IS_NUMBER or IS_DATE")


def header_key(v):
    """S9: how COL matches headers: canonical text, line breaks and other whitespace become one
    space, runs collapse, ends trimmed, case ignored. "Required\\nQty" matches "required qty"."""
    s = text_of(v)
    for ch in WHITESPACE:
        s = s.replace(ch, " ")
    return " ".join(s.split()).casefold()


def col(headers, name):
    """COL: 1-based position of the first header that matches name, or 0."""
    want = header_key(name)
    for i, h in enumerate(headers, start=1):
        if header_key(h) == want:
            return i
    return 0


def dict_key(v):
    """S10: dictionary keys compare as canonical text, ignoring case: 1001, "1001" and "1001 "
    (after TRIM) are the same key. Without this, VBA's Dictionary treats 1001 and "1001" as
    different keys."""
    return text_of(v).casefold()


def last_row(column_values):
    """S11: like Excel's End(xlUp): the last row holding any value, including spaces or text
    from formulas. A column that is completely empty gives 1 (the header row)."""
    for i in range(len(column_values), 0, -1):
        if column_values[i - 1] is not None:
            return i
    return 1
