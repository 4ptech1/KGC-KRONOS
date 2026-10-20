"""Reference semantics for the VCL core value functions (VCL 0.1, plus the VCL 0.2 value functions).

These functions ARE the meaning of VCL. ForgeLib (VBA) must behave identically; the
conformance table in tests/test_semantics.py is shared by both and is the contract.

Cell values in the model: None (empty), str, int/float (numbers), bool, datetime.date, XlError.
"""
import datetime as _dt
import math as _math
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
    """S5 (VCL 0.2): true for Excel error values (#N/A, #DIV/0!, ...); false for everything else,
    including the text "#N/A"."""
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


# ---------------------------------------------------------------- VCL 0.2 value functions

# FORMAT fill names: each is the colour of the VBA constant of that name (vbRed = RGB(255, 0, 0)).
# "none" removes the fill, so it has no colour here. The validator's FILL_NAMES lists the same names.
FILL_COLORS = {"black": "#000000", "white": "#FFFFFF", "red": "#FF0000", "green": "#00FF00",
               "blue": "#0000FF", "yellow": "#FFFF00", "magenta": "#FF00FF", "cyan": "#00FFFF"}

DATE_FORMATS = ("dd/mm/yyyy", "dd-mm-yyyy", "dd.mm.yyyy", "mm/dd/yyyy", "yyyy-mm-dd")


def same(a, b):
    """S12 (VCL 0.2): SAME(a, b) is true when TRIM(a) and TRIM(b) are the same text, ignoring
    case (letter by letter, as LOWER: "ß" is not "ss"). Numbers are compared as text, so 25 and
    "25" are the same but 25 and "25.0" are not."""
    return trim(a).lower() == trim(b).lower()


def clean_text(v):
    """S13 (VCL 0.2): CLEAN_TEXT removes characters below code 32 (other than tabs and line
    breaks), turns every run of spaces, non-breaking spaces, tabs and line breaks into one
    space, and trims both ends. Case is kept."""
    out, pending = [], False
    for ch in text_of(v):
        if ch in WHITESPACE:
            pending = True
        elif ord(ch) >= 32:
            if pending and out:
                out.append(" ")
            pending = False
            out.append(ch)
    return "".join(out)


def _digits(s, fewest, most):
    return fewest <= len(s) <= most and all("0" <= ch <= "9" for ch in s)


def parse_date(v, fmt):
    """S14 (VCL 0.2): PARSE_DATE(x, format) gives a real date, or empty when x is not a date in
    that format. A real date passes through (time of day dropped). Text is TRIMmed and must
    match the format exactly: day and month 1 or 2 digits, year 4 digits from 1900 to 9999,
    and the date must exist (31/02/2026 gives empty). Numbers and everything else give empty.
    The format is required, because "03/04/2026" means a different day on different PCs."""
    if fmt not in DATE_FORMATS:
        raise ValueError(f"unknown date format {fmt!r}")
    if isinstance(v, _dt.datetime):
        return v.date()
    if isinstance(v, _dt.date):
        return v
    if not isinstance(v, str):
        return None
    sep = fmt[4] if fmt.startswith("yyyy") else fmt[2]
    parts = trim(v).split(sep)
    if len(parts) != 3:
        return None
    order = fmt.split(sep)
    found = dict(zip(order, parts))
    if not (_digits(found["dd"], 1, 2) and _digits(found["mm"], 1, 2) and _digits(found["yyyy"], 4, 4)):
        return None
    y, m, d = int(found["yyyy"]), int(found["mm"]), int(found["dd"])
    if y < 1900:
        return None
    try:
        return _dt.date(y, m, d)
    except ValueError:
        return None


def iif(cond, a, b):
    """S15 (VCL 0.2): IIF(cond, a, b) gives a when cond is TRUE, else b. Both a and b are always
    evaluated (as VBA's IIf), so an IIF branch is never a guard."""
    return a if cond else b


def mod(a, b):
    """S16 (VCL 0.2): MOD(a, b) is a - b * INT(a / b), as Excel's MOD: the result has the sign
    of b, and decimals are kept. VBA's Mod operator rounds both sides first and gives the sign
    of a, so it is not used. MOD(a, 0) is a run-time error."""
    return a - b * _math.floor(a / b)


def needs_text_guard(text):
    """S17 (VCL 0.2): SET_VALUE, WRITE_ROW and COPY_ROW store text as exactly that text. Excel
    would turn some text into something else when it is written to a cell: a number, date,
    time, TRUE/FALSE, error value or formula. Such text is written with a ' prefix, which Excel
    hides, so the cell holds the text unchanged. Being careful, the prefix goes on any text with
    a digit or a character above code 127, any text starting with = + - @ ' or #, and TRUE or
    FALSE after TRIM in any case. "" is never prefixed: it empties the cell."""
    if text == "":
        return False
    return (any("0" <= ch <= "9" or ord(ch) > 127 for ch in text) or text[0] in "=+-@'#"
            or text.strip(WHITESPACE).upper() in ("TRUE", "FALSE"))


# ---------------------------------------------------------------- VCL 0.3 path functions

NAME_FORBIDDEN = '\\/:*?"<>|'
RESERVED_NAMES = {"CON", "PRN", "AUX", "NUL"} | {f"COM{i}" for i in range(1, 10)} | {f"LPT{i}" for i in range(1, 10)}
MAX_NAME = 100


def safe_name(v):
    """S18 (VCL 0.3): SAFE_NAME(x) is CLEAN_TEXT(x) when that is a safe Windows file or folder name,
    else "" (empty). Unsafe: empty, longer than 100 characters, any of \\ / : * ? " < > |, the
    names . and .., a trailing dot, or a reserved device name (CON, PRN, AUX, NUL, COM1 to COM9,
    LPT1 to LPT9), ignoring case and any extension (con.txt is reserved). A name can never climb out
    of its folder, because it can't contain a separator or be .."""
    s = clean_text(v)
    if not s or len(s) > MAX_NAME or any(ch in NAME_FORBIDDEN for ch in s):
        return ""
    if s in (".", "..") or s.endswith("."):
        return ""
    if s.split(".")[0].rstrip(" ").upper() in RESERVED_NAMES:
        return ""
    return s


def path_join(base, *names):
    """S19 (VCL 0.3): PATH_JOIN(base, name, ...) is base, then each name, separated by \\ (a
    trailing \\ characters on base are dropped). It is "" (empty) when base is empty or any name is not
    already a safe name (SAFE_NAME(name) = name), so an empty or unsafe part can never become a
    path; a file op given "" reports Bad name."""
    root = text_of(base).rstrip("\\")
    parts = [text_of(n) for n in names]
    if not root or not parts or any(not n or safe_name(n) != n for n in parts):
        return ""
    return "\\".join([root] + parts)


# ---------------------------------------------------------------- revision 3 functions (VEP 0.4)

def file_name(path):
    """S20 (VEP 0.4, D36): FILE_NAME(path) is the text after the last backslash, or all of the text
    when it has none; a path ending in a backslash gives "". Only \\ separates: a / is part of
    the name."""
    return text_of(path).rsplit("\\", 1)[-1]


def row_key(values):
    """The DEDUPE key of a row from its key cells' values (S21), or None when it has none."""
    if any(is_empty(v) or is_error(v) for v in values):
        return None
    return "".join(f"{len(k)}:{k}" for k in map(dict_key, values))   # the lengths keep "ab","c" from "a","bc"


def same_row_key(a, b):
    """S21 (VEP 0.4, D38): DEDUPE treats two rows as the same when both have a key and the keys are
    equal. A row's key is the values of its key cells, each compared as a dictionary key (S10: its
    canonical text, case ignored, no TRIM). A row with an empty or error key cell has no key, so
    DEDUPE never deletes a row because data is missing."""
    key = row_key(a)
    return key is not None and key == row_key(b)
