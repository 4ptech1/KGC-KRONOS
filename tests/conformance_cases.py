"""Conformance cases for VCL core semantics: (function, args, expected, why).
Shared by tests/test_semantics.py, docs generation and (later) the Excel conformance batch."""
import datetime as dt
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "vcl"))
from semantics import XlError  # noqa: E402

NA = XlError("#N/A")
D = dt.date(2026, 10, 15)
NBSP = "\u00a0"

CASES = [
    # IS_EMPTY
    ("is_empty", [None], True, "empty cell"),
    ("is_empty", [""], True, "empty text"),
    ("is_empty", ["   "], True, "spaces only"),
    ("is_empty", [NBSP + " "], True, "non-breaking space from copy-paste"),
    ("is_empty", ["\n"], True, "line break only (Alt+Enter)"),
    ("is_empty", [0], False, "zero is a value"),
    ("is_empty", [False], False, "FALSE is a value"),
    ("is_empty", [NA], False, "an error cell is not empty"),
    ("is_empty", [" x "], False, "text with spaces around it"),
    # IS_NUMBER
    ("is_number", [25], True, "a number"),
    ("is_number", [2.5], True, "a decimal"),
    ("is_number", ["25"], True, "numeric text"),
    ("is_number", [" -3.5 "], True, "numeric text with spaces and a sign"),
    ("is_number", ["007"], True, "leading zeros are still a number (use TRIM for IDs)"),
    ("is_number", [None], False, "blank is not 0"),
    ("is_number", [""], False, "empty text is not 0"),
    ("is_number", ["1,250"], False, "thousands separator: regional, so rejected"),
    ("is_number", ["1e3"], False, "scientific notation text rejected"),
    ("is_number", ["25%"], False, "percent text rejected"),
    ("is_number", ["(25)"], False, "accounting negative rejected"),
    ("is_number", [True], False, "TRUE is not a number"),
    ("is_number", [D], False, "a date is a date, not a number"),
    ("is_number", [NA], False, "an error is not a number"),
    # IS_DATE
    ("is_date", [D], True, "a real date"),
    ("is_date", ["15/10/2026"], False, "text that looks like a date"),
    ("is_date", [46310], False, "a serial number is a number"),
    ("is_date", [None], False, "blank"),
    ("is_date", [NA], False, "error"),
    # IS_ERROR (0.2 candidate)
    ("is_error", [NA], True, "#N/A"),
    ("is_error", [XlError("#DIV/0!")], True, "#DIV/0!"),
    ("is_error", ["#N/A"], False, "the text '#N/A' is just text"),
    # TRIM
    ("trim", ["  W-01  "], "W-01", "spaces"),
    ("trim", [NBSP + "W-01" + NBSP], "W-01", "non-breaking spaces (VBA Trim misses these)"),
    ("trim", ["W-01\n"], "W-01", "trailing line break"),
    ("trim", ["W  01"], "W  01", "inner spaces kept"),
    ("trim", [1001], "1001", "number becomes its text"),
    ("trim", [None], "", "empty"),
    # text_of (used by TRIM, UPPER, &, dictionary keys)
    ("text_of", [25.0], "25", "whole number without .0"),
    ("text_of", [2.5], "2.5", "decimal point is always '.'"),
    ("text_of", [D], "2026-10-15", "dates are yyyy-mm-dd on every PC"),
    ("text_of", [True], "TRUE", "boolean"),
    ("text_of", [NA], "#N/A", "error text, not VBA's 'Error 2042'"),
    # UPPER / LEN / LEFT / RIGHT
    ("upper", ["Open"], "OPEN", "case"),
    ("length", ["W-01"], 4, "length"),
    ("length", [None], 0, "empty"),
    ("left", ["Fit-up", 3], "Fit", "left part"),
    ("left", ["Fit", 10], "Fit", "longer than the text"),
    ("left", ["Fit", -1], "", "negative counts as 0"),
    ("right", ["W-01", 2], "01", "right part"),
    # equals (=)
    ("equals", [25, "25"], True, "number cell equals its text (plain VBA says False)"),
    ("equals", ["Open", "Open"], True, "same text"),
    ("equals", ["open", "Open"], False, "= is case-sensitive; use UPPER or SAME"),
    ("equals", ["Open ", "Open"], False, "= doesn't trim; use TRIM or SAME"),
    ("equals", [None, ""], True, "empty cell equals empty text"),
    ("equals", [2, 2.0], True, "numbers"),
    # COL / header matching
    ("col", [["Part No", "Required Qty"], "required qty"], 2, "case ignored"),
    ("col", [["Part No", "Required\nQty"], "Required Qty"], 2, "line break in header (Alt+Enter)"),
    ("col", [["Part  No ", "Qty"], "Part No"], 1, "double and trailing spaces"),
    ("col", [["Part No", "Part No"], "Part No"], 1, "duplicate header: first wins"),
    ("col", [["Part No"], "Qty"], 0, "missing header gives 0"),
    # dictionary keys
    ("dict_key", [1001], "1001", "number key"),
    ("dict_key", ["1001"], "1001", "text key equals the number key"),
    ("dict_key", ["W-01"], "w-01", "case ignored"),
    # LASTROW
    ("last_row", [["Hdr", "a", None, "b", None]], 4, "last value, gaps allowed"),
    ("last_row", [["Hdr", "a", "  "]], 3, "a spaces-only cell still counts (like End(xlUp))"),
    ("last_row", [[None, None]], 1, "empty column gives the header row"),
]
