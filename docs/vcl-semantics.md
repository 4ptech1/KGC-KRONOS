# VCL core semantics (reference)

Generated from `vcl/semantics.py` and `tests/conformance_cases.py`. Do not edit by hand.
ForgeLib (VBA) must give exactly these results; the same table becomes the Excel conformance batch.
In the tables, ⍽ is a non-breaking space (character 160).

## Decisions

| Id | Function | Meaning |
| --- | --- | --- |
| S1 | `is_empty` | empty cell, empty text, or text made only of spaces, non-breaking spaces, tabs or line breaks. |
| S2 | `trim` | text of v with whitespace (incl. non-breaking spaces) removed at both ends. VBA Trim only removes spaces. |
| S3 | `is_number` | true for numbers and for text that is a plain number after TRIM ("25", " -3.5 "). False for blanks, booleans, dates, errors, and text like "1,250", "1e3", "25%", "(25)". |
| S4 | `is_date` | true only for real date values. Text that looks like a date ("15/10/2026") is false. |
| S5 (VCL 0.2 candidate) | `is_error` | true for Excel error values (#N/A, #DIV/0!, ...). |
| S6 | `text_of` | the canonical text of any value, independent of the PC's regional settings. empty -> ""; whole numbers without ".0"; "." as decimal point; dates -> yyyy-mm-dd; booleans -> TRUE/FALSE; errors -> their Excel text (#N/A). |
| S7 | `equals` | '=' between a cell and text compares the cell's canonical text, exactly (case and spaces count). So a cell holding the number 25 equals "25". Plain VBA would say False. |
| S8 | `left` | n below 0 counts as 0; n beyond the length returns the whole text. |
| S9 | `header_key` | how COL matches headers: canonical text, line breaks and other whitespace become one space, runs collapse, ends trimmed, case ignored. "Required\nQty" matches "required qty". |
| S10 | `dict_key` | dictionary keys compare as canonical text, ignoring case: 1001, "1001" and "1001 " (after TRIM) are the same key. Without this, VBA's Dictionary treats 1001 and "1001" as different keys. |
| S11 | `last_row` | like Excel's End(xlUp): the last row holding any value, including spaces or text from formulas. A column that is completely empty gives 1 (the header row). |

## Conformance cases


### IS_EMPTY

| Input | Result | Why |
| --- | --- | --- |
| (empty) | True | empty cell |
| "" | True | empty text |
| "   " | True | spaces only |
| "⍽ " | True | non-breaking space from copy-paste |
| "\n" | True | line break only (Alt+Enter) |
| 0 | False | zero is a value |
| False | False | FALSE is a value |
| #N/A | False | an error cell is not empty |
| " x " | False | text with spaces around it |

### IS_NUMBER

| Input | Result | Why |
| --- | --- | --- |
| 25 | True | a number |
| 2.5 | True | a decimal |
| "25" | True | numeric text |
| " -3.5 " | True | numeric text with spaces and a sign |
| "007" | True | leading zeros are still a number (use TRIM for IDs) |
| (empty) | False | blank is not 0 |
| "" | False | empty text is not 0 |
| "1,250" | False | thousands separator: regional, so rejected |
| "1e3" | False | scientific notation text rejected |
| "25%" | False | percent text rejected |
| "(25)" | False | accounting negative rejected |
| True | False | TRUE is not a number |
| date 2026-10-15 | False | a date is a date, not a number |
| #N/A | False | an error is not a number |

### IS_DATE

| Input | Result | Why |
| --- | --- | --- |
| date 2026-10-15 | True | a real date |
| "15/10/2026" | False | text that looks like a date |
| 46310 | False | a serial number is a number |
| (empty) | False | blank |
| #N/A | False | error |

### IS_ERROR (0.2 candidate)

| Input | Result | Why |
| --- | --- | --- |
| #N/A | True | #N/A |
| #DIV/0! | True | #DIV/0! |
| "#N/A" | False | the text '#N/A' is just text |

### TRIM

| Input | Result | Why |
| --- | --- | --- |
| "  W-01  " | "W-01" | spaces |
| "⍽W-01⍽" | "W-01" | non-breaking spaces (VBA Trim misses these) |
| "W-01\n" | "W-01" | trailing line break |
| "W  01" | "W  01" | inner spaces kept |
| 1001 | "1001" | number becomes its text |
| (empty) | "" | empty |

### Text conversion (TRIM, UPPER, &, keys)

| Input | Result | Why |
| --- | --- | --- |
| 25.0 | "25" | whole number without .0 |
| 2.5 | "2.5" | decimal point is always '.' |
| date 2026-10-15 | "2026-10-15" | dates are yyyy-mm-dd on every PC |
| True | "TRUE" | boolean |
| #N/A | "#N/A" | error text, not VBA's 'Error 2042' |

### UPPER

| Input | Result | Why |
| --- | --- | --- |
| "Open" | "OPEN" | case |

### LEN

| Input | Result | Why |
| --- | --- | --- |
| "W-01" | 4 | length |
| (empty) | 0 | empty |

### LEFT

| Input | Result | Why |
| --- | --- | --- |
| "Fit-up", 3 | "Fit" | left part |
| "Fit", 10 | "Fit" | longer than the text |
| "Fit", -1 | "" | negative counts as 0 |

### RIGHT

| Input | Result | Why |
| --- | --- | --- |
| "W-01", 2 | "01" | right part |

### = with text

| Input | Result | Why |
| --- | --- | --- |
| 25, "25" | True | number cell equals its text (plain VBA says False) |
| "Open", "Open" | True | same text |
| "open", "Open" | False | = is case-sensitive; use UPPER or SAME |
| "Open ", "Open" | False | = doesn't trim; use TRIM or SAME |
| (empty), "" | True | empty cell equals empty text |
| 2, 2.0 | True | numbers |

### COL

| Input | Result | Why |
| --- | --- | --- |
| ["Part No", "Required Qty"], "required qty" | 2 | case ignored |
| ["Part No", "Required\nQty"], "Required Qty" | 2 | line break in header (Alt+Enter) |
| ["Part  No ", "Qty"], "Part No" | 1 | double and trailing spaces |
| ["Part No", "Part No"], "Part No" | 1 | duplicate header: first wins |
| ["Part No"], "Qty" | 0 | missing header gives 0 |

### Dictionary keys

| Input | Result | Why |
| --- | --- | --- |
| 1001 | "1001" | number key |
| "1001" | "1001" | text key equals the number key |
| "W-01" | "w-01" | case ignored |

### LASTROW

| Input | Result | Why |
| --- | --- | --- |
| ["Hdr", "a", (empty), "b", (empty)] | 4 | last value, gaps allowed |
| ["Hdr", "a", "  "] | 3 | a spaces-only cell still counts (like End(xlUp)) |
| [(empty), (empty)] | 1 | empty column gives the header row |
