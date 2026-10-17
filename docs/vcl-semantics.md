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
| S5 (VCL 0.2) | `is_error` | true for Excel error values (#N/A, #DIV/0!, ...); false for everything else, including the text "#N/A". |
| S6 | `text_of` | the canonical text of any value, independent of the PC's regional settings. empty -> ""; whole numbers without ".0"; "." as decimal point; dates -> yyyy-mm-dd; booleans -> TRUE/FALSE; errors -> their Excel text (#N/A). |
| S7 | `equals` | '=' between a cell and text compares the cell's canonical text, exactly (case and spaces count). So a cell holding the number 25 equals "25". Plain VBA would say False. |
| S8 | `left` | n below 0 counts as 0; n beyond the length returns the whole text. |
| S9 | `header_key` | how COL matches headers: canonical text, line breaks and other whitespace become one space, runs collapse, ends trimmed, case ignored. "Required\nQty" matches "required qty". |
| S10 | `dict_key` | dictionary keys compare as canonical text, ignoring case: 1001, "1001" and "1001 " (after TRIM) are the same key. Without this, VBA's Dictionary treats 1001 and "1001" as different keys. |
| S11 | `last_row` | like Excel's End(xlUp): the last row holding any value, including spaces or text from formulas. A column that is completely empty gives 1 (the header row). |
| S12 (VCL 0.2) | `same` | SAME(a, b) is true when TRIM(a) and TRIM(b) are the same text, ignoring case (letter by letter, as LOWER: "ß" is not "ss"). Numbers are compared as text, so 25 and "25" are the same but 25 and "25.0" are not. |
| S13 (VCL 0.2) | `clean_text` | CLEAN_TEXT removes characters below code 32 (other than tabs and line breaks), turns every run of spaces, non-breaking spaces, tabs and line breaks into one space, and trims both ends. Case is kept. |
| S14 (VCL 0.2) | `parse_date` | PARSE_DATE(x, format) gives a real date, or empty when x is not a date in that format. A real date passes through (time of day dropped). Text is TRIMmed and must match the format exactly: day and month 1 or 2 digits, year 4 digits from 1900 to 9999, and the date must exist (31/02/2026 gives empty). Numbers and everything else give empty. The format is required, because "03/04/2026" means a different day on different PCs. |
| S15 (VCL 0.2) | `iif` | IIF(cond, a, b) gives a when cond is TRUE, else b. Both a and b are always evaluated (as VBA's IIf), so an IIF branch is never a guard. |
| S16 (VCL 0.2) | `mod` | MOD(a, b) is a - b * INT(a / b), as Excel's MOD: the result has the sign of b, and decimals are kept. VBA's Mod operator rounds both sides first and gives the sign of a, so it is not used. MOD(a, 0) is a run-time error. |

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

### IS_ERROR (VCL 0.2)

| Input | Result | Why |
| --- | --- | --- |
| #N/A | True | #N/A |
| #DIV/0! | True | #DIV/0! |
| "#N/A" | False | the text '#N/A' is just text |
| (empty) | False | an empty cell is not an error |
| 0 | False | zero is not an error |

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

### SAME (VCL 0.2)

| Input | Result | Why |
| --- | --- | --- |
| " ready ", "READY" | True | spaces at the ends and case are ignored |
| "⍽Ready\n", "ready" | True | non-breaking spaces and line breaks too, as TRIM |
| "Re ady", "Ready" | False | inner spaces count |
| 25, "25" | True | a number compares as its text |
| 25, "25.0" | False | not compared as numbers: use IS_NUMBER and = for that |
| 2.5, " 2.5 " | True | decimal point is always '.' |
| (empty), "" | True | an empty cell is the same as empty text |
| (empty), "   " | True | spaces only, after TRIM, is empty text |
| date 2026-10-15, "2026-10-15" | True | a date compares as yyyy-mm-dd |
| True, "true" | True | TRUE compares as its text, case ignored |
| #N/A, "#n/a" | True | an error compares as its text: use IS_ERROR to tell them apart |
| "straße", "STRASSE" | False | letter by letter, as LOWER: ß is not ss |

### CLEAN_TEXT (VCL 0.2)

| Input | Result | Why |
| --- | --- | --- |
| "  W-01 ⍽\n  A " | "W-01 A" | runs of any whitespace become one space |
| "W-01\tA" | "W-01 A" | a tab inside becomes a space |
| "Fit\x01up" | "Fitup" | characters below code 32 are removed |
| "a \x07 b" | "a b" | removed first, so the spaces around it join |
| "Fit-Up" | "Fit-Up" | case is kept |
| 1001 | "1001" | a number becomes its text |
| (empty) | "" | empty |
| "   " | "" | spaces only |

### PARSE_DATE (VCL 0.2)

| Input | Result | Why |
| --- | --- | --- |
| "15/10/2026", "dd/mm/yyyy" | date 2026-10-15 | day first |
| "10/15/2026", "mm/dd/yyyy" | date 2026-10-15 | month first |
| "2026-10-15", "yyyy-mm-dd" | date 2026-10-15 | ISO |
| "15-10-2026", "dd-mm-yyyy" | date 2026-10-15 | dashes |
| "15.10.2026", "dd.mm.yyyy" | date 2026-10-15 | dots |
| " 5/3/2026 ", "dd/mm/yyyy" | date 2026-03-05 | trimmed; 1-digit day and month |
| "03/04/2026", "dd/mm/yyyy" | date 2026-04-03 | the format decides, never the PC |
| "03/04/2026", "mm/dd/yyyy" | date 2026-03-04 | the same text, the other format |
| "31/02/2026", "dd/mm/yyyy" | (empty) | the date must exist |
| "29/02/2028", "dd/mm/yyyy" | date 2028-02-29 | leap day |
| "15/10/26", "dd/mm/yyyy" | (empty) | the year needs 4 digits |
| "15/10/0026", "dd/mm/yyyy" | (empty) | years before 1900 give empty (Excel can't hold them) |
| "015/10/2026", "dd/mm/yyyy" | (empty) | day takes at most 2 digits |
| "15-10-2026", "dd/mm/yyyy" | (empty) | other separators give empty |
| "15 / 10 / 2026", "dd/mm/yyyy" | (empty) | no spaces inside |
| "15/10/2026 08:30", "dd/mm/yyyy" | (empty) | no time of day in text |
| date 2026-10-15, "mm/dd/yyyy" | date 2026-10-15 | a real date passes through |
| date and time 2026-10-15 08:30, "dd/mm/yyyy" | date 2026-10-15 | time of day dropped |
| 46310, "dd/mm/yyyy" | (empty) | a number is not a date |
| (empty), "dd/mm/yyyy" | (empty) | empty gives empty |
| #N/A, "dd/mm/yyyy" | (empty) | an error gives empty |

### IIF (VCL 0.2)

| Input | Result | Why |
| --- | --- | --- |
| True, "Short", "OK" | "Short" | TRUE gives the first value |
| False, "Short", "OK" | "OK" | FALSE gives the second |
| False, 1, (empty) | (empty) | a value can be empty |

### MOD (VCL 0.2)

| Input | Result | Why |
| --- | --- | --- |
| 7, 3 | 1 | remainder |
| -7, 3 | 2 | sign of b, as Excel (VBA's Mod gives -1) |
| 7, -3 | -2 | sign of b |
| 7.5, 2 | 1.5 | decimals kept (VBA's Mod rounds them away) |
| 6, 3 | 0 | divides exactly |
| 0, 5 | 0 | zero |
