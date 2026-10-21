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
| S17 (VCL 0.2) | `needs_text_guard` | SET_VALUE, WRITE_ROW and COPY_ROW store text as exactly that text. Excel would turn some text into something else when it is written to a cell: a number, date, time, TRUE/FALSE, error value or formula. Such text is written with a ' prefix, which Excel hides, so the cell holds the text unchanged. Being careful, the prefix goes on any text with a digit or a character above code 127, any text starting with = + - @ ' or #, and TRUE or FALSE after TRIM in any case. "" is never prefixed: it empties the cell. |
| S18 (VCL 0.3) | `safe_name` | SAFE_NAME(x) is CLEAN_TEXT(x) when that is a safe Windows file or folder name, else "" (empty). Unsafe: empty, longer than 100 characters, any of \ / : * ? " < > /, the names . and .., a trailing dot, or a reserved device name (CON, PRN, AUX, NUL, COM1 to COM9, LPT1 to LPT9), ignoring case and any extension (con.txt is reserved). A name can never climb out of its folder, because it can't contain a separator or be .. |
| S19 (VCL 0.3) | `path_join` | PATH_JOIN(base, name, ...) is base, then each name, separated by \ (a trailing \ characters on base are dropped). It is "" (empty) when base is empty or any name is not already a safe name (SAFE_NAME(name) = name), so an empty or unsafe part can never become a path; a file op given "" reports Bad name. |
| S20 (VEP 0.4, D36) | `file_name` | FILE_NAME(path) is the text after the last backslash, or all of the text when it has none; a path ending in a backslash gives "". Only \ separates: a / is part of the name. |
| S21 (VEP 0.4, D38) | `same_row_key` | DEDUPE treats two rows as the same when both have a key and the keys are equal. A row's key is the values of its key cells, each compared as a dictionary key (S10: its canonical text, case ignored, no TRIM). A row with an empty or error key cell has no key, so DEDUPE never deletes a row because data is missing. |
| S22 (VEP 0.4, D39) | `find` | FIND(range, x, match=...) looks at the cells row by row, left to right, and gives the first that matches, or NOTHING. match=whole (the default): the cell is SAME as x (S12). match=part: the cell's text holds TRIM(x), ignoring case (letter by letter, as LOWER); an empty TRIM(x) finds NOTHING, because every cell would hold it. Here: the 1-based position of that cell among values, or 0 for NOTHING. |
| S23 (VEP 0.4, D40) | `sheet_name_ok` | a name ADD_SHEET may give a new sheet: 1 to 31 characters, no spaces at the ends (TRIM leaves it as it is), none of \ / ? * [ ] :, no apostrophe first or last, no character below code 32, and not History, which Excel keeps for itself (case ignored). Stricter than Excel where Excel allows lookalikes, so a name never differs from what a TEST asserts. |
| S24 (VEP 0.4, D44) | `list_files` | LIST_FILES gives the names of the files directly in the folder that match the pattern, sorted by LOWER(name) and then by the name as written, so the order is the same on every PC. Hidden and system files are left out (D47; FIXTURE folders hold none). Here: from the names of the files in the folder. |
| S25 (VEP 0.4, D45) | `text_ok` | text that READ_TEXT and WRITE_TEXT handle: printable ASCII (codes 32 to 126), tabs and line breaks, at most 32,767 characters. Any other character would read or write differently with each PC's code page. |
| S25 (VEP 0.4, D45) | `file_text` | what READ_TEXT gives for a file's contents: the text with one final line break (CR LF, or LF) dropped, as editors add one; None when it isn't text (S25), which READ_TEXT reports as Not text. |

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

### Text written by SET_VALUE, WRITE_ROW and COPY_ROW (VCL 0.2)

| Input | Result | Why |
| --- | --- | --- |
| "Short" | False | plain words are stored as typed |
| "Days Open" | False | spaces are fine |
| "" | False | empty text empties the cell |
| "1001" | True | Excel would store the number 1001 |
| "1-2" | True | Excel would store a date |
| "W-01" | True | any digit counts, to be safe |
| "TRUE" | True | Excel would store the logical TRUE |
| " false " | True | after TRIM, in any case |
| "=A1" | True | Excel would store a formula |
| "-" | True | a leading minus could start a formula |
| "#N/A" | True | Excel would store the error value |
| "'quoted" | True | a leading ' would be hidden |
| "café" | True | a character above code 127 counts, to be safe |

### SAFE_NAME (VCL 0.3)

| Input | Result | Why |
| --- | --- | --- |
| "J-101" | "J-101" | an ordinary job number |
| "  J-101 ⍽" | "J-101" | tidied like CLEAN_TEXT |
| "Report  2026.pdf" | "Report 2026.pdf" | inner spaces joined, extension kept |
| 1001 | "1001" | a number becomes its text |
| (empty) | "" | empty |
| ".." | "" | can't climb out of the folder |
| "." | "" | the folder itself |
| "..\..\Windows" | "" | a separator is never allowed |
| "A/B" | "" | neither is / |
| "J*1" | "" | wildcards are refused |
| "J?1" | "" | wildcards are refused |
| "C:x" | "" | a drive colon is refused |
| "J"1" | "" | quotes are refused |
| "J<1>" | "" | angle brackets are refused |
| "J|1" | "" | a pipe is refused |
| "J-1." | "" | Windows drops a trailing dot, so it is refused |
| "CON" | "" | a reserved device name |
| "con.txt" | "" | reserved with any extension and any case |
| "COM1" | "" | COM1 to COM9 are reserved |
| "LPT9.pdf" | "" | LPT1 to LPT9 are reserved |
| "CONSOLE" | "CONSOLE" | only the exact device names are reserved |
| "COM10" | "COM10" | COM10 is an ordinary name |
| "J\x01-1" | "J-1" | control characters are removed, as CLEAN_TEXT |
| "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx" | "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx" | 100 characters is the limit |
| "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx" | "" | over 100 characters is refused |

### PATH_JOIN (VCL 0.3)

| Input | Result | Why |
| --- | --- | --- |
| "C:\Jobs", "J-101", "r1.pdf" | "C:\Jobs\J-101\r1.pdf" | base, then each name |
| "C:\Jobs\", "J-101" | "C:\Jobs\J-101" | a trailing separator on base is dropped |
| "\\server\share", "J-101" | "\\server\share\J-101" | a network (UNC) root works |
| "C:\Jobs", "" | "" | an empty part gives an empty path |
| "C:\Jobs", ".." | "" | an unsafe part gives an empty path |
| "C:\Jobs", "A\B" | "" | a part may not hide a separator |
| "", "J-101" | "" | no root, no path |

### FILE_NAME (revision 3, VEP 0.4)

| Input | Result | Why |
| --- | --- | --- |
| "C:\Jobs\J-101\r1.pdf" | "r1.pdf" | the part after the last backslash |
| "r1.pdf" | "r1.pdf" | no backslash: all of it |
| "C:\Jobs\" | "" | a folder path ending in a backslash |
| "\\server\share\a b.pdf" | "a b.pdf" | a network path; spaces are kept |
| "C:/Jobs/r1.pdf" | "C:/Jobs/r1.pdf" | only a backslash separates |
| (empty) | "" | empty |

### DEDUPE keys (VCL 0.4)

| Input | Result | Why |
| --- | --- | --- |
| ["S1", "W1"], ["S1", "W1"] | True | the same values in every key column |
| ["S1", "W1"], ["s1", "w1"] | True | case is ignored, as in dictionary keys (S10) |
| [1001], ["1001"] | True | a number and the same number as text |
| [2.5], ["2.50"] | False | text is compared as typed: 2.50 is not 2.5 |
| ["S1"], ["S1 "] | False | no TRIM: a trailing space makes another key |
| ["S1", "W1"], ["S1", "W2"] | False | one key column differs |
| ["ab", "c"], ["a", "bc"] | False | the columns stay apart |
| [(empty)], [(empty)] | False | an empty key: never a duplicate (data is missing) |
| ["S1", "  "], ["S1", "  "] | False | a key cell of spaces only counts as empty |
| [#N/A], [#N/A] | False | an error in a key cell: never a duplicate |
| [date 2026-10-15], [date 2026-10-15] | True | the same date |

### FIND (revision 3, VEP 0.4)

| Input | Result | Why |
| --- | --- | --- |
| ["Seam", "Welder"], "welder", "whole" | 2 | whole: SAME, so case is ignored |
| [" W1 ", "W1"], "W1", "whole" | 1 | whole: spaces at the ends are ignored |
| ["W12", "W1"], "W1", "whole" | 2 | whole: all of the text must match |
| [25, "25"], "25.0", "whole" | 0 | numbers compare as text: 25.0 is not 25 |
| [25], "25", "whole" | 1 | a number and the same number as text |
| ["a", (empty), "b"], "", "whole" | 2 | whole: an empty x finds the first empty cell |
| ["a", "b"], "c", "whole" | 0 | not there: NOTHING |
| ["W12", "W1"], "W1", "part" | 1 | part: the first cell that holds the text |
| ["Seam S-12 repair"], " s-12 ", "part" | 1 | part: TRIM(x), case ignored |
| [1250], 25, "part" | 1 | part: a number is its text |
| ["a", "b"], "  ", "part" | 0 | part: nothing to look for finds NOTHING |
| [date 2026-10-15], "2026-10-15", "whole" | 1 | a date is its yyyy-mm-dd text |

### Sheet names for ADD_SHEET (VCL 0.4)

| Input | Result | Why |
| --- | --- | --- |
| "Summary" | True | a plain name |
| "NCR Log 2026-10" | True | spaces inside, digits and a dash |
| "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" | True | 31 characters, the most Excel allows |
| "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" | False | 32 characters is one too many |
| "" | False | empty |
| " Summary" | False | a space at the start |
| "Summary " | False | a space at the end |
| "Q1/Q2" | False | a slash |
| "Rev[2]" | False | square brackets |
| "Time 10:30" | False | a colon |
| "What?" | False | a question mark |
| "'Draft" | False | an apostrophe first |
| "Bob's" | True | an apostrophe inside is fine |
| "history" | False | History is kept by Excel, in any case |
| 2026 | True | a number is its text |

### LIST_FILES (VCL 0.4)

| Input | Result | Why |
| --- | --- | --- |
| ["b.pdf", "A.PDF", "c.txt"], "*.pdf" | ["A.PDF", "b.pdf"] | case ignored, then name order |
| ["x.pdfa", "x.pdf"], "*.pdf" | ["x.pdf"] | the whole name: *.pdf is not x.pdfa |
| ["r1.pdf", "r10.pdf", "r2.pdf"], "r?.pdf" | ["r1.pdf", "r2.pdf"] | ? is exactly one character |
| ["r1.pdf", "r10.pdf", "r2.pdf"], "r*.pdf" | ["r1.pdf", "r10.pdf", "r2.pdf"] | * is any run; the order is by character, not by number |
| ["a.pdf", "A.pdf"], "*" | ["A.pdf", "a.pdf"] | the same name in two cases: as written breaks the tie |
| ["ab[1].txt", "ab1.txt"], "ab[1].txt" | ["ab[1].txt"] | [ ] are themselves, not a set |
| ["a.pdf"], "*.txt" | [] | nothing matches: an empty list |
| [], "*" | [] | an empty folder |

### READ_TEXT (VCL 0.4)

| Input | Result | Why |
| --- | --- | --- |
| "J-101 Shell" | "J-101 Shell" | one line, as it is |
| "J-101\x0d\n" | "J-101" | one final CR LF is dropped |
| "J-101\n" | "J-101" | one final LF is dropped |
| "a\x0d\nb\x0d\n\x0d\n" | "a\x0d\nb\x0d\n" | only one: the inner and the other line breaks stay |
| "a\tb" | "a\tb" | a tab is text |
| "" | "" | an empty file |
| "café" | (empty) | an accented letter is not text: it reads differently with each code page |
| "x\x00y" | (empty) | a NUL byte is not text |

### Text for READ_TEXT and WRITE_TEXT (VCL 0.4)

| Input | Result | Why |
| --- | --- | --- |
| "done: J-101" | True | printable ASCII |
| "a\x0d\nb" | True | line breaks |
| "€ 5" | False | the euro sign is not ASCII |
