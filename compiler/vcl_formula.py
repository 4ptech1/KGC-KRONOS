"""SET_FORMULA's formula check (VEP 0.4 Phase 6, D43).

formula_problems(text) -> a list of messages, empty when the formula may be written. The rules:
  - it starts with "=" and is at most 8192 characters (Excel's limit);
  - every function is on EXCEL_2016, a list of Excel 2016 functions that read nothing outside the
    workbook and mean the same on every PC (an allowlist, so a Microsoft 365 function such as
    XLOOKUP, FILTER, LET or TEXTJOIN is refused, and so is any function nobody listed yet);
  - no reference to another workbook or a table ("[" outside text; with style=r1c1 an offset such
    as RC[-2] is a reference, not a bracket), no "@" (implicit intersection, Microsoft 365) and no
    spill reference (A1#);
  - no _xlfn. or _xlws. prefix (how Excel stores functions newer than the file format).
Text in "double quotes" and sheet names in 'single quotes' are not looked into.
"""
import re

EXCEL_2016 = frozenset("""
ABS AGGREGATE AND AVERAGE AVERAGEA AVERAGEIF AVERAGEIFS CEILING CEILING.MATH CHAR CHOOSE CLEAN CODE
COLUMN COLUMNS CONCATENATE COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS DATE DATEDIF DAY DAYS EDATE EOMONTH
ERROR.TYPE EVEN EXACT EXP FALSE FIND FLOOR FLOOR.MATH HLOOKUP HOUR IF IFERROR IFNA INDEX INT ISBLANK
ISERR ISERROR ISEVEN ISLOGICAL ISNA ISNONTEXT ISNUMBER ISODD ISOWEEKNUM ISTEXT LARGE LEFT LEN LN LOG
LOG10 LOOKUP LOWER MATCH MAX MAXA MEDIAN MID MIN MINA MINUTE MOD MONTH MROUND N NA NETWORKDAYS
NETWORKDAYS.INTL NOT NOW NUMBERVALUE ODD OR PERCENTILE.INC PI POWER PRODUCT PROPER QUARTILE.INC
RANK.EQ REPLACE REPT RIGHT ROUND ROUNDDOWN ROUNDUP ROW ROWS SEARCH SECOND SIGN SMALL SQRT STDEV.P
STDEV.S SUBSTITUTE SUBTOTAL SUM SUMIF SUMIFS SUMPRODUCT T TIME TODAY TRIM TRUE TRUNC UPPER VAR.P VAR.S
VLOOKUP WEEKDAY WEEKNUM WORKDAY WORKDAY.INTL XOR YEAR YEARFRAC
""".split())
TEXT_OR_SHEET = re.compile(r'"(?:[^"]|"")*"|\'(?:[^\']|\'\')*\'')
FUNCTION = re.compile(r"([A-Za-z_][A-Za-z0-9_.]*)\s*\(")
SPILL = re.compile(r"[A-Za-z0-9_$)]#")
R1C1_OFFSET = re.compile(r"(?i)(?<=[RC])\[-?\d+\]")


def formula_problems(text, r1c1=False):
    """What stops SET_FORMULA from writing this formula, as messages (none: it may be written)."""
    problems = []
    if not text.startswith("="):
        problems.append("a formula starts with =")
    if len(text) > 8192:
        problems.append("a formula has at most 8192 characters")
    bare = TEXT_OR_SHEET.sub('""', text)
    if r1c1:
        bare = R1C1_OFFSET.sub("", bare)
    if re.search(r"(?i)_xl(fn|ws|udf)\.", bare):
        problems.append("_xlfn., _xlws. and _xludf. mark functions newer than Excel 2016")
    if "[" in bare:
        problems.append("[ ... ] refers to another workbook or a table; a formula reads this workbook only")
    if "@" in bare:
        problems.append("@ is Microsoft 365 (implicit intersection)")
    if SPILL.search(bare):
        problems.append("A1# is a spill reference (Microsoft 365)")
    for name in FUNCTION.findall(bare):
        if name.upper() not in EXCEL_2016:
            problems.append(f"{name.upper()} is not on the Excel 2016 list of SET_FORMULA functions (D43)")
    return problems
