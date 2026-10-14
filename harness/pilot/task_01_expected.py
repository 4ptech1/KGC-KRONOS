"""Expected results of the pilot batch (specs/task_01.vcl), computed from the fixture data.

It mirrors BuildShortageList in Python, takes the same measures as the VBA test procedures in
harness/pilot/task_01.bas, and evaluates each TEST's assertion on them. The result is the
expected Results sheet: compare it with what Excel produces.
"""

SPEC = "task_01"


def _is_empty(v):
    return v is None or (isinstance(v, str) and not v.strip())


def _is_number(v):
    """IS_NUMBER: numbers, and text that is a plain number; false for blanks."""
    if isinstance(v, bool):
        return False
    if isinstance(v, (int, float)):
        return True
    if isinstance(v, str):
        s = v.strip()
        body = s[1:] if s[:1] in "+-" else s
        return body.replace(".", "", 1).isdigit()
    return False


def _cstr(v):
    """VBA CStr of a cell value: Excel stores numbers as Double, and 1001# prints as 1001."""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _column(sheet, header):
    return [row[sheet.headers.index(header)] for row in sheet.rows]


def build_shortage_list(bom, grn):
    """Python mirror of BuildShortageList: the Shortages rows, as [part, required, received, shortage, remark]."""
    received = {}
    for part, qty in zip(_column(grn, "Part No"), _column(grn, "Received Qty")):
        if not _is_empty(part) and _is_number(qty):
            key = _cstr(part).strip().casefold()
            received[key] = received.get(key, 0.0) + float(qty)
    out = []
    for part, req in zip(_column(bom, "Part No"), _column(bom, "Required Qty")):
        if _is_empty(part):
            continue
        part_no = _cstr(part).strip()
        got = received.get(part_no.casefold(), 0.0)
        if not _is_number(req):
            out.append([part_no, None, got, None, "Required qty missing"])
        elif got < float(req):
            remark = "Not received" if got == 0 else "Short"
            out.append([part_no, req, got, float(req) - got, remark])
    return out


def _numeric_sum(values):
    return sum(v for v in values if isinstance(v, (int, float)) and not isinstance(v, bool))


def expected_tests(sheets):
    """One dict per TEST: id, spec line, expected outcome and the measures the VBA records."""
    bom, grn, shortages = sheets["BOM"], sheets["GRN"], sheets["Shortages"]
    rows = build_shortage_list(bom, grn)
    result = len(rows)
    counta_part = 1 + sum(1 for row in rows if not _is_empty(row[0]))  # header + one part per row
    negative = sum(1 for row in rows if isinstance(row[3], float) and row[3] < 0)
    missing = sum(1 for row in rows if row[4] == "Required qty missing")
    bom_sum = _numeric_sum(_column(bom, "Required Qty"))
    grn_sum = _numeric_sum(_column(grn, "Received Qty"))
    pre_counta_short = 1 + sum(1 for v in _column(shortages, "Part No") if not _is_empty(v))

    tests = [
        {"id": "B000-T01", "line": 60, "measures": {
            "RESULT": result, "ERROR_COUNT": 0, "COUNTA(Shortages![Part No])": counta_part,
            'COUNTIF(Shortages![Shortage], "<0")': negative}},
        {"id": "B000-T02", "line": 61, "measures": {
            "RESULT": result, "ERROR_COUNT": 0,
            'COUNTIF(Shortages![Remark], "Required qty missing")': missing}},
        {"id": "B000-T03", "line": 62, "measures": {
            "RESULT": result, "ERROR_COUNT": 0,
            "PRE_SUM(BOM![Required Qty])": bom_sum, "PRE_SUM(GRN![Received Qty])": grn_sum,
            "SUM(BOM![Required Qty])": bom_sum, "SUM(GRN![Received Qty])": grn_sum}},
        {"id": "B000-T04", "line": 63, "measures": {
            "RESULT": -1, "ERROR_COUNT": 1,
            "PRE_COUNTA(Shortages![Part No])": pre_counta_short,
            "COUNTA(Shortages![Part No])": pre_counta_short}},
    ]
    m = [t["measures"] for t in tests]
    checks = [
        m[0]["RESULT"] >= 5 and m[0]["COUNTA(Shortages![Part No])"] == m[0]["RESULT"] + 1
        and m[0]['COUNTIF(Shortages![Shortage], "<0")'] == 0 and m[0]["ERROR_COUNT"] == 0,
        m[1]['COUNTIF(Shortages![Remark], "Required qty missing")'] == 3,
        m[2]["SUM(BOM![Required Qty])"] == m[2]["PRE_SUM(BOM![Required Qty])"]
        and m[2]["SUM(GRN![Received Qty])"] == m[2]["PRE_SUM(GRN![Received Qty])"],
        m[3]["RESULT"] == -1 and m[3]["ERROR_COUNT"] == 1
        and m[3]["COUNTA(Shortages![Part No])"] == m[3]["PRE_COUNTA(Shortages![Part No])"],
    ]
    for test, ok in zip(tests, checks):
        test["spec"] = SPEC
        test["outcome"] = "PASS" if ok else "FAIL"
    info = {"not_received": sum(1 for row in rows if row[4] == "Not received"),
            "short": sum(1 for row in rows if row[4] == "Short"),
            "missing": missing}
    return tests, info
