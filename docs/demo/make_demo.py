"""The KRONOS demo (D88): an invented NCR register, and what KRONOS makes of it with the three demo recipes.

    python docs/demo/make_demo.py

writes, next to this file:
  NCR_Register_Demo.xlsx              the workbook a person gives KRONOS (invented data only: no real jobs, people or
                                      customers)
  NCR_Register_Demo_after_<n>.xlsx    the changed copy each recipe gives, made by compiler/kronos_book.py, the code
                                      KRONOS Offline's page runs
  facts.json                          what each run did (the numbers the Word guide quotes)

The Word guide, KRONOS_Demo_NCR_Register.docx, is made from facts.json by make_guide.js. tests/test_demo.py checks
that every file here is what this script makes today.
"""
import datetime as dt
import io
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent
sys.path.insert(0, str(ROOT / "compiler"))

import openpyxl  # noqa: E402
from openpyxl.styles import Alignment, Font, PatternFill  # noqa: E402

WORKBOOK = "NCR_Register_Demo.xlsx"
SHEET = "NCR Register"
HEADERS = ["NCR No", "Date Raised", "Job No", "Item", "Description", "Severity", "Department", "Status", "Date Closed"]
RECIPES = [   # (file, the arguments the person picks on the page, the after file)
    ("recipe_1_highlight.vcl", {"ws": SHEET}, "NCR_Register_Demo_after_1.xlsx"),
    ("recipe_2_tidy.vcl", {"ws": SHEET}, "NCR_Register_Demo_after_2.xlsx"),
    ("recipe_3_summary.vcl", {"book": "", "ws": SHEET}, "NCR_Register_Demo_after_3.xlsx"),
]
TODAY = dt.date(2026, 11, 16)   # pins the practice tests' TODAY (none of the recipes reads it)

# Invented values only. Jobs, items and defects are made up for the demo; departments, not people, raise NCRs.
JOBS = ["J-9001", "J-9002", "J-9003", "J-9004", "J-9005", "J-9006"]
ITEMS = ["Shell course 2", "Nozzle N3", "Dished end", "Saddle support", "Flange F-12", "Manway cover", "Tube sheet",
         "Lifting lug", "Skirt", "Nameplate bracket"]
DEFECTS = ["Weld undercut over 0.5 mm", "Porosity found by RT", "Wrong material grade on the certificate",
           "Ovality out of tolerance", "Bolt holes misaligned", "Hydrotest stamp missing", "Paint thickness below spec",
           "Dimension 3 mm off the drawing", "Arc strike on the shell", "Lamination at the plate edge", "Root gap too wide",
           "Visual: spatter not removed"]
DEPARTMENTS = ["Welding", "QA", "Machining", "Fabrication", "Stores"]
# The status as it was typed: tidy words, and a few typed in a hurry (the Tidy recipe cleans them)
STATUSES = ["Open", "Closed", "Closed", "On Hold", " open", "Closed", "Open", "OPEN ", "closed", "Closed", "Open",
            "On hold "]


def rows():
    """The 38 NCRs, then two of them typed a second time: 40 rows, in the order they were entered (not by date)."""
    out = []
    for i in range(38):
        raised = dt.date(2026, 7, 1) + dt.timedelta(days=(i * 37) % 110)
        status = STATUSES[i % len(STATUSES)]
        severity = "Major" if i % 3 == 0 or i % 12 in (4, 7) else "Minor"
        closed = raised + dt.timedelta(days=5 + (i * 7) % 20) if status.strip().lower() == "closed" else None
        out.append([f"NCR-2026-{i + 1:03d}", raised, JOBS[i % len(JOBS)], ITEMS[(i * 3) % len(ITEMS)],
                    DEFECTS[(i * 5) % len(DEFECTS)], severity, DEPARTMENTS[(i * 2) % len(DEPARTMENTS)], status, closed])
    out.insert(15, list(out[6]))    # NCR-2026-007 typed again
    out.insert(31, list(out[20]))   # NCR-2026-020 typed again
    return out


def build():
    """The demo workbook's bytes."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET
    ws.append(HEADERS)
    navy = PatternFill(fill_type="solid", start_color="FF0B2A6F", end_color="FF0B2A6F")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFFFF")
        cell.fill = navy
        cell.alignment = Alignment(vertical="center")
    for row in rows():
        ws.append(row)
    for r in range(2, ws.max_row + 1):
        for c in (2, 9):
            ws.cell(r, c).number_format = "dd-mmm-yyyy"
    for letter, width in zip("ABCDEFGHI", (15, 13, 9, 18, 38, 9, 13, 11, 13)):
        ws.column_dimensions[letter].width = width
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"
    about = wb.create_sheet("About")
    lines = ["KRONOS demo workbook", "",
             "Everything in this workbook is invented for a demonstration of KGC KRONOS.",
             "No job, item, defect or date in it is real.",
             "The sheet NCR Register is a made-up non-conformance (NCR) register of 40 rows.",
             "Some Status cells were typed in a hurry, two NCRs were entered twice, and the rows are not in date order:",
             "the demo recipes find and fix exactly that."]
    for i, text in enumerate(lines, start=1):
        about.cell(i, 1).value = text or None
    about["A1"].font = Font(bold=True, size=14)
    about.column_dimensions["A"].width = 100
    stamp = dt.datetime(2026, 11, 16, 12, 0, 0)
    wb.properties.creator = "KGC KRONOS demo"
    wb.properties.created = stamp
    wb.properties.modified = stamp
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def run(data):
    """[(recipe file, the result KRONOS shows, the changed copy's bytes)] for the three recipes on the demo workbook."""
    import kronos_book
    found = []
    for name, args, _ in RECIPES:
        text = (HERE / name).read_text(encoding="utf-8")
        result, copy = kronos_book.apply(text, data, WORKBOOK, args, TODAY)
        if not result["ok"]:
            raise SystemExit(f"{name}: {result['message']}")
        found.append((name, result, copy))
    return found


def facts(data, results):
    """What the Word guide says about the workbook and each run, as plain data."""
    wb = openpyxl.load_workbook(io.BytesIO(data))
    ws = wb[SHEET]
    table = [[("" if v is None else v.strftime("%d-%b-%Y") if isinstance(v, (dt.date, dt.datetime)) else str(v))
              for v in row] for row in ws.iter_rows(min_row=1, max_row=9, values_only=True)]
    out = {"workbook": WORKBOOK, "sheet": SHEET, "rows": ws.max_row - 1, "headers": HEADERS, "first_rows": table,
           "runs": []}
    for (name, args, after), (_, result, copy) in zip(RECIPES, results):
        book = openpyxl.load_workbook(io.BytesIO(copy))
        entry = {"recipe": name, "after": after, "result": result["result"], "sheets": result["sheets"],
                 "changes_total": result["changes_total"], "warnings": [w["text"] for w in result["warnings"]],
                 "args": args}
        if "NCR Summary" in book.sheetnames:
            entry["summary"] = [["" if v is None else str(v) for v in row]
                                for row in book["NCR Summary"].iter_rows(values_only=True)]
        if name.startswith("recipe_2"):
            entry["first_rows_after"] = [[("" if v is None else v.strftime("%d-%b-%Y") if isinstance(v, (dt.date, dt.datetime))
                                           else str(v)) for v in row]
                                         for row in book[SHEET].iter_rows(min_row=1, max_row=9, values_only=True)]
        if name.startswith("recipe_1"):
            entry["highlighted"] = [[str(book[SHEET].cell(r, c).value) for c in (1, 6, 7, 8)]
                                    for r in range(2, book[SHEET].max_row + 1)
                                    if book[SHEET].cell(r, 1).fill.fill_type == "solid"]
        out["runs"].append(entry)
    # chained: Tidy first, then Summary on the tidy copy (what the guide suggests for real use)
    import kronos_book
    tidy = results[1][2]
    text = (HERE / RECIPES[2][0]).read_text(encoding="utf-8")
    result, copy = kronos_book.apply(text, tidy, WORKBOOK, RECIPES[2][1], TODAY)
    book = openpyxl.load_workbook(io.BytesIO(copy))
    out["chained_summary"] = {"result": result["result"],
                              "summary": [["" if v is None else str(v) for v in row]
                                          for row in book["NCR Summary"].iter_rows(values_only=True)]}
    return out


def main():
    data = build()
    (HERE / WORKBOOK).write_bytes(data)
    results = run(data)
    for (_, _, after), (_, _, copy) in zip(RECIPES, results):
        (HERE / after).write_bytes(copy)
    (HERE / "facts.json").write_text(json.dumps(facts(data, results), indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    for name, result, _ in results:
        print(f"{name}: RESULT {result['result']}, {result['changes_total']} cells changed, "
              f"{len(result['warnings'])} warnings")


if __name__ == "__main__":
    main()
