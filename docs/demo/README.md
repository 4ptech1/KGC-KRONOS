# KRONOS demo: an invented NCR register (D88)

A ready-made demo of KRONOS Offline (version 3 or later): a made-up register of 40 non-conformance reports, three
recipes that do everyday jobs on it, what KRONOS gives back, and a Word guide for colleagues and managers. Everything
here is invented: no job, item, defect or date is real.

| File | What it is |
| --- | --- |
| [`KRONOS_Demo_NCR_Register.docx`](https://github.com/4ptech1/KGC-KRONOS/raw/main/docs/demo/KRONOS_Demo_NCR_Register.docx) | The guide, in plain words: the workbook, what KRONOS does, the three results |
| [`NCR_Register_Demo.xlsx`](https://github.com/4ptech1/KGC-KRONOS/raw/main/docs/demo/NCR_Register_Demo.xlsx) | The workbook to give KRONOS |
| [`recipe_1_highlight.vcl`](recipe_1_highlight.vcl) | Job 1: colour every NCR that is Major and still Open |
| [`recipe_2_tidy.vcl`](recipe_2_tidy.vcl) | Job 2: clean the Status words, remove NCRs typed twice, sort newest first |
| [`recipe_3_summary.vcl`](recipe_3_summary.vcl) | Job 3: add an NCR Summary sheet with counts by status and department |
| `NCR_Register_Demo_after_1.xlsx` … `_after_3.xlsx` | What KRONOS gave back for each job, to compare |
| `facts.json` | The numbers the guide quotes, from KRONOS's own runs |
| `shots/` | The screenshots in the guide |

How it is made (from the repo root):

```
python docs/demo/make_demo.py            # the workbook, the three copies and facts.json
python packaging/build_offline.py        # KRONOS Offline, for the screenshots
python docs/demo/take_screenshots.py     # shots/ (needs Playwright, packaging/offline_test/)
node docs/demo/make_guide.js             # the guide (needs the npm package docx)
```

`tests/test_demo.py` checks that the recipes are PROVEN and canonical, and that every committed file is what the
scripts make today. In KRONOS Offline: choose the workbook, paste a recipe, Check the recipe, choose the sheet
NCR Register, Make the changes on a copy, Download the changed copy.
