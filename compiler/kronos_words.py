"""KRONOS's own words for people, shared by the window (app/kronos_app.py) and KRONOS Offline's page (D89).

Kept apart from the window because the window imports tkinter, which a browser's Python (Pyodide) doesn't have.
"""
EXAMPLE = ("In the Calibration sheet, list every instrument due for calibration in the next 15 days on a "
           "sheet called Due List.")
TIPS = ["Name the sheets and the column headers it should use.",
        "Say what must change, and what must stay as it is.",
        "Practice requests with invented data only, for now."]
NEVER = ["delete or move files (it only adds copies)", "use the internet or start other programs",
         "change Excel's security settings", "change anything outside what it lists above"]
