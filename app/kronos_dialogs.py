"""The KGC KRONOS window's dialogs (item 5 Phase 2c, docs/kronos-app.md section 2, D77): Settings, the
Recipe window, History and Help, and Check a practice run (D81). Like the window, they hold no logic: they show what the controller says
and call its methods. Each is built with the App (app/kronos_app.py), which keeps one of each open.
"""
import os
import tkinter as tk
from tkinter import filedialog, ttk

from kronos_ai import CODEX_COMMAND

STATUS_WORDS = {"proven": "Proven", "not_proven": "Not proven", "question": "Waiting for an answer",
                "ai_error": "No reply from the AI", "cancelled": "Cancelled", "new": "Not started"}
HELP = [
    ("What KRONOS does", "You describe an Excel job in plain words. The AI writes a recipe, KRONOS tests the "
                         "recipe on practice data, and you get a macro file to import into Excel."),
    ("The four steps", "1 Ask: type what Excel should do.\n"
                       "2 Questions: only if the AI needs to ask you something.\n"
                       "3 Review: read what the macro does, assumes and changes, and tick that you checked the "
                       "assumptions.\n"
                       "4 Save: type your real sheet names and save the macro file. A practice run is saved "
                       "beside it: try the macro in a new, empty workbook first, then Check a practice run."),
    ("Staying safe", "KRONOS never opens, changes or runs your workbooks. You run the macro yourself in Excel, "
                     "on a copy of your workbook first.\n"
                     "Use practice requests with invented data only, for now. Keep \"Improve the model for "
                     "everyone\" turned off in ChatGPT's Data Controls."),
    ("Where things are kept", "Every run is saved in the runs folder of the data folder shown below: your "
                              "request, the conversation, the recipes and the report. Nothing goes into the "
                              "project's public code."),
    ("The setup guide", "The setup guide, with every click to install and start KRONOS, comes with KRONOS.exe."),
]


def can_open_folders():
    return hasattr(os, "startfile")   # Windows


class Dialog:
    """A Toplevel that the App tracks; close() tells the App it is gone."""

    def __init__(self, app, title):
        self.app, self.c, self.colours = app, app.c, app.colours
        self.win = tk.Toplevel(app.root, bg=self.colours["panel"], padx=22, pady=16)
        self.win.title(title)
        self.win.transient(app.root)
        self.win.protocol("WM_DELETE_WINDOW", self.close)
        self.win.bind("<Escape>", lambda e: self.close())

    def close(self):
        self.win.destroy()
        self.app.forget(self)

    def sync(self):
        """Called by the App after every refresh: follow the controller."""

    def heading(self, text):
        ttk.Label(self.win, text=text, style="Head.TLabel").pack(anchor="w")

    def muted(self, text, wrap=560):
        return self.app.label(self.win, text, "Muted.TLabel", wrap=wrap)


class SettingsDialog(Dialog):
    def __init__(self, app):
        super().__init__(app, "Settings")
        c, s = self.c, self.c.settings
        self.heading("Settings")
        grid = ttk.Frame(self.win, style="Panel.TFrame")
        grid.pack(fill="x", pady=(10, 0))
        self.win.columnconfigure(0, weight=1)

        ttk.Label(grid, text="Data folder", style="Panel.TLabel").grid(row=0, column=0, sticky="nw", pady=4)
        folder = ttk.Frame(grid, style="Panel.TFrame")
        folder.grid(row=0, column=1, sticky="w", padx=10)
        app.label(folder, str(c.data), "Muted.TLabel", wrap=420).pack(side="left")
        if can_open_folders():
            ttk.Button(folder, text="Open the folder", command=lambda: os.startfile(str(c.data))
                       ).pack(side="left", padx=8)
        ttk.Label(grid, text="(to move it, start KRONOS with --data FOLDER)", style="Muted.TLabel"
                  ).grid(row=1, column=1, sticky="w", padx=10)

        ttk.Label(grid, text="Save macros in", style="Panel.TLabel").grid(row=2, column=0, sticky="w", pady=(10, 4))
        self.macro_folder = tk.StringVar(value=s["macro_folder"])
        row = ttk.Frame(grid, style="Panel.TFrame")
        row.grid(row=2, column=1, sticky="w", padx=10, pady=(10, 4))
        ttk.Entry(row, textvariable=self.macro_folder, width=40, font=app.f_body).pack(side="left")
        ttk.Button(row, text="Browse...", command=self.browse).pack(side="left", padx=6)
        ttk.Label(grid, text="(empty: the macros folder inside the data folder)", style="Muted.TLabel"
                  ).grid(row=3, column=1, sticky="w", padx=10)

        ttk.Label(grid, text="Wait for Codex up to", style="Panel.TLabel").grid(row=4, column=0, sticky="w",
                                                                                pady=(10, 4))
        self.timeout = tk.StringVar(value=str(s["codex_timeout"]))
        row = ttk.Frame(grid, style="Panel.TFrame")
        row.grid(row=4, column=1, sticky="w", padx=10, pady=(10, 4))
        ttk.Spinbox(row, from_=10, to=3600, increment=30, textvariable=self.timeout, width=7,
                    font=app.f_body).pack(side="left")
        ttk.Label(row, text=" seconds", style="Panel.TLabel").pack(side="left")

        self.show_recipe = tk.BooleanVar(value=s["show_recipe"])
        self.strength = tk.BooleanVar(value=s["measure_strength"])
        ttk.Checkbutton(self.win, variable=self.show_recipe, text="Show the recipe (advanced): lets you read and "
                        "edit the recipe behind each macro").pack(anchor="w", pady=(12, 0))
        ttk.Checkbutton(self.win, variable=self.strength, text="Measure how strong the tests are (slower)"
                        ).pack(anchor="w", pady=(4, 0))

        ttk.Label(self.win, text="Codex command line (advanced), one part per line", style="Section.TLabel"
                  ).pack(anchor="w", pady=(14, 2))
        self.command = app.text_box(self.win, 8, mono=True)
        self.command.insert("1.0", "\n".join(s["codex_command"]))
        self.command.pack(fill="x")
        ttk.Button(self.win, text="Reset to default", command=self.reset_command).pack(anchor="w", pady=(4, 0))

        test = ttk.Frame(self.win, style="Panel.TFrame")
        test.pack(fill="x", pady=(14, 0))
        self.test_btn = ttk.Button(test, text=f"Test {c.ai_label}", command=self.test)
        self.test_btn.pack(side="left")
        ttk.Button(test, text="Run the checklist again", command=app.checklist).pack(side="left", padx=8)
        self.test_label = app.label(self.win, "", "Panel.TLabel", wrap=560)
        self.test_label.pack(anchor="w", pady=(6, 0))
        self.error = app.label(self.win, "", "Panel.TLabel", wrap=560)
        self.error.pack(anchor="w", pady=(6, 0))
        self.error.configure(foreground=self.colours["note"])

        buttons = ttk.Frame(self.win, style="Panel.TFrame")
        buttons.pack(fill="x", pady=(12, 0))
        self.save_btn = ttk.Button(buttons, text="Save", style="Primary.TButton", command=self.save)
        self.save_btn.pack(side="right")
        ttk.Button(buttons, text="Cancel", command=self.close).pack(side="right", padx=(0, 8))

    def browse(self):
        chosen = filedialog.askdirectory(parent=self.win, title="Where should KRONOS suggest saving macros?")
        if chosen:
            self.macro_folder.set(chosen)

    def reset_command(self):
        self.command.delete("1.0", "end")
        self.command.insert("1.0", "\n".join(CODEX_COMMAND))

    def changes(self):
        """What the person set, as the controller takes it; ValueError for what is not a whole number."""
        try:
            timeout = int(self.timeout.get().strip())
        except ValueError:
            raise ValueError("Wait for Codex: type a whole number of seconds, from 10 to 3600.") from None
        parts = [line.strip() for line in self.command.get("1.0", "end").splitlines() if line.strip()]
        return {"macro_folder": self.macro_folder.get().strip(), "codex_timeout": timeout,
                "show_recipe": self.show_recipe.get(), "measure_strength": self.strength.get(),
                "codex_command": parts or list(CODEX_COMMAND)}

    def save(self):
        try:
            self.c.apply_settings(self.changes())
        except ValueError as exc:
            self.error.configure(text=str(exc))
            return
        self.close()
        self.app.refresh()

    def test(self):
        self.app.run_test(self.test_label, self.test_btn)

    def sync(self):
        busy = self.c.working
        self.save_btn.configure(state="disabled" if busy else "normal")
        if busy:
            self.error.configure(text="Wait for the work to finish, or cancel it, before saving settings.")


class RecipeWindow(Dialog):
    """The recipe (VCL) behind the macro shown: only with Show the recipe on (D74)."""

    def __init__(self, app):
        super().__init__(app, "Recipe (advanced)")
        self.win.geometry("900x640")
        self.was_working = self.c.working
        self.heading("Recipe (advanced)")
        self.muted("This is the recipe KRONOS tests. Change it only if you know the recipe language. Re-check "
                   "tests your change; Ask Codex to fix sends it back with the problems.", wrap=820).pack(
            anchor="w", pady=(2, 8))
        body = ttk.Frame(self.win, style="Panel.TFrame")
        body.pack(fill="both", expand=True)
        self.gutter = tk.Text(body, width=5, font=app.f_mono, relief="flat", bg=self.colours["bg"],
                              fg=self.colours["muted"], state="disabled", padx=4, pady=6, takefocus=0,
                              highlightthickness=0)
        self.gutter.tag_configure("right", justify="right")
        self.gutter.pack(side="left", fill="y")
        self.text = tk.Text(body, wrap="none", font=app.f_mono, undo=True, relief="solid", borderwidth=1,
                            padx=6, pady=6, highlightthickness=1, highlightbackground=self.colours["line"],
                            yscrollcommand=self.on_scroll)
        self.text.tag_configure("problem", background=self.colours["note_bg"])
        self.text.pack(side="left", fill="both", expand=True)
        self.scroll = ttk.Scrollbar(body, orient="vertical", command=self.scroll_both)
        self.scroll.pack(side="right", fill="y")
        self.text.bind("<KeyRelease>", lambda e: self.number_lines())

        ttk.Label(self.win, text="Problems (click one to go to its line)", style="Section.TLabel"
                  ).pack(anchor="w", pady=(10, 2))
        self.problems = tk.Listbox(self.win, height=5, font=app.f_mono, activestyle="none", relief="solid",
                                   borderwidth=1, exportselection=False, highlightthickness=0)
        self.problems.pack(fill="x")
        self.problems.bind("<<ListboxSelect>>", lambda e: self.jump())
        self.lines = []

        bottom = ttk.Frame(self.win, style="Panel.TFrame")
        bottom.pack(fill="x", pady=(10, 0))
        self.status = app.label(bottom, "", "Panel.TLabel", wrap=420)
        self.status.pack(side="left")
        ttk.Button(bottom, text="Close", command=self.close).pack(side="right")
        self.fix_btn = ttk.Button(bottom, text="Ask Codex to fix", command=self.fix)
        self.fix_btn.pack(side="right", padx=8)
        self.tidy_btn = ttk.Button(bottom, text="Tidy layout", command=self.tidy)
        self.tidy_btn.pack(side="right")
        self.recheck_btn = ttk.Button(bottom, text="Re-check", style="Primary.TButton", command=self.recheck)
        self.recheck_btn.pack(side="right", padx=8)
        self.load()
        self.buttons()

    # -- the text and its line numbers --

    def on_scroll(self, first, last):
        self.scroll.set(first, last)
        self.gutter.yview_moveto(first)

    def scroll_both(self, *args):
        self.text.yview(*args)
        self.gutter.yview(*args)

    def number_lines(self):
        count = int(self.text.index("end-1c").split(".")[0])
        self.gutter.configure(state="normal")
        self.gutter.delete("1.0", "end")
        self.gutter.insert("1.0", "\n".join(str(n) for n in range(1, count + 1)), "right")
        self.gutter.configure(state="disabled")
        self.gutter.yview_moveto(self.text.yview()[0])

    def load(self):
        """Show the recipe and the problems the controller now has."""
        self.text.delete("1.0", "end")
        self.text.insert("1.0", self.c.recipe())
        self.text.edit_reset()
        self.number_lines()
        self.problems.delete(0, "end")
        self.lines = []
        for line, message in self.c.problems():
            self.problems.insert("end", f"line {line}: {message}" if line else message)
            self.lines.append(line)
        self.text.tag_remove("problem", "1.0", "end")
        for line in {n for n in self.lines if n}:
            self.text.tag_add("problem", f"{line}.0", f"{line}.end+1c")
        proven = self.c.banner == "proven"
        count = len(self.lines)
        self.status.configure(text="Proven: it passed every check and test." if proven else
                              self.c.notice or f"Not proven: {count} problem{'' if count == 1 else 's'}.")

    def jump(self):
        picked = self.problems.curselection()
        if picked and self.lines[picked[0]]:
            self.text.see(f"{self.lines[picked[0]]}.0")
            self.text.mark_set("insert", f"{self.lines[picked[0]]}.0")

    # -- the buttons --

    def buttons(self):
        state = "normal" if self.c.settings["show_recipe"] and not self.c.working else "disabled"
        for b in (self.recheck_btn, self.tidy_btn, self.fix_btn):
            b.configure(state=state)

    def typed(self):
        return self.text.get("1.0", "end-1c")

    def work(self, action):
        try:
            action(self.typed())
        except ValueError as exc:
            self.status.configure(text=str(exc))
            return
        self.was_working = True
        self.status.configure(text="Working...")
        self.app.refresh()

    def recheck(self):
        self.work(self.c.recheck)

    def fix(self):
        self.work(self.c.fix)

    def tidy(self):
        try:
            tidy = self.c.tidy(self.typed())
        except ValueError as exc:
            self.status.configure(text=f"Can't tidy yet: {exc}")
            return
        self.text.delete("1.0", "end")
        self.text.insert("1.0", tidy)
        self.number_lines()
        self.status.configure(text="Layout tidied.")

    def sync(self):
        if self.was_working and not self.c.working:
            self.was_working = False
            self.load()
        self.buttons()


class HistoryDialog(Dialog):
    def __init__(self, app):
        super().__init__(app, "History")
        self.win.geometry("900x560")
        self.heading("History")
        self.muted("Every request you have made on this PC. Select one to read its report.", wrap=820).pack(
            anchor="w", pady=(2, 8))
        self.table = ttk.Treeview(self.win, columns=("when", "request", "status"), show="headings", height=8,
                                  selectmode="browse")
        for name, title, width in (("when", "When", 140), ("request", "Request", 520), ("status", "Status", 180)):
            self.table.heading(name, text=title)
            self.table.column(name, width=width, stretch=name == "request")
        self.table.pack(fill="x")
        self.rows = {}
        for row in self.c.history():
            request = row["request"].replace("\n", " ")
            item = self.table.insert("", "end", values=(row["when"], request[:80] + ("..." if len(request) > 80
                                                                                    else ""),
                                                        STATUS_WORDS.get(row["status"], row["status"])))
            self.rows[item] = row
        self.table.bind("<<TreeviewSelect>>", lambda e: self.show())
        self.report = app.text_box(self.win, 10, mono=True)
        self.report.configure(state="disabled")
        self.report.pack(fill="both", expand=True, pady=(8, 0))
        bottom = ttk.Frame(self.win, style="Panel.TFrame")
        bottom.pack(fill="x", pady=(10, 0))
        self.status = app.label(bottom, "", "Panel.TLabel", wrap=420)
        self.status.pack(side="left")
        ttk.Button(bottom, text="Close", command=self.close).pack(side="right")
        self.again_btn = ttk.Button(bottom, text="Use this macro again", style="Primary.TButton",
                                    command=self.use_again, state="disabled")
        self.again_btn.pack(side="right", padx=8)
        if can_open_folders():
            ttk.Button(bottom, text="Open the folder", command=self.open_folder).pack(side="right")
        if not self.rows:
            self.status.configure(text="Nothing here yet: your requests will be listed once you make one.")

    def chosen(self):
        picked = self.table.selection()
        return self.rows[picked[0]] if picked else None

    def show(self):
        row = self.chosen()
        self.report.configure(state="normal")
        self.report.delete("1.0", "end")
        if row:
            try:
                self.report.insert("1.0", self.c.history_report(row["name"]))
            except (OSError, ValueError) as exc:
                self.report.insert("1.0", f"This run has no report: {exc}")
        self.report.configure(state="disabled")
        self.again_btn.configure(state="normal" if row and row["status"] in ("proven", "not_proven")
                                 and not self.c.working else "disabled")
        self.status.configure(text="")

    def use_again(self):
        row = self.chosen()
        if not row:
            return
        try:
            self.c.use_again(row["name"])
        except ValueError as exc:
            self.status.configure(text=str(exc))
            return
        self.close()
        self.app.refresh()

    def open_folder(self):
        row = self.chosen()
        os.startfile(str(self.c.data / "runs" / row["name"]) if row else str(self.c.data / "runs"))


class HelpDialog(Dialog):
    def __init__(self, app, version):
        super().__init__(app, "About KGC KRONOS")
        self.heading(f"KGC KRONOS {version}")
        for title, text in HELP:
            ttk.Label(self.win, text=title, style="Section.TLabel").pack(anchor="w", pady=(12, 0))
            app.label(self.win, text, wrap=560).pack(anchor="w")
        ttk.Separator(self.win).pack(fill="x", pady=(14, 8))
        self.muted(f"AI in use: {self.c.ai_label}\nData folder: {self.c.data}").pack(anchor="w")
        ttk.Button(self.win, text="Close", command=self.close).pack(anchor="e", pady=(10, 0))


class PracticeDialog(Dialog):
    """Check a practice run (D81): the steps, a box for column A of the sheet ForgePilot, and the verdict."""

    def __init__(self, app):
        super().__init__(app, "Check a practice run")
        self.win.geometry("820x620")
        self.heading("Check a practice run")
        app.label(self.win, "\n".join(self.c.practice_steps()), wrap=760).pack(anchor="w", pady=(6, 10))
        ttk.Label(self.win, text="Paste column A of the sheet ForgePilot here", style="Section.TLabel").pack(
            anchor="w")
        self.box = app.text_box(self.win, 9, mono=True)
        self.box.pack(fill="both", expand=True, pady=(4, 8))
        self.verdict = app.label(self.win, "", wrap=760)
        self.verdict.pack(anchor="w")
        self.details = self.muted("", wrap=760)
        self.details.pack(anchor="w", pady=(2, 0))
        bottom = ttk.Frame(self.win, style="Panel.TFrame")
        bottom.pack(fill="x", pady=(10, 0))
        ttk.Button(bottom, text="Close", command=self.close).pack(side="right")
        ttk.Button(bottom, text="Check", style="Primary.TButton", command=self.check).pack(side="right", padx=8)
        if self.c.practice_result:
            self.show(*self.c.practice_result)

    def check(self):
        try:
            self.show(*self.c.check_practice(self.box.get("1.0", "end")))
        except ValueError as exc:
            self.show(False, str(exc), "")

    def show(self, ok, words, details):
        self.verdict.configure(text=("\u2713  " if ok else "\u2717  ") + words,
                               foreground=self.colours["ok"] if ok else self.colours["note"])
        self.details.configure(text=details)
