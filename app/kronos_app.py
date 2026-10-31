"""KGC KRONOS: the window (item 5 Phase 2b, docs/kronos-app.md section 2, D74 to D76).

    python app/kronos_app.py [--data DIR] [--ai fake --replies FILE] [--today YYYY-MM-DD]
    python app/kronos_app.py --smoke     one practice request through all four steps, then close

A four-step wizard: 1 Ask, 2 Questions (only when Codex asks), 3 Review, 4 Save. The window holds no
logic: it shows what app/kronos_controller.py says and calls its methods. While the AI or the checks
work, a panel covers the step with what is happening and a Cancel button; the window reads the
controller's events every 100 ms, so it never freezes. --smoke opens the window with the practice AI,
answers its question, ticks the assumptions, saves the module and closes: exit 0 when all of that
worked (CI runs it under a virtual display; KRONOS.exe's self-test will too).
"""
import argparse
import datetime
import json
import os
import sys
import tempfile
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import tkinter.font as tkfont

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from kronos_controller import Controller, ROOT  # noqa: E402
from kronos_ai import FakeAI  # noqa: E402

VERSION = "0.2"
EXAMPLE = ("In the Calibration sheet, list every instrument due for calibration in the next 15 days on a "
           "sheet called Due List.")
TIPS = ["Name the sheets and the column headers it should use.",
        "Say what must change, and what must stay as it is.",
        "Practice requests with invented data only, for now."]
NEVER = ["delete or move files (it only adds copies)", "use the internet or start other programs",
         "change Excel's security settings", "change anything outside what it lists above"]
STEP_NAMES = {"ask": "Ask", "question": "Questions", "review": "Review", "save": "Save"}
COLOURS = {"bg": "#f4f6f8", "panel": "#ffffff", "text": "#1d2733", "muted": "#5b6a79", "line": "#d4dbe2",
           "accent": "#1f5f99", "ok_bg": "#e5f3ea", "ok": "#1e6b3a", "warn_bg": "#fdf3df", "warn": "#8a5300",
           "note_bg": "#fbe9e7", "note": "#a3321d"}


def pick(families, wanted, fallback):
    return next((name for name in wanted if name in families), fallback)


class App:
    def __init__(self, root, controller):
        self.root, self.c = root, controller
        self.saved_once = False
        root.title(f"KGC KRONOS {VERSION}")
        root.minsize(900, 640)
        root.geometry("1040x720")
        root.configure(bg=COLOURS["bg"])
        self.fonts()
        self.style()
        self.build()
        root.bind("<Control-Return>", lambda e: self.primary())
        self.refresh()
        root.after(100, self.tick)
        if not self.c.checklist_done:
            root.after(200, self.checklist)

    # ---------- look ----------

    def fonts(self):
        families = set(tkfont.families(self.root))
        sans = pick(families, ("Segoe UI", "Noto Sans", "DejaVu Sans", "Liberation Sans"), "TkDefaultFont")
        mono = pick(families, ("Consolas", "DejaVu Sans Mono", "Liberation Mono"), "TkFixedFont")
        self.f_body = (sans, 10)
        self.f_small = (sans, 9)
        self.f_bold = (sans, 10, "bold")
        self.f_title = (sans, 15, "bold")
        self.f_head = (sans, 12, "bold")
        self.f_mono = (mono, 10)
        tkfont.nametofont("TkDefaultFont").configure(family=sans, size=10)

    def style(self):
        s = ttk.Style(self.root)
        s.theme_use("vista" if "vista" in s.theme_names() else "clam")
        bg, panel = COLOURS["bg"], COLOURS["panel"]
        s.configure("TFrame", background=bg)
        s.configure("Panel.TFrame", background=panel)
        s.configure("TLabel", background=bg, foreground=COLOURS["text"], font=self.f_body)
        s.configure("Panel.TLabel", background=panel)
        s.configure("Muted.TLabel", background=panel, foreground=COLOURS["muted"], font=self.f_small)
        s.configure("Title.TLabel", font=self.f_title, foreground=COLOURS["accent"])
        s.configure("Head.TLabel", background=panel, font=self.f_head)
        s.configure("Section.TLabel", background=panel, font=self.f_bold)
        s.configure("Step.TLabel", font=self.f_body, foreground=COLOURS["muted"])
        s.configure("StepNow.TLabel", font=self.f_bold, foreground=COLOURS["accent"])
        s.configure("StepDone.TLabel", font=self.f_body, foreground=COLOURS["ok"])
        s.configure("TCheckbutton", background=panel, font=self.f_body)
        s.configure("Primary.TButton", font=self.f_bold)
        s.configure("TButton", font=self.f_body)

    def label(self, parent, text="", style="Panel.TLabel", wrap=760, **kw):
        return ttk.Label(parent, text=text, style=style, wraplength=wrap, justify="left", **kw)

    def text_box(self, parent, height, mono=False):
        box = tk.Text(parent, height=height, wrap="word", font=self.f_mono if mono else self.f_body, relief="solid",
                      borderwidth=1, highlightthickness=1, highlightcolor=COLOURS["accent"],
                      highlightbackground=COLOURS["line"], padx=8, pady=6, undo=True)
        return box

    # ---------- the frame ----------

    def build(self):
        top = ttk.Frame(self.root, padding=(18, 12, 18, 4))
        top.pack(fill="x")
        ttk.Label(top, text="KGC KRONOS", style="Title.TLabel").pack(side="left")
        ttk.Label(top, text=f"  {VERSION}", style="Step.TLabel").pack(side="left", pady=(6, 0))
        self.bar = ttk.Frame(self.root, padding=(18, 2, 18, 8))
        self.bar.pack(fill="x")
        self.bar_labels = {}
        for i, step in enumerate(("ask", "question", "review", "save")):
            if i:
                ttk.Label(self.bar, text="  >  ", style="Step.TLabel").pack(side="left")
            self.bar_labels[step] = ttk.Label(self.bar, style="Step.TLabel")
            self.bar_labels[step].pack(side="left")
        ttk.Separator(self.root).pack(fill="x")

        self.body = ttk.Frame(self.root, style="Panel.TFrame")
        self.body.pack(fill="both", expand=True)
        self.foot = ttk.Frame(self.root, padding=(18, 10))
        self.foot.pack(fill="x", side="bottom")
        self.primary_btn = ttk.Button(self.foot, style="Primary.TButton", command=self.primary)
        self.primary_btn.pack(side="right")
        self.back_btn = ttk.Button(self.foot, text="Back", command=self.back)
        self.back_btn.pack(side="right", padx=(0, 8))

        self.steps = {"ask": self.build_ask(), "question": self.build_question(), "review": self.build_review(),
                      "save": self.build_save()}
        self.working = self.build_working()
        self.shown_step = None

    def page(self):
        frame = ttk.Frame(self.body, style="Panel.TFrame", padding=(28, 20))
        return frame

    def build_ask(self):
        f = self.page()
        ttk.Label(f, text="What should Excel do?", style="Head.TLabel").pack(anchor="w")
        self.label(f, "Describe the job in plain words. Codex writes a recipe, KRONOS tests it on practice "
                      "data, and you get a macro to import into Excel.", "Muted.TLabel").pack(anchor="w", pady=(2, 10))
        self.request_box = self.text_box(f, 7)
        self.request_box.pack(fill="x")
        self.placeholder(self.request_box, "For example: " + EXAMPLE)
        self.request_box.bind("<KeyRelease>", lambda e: self.refresh_footer())
        tips = ttk.Frame(f, style="Panel.TFrame")
        tips.pack(anchor="w", pady=(12, 0), fill="x")
        ttk.Label(tips, text="Tips", style="Section.TLabel").pack(anchor="w")
        for tip in TIPS:
            self.label(tips, "  -  " + tip).pack(anchor="w")
        self.ask_notice = self.notice_row(f)
        return f

    def notice_row(self, parent):
        row = tk.Frame(parent, bg=COLOURS["note_bg"], padx=12, pady=8)
        text = tk.Label(row, bg=COLOURS["note_bg"], fg=COLOURS["note"], font=self.f_body, justify="left",
                        wraplength=640)
        text.pack(side="left")
        again = ttk.Button(row, text="Try again", command=self.retry)
        again.pack(side="right")
        row.text, row.again = text, again
        return row

    def build_question(self):
        f = self.page()
        self.question_head = ttk.Label(f, style="Head.TLabel")
        self.question_head.pack(anchor="w")
        box = tk.Frame(f, bg=COLOURS["warn_bg"], padx=14, pady=10)
        box.pack(fill="x", pady=(10, 14))
        self.question_text = tk.Label(box, bg=COLOURS["warn_bg"], fg=COLOURS["text"], font=self.f_bold,
                                      justify="left", wraplength=820)
        self.question_text.pack(anchor="w")
        ttk.Label(f, text="Your answer", style="Section.TLabel").pack(anchor="w")
        self.answer_box = self.text_box(f, 4)
        self.answer_box.pack(fill="x", pady=(4, 8))
        self.answer_box.bind("<KeyRelease>", lambda e: self.refresh_footer())
        self.convo_btn = ttk.Button(f, text="Show the conversation", command=self.toggle_conversation)
        self.convo_btn.pack(anchor="w")
        self.convo = self.text_box(f, 8)
        self.question_notice = self.notice_row(f)
        return f

    def build_review(self):
        outer = ttk.Frame(self.body, style="Panel.TFrame")
        canvas = tk.Canvas(outer, bg=COLOURS["panel"], highlightthickness=0)
        scroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        f = ttk.Frame(canvas, style="Panel.TFrame", padding=(28, 18))
        f.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        window = canvas.create_window((0, 0), window=f, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(window, width=e.width))
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))
        self.review_canvas = canvas
        self.banner = tk.Frame(f, padx=16, pady=12)
        self.banner.pack(fill="x")
        self.banner_title = tk.Label(self.banner, font=self.f_head, justify="left")
        self.banner_title.pack(anchor="w")
        self.banner_text = tk.Label(self.banner, font=self.f_body, justify="left", wraplength=820)
        self.banner_text.pack(anchor="w")
        self.banner_btns = tk.Frame(self.banner)
        self.banner_btns.pack(anchor="w", pady=(8, 0))
        ttk.Button(self.banner_btns, text="Try again", command=self.again).pack(side="left")
        ttk.Button(self.banner_btns, text="Change my request", command=self.back).pack(side="left", padx=8)
        self.review_body = ttk.Frame(f, style="Panel.TFrame")
        self.review_body.pack(fill="x", pady=(14, 0))
        change = ttk.Frame(f, style="Panel.TFrame")
        change.pack(fill="x", pady=(18, 0))
        ttk.Label(change, text="Ask for a change", style="Section.TLabel").pack(anchor="w")
        self.change_hint = self.label(change, "", "Muted.TLabel")
        self.change_hint.pack(anchor="w")
        row = ttk.Frame(change, style="Panel.TFrame")
        row.pack(fill="x", pady=(4, 0))
        self.change_box = self.text_box(row, 2)
        self.change_box.pack(side="left", fill="x", expand=True)
        self.change_btn = ttk.Button(row, text="Send change", command=self.send_change)
        self.change_btn.pack(side="left", padx=(8, 0), anchor="n")
        self.review_notice = self.notice_row(f)
        return outer

    def build_save(self):
        f = self.page()
        ttk.Label(f, text="Save the macro", style="Head.TLabel").pack(anchor="w")
        self.save_intro = self.label(f, "", "Muted.TLabel")
        self.save_intro.pack(anchor="w", pady=(2, 10))
        self.fields_frame = ttk.Frame(f, style="Panel.TFrame")
        self.fields_frame.pack(fill="x")
        row = ttk.Frame(f, style="Panel.TFrame")
        row.pack(fill="x", pady=(12, 0))
        self.save_btn = ttk.Button(row, text="Save macro file...", style="Primary.TButton", command=self.save)
        self.save_btn.pack(side="left")
        self.saved_label = self.label(row, "", "Panel.TLabel", wrap=640)
        self.saved_label.pack(side="left", padx=12)
        ttk.Label(f, text="How to use it in Excel", style="Section.TLabel").pack(anchor="w", pady=(20, 2))
        self.steps_label = self.label(f, "")
        self.steps_label.pack(anchor="w")
        ttk.Button(f, text="Copy steps", command=self.copy_steps).pack(anchor="w", pady=(8, 0))
        self.field_widgets = {}
        self.entered = {}   # what the person typed on Save, kept when the step is shown again
        return f

    def build_working(self):
        f = ttk.Frame(self.body, style="Panel.TFrame", padding=(28, 40))
        self.stage_label = ttk.Label(f, style="Head.TLabel")
        self.stage_label.pack(anchor="w")
        row = ttk.Frame(f, style="Panel.TFrame")
        row.pack(fill="x", pady=(12, 4))
        self.progress = ttk.Progressbar(row, mode="indeterminate", length=420)
        self.progress.pack(side="left")
        self.elapsed_label = ttk.Label(row, style="Panel.TLabel")
        self.elapsed_label.pack(side="left", padx=12)
        self.tries_label = self.label(f, "")
        self.tries_label.pack(anchor="w", pady=(14, 0))
        bottom = ttk.Frame(f, style="Panel.TFrame")
        bottom.pack(fill="x", pady=(18, 0))
        self.calls_label = ttk.Label(bottom, style="Muted.TLabel")
        self.calls_label.pack(side="left")
        ttk.Button(bottom, text="Cancel", command=self.c.cancel).pack(side="right")
        return f

    def placeholder(self, box, text):
        box.insert("1.0", text)
        box.configure(fg=COLOURS["muted"])
        box.has_placeholder = True

        def clear(event):
            if getattr(box, "has_placeholder", False):
                box.delete("1.0", "end")
                box.configure(fg=COLOURS["text"])
                box.has_placeholder = False

        box.bind("<FocusIn>", clear)

    def typed(self, box):
        return "" if getattr(box, "has_placeholder", False) else box.get("1.0", "end").strip()

    def set_text(self, box, text):
        box.has_placeholder = False
        box.configure(fg=COLOURS["text"])
        box.delete("1.0", "end")
        box.insert("1.0", text)

    # ---------- showing the state ----------

    def refresh(self):
        c = self.c
        for step, state in c.step_bar():
            number = "1234"[("ask", "question", "review", "save").index(step)]
            text = f"{number}  {STEP_NAMES[step]}"
            if state == "done":
                text = f"✓ {STEP_NAMES[step]}"
            elif state == "skipped":
                text = f"{number}  {STEP_NAMES[step]} (none)"
            self.bar_labels[step].configure(text=text, style={"current": "StepNow.TLabel", "done": "StepDone.TLabel"}
                                            .get(state, "Step.TLabel"))
        target = self.working if c.working else self.steps[c.step]
        if self.shown_step is not target:
            if self.shown_step is not None:
                self.shown_step.pack_forget()
            target.pack(fill="both", expand=True)
            self.shown_step = target
            if c.working:
                self.progress.start(12)
            else:
                self.progress.stop()
        if not c.working:
            getattr(self, f"show_{c.step}")()
        self.refresh_working()
        self.refresh_footer()

    def show_notice(self, row):
        if self.c.notice:
            row.text.configure(text=self.c.notice)
            if self.c.run is not None and self.c.run.status in ("ai_error", "cancelled"):
                row.again.pack(side="right")
            else:
                row.again.pack_forget()
            row.pack(fill="x", pady=(14, 0))
        else:
            row.pack_forget()

    def show_ask(self):
        if self.c.request and self.typed(self.request_box) != self.c.request:
            self.set_text(self.request_box, self.c.request)
        self.show_notice(self.ask_notice)

    def show_question(self):
        self.question_head.configure(text=f"{self.c.ai_title} has a question before it can write the macro")
        self.question_text.configure(text=self.c.question)
        self.show_notice(self.question_notice)

    def show_review(self):
        c = self.c
        review = c.review()
        proven = review["banner"] == "proven"
        bg, fg = (COLOURS["ok_bg"], COLOURS["ok"]) if proven else (COLOURS["warn_bg"], COLOURS["warn"])
        for w in (self.banner, self.banner_title, self.banner_text, self.banner_btns):
            w.configure(bg=bg)
        self.banner_title.configure(fg=fg, text="PROVEN" if proven else "NOT PROVEN")
        tests = review["tests"]
        passed = sum(t["passed"] for t in tests)
        if proven:
            self.banner_text.configure(fg=COLOURS["text"], text=f"Passed every check and {passed} of {len(tests)} "
                                                                f"practice tests.")
            self.banner_btns.pack_forget()
        else:
            self.banner_text.configure(fg=COLOURS["text"], text=f"{c.ai_title} couldn't get it right in 3 tries. "
                                                                "Try again, or change your request.")
            self.banner_btns.pack(anchor="w", pady=(8, 0))
        self.change_hint.configure(text=f"Something not right? Say what to change, and {c.ai_label} rewrites "
                                        "the recipe.")
        for w in self.review_body.winfo_children():
            w.destroy()
        words = review["summary"]
        if words["does"]:
            self.section("What it does", [words["does"]] + [f"Step: {s}" for s in words["steps"]])
        if proven:
            self.section("What it assumes (check these match your workbook)", words["assumes"] or ["Nothing special."])
            self.assumed = tk.BooleanVar(value=c.assumptions_ok)
            ttk.Checkbutton(self.review_body, text="I checked these assumptions", variable=self.assumed,
                            command=lambda: (c.set_assumptions(self.assumed.get()), self.refresh_footer())
                            ).pack(anchor="w", pady=(6, 0))
            self.section("What it changes", words["changes"] or ["Nothing in the workbook: it only reads."])
            self.section("What it never does", NEVER)
        lines = [f"{passed} of {len(tests)} practice tests passed."] if tests else []
        lines += review["tries"]
        if review["strength"]:
            lines.append(f"Test strength: the tests caught {review['strength'][0]} of {review['strength'][1]} "
                         "mistakes planted on purpose.")
        self.section("How it was tested", lines or [review["note"]])
        if tests:
            self.tests_btn = ttk.Button(self.review_body, text="Show the tests", command=self.toggle_tests)
            self.tests_btn.pack(anchor="w", pady=(6, 0))
            self.tests_table = self.build_tests_table(tests)
        self.show_notice(self.review_notice)
        self.review_canvas.yview_moveto(0)

    def section(self, title, lines):
        ttk.Label(self.review_body, text=title, style="Section.TLabel").pack(anchor="w", pady=(12, 2))
        for line in lines:
            self.label(self.review_body, "  -  " + line, wrap=840).pack(anchor="w")

    def build_tests_table(self, tests):
        frame = ttk.Frame(self.review_body, style="Panel.TFrame")
        table = ttk.Treeview(frame, columns=("result", "checks"), show="headings", height=min(len(tests), 6))
        table.heading("result", text="Result")
        table.heading("checks", text="What it checks")
        table.column("result", width=80, stretch=False)
        table.column("checks", width=760)
        for t in tests:
            table.insert("", "end", values=("passed" if t["passed"] else "FAILED",
                                            t["checks"] + ("  -  " + "; ".join(t["why"]) if t["why"] else "")))
        table.pack(fill="x")
        return frame

    def toggle_tests(self):
        if self.tests_table.winfo_ismapped():
            self.tests_table.pack_forget()
            self.tests_btn.configure(text="Show the tests")
        else:
            self.tests_table.pack(fill="x", pady=(6, 0), after=self.tests_btn)
            self.tests_btn.configure(text="Hide the tests")

    def toggle_conversation(self):
        if self.convo.winfo_ismapped():
            self.convo.pack_forget()
            self.convo_btn.configure(text="Show the conversation")
            return
        who = {"user": "You", "assistant": self.c.ai_label}
        self.convo.configure(state="normal")
        self.convo.delete("1.0", "end")
        for m in self.c.conversation():
            self.convo.insert("end", f"{who[m['role']]}:\n{m['text'].strip()}\n\n")
        self.convo.configure(state="disabled")
        self.convo.pack(fill="both", expand=True, pady=(6, 0))
        self.convo_btn.configure(text="Hide the conversation")

    def show_save(self):
        for w in self.fields_frame.winfo_children():
            w.destroy()
        self.field_widgets = {}
        fields = self.c.sheet_fields()
        if fields is None:
            self.save_intro.configure(text="This tool already names the sheets it works on.")
        else:
            self.save_intro.configure(text="Type the real names of your sheets, exactly as they appear on the "
                                           "tabs in Excel.")
            for i, (name, typ, hint) in enumerate(fields):
                what = "sheet" if typ == "Worksheet" else {"String": "text", "Long": "whole number",
                                                          "Double": "number", "Boolean": "TRUE or FALSE"}[typ]
                ttk.Label(self.fields_frame, text=f'The {what} called "{name}" in the recipe',
                          style="Panel.TLabel").grid(row=i, column=0, sticky="w", pady=4)
                var = tk.StringVar(value=self.entered.get(name, ""))
                entry = ttk.Entry(self.fields_frame, textvariable=var, width=34, font=self.f_body)
                entry.grid(row=i, column=1, sticky="w", padx=10)
                problem = ttk.Label(self.fields_frame, style="Muted.TLabel")
                problem.grid(row=i, column=2, sticky="w")
                if hint:
                    ttk.Label(self.fields_frame, text=f"(the practice tests called it {hint})",
                              style="Muted.TLabel").grid(row=i, column=3, sticky="w")
                var.trace_add("write", lambda *a, n=name, t=typ, v=var, p=problem: (
                    self.entered.__setitem__(n, v.get()),
                    p.configure(text=self.c.check_value(t, v.get()) if v.get() else ""), self.refresh_footer()))
                self.field_widgets[name] = (typ, var)
        self.saved_label.configure(text=f"Saved: {self.c.saved}" if self.c.saved else "")
        self.steps_label.configure(text="\n".join(self.c.import_steps()))

    def refresh_working(self):
        c = self.c
        if not c.working:
            return
        self.stage_label.configure(text=c.stage or f"Asking {c.ai_label}")
        minutes, seconds = divmod(c.elapsed(), 60)
        self.elapsed_label.configure(text=f"{minutes}:{seconds:02d}")
        self.tries_label.configure(text="\n".join(c.lines))
        self.calls_label.configure(text=f"Calls to {c.ai_label} so far: {c.calls}")

    def refresh_footer(self):
        c = self.c
        step = c.step
        self.back_btn.configure(state="disabled" if c.working or step == "ask" else "normal")
        if c.working:
            self.primary_btn.configure(text="Working...", state="disabled")
            return
        label, enabled = {
            "ask": (f"Ask {c.ai_label}", c.can_ask(self.typed(self.request_box))),
            "question": ("Send answer", bool(self.typed(self.answer_box))),
            "review": ("Next: Save the macro", c.can_save()),
            "save": ("Start a new request", True),
        }[step]
        self.primary_btn.configure(text=label, state="normal" if enabled else "disabled")
        if step == "save":
            ok = all(not c.check_value(typ, var.get()) for typ, var in self.field_widgets.values())
            self.save_btn.configure(state="normal" if ok else "disabled")
        if step == "review":
            self.change_btn.configure(state="normal")

    def tick(self):
        events = self.c.poll()
        if any(e["type"] in ("done", "failed") for e in events):
            self.refresh()
        elif events or self.c.working:
            if self.c.working and self.shown_step is not self.working:
                self.refresh()
            self.refresh_working()
        self.root.after(100, self.tick)

    # ---------- actions ----------

    def primary(self):
        c = self.c
        if c.working:
            return
        if c.step == "ask" and c.can_ask(self.typed(self.request_box)):
            if c.run is not None and self.typed(self.request_box) != c.request and not messagebox.askyesno(
                    "New request", "This starts a new request. The old one stays in History. Go on?",
                    parent=self.root):
                return
            c.ask(self.typed(self.request_box))
        elif c.step == "question" and self.typed(self.answer_box):
            c.answer(self.typed(self.answer_box))
            self.answer_box.delete("1.0", "end")
        elif c.step == "review" and c.can_save():
            c.go_save()
        elif c.step == "save":
            if not c.saved and not messagebox.askyesno("Start a new request",
                                                       "You haven't saved the macro. Start a new request anyway?",
                                                       parent=self.root):
                return
            c.new_request()
            self.entered = {}
            self.set_text(self.request_box, "")
        self.refresh()

    def back(self):
        self.c.back()
        self.refresh()

    def retry(self):
        self.c.retry()
        self.refresh()

    def again(self):
        self.c.again()
        self.refresh()

    def send_change(self):
        text = self.typed(self.change_box)
        if text and not self.c.working and self.c.step == "review":
            self.c.change(text)
            self.change_box.delete("1.0", "end")
            self.refresh()

    def values(self):
        return {name: var.get() for name, (typ, var) in self.field_widgets.items()}

    def save(self, dest=None):
        c = self.c
        if dest is None:
            default = c.default_save_path()
            default.parent.mkdir(parents=True, exist_ok=True)
            chosen = filedialog.asksaveasfilename(parent=self.root, title="Save the macro file",
                                                  initialdir=str(default.parent), initialfile=default.name,
                                                  defaultextension=".bas", filetypes=[("VBA module", "*.bas")])
            if not chosen:
                return
            dest = chosen
        try:
            c.save(self.values(), dest)
        except (ValueError, OSError) as exc:
            messagebox.showerror("Not saved", str(exc), parent=self.root)
            return
        self.refresh()

    def copy_steps(self):
        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(self.c.import_steps()))

    # ---------- the first-run checklist ----------

    def checklist(self):
        c = self.c
        win = tk.Toplevel(self.root, bg=COLOURS["panel"], padx=24, pady=18)
        win.title("Before you start")
        win.transient(self.root)
        win.resizable(False, False)
        ttk.Label(win, text="Before you start", style="Head.TLabel").pack(anchor="w")
        items = dict((item, (ok, detail)) for item, ok, detail in c.checklist())
        found, detail = items["codex_found"]
        flags_ok, flags = items["codex_flags"]
        mark = lambda ok: "✓" if ok else "✗"   # noqa: E731
        if c.ai_label == "Codex":
            self.label(win, f"{mark(found)}  Codex is installed: {detail}", wrap=560).pack(anchor="w", pady=(10, 0))
        else:
            self.label(win, "✓  The practice AI is ready (practice runs need no Codex)", wrap=560
                       ).pack(anchor="w", pady=(10, 0))
        if found and c.ai_label == "Codex":
            self.label(win, f"{mark(flags_ok)}  Its command line works with this version"
                            + ("" if flags_ok else f": {flags}"), wrap=560).pack(anchor="w")
        controls = tk.BooleanVar(value=c.settings["data_controls_confirmed"])
        practice = tk.BooleanVar(value=c.settings["practice_confirmed"])
        ttk.Checkbutton(win, variable=controls, text='I turned off "Improve the model for everyone" (ChatGPT > '
                        'Settings > Data Controls)').pack(anchor="w", pady=(12, 0))
        self.label(win, f"KRONOS keeps its records in {items['data_folder'][1]}", "Muted.TLabel",
                   wrap=560).pack(anchor="w", pady=(8, 0))
        ttk.Checkbutton(win, variable=practice, text="I'll use practice requests with invented data only, for now"
                        ).pack(anchor="w", pady=(8, 0))
        done = ttk.Button(win, text="Done", style="Primary.TButton")
        done.pack(anchor="e", pady=(16, 0))

        def update(*args):
            done.configure(state="normal" if controls.get() and practice.get() else "disabled")

        def finish():
            c.confirm_checklist(controls.get(), practice.get())
            win.destroy()
            self.refresh()

        controls.trace_add("write", update)
        practice.trace_add("write", update)
        done.configure(command=finish)
        update()
        self.checklist_win = (win, controls, practice, done)


# ---------- --smoke: one practice request through all four steps ----------

def smoke(app, out):
    """Drives the window as a person would; out["ok"] is set when the module is saved."""
    c = app.c
    plan = [
        ("checklist", lambda: hasattr(app, "checklist_win"), lambda: (
            app.checklist_win[1].set(True), app.checklist_win[2].set(True), app.checklist_win[3].invoke())),
        ("ask", lambda: c.checklist_done and c.step == "ask" and not c.working, lambda: (
            app.set_text(app.request_box, "In the Orders sheet, delete every cancelled order, then sort by date."),
            app.primary())),
        ("question", lambda: c.step == "question" and not c.working, lambda: (
            app.set_text(app.answer_box, "The sheet is called Orders."), app.primary())),
        ("review", lambda: c.step == "review" and not c.working and c.banner == "proven", lambda: (
            app.assumed.set(True), c.set_assumptions(True), app.primary())),
        ("save", lambda: c.step == "save", lambda: (
            app.field_widgets["ws"][1].set("Orders"), app.save(c.data / "smoke" / "smoke.bas"))),
        ("saved", lambda: c.saved is not None and c.saved.exists(), lambda: out.update(ok=True)),
    ]
    out.setdefault("done", [])

    def step():
        if plan and plan[0][1]():
            name, cond, act = plan.pop(0)
            act()
            out["done"].append(name)
        if plan:
            app.root.after(50, step)
        else:
            app.root.after(200, app.root.destroy)

    app.root.after(300, step)
    app.root.after(60000, app.root.destroy)   # never hang CI


def main(argv=None):
    ap = argparse.ArgumentParser(prog="kronos_app", description="KGC KRONOS: a request in plain words to a proven "
                                                                "Excel macro.")
    ap.add_argument("--data", help="the data folder (default: Documents/KGC KRONOS)")
    ap.add_argument("--ai", choices=("codex", "fake"))
    ap.add_argument("--replies", help="the practice AI's replies: a JSON list of strings")
    ap.add_argument("--today", type=datetime.date.fromisoformat, help="pins TODAY for the tests")
    ap.add_argument("--smoke", action="store_true", help="run one practice request through every step, then close")
    args = ap.parse_args(argv)
    if os.name == "nt":
        try:   # sharp text on scaled Windows displays
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except (AttributeError, OSError):
            pass
    out = {}
    if args.smoke:
        good = (ROOT / "vcl" / "examples" / "example1_rev3.vcl").read_text(encoding="utf-8")
        ai = FakeAI(["```vcl\n# QUESTION: Which sheet holds the orders?\n```", f"```vcl\n{good}```"])
        controller = Controller(Path(tempfile.mkdtemp()) / "KGC KRONOS", ai=ai, today=datetime.date(2026, 10, 21))
    else:
        from kronos_engine import default_data_folder
        data = Path(args.data) if args.data else default_data_folder()
        ai = FakeAI(json.loads(Path(args.replies).read_text(encoding="utf-8"))) if args.ai == "fake" and args.replies \
            else None
        controller = Controller(data, ai=ai, today=args.today)
    root = tk.Tk()
    app = App(root, controller)
    if args.smoke:
        smoke(app, out)
    root.mainloop()
    if args.smoke:
        print("smoke: " + ("OK, all four steps and the saved module" if out.get("ok") else
                           f"FAILED after {out.get('done')}"))
        return 0 if out.get("ok") else 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
