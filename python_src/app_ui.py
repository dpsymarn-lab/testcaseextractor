#!/usr/bin/env python3
"""
CBAM Test Case Extractor - GUI wrapper around extract_test_cases.py
Run with:  python app_ui.py
Requires:  pip install python-docx (for .docx input)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import sys
import io

# ── make sure the extractor module is importable from the same folder ──────
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from extract_test_cases import load_file, extract_test_cases, write_zephyr_csv

# ── colour palette ──────────────────────────────────────────────────────────
BG        = "#0f1117"
SURFACE   = "#1a1d27"
CARD      = "#22263a"
ACCENT    = "#4f8ef7"
ACCENT2   = "#3ecf8e"
TEXT      = "#e8eaf0"
MUTED     = "#6b7280"
DANGER    = "#f87171"
BORDER    = "#2e3347"
FONT_BODY = ("Helvetica Neue", 11)
FONT_MONO = ("Menlo", 10)
FONT_H1   = ("Helvetica Neue", 20, "bold")
FONT_H2   = ("Helvetica Neue", 13, "bold")
FONT_SMALL= ("Helvetica Neue", 9)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CBAM Test Case Extractor")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.geometry("780x680")
        self.minsize(640, 540)

        self.files = []          # list of selected input paths
        self.output_path = tk.StringVar(value="test_cases_zephyr.csv")

        self._build()

    # ── Layout ───────────────────────────────────────────────────────────────

    def _build(self):
        # ── header bar ──
        hdr = tk.Frame(self, bg=ACCENT, height=4)
        hdr.pack(fill="x")

        top = tk.Frame(self, bg=BG, padx=32, pady=20)
        top.pack(fill="x")

        tk.Label(top, text="CBAM Test Case Extractor",
                 font=FONT_H1, fg=TEXT, bg=BG).pack(anchor="w")
        tk.Label(top, text="Extract E2E test cases from .docx files → Jira Zephyr Scale CSV",
                 font=FONT_SMALL, fg=MUTED, bg=BG).pack(anchor="w", pady=(2,0))

        divider = tk.Frame(self, bg=BORDER, height=1)
        divider.pack(fill="x", padx=32)

        body = tk.Frame(self, bg=BG, padx=32, pady=20)
        body.pack(fill="both", expand=True)

        # ── Step 1: input files ──
        self._section(body, "1  Select input files")

        file_row = tk.Frame(body, bg=BG)
        file_row.pack(fill="x", pady=(6,0))

        self.file_list = tk.Listbox(
            file_row, bg=CARD, fg=TEXT, selectbackground=ACCENT,
            font=FONT_MONO, borderwidth=0, highlightthickness=1,
            highlightbackground=BORDER, highlightcolor=ACCENT,
            height=5, activestyle="none"
        )
        self.file_list.pack(side="left", fill="both", expand=True)

        sb = ttk.Scrollbar(file_row, orient="vertical",
                           command=self.file_list.yview)
        sb.pack(side="left", fill="y")
        self.file_list.configure(yscrollcommand=sb.set)

        btn_col = tk.Frame(body, bg=BG)
        btn_col.pack(fill="x", pady=(8,0))
        self._btn(btn_col, "+ Add Files", self._add_files, ACCENT).pack(side="left")
        self._btn(btn_col, "Remove Selected", self._remove_file, SURFACE,
                  fg=MUTED).pack(side="left", padx=(8,0))
        self._btn(btn_col, "Clear All", self._clear_files, SURFACE,
                  fg=DANGER).pack(side="left", padx=(8,0))

        # ── Step 2: output ──
        tk.Frame(body, bg=BG, height=16).pack()
        self._section(body, "2  Output file")

        out_row = tk.Frame(body, bg=BG)
        out_row.pack(fill="x", pady=(6,0))

        self.out_entry = tk.Entry(
            out_row, textvariable=self.output_path,
            bg=CARD, fg=TEXT, insertbackground=TEXT,
            font=FONT_MONO, borderwidth=0,
            highlightthickness=1, highlightbackground=BORDER,
            highlightcolor=ACCENT, relief="flat"
        )
        self.out_entry.pack(side="left", fill="x", expand=True, ipady=6, padx=(0,8))
        self._btn(out_row, "Browse…", self._browse_output, SURFACE,
                  fg=TEXT).pack(side="left")

        # ── Run button ──
        tk.Frame(body, bg=BG, height=20).pack()
        self.run_btn = self._btn(body, "⚡  Extract Test Cases",
                                 self._run, ACCENT, fg="white",
                                 font=("Helvetica Neue", 13, "bold"),
                                 padx=28, pady=10)
        self.run_btn.pack(anchor="center")

        # ── Progress bar ──
        self.progress = ttk.Progressbar(body, mode="indeterminate", length=300)

        # ── Log ──
        tk.Frame(body, bg=BG, height=16).pack()
        self._section(body, "Log")

        log_frame = tk.Frame(body, bg=BG)
        log_frame.pack(fill="both", expand=True, pady=(6,0))

        self.log = tk.Text(
            log_frame, bg=CARD, fg=TEXT, insertbackground=TEXT,
            font=FONT_MONO, borderwidth=0, relief="flat",
            highlightthickness=1, highlightbackground=BORDER,
            highlightcolor=ACCENT, state="disabled", height=8,
            wrap="word"
        )
        self.log.pack(side="left", fill="both", expand=True)
        lsb = ttk.Scrollbar(log_frame, orient="vertical",
                             command=self.log.yview)
        lsb.pack(side="left", fill="y")
        self.log.configure(yscrollcommand=lsb.set)

        self.log.tag_config("ok",    foreground=ACCENT2)
        self.log.tag_config("err",   foreground=DANGER)
        self.log.tag_config("info",  foreground=ACCENT)
        self.log.tag_config("muted", foreground=MUTED)

        self._log("Ready. Add .docx files above and click Extract.", "muted")

    # ── Helpers ──────────────────────────────────────────────────────────────

    def _section(self, parent, title):
        tk.Label(parent, text=title, font=FONT_H2,
                 fg=TEXT, bg=BG).pack(anchor="w")

    def _btn(self, parent, text, cmd, bg, fg=TEXT,
             font=FONT_BODY, padx=14, pady=6):
        b = tk.Button(
            parent, text=text, command=cmd,
            bg=bg, fg=fg, activebackground=ACCENT,
            activeforeground="white", font=font,
            relief="flat", cursor="hand2",
            padx=padx, pady=pady, borderwidth=0
        )
        b.bind("<Enter>", lambda e: b.configure(bg=ACCENT if bg==ACCENT else BORDER))
        b.bind("<Leave>", lambda e: b.configure(bg=bg))
        return b

    def _log(self, msg, tag=""):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    # ── Actions ──────────────────────────────────────────────────────────────

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select input files",
            filetypes=[("Word / Text files", "*.docx *.txt *.md"),
                       ("All files", "*.*")]
        )
        for p in paths:
            if p not in self.files:
                self.files.append(p)
                self.file_list.insert("end", os.path.basename(p))
        if paths:
            self._log(f"Added {len(paths)} file(s).", "info")

    def _remove_file(self):
        sel = self.file_list.curselection()
        if not sel:
            return
        idx = sel[0]
        self.file_list.delete(idx)
        removed = self.files.pop(idx)
        self._log(f"Removed: {os.path.basename(removed)}", "muted")

    def _clear_files(self):
        self.files.clear()
        self.file_list.delete(0, "end")
        self._log("File list cleared.", "muted")

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Save CSV as",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if path:
            self.output_path.set(path)

    def _run(self):
        if not self.files:
            messagebox.showwarning("No files", "Please add at least one input file.")
            return
        out = self.output_path.get().strip()
        if not out:
            messagebox.showwarning("No output", "Please specify an output file path.")
            return
        # disable UI during run
        self.run_btn.configure(state="disabled", text="Processing…")
        self.progress.pack(pady=(12,0))
        self.progress.start(12)
        threading.Thread(target=self._worker, args=(list(self.files), out),
                         daemon=True).start()

    def _worker(self, files, out_path):
        try:
            all_cases = []
            for path in files:
                self._log(f"▶ Processing: {os.path.basename(path)}", "info")
                text, source = load_file(path)
                cases = extract_test_cases(text, source)
                self._log(f"  ✓ {len(cases)} test cases found", "ok")
                all_cases.extend(cases)

            if not all_cases:
                self._log("✗ No test cases found in any file.", "err")
                return

            total_steps = sum(len(tc['Steps']) for tc in all_cases)
            write_zephyr_csv(all_cases, out_path)

            self._log("─" * 48, "muted")
            self._log(f"  Total test cases : {len(all_cases)}", "ok")
            self._log(f"  Total steps      : {total_steps}", "ok")
            self._log(f"  CSV saved to     : {out_path}", "ok")
            self._log("─" * 48, "muted")
            self.after(0, lambda: messagebox.showinfo(
                "Done",
                f"Extracted {len(all_cases)} test cases ({total_steps} steps).\n\nSaved to:\n{out_path}"
            ))
        except Exception as ex:
            self._log(f"✗ Error: {ex}", "err")
            self.after(0, lambda: messagebox.showerror("Error", str(ex)))
        finally:
            self.after(0, self._done)

    def _done(self):
        self.progress.stop()
        self.progress.pack_forget()
        self.run_btn.configure(state="normal", text="⚡  Extract Test Cases")


if __name__ == "__main__":
    app = App()
    app.mainloop()
