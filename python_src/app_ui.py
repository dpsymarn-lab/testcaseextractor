#!/usr/bin/env python3
"""
CBAM Test Case Extractor - GUI entry point for extract_test_cases.py

Run with:
    python app_ui.py

Requires:
    pip install python-docx

This UI is intentionally tolerant of extractor API changes. It supports both:

Old extractor API:
    load_file(path) -> (text, source_name)
    extract_test_cases(text, source_name) -> list[dict]
    write_zephyr_csv(test_cases, output_path)

New extractor API:
    load_file(path) -> list[dict]
    or extract_test_cases_from_docx(path, source_name) -> list[dict]
    or extract_test_cases_from_text(text, source_name) -> list[dict]
    write_zephyr_csv(test_cases, output_path)
"""

import importlib
import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Make sure the extractor module is importable from the same folder as this UI.
APP_DIR = os.path.dirname(os.path.abspath(__file__))
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)


# ---------------------------------------------------------------------------
# Extractor adapter
# ---------------------------------------------------------------------------

class ExtractorAdapter:
    """Small compatibility layer around extract_test_cases.py.

    The GUI should not depend on one exact internal extractor function name.
    This adapter supports the original flat-text API and the newer direct-docx API.
    """

    def __init__(self, module_name="extract_test_cases"):
        try:
            self.module = importlib.import_module(module_name)
        except Exception as exc:
            raise RuntimeError(
                f"Could not import {module_name}.py from:\n{APP_DIR}\n\n"
                "Put app_ui.py and extract_test_cases.py in the same folder.\n\n"
                f"Original error: {exc}"
            ) from exc

        if not hasattr(self.module, "write_zephyr_csv"):
            raise RuntimeError(
                "extract_test_cases.py was imported, but it does not expose "
                "write_zephyr_csv(test_cases, output_path)."
            )

    @property
    def module_path(self):
        return getattr(self.module, "__file__", "unknown")

    def extract_cases(self, path):
        """Return extracted test cases for one input file."""
        source_name = os.path.splitext(os.path.basename(path))[0]
        ext = os.path.splitext(path)[1].lower()

        # Newer direct APIs, if available.
        if ext == ".docx" and hasattr(self.module, "extract_test_cases_from_docx"):
            return self.module.extract_test_cases_from_docx(path, source_name)

        if ext in {".txt", ".md"} and hasattr(self.module, "extract_test_cases_from_text"):
            with open(path, "r", encoding="utf-8") as f:
                return self.module.extract_test_cases_from_text(f.read(), source_name)

        # Possible one-shot APIs some versions may expose.
        for fn_name in (
            "extract_file",
            "extract_cases_from_file",
            "extract_test_cases_from_file",
            "extract_from_file",
        ):
            fn = getattr(self.module, fn_name, None)
            if callable(fn):
                result = fn(path)
                return self._normalize_result(result, source_name)

        # Original API or newer load_file API.
        load_file = getattr(self.module, "load_file", None)
        if callable(load_file):
            result = load_file(path)

            # Newer v2 API: load_file(path) already returns list[dict].
            if isinstance(result, list):
                return result

            # Original API: load_file(path) returns (text, source_name).
            if isinstance(result, tuple) and len(result) >= 2:
                text_or_cases, detected_source = result[0], result[1]

                if isinstance(text_or_cases, list):
                    return text_or_cases

                extract_test_cases = getattr(self.module, "extract_test_cases", None)
                if callable(extract_test_cases):
                    return extract_test_cases(text_or_cases, detected_source)

                extract_text = getattr(self.module, "extract_test_cases_from_text", None)
                if callable(extract_text):
                    return extract_text(text_or_cases, detected_source)

                raise RuntimeError(
                    "load_file(path) returned text, but no compatible parser was found. "
                    "Expected extract_test_cases(text, source) or "
                    "extract_test_cases_from_text(text, source)."
                )

        raise RuntimeError(
            "No compatible extraction API found in extract_test_cases.py. "
            "Expected one of: extract_test_cases_from_docx, load_file, "
            "extract_file, extract_cases_from_file, or extract_test_cases_from_file."
        )

    def write_csv(self, test_cases, output_path):
        self.module.write_zephyr_csv(test_cases, output_path)

    def _normalize_result(self, result, source_name):
        if isinstance(result, list):
            return result
        if isinstance(result, tuple) and len(result) >= 2:
            text_or_cases, detected_source = result[0], result[1]
            if isinstance(text_or_cases, list):
                return text_or_cases
            extract_test_cases = getattr(self.module, "extract_test_cases", None)
            if callable(extract_test_cases):
                return extract_test_cases(text_or_cases, detected_source)
        raise RuntimeError(
            f"Extractor returned an unsupported result for source {source_name!r}: "
            f"{type(result).__name__}"
        )


# ---------------------------------------------------------------------------
# Colour palette
# ---------------------------------------------------------------------------

BG = "#0f1117"
SURFACE = "#1a1d27"
CARD = "#22263a"
ACCENT = "#4f8ef7"
ACCENT2 = "#3ecf8e"
TEXT = "#e8eaf0"
MUTED = "#6b7280"
DANGER = "#f87171"
BORDER = "#2e3347"
FONT_BODY = ("Helvetica Neue", 11)
FONT_MONO = ("Menlo", 10)
FONT_H1 = ("Helvetica Neue", 20, "bold")
FONT_H2 = ("Helvetica Neue", 13, "bold")
FONT_SMALL = ("Helvetica Neue", 9)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CBAM Test Case Extractor")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.geometry("780x680")
        self.minsize(640, 540)

        self.files = []
        self.output_path = tk.StringVar(value="test_cases_zephyr.csv")

        try:
            self.extractor = ExtractorAdapter()
            self.extractor_error = None
        except Exception as exc:
            self.extractor = None
            self.extractor_error = str(exc)

        self._build()

    # ------------------------------------------------------------------
    # Layout
    # ------------------------------------------------------------------

    def _build(self):
        hdr = tk.Frame(self, bg=ACCENT, height=4)
        hdr.pack(fill="x")

        top = tk.Frame(self, bg=BG, padx=32, pady=20)
        top.pack(fill="x")

        tk.Label(top, text="CBAM Test Case Extractor", font=FONT_H1, fg=TEXT, bg=BG).pack(anchor="w")
        tk.Label(
            top,
            text="Extract E2E test cases from .docx/.txt/.md files to Jira Zephyr Scale CSV",
            font=FONT_SMALL,
            fg=MUTED,
            bg=BG,
        ).pack(anchor="w", pady=(2, 0))

        divider = tk.Frame(self, bg=BORDER, height=1)
        divider.pack(fill="x", padx=32)

        body = tk.Frame(self, bg=BG, padx=32, pady=20)
        body.pack(fill="both", expand=True)

        self._section(body, "1  Select input files")

        file_row = tk.Frame(body, bg=BG)
        file_row.pack(fill="x", pady=(6, 0))

        self.file_list = tk.Listbox(
            file_row,
            bg=CARD,
            fg=TEXT,
            selectbackground=ACCENT,
            font=FONT_MONO,
            borderwidth=0,
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=ACCENT,
            height=5,
            activestyle="none",
        )
        self.file_list.pack(side="left", fill="both", expand=True)

        sb = ttk.Scrollbar(file_row, orient="vertical", command=self.file_list.yview)
        sb.pack(side="left", fill="y")
        self.file_list.configure(yscrollcommand=sb.set)

        btn_col = tk.Frame(body, bg=BG)
        btn_col.pack(fill="x", pady=(8, 0))
        self._btn(btn_col, "+ Add Files", self._add_files, ACCENT).pack(side="left")
        self._btn(btn_col, "Remove Selected", self._remove_file, SURFACE, fg=MUTED).pack(side="left", padx=(8, 0))
        self._btn(btn_col, "Clear All", self._clear_files, SURFACE, fg=DANGER).pack(side="left", padx=(8, 0))

        tk.Frame(body, bg=BG, height=16).pack()
        self._section(body, "2  Output file")

        out_row = tk.Frame(body, bg=BG)
        out_row.pack(fill="x", pady=(6, 0))

        self.out_entry = tk.Entry(
            out_row,
            textvariable=self.output_path,
            bg=CARD,
            fg=TEXT,
            insertbackground=TEXT,
            font=FONT_MONO,
            borderwidth=0,
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=ACCENT,
            relief="flat",
        )
        self.out_entry.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 8))
        self._btn(out_row, "Browse...", self._browse_output, SURFACE, fg=TEXT).pack(side="left")

        tk.Frame(body, bg=BG, height=20).pack()
        self.run_btn = self._btn(
            body,
            "Extract Test Cases",
            self._run,
            ACCENT,
            fg="white",
            font=("Helvetica Neue", 13, "bold"),
            padx=28,
            pady=10,
        )
        self.run_btn.pack(anchor="center")

        self.progress = ttk.Progressbar(body, mode="indeterminate", length=300)

        tk.Frame(body, bg=BG, height=16).pack()
        self._section(body, "Log")

        log_frame = tk.Frame(body, bg=BG)
        log_frame.pack(fill="both", expand=True, pady=(6, 0))

        self.log = tk.Text(
            log_frame,
            bg=CARD,
            fg=TEXT,
            insertbackground=TEXT,
            font=FONT_MONO,
            borderwidth=0,
            relief="flat",
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=ACCENT,
            state="disabled",
            height=8,
            wrap="word",
        )
        self.log.pack(side="left", fill="both", expand=True)
        lsb = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        lsb.pack(side="left", fill="y")
        self.log.configure(yscrollcommand=lsb.set)

        self.log.tag_config("ok", foreground=ACCENT2)
        self.log.tag_config("err", foreground=DANGER)
        self.log.tag_config("info", foreground=ACCENT)
        self.log.tag_config("muted", foreground=MUTED)

        if self.extractor_error:
            self._log("Extractor import failed.", "err")
            self._log(self.extractor_error, "err")
        else:
            self._log("Ready. Add input files above and click Extract.", "muted")
            self._log(f"Using extractor: {self.extractor.module_path}", "muted")

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _section(self, parent, title):
        tk.Label(parent, text=title, font=FONT_H2, fg=TEXT, bg=BG).pack(anchor="w")

    def _btn(self, parent, text, cmd, bg, fg=TEXT, font=FONT_BODY, padx=14, pady=6):
        b = tk.Button(
            parent,
            text=text,
            command=cmd,
            bg=bg,
            fg=fg,
            activebackground=ACCENT,
            activeforeground="white",
            font=font,
            relief="flat",
            cursor="hand2",
            padx=padx,
            pady=pady,
            borderwidth=0,
        )
        b.bind("<Enter>", lambda e: b.configure(bg=ACCENT if bg == ACCENT else BORDER))
        b.bind("<Leave>", lambda e: b.configure(bg=bg))
        return b

    def _log(self, msg, tag=""):
        # Tkinter widgets must be updated from the main thread.
        if threading.current_thread() is not threading.main_thread():
            self.after(0, lambda m=msg, t=tag: self._log(m, t))
            return

        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select input files",
            filetypes=[("Word / Text files", "*.docx *.txt *.md"), ("All files", "*.*")],
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
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if path:
            self.output_path.set(path)

    def _run(self):
        if self.extractor_error:
            messagebox.showerror("Extractor error", self.extractor_error)
            return
        if not self.files:
            messagebox.showwarning("No files", "Please add at least one input file.")
            return
        out = self.output_path.get().strip()
        if not out:
            messagebox.showwarning("No output", "Please specify an output file path.")
            return

        self.run_btn.configure(state="disabled", text="Processing...")
        self.progress.pack(pady=(12, 0))
        self.progress.start(12)
        threading.Thread(target=self._worker, args=(list(self.files), out), daemon=True).start()

    def _worker(self, files, out_path):
        try:
            all_cases = []
            for path in files:
                self._log(f"Processing: {os.path.basename(path)}", "info")
                cases = self.extractor.extract_cases(path)
                steps = sum(len(tc.get("Steps", [])) for tc in cases)
                self._log(f"  OK: {len(cases)} test cases found, {steps} steps found", "ok")

                for tc in cases:
                    if not tc.get("Steps"):
                        self._log(f"  Warning: no steps found for {tc.get('ID', '<unknown ID>')}", "muted")

                all_cases.extend(cases)

            if not all_cases:
                self._log("No test cases found in any file.", "err")
                return

            total_steps = sum(len(tc.get("Steps", [])) for tc in all_cases)
            self.extractor.write_csv(all_cases, out_path)

            self._log("-" * 48, "muted")
            self._log(f"  Total test cases : {len(all_cases)}", "ok")
            self._log(f"  Total steps      : {total_steps}", "ok")
            self._log(f"  CSV saved to     : {out_path}", "ok")
            self._log("-" * 48, "muted")

            self.after(
                0,
                lambda: messagebox.showinfo(
                    "Done",
                    f"Extracted {len(all_cases)} test cases ({total_steps} steps).\n\nSaved to:\n{out_path}",
                ),
            )
        except Exception as exc:
            err = str(exc)
            self._log(f"Error: {err}", "err")
            self.after(0, lambda msg=err: messagebox.showerror("Error", msg))
        finally:
            self.after(0, self._done)

    def _done(self):
        self.progress.stop()
        self.progress.pack_forget()
        self.run_btn.configure(state="normal", text="Extract Test Cases")


if __name__ == "__main__":
    app = App()
    app.mainloop()
