"""
NJN Comparator — Desktop App
Cross-platform (macOS + Windows). Requires: customtkinter, openpyxl
"""

import os, sys, threading
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk

# So app.py can import comparator.py from the same folder
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from comparator import compare_and_export, rev_label

# ── Theme ─────────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("System")   # follows macOS/Windows dark-mode setting
ctk.set_default_color_theme("blue")

ACCENT      = "#1F6FEB"
GREEN_BTN   = "#2D9F47"
GREEN_HOVER = "#236B35"


# ── File row widget ───────────────────────────────────────────────────────────
class FileRow(ctk.CTkFrame):
    def __init__(self, parent, path, on_remove, **kw):
        super().__init__(parent, fg_color="transparent", **kw)
        self.columnconfigure(0, weight=1)

        name = os.path.basename(path)
        ctk.CTkLabel(
            self, text=f"📄  {name}", anchor="w",
            font=ctk.CTkFont(size=12), wraplength=500
        ).grid(row=0, column=0, sticky="w", padx=(6, 0))

        ctk.CTkButton(
            self, text="✕", width=28, height=26,
            command=on_remove,
            fg_color="transparent",
            hover_color=("#FFCCCC", "#8B1A1A"),
            text_color=("gray40", "gray60"),
            font=ctk.CTkFont(size=13, weight="bold")
        ).grid(row=0, column=1, padx=(6, 4))


# ── Main App ──────────────────────────────────────────────────────────────────
class NJNComparatorApp(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("NJN Comparator")
        self.geometry("700x640")
        self.minsize(620, 520)
        self._files = []          # list of (path, FileRow widget)
        self._out_dir = tk.StringVar(value=os.path.expanduser("~/Desktop"))
        self._build_ui()

    # ── UI construction ───────────────────────────────────────────────────────
    def _build_ui(self):
        self.rowconfigure(1, weight=1)
        self.columnconfigure(0, weight=1)

        # ── Header bar ──────────────────────────────────────────────────────
        bar = ctk.CTkFrame(self, corner_radius=0, height=64,
                           fg_color=("#1A1A2E", "#1A1A2E"))
        bar.grid(row=0, column=0, sticky="ew")
        bar.grid_propagate(False)
        bar.columnconfigure(1, weight=1)

        ctk.CTkLabel(
            bar, text="NJN COMPARATOR",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color="#FFFFFF"
        ).grid(row=0, column=0, padx=20, pady=10, sticky="w")

        ctk.CTkLabel(
            bar, text="Compare PPAP NJN documents and highlight every change",
            font=ctk.CTkFont(size=11), text_color="#8899BB"
        ).grid(row=0, column=1, padx=4, sticky="w")

        # ── Scrollable body ──────────────────────────────────────────────────
        body = ctk.CTkScrollableFrame(self, corner_radius=0)
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)

        pad = {"padx": 20}

        # ── Section: Files ───────────────────────────────────────────────────
        self._section_label(body, row=0, text="NJN Files to Compare")

        files_header = ctk.CTkFrame(body, fg_color="transparent")
        files_header.grid(row=1, column=0, sticky="ew", **pad, pady=(0, 6))
        files_header.columnconfigure(0, weight=1)

        ctk.CTkLabel(
            files_header,
            text="Add 2 or more NJN revision files — the app compares every pair.",
            font=ctk.CTkFont(size=11), text_color=("gray40", "gray60")
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkButton(
            files_header, text="+ Add Files", width=120,
            command=self._add_files
        ).grid(row=0, column=1)

        # File list card
        self._file_card = ctk.CTkFrame(body, corner_radius=10,
                                        fg_color=("gray93", "gray17"))
        self._file_card.grid(row=2, column=0, sticky="ew", **pad, pady=(0, 18))
        self._file_card.columnconfigure(0, weight=1)

        self._empty_label = ctk.CTkLabel(
            self._file_card,
            text="No files added yet.\nClick  '+ Add Files'  to browse for .xlsx files.",
            font=ctk.CTkFont(size=12), text_color=("gray50", "gray55")
        )
        self._empty_label.grid(row=0, column=0, pady=28)

        # ── Section: Output folder ───────────────────────────────────────────
        self._section_label(body, row=3, text="Output Folder")

        out_row = ctk.CTkFrame(body, fg_color="transparent")
        out_row.grid(row=4, column=0, sticky="ew", **pad, pady=(0, 20))
        out_row.columnconfigure(0, weight=1)

        ctk.CTkEntry(out_row, textvariable=self._out_dir,
                     font=ctk.CTkFont(size=12)).grid(row=0, column=0, sticky="ew",
                                                      padx=(0, 10))
        ctk.CTkButton(out_row, text="Browse…", width=100,
                      command=self._browse_out).grid(row=0, column=1)

        # ── Compare button ───────────────────────────────────────────────────
        self._run_btn = ctk.CTkButton(
            body, text="▶   COMPARE NOW",
            font=ctk.CTkFont(size=17, weight="bold"),
            height=52, corner_radius=10,
            fg_color=ACCENT, hover_color="#0D5BC4",
            command=self._start
        )
        self._run_btn.grid(row=5, column=0, sticky="ew", **pad, pady=(0, 10))

        # ── Progress bar ─────────────────────────────────────────────────────
        self._prog = ctk.CTkProgressBar(body, mode="indeterminate", height=6)
        self._prog.grid(row=6, column=0, sticky="ew", **pad, pady=(0, 6))
        self._prog.grid_remove()

        # ── Results card ─────────────────────────────────────────────────────
        self._result_card = ctk.CTkFrame(body, corner_radius=10,
                                          fg_color=("gray93", "gray17"))
        self._result_card.grid(row=7, column=0, sticky="ew", **pad, pady=(4, 10))
        self._result_card.columnconfigure(0, weight=1)
        self._result_card.grid_remove()

        self._log_box = ctk.CTkTextbox(
            self._result_card, height=160,
            font=ctk.CTkFont(family="Courier", size=12),
            state="disabled", wrap="word",
            fg_color="transparent"
        )
        self._log_box.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        # ── Open folder button ───────────────────────────────────────────────
        self._open_btn = ctk.CTkButton(
            body, text="📂   Open Output Folder",
            font=ctk.CTkFont(size=13, weight="bold"),
            height=40, corner_radius=8,
            fg_color=GREEN_BTN, hover_color=GREEN_HOVER,
            command=self._open_out
        )
        self._open_btn.grid(row=8, column=0, pady=(0, 24))
        self._open_btn.grid_remove()

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _section_label(self, parent, row, text):
        f = ctk.CTkFrame(parent, fg_color="transparent", height=32)
        f.grid(row=row, column=0, sticky="ew", padx=20, pady=(14, 2))
        f.columnconfigure(1, weight=1)
        ctk.CTkLabel(f, text=text,
                     font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, sticky="w")
        sep = ctk.CTkFrame(f, height=2, fg_color=("gray80", "gray30"))
        sep.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=8)

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select NJN Excel files",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        for p in paths:
            if p not in [f[0] for f in self._files]:
                self._append_file(p)

    def _append_file(self, path):
        if not self._files:
            self._empty_label.grid_remove()

        idx = len(self._files)

        def remove(p=path):
            # find and destroy
            for i, (fp, row_w) in enumerate(self._files):
                if fp == p:
                    row_w.grid_forget()
                    row_w.destroy()
                    self._files.pop(i)
                    break
            if not self._files:
                self._empty_label.grid()

        row_widget = FileRow(self._file_card, path, on_remove=remove)
        row_widget.grid(row=idx + 1, column=0, sticky="ew", padx=6, pady=2)
        self._files.append((path, row_widget))

    def _browse_out(self):
        d = filedialog.askdirectory(title="Choose output folder",
                                    initialdir=self._out_dir.get())
        if d:
            self._out_dir.set(d)

    def _open_out(self):
        folder = self._out_dir.get()
        if sys.platform == "darwin":
            os.system(f'open "{folder}"')
        elif sys.platform == "win32":
            os.startfile(folder)
        else:
            os.system(f'xdg-open "{folder}"')

    def _log(self, text):
        self._log_box.configure(state="normal")
        self._log_box.insert("end", text + "\n")
        self._log_box.configure(state="disabled")
        self._log_box.see("end")

    # ── Comparison runner ─────────────────────────────────────────────────────
    def _start(self):
        paths = [f[0] for f in self._files]
        if len(paths) < 2:
            messagebox.showwarning("Need more files",
                                   "Please add at least 2 NJN files to compare.")
            return
        out_dir = self._out_dir.get()
        if not os.path.isdir(out_dir):
            messagebox.showerror("Folder not found",
                                 f"Output folder does not exist:\n{out_dir}")
            return

        # Reset UI
        self._run_btn.configure(state="disabled", text="Running…")
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")
        self._result_card.grid()
        self._open_btn.grid_remove()
        self._prog.grid()
        self._prog.start()

        threading.Thread(target=self._worker,
                         args=(paths, out_dir), daemon=True).start()

    def _worker(self, paths, out_dir):
        from datetime import datetime
        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        files = [(p,) + rev_label(p) for p in paths]
        pairs = [(i, j) for i in range(len(files))
                         for j in range(i + 1, len(files))]
        errors = []

        for i, j in pairs:
            pa, rev_a, lbl_a = files[i]
            pb, rev_b, lbl_b = files[j]
            fname    = f"NJN_{rev_a}_vs_{rev_b}_{stamp}.xlsx"
            out_path = os.path.join(out_dir, fname)
            try:
                n_chg, n_add, n_del = compare_and_export(
                    pa, pb, out_path, lbl_a, lbl_b
                )
                self.after(0, self._log,
                           f"✓  {lbl_a}  →  {lbl_b}\n"
                           f"   {n_chg} cells changed · "
                           f"{n_add} rows added · {n_del} rows removed\n"
                           f"   📄 {fname}\n")
            except Exception as e:
                errors.append(str(e))
                self.after(0, self._log, f"✗  {lbl_a} → {lbl_b}: ERROR — {e}\n")

        self.after(0, self._finish, errors)

    def _finish(self, errors):
        self._prog.stop()
        self._prog.grid_remove()
        self._run_btn.configure(state="normal", text="▶   COMPARE NOW")
        self._open_btn.grid()
        if errors:
            messagebox.showerror("Errors during comparison", "\n\n".join(errors))


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = NJNComparatorApp()
    app.mainloop()
