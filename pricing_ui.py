"""Drag-and-drop UI for the LNDI drug pricing tool.

Drop a workbook of molecule names on the window, pick the sheet, and it writes an
Excel file next to the source containing:
  * "Results"    - every registered product, priced
  * "Exceptions" - everything that could not be priced, with the reason

Brand rows carry both a 'Final -30%' and a 'Final -40%' column, so no
reference-country choice is needed. Every calculated column is a live Excel
formula rather than a fixed number.

The pricing logic lives in lndi.py (unchanged); this file is only the window.
"""

import os
import subprocess
import sys
import threading
import traceback

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _DND = True
except Exception:
    _DND = False

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import lndi  # noqa: E402

# theme
BG = "#101c2b"
PANEL = "#17273b"
ACCENT = "#2f7fd6"
ACCENT_2 = "#41ab5d"
TEXT = "#e8f0f9"
MUTED = "#93a9c2"
DROP_IDLE = "#1d3552"
DROP_HOVER = "#245079"
DROP_SET = "#1f4a36"

EXTS = (".xls", ".xlsx", ".xlsm", ".xltx", ".csv")

# "Show in Finder" is macOS wording; on Windows/Linux the file manager differs.
REVEAL_LABEL = {"darwin": "Show in Finder", "win32": "Show in Folder"}.get(sys.platform, "Open Folder")


def reveal_in_file_manager(path):
    """Open the OS file manager with `path` selected (macOS/Windows/Linux)."""
    if not (path and os.path.exists(path)):
        return
    if sys.platform == "darwin":
        subprocess.run(["open", "-R", path])
    elif sys.platform == "win32":
        # "/select," and the path must arrive as ONE argument. Passed separately,
        # Windows inserts a space after the comma, explorer cannot parse it, and
        # it silently opens Documents instead of selecting the file.
        # explorer returns exit code 1 even on success, so don't check it.
        subprocess.run(["explorer", f"/select,{os.path.normpath(path)}"])
    else:
        subprocess.run(["xdg-open", os.path.dirname(path)])


class App:
    def __init__(self, root):
        self.root = root
        root.title("Drug Prices")
        root.configure(bg=BG)
        root.geometry("560x600")
        root.minsize(500, 540)

        self.path = None
        self.stratum_path = None
        self.reveal_path = None

        self._style()
        self._build()

    def _style(self):
        s = ttk.Style()
        try:
            s.theme_use("clam")
        except tk.TclError:
            pass
        s.configure("TLabel", background=BG, foreground=TEXT, font=("Helvetica", 12))
        s.configure("Head.TLabel", background=BG, foreground=TEXT, font=("Helvetica", 22, "bold"))
        s.configure("Sub.TLabel", background=BG, foreground=MUTED, font=("Helvetica", 11))
        s.configure("TCombobox", fieldbackground="white", font=("Helvetica", 11))

    def _build(self):
        outer = tk.Frame(self.root, bg=BG)
        outer.pack(fill="both", expand=True, padx=20, pady=18)

        ttk.Label(outer, text="Drug Prices", style="Head.TLabel").pack(anchor="w")
        ttk.Label(outer,
                  text="Drop a file of molecule names below, choose the sheet, then "
                       "price them against Lebanon's National Drug Index.",
                  style="Sub.TLabel", wraplength=500, justify="left").pack(anchor="w", pady=(2, 14))

        row = tk.Frame(outer, bg=BG)
        row.pack(fill="x", pady=(0, 10))
        self.refresh = tk.BooleanVar(value=False)
        tk.Checkbutton(row, text="Fetch today's prices from the website",
                       variable=self.refresh, bg=BG, fg=TEXT, selectcolor=DROP_IDLE,
                       activebackground=BG, activeforeground=TEXT, bd=0,
                       highlightthickness=0, font=("Helvetica", 11)).pack(side="left")

        # stratum rates: found automatically, but selectable for a packaged build
        # where Data/ is not alongside the app
        rates = tk.Frame(outer, bg=BG)
        rates.pack(fill="x", pady=(0, 12))
        self.rates_lbl = ttk.Label(rates, text="", style="Sub.TLabel")
        self.rates_lbl.pack(side="left")
        tk.Button(rates, text="Change…", command=self.pick_rates, bg=PANEL, fg=TEXT,
                  relief="flat", bd=0, font=("Helvetica", 10), cursor="hand2",
                  activebackground=DROP_HOVER, activeforeground=TEXT).pack(side="right")
        self._show_rates()

        # drop zone
        self.drop = tk.Label(
            outer, text="⬇  Drop the molecule file here\n(or click to browse)",
            bg=DROP_IDLE, fg=TEXT, font=("Helvetica", 14), height=6, cursor="hand2", bd=2,
        )
        self.drop.pack(fill="both", expand=True, pady=(0, 8))
        self.drop.bind("<Button-1>", lambda e: self.browse())
        if _DND:
            self.drop.drop_target_register(DND_FILES)
            self.drop.dnd_bind("<<Drop>>", self.on_drop)
            self.drop.dnd_bind("<<DragEnter>>", lambda e: self.drop.configure(bg=DROP_HOVER))
            self.drop.dnd_bind("<<DragLeave>>", lambda e: self.drop.configure(bg=self._idle_bg()))

        # sheet chooser (shown only once a file with selectable sheets is loaded)
        self.sheet_row = tk.Frame(outer, bg=BG)
        tk.Label(self.sheet_row, text="Sheet:", bg=BG, fg=MUTED,
                 font=("Helvetica", 11)).pack(side="left")
        self.sheet_cb = ttk.Combobox(self.sheet_row, state="readonly", width=32)
        self.sheet_cb.pack(side="left", padx=(6, 0))
        self.sheet_cb.bind("<<ComboboxSelected>>", lambda e: self.refresh_state())

        self.button = tk.Button(
            outer, text="Get Prices", command=self.run,
            bg=ACCENT, fg="white", activebackground=ACCENT_2, activeforeground="white",
            disabledforeground="#7f97b3", font=("Helvetica", 15, "bold"),
            relief="flat", bd=0, height=2, state="disabled", cursor="hand2",
        )
        self.button.pack(fill="x", pady=(10, 8))

        self.progress = ttk.Progressbar(outer, mode="determinate")
        self.status = ttk.Label(outer, text="Waiting for a file…", style="Sub.TLabel",
                                wraplength=500, justify="left")
        self.status.pack(anchor="w")

        self.reveal_btn = tk.Button(
            outer, text=REVEAL_LABEL, command=self.reveal, bg=ACCENT_2, fg="white",
            relief="flat", bd=0, font=("Helvetica", 12, "bold"), cursor="hand2",
        )

    def _idle_bg(self):
        return DROP_SET if self.path else DROP_IDLE

    # --- stratum rates ---
    def _show_rates(self):
        try:
            _, path = lndi.read_stratum(self.stratum_path)
            self.rates_lbl.configure(text=f"Stratum rates: {os.path.basename(str(path))}")
        except lndi.InputError:
            self.rates_lbl.configure(text="Stratum rates: not found — click Change…")

    def pick_rates(self):
        path = filedialog.askopenfilename(
            title="Choose the stratum rates table",
            filetypes=[("Excel / CSV", "*.xls *.xlsx *.xlsm *.xltx *.csv"), ("All files", "*.*")],
        )
        if path:
            self.stratum_path = path
            self._show_rates()

    # --- intake ---
    def browse(self):
        path = filedialog.askopenfilename(
            title="Choose a file of molecule names",
            filetypes=[("Excel / CSV", "*.xls *.xlsx *.xlsm *.xltx *.csv"), ("All files", "*.*")],
        )
        if path:
            self.set_path(path)

    def on_drop(self, event):
        paths = self.root.tk.splitlist(event.data)
        files = [p for p in paths if p.lower().endswith(EXTS)]
        if not files:
            messagebox.showwarning("Not a spreadsheet",
                                   "Please drop an Excel or CSV file of molecule names.")
            self.drop.configure(bg=self._idle_bg())
            return
        self.set_path(files[0])

    def set_path(self, path):
        self.path = path
        self.drop.configure(text="📄  " + os.path.basename(path), bg=DROP_SET)
        self.reveal_btn.pack_forget()

        # populate the sheet chooser for real multi-sheet workbooks
        sheets = lndi.list_sheets(path)
        if sheets:
            self.sheet_cb["values"] = sheets
            self.sheet_cb.set(lndi.best_sheet(path) or sheets[0])
            self.sheet_row.pack(fill="x", pady=(2, 0), before=self.button)
        else:
            self.sheet_cb.set("")
            self.sheet_row.pack_forget()   # CSV or single table: nothing to choose
        self.refresh_state()

    def sheet(self):
        """The chosen sheet name, or None to auto-detect."""
        return self.sheet_cb.get() or None

    def refresh_state(self):
        ready = bool(self.path)
        self.button.configure(state="normal" if ready else "disabled")
        if ready:
            where = f" from '{self.sheet()}'" if self.sheet() else ""
            self.status.configure(text=f"Ready. Click Get Prices{where}.")

    # --- work ---
    def run(self):
        self.reveal_btn.pack_forget()
        self.button.configure(state="disabled")
        self.status.configure(text=f"Reading {os.path.basename(self.path)}…")
        # Nothing to count until the molecule list is known, so sweep until then.
        self.progress.configure(mode="indeterminate", value=0)
        self.progress.pack(fill="x", pady=(4, 8))
        self.progress.start(12)
        lndi.REFRESH = bool(self.refresh.get())
        threading.Thread(target=self._work, args=(self.path, self.sheet()),
                         daemon=True).start()

    def _work(self, path, sheet):
        try:
            stem = os.path.splitext(os.path.basename(path))[0]
            # Named after the input, beside the input: a file the system has not
            # seen gets its own workbook, and re-running the same file refreshes
            # that same workbook rather than piling up copies.
            out = os.path.join(os.path.dirname(path), f"{stem} - Prices.xlsx")
            res, exc = lndi.run(lndi.Path(path), sheet, stratum=self.stratum_path,
                                on_progress=self._tick)
            try:
                lndi.write_output(res, exc, out)
            except PermissionError:
                # Windows locks a file while Excel has it open, so overwriting the
                # previous report fails outright. Say which file and why, rather
                # than showing a traceback for something the user can just close.
                raise lndi.InputError(
                    f"Could not save to:\n{out}\n\n"
                    "That file is open in Excel (or another program). Close it, "
                    "then click Get Prices again — the lookups are already "
                    "cached, so it will finish straight away.")
            self.root.after(0, lambda o=str(out): self._done(o, len(res), len(exc)))
        except PermissionError as exc:
            self.root.after(0, lambda m=str(exc): self._problem(
                f"Windows denied access to a file:\n{m}\n\n"
                "Check the file is not open elsewhere, and that the folder is "
                "not read-only or still syncing to OneDrive."))
        except lndi.InputError as exc:
            # Bind the values now: Python unbinds `exc` when the except block
            # ends, so a bare closure would raise NameError on the main thread
            # and the failure would never reach the window.
            self.root.after(0, lambda m=str(exc): self._problem(m))
        except Exception as exc:
            tb = traceback.format_exc()
            self.root.after(0, lambda e=exc: self._error(e, tb))

    def _tick(self, i, total, label):
        self.root.after(0, self._set_progress, i, total, label)

    def _set_progress(self, i, total, label):
        if str(self.progress.cget("mode")) == "indeterminate":
            self.progress.stop()
            self.progress.configure(mode="determinate")
        self.progress.configure(maximum=total, value=i)
        self.status.configure(text=f"Looking up {label}   ({i} of {total})")

    def _finish(self):
        self.progress.stop()
        self.progress.pack_forget()
        self.button.configure(state="normal")

    def _done(self, out, priced, skipped):
        self._finish()
        self.reveal_path = out
        note = (f"✓ Priced {priced} product(s); {skipped} could not be priced.\n"
                f"Saved to: {out}\n"
                "Sheets: 'Results' and 'Exceptions'.")
        if priced:
            note += ("\nBrand rows carry both 'Final -30%' and 'Final -40%' — "
                     "read whichever applies.")
        self.status.configure(text=note)
        self.reveal_btn.pack(anchor="w", pady=(10, 0))

    def _problem(self, message):
        """A problem with the user's files, not a crash."""
        self._finish()
        self.status.configure(text=f"⚠ {message}")
        messagebox.showwarning("Check your files", message)

    def _error(self, exc, tb):
        self._finish()
        self.status.configure(text="Something went wrong.")
        messagebox.showerror("Error", f"{exc}\n\n{tb}")

    def reveal(self):
        reveal_in_file_manager(self.reveal_path)


def main():
    root = TkinterDnD.Tk() if _DND else tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
