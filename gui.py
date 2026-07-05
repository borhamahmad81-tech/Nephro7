"""
gui.py
Medvision Nephro EMR Sync Engine - GUI
Slate Dark theme using standard tkinter/ttk.

Run: python gui.py
"""

import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

from docx_parser import parse_docx
from automation_engine import AutomationEngine, IMAGE_ASSETS

BG = "#0f172a"
BG_PANEL = "#1e293b"
FG = "#e2e8f0"
FG_DIM = "#94a3b8"
GREEN = "#22c55e"
RED = "#ef4444"
YELLOW = "#eab308"
FONT_MONO = ("Consolas", 10)
FONT_UI = ("Segoe UI", 10)


class MedvisionApp:
    def __init__(self, root):
        self.root = root
        root.title("Medvision Nephro - EMR Sync Engine")
        root.geometry("880x720")
        root.configure(bg=BG)

        self.records = []
        self.docx_path = None
        self.engine = AutomationEngine(log_callback=self._log, image_dir=os.getcwd())
        self.engine.on_ready_to_sign = self._on_ready_to_sign
        self.worker_thread = None

        self._build_style()
        self._build_layout()
        self._refresh_asset_grid()

    def _build_style(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("TFrame", background=BG)
        style.configure("Panel.TFrame", background=BG_PANEL)
        style.configure("TLabel", background=BG, foreground=FG, font=FONT_UI)
        style.configure("Panel.TLabel", background=BG_PANEL, foreground=FG, font=FONT_UI)
        style.configure("Dim.TLabel", background=BG_PANEL, foreground=FG_DIM, font=FONT_UI)
        style.configure("TScale", background=BG_PANEL)
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=6)
        style.configure("TRadiobutton", background=BG_PANEL, foreground=FG, font=FONT_UI)
        style.configure("TCheckbutton", background=BG_PANEL, foreground=FG, font=FONT_UI)

    def _build_layout(self):
        pad = 10

        # --- File selector ---
        top = ttk.Frame(self.root, style="Panel.TFrame")
        top.pack(fill="x", padx=pad, pady=(pad, 5))
        self.file_label = ttk.Label(top, text="No file loaded", style="Panel.TLabel")
        self.file_label.pack(side="left", padx=10, pady=8)
        ttk.Button(top, text="Browse .docx", command=self._browse_file).pack(
            side="right", padx=10, pady=8
        )

        # --- Timeout ceilings (adaptive polling - these are MAX waits, not fixed delays) ---
        timing = ttk.Frame(self.root, style="Panel.TFrame")
        timing.pack(fill="x", padx=pad, pady=5)
        ttk.Label(
            timing, text="Timeout ceilings (proceeds as soon as ready, waits up to this max):",
            style="Panel.TLabel"
        ).pack(anchor="w", padx=10, pady=(8, 0))
        self.patient_load_timeout_var = tk.DoubleVar(value=90)
        self.save_commit_timeout_var = tk.DoubleVar(value=90)
        self._add_slider(timing, "Patient Load Max Wait (s)", self.patient_load_timeout_var, 10, 180)
        self._add_slider(timing, "Save Commit Max Wait (s)", self.save_commit_timeout_var, 10, 180)

        # --- Mode + tab assumption + password ---
        opts = ttk.Frame(self.root, style="Panel.TFrame")
        opts.pack(fill="x", padx=pad, pady=5)

        mode_row = ttk.Frame(opts, style="Panel.TFrame")
        mode_row.pack(fill="x", padx=10, pady=(8, 2))
        ttk.Label(mode_row, text="Mode:", style="Panel.TLabel").pack(side="left")
        self.mode_var = tk.StringVar(value="guided")
        ttk.Radiobutton(
            mode_row, text="Guided (pause before each sign)",
            variable=self.mode_var, value="guided"
        ).pack(side="left", padx=10)
        ttk.Radiobutton(
            mode_row, text="Full Auto", variable=self.mode_var, value="auto"
        ).pack(side="left", padx=10)

        self.tab_active_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            opts, text="Progress Notes tab is already active by default",
            variable=self.tab_active_var
        ).pack(anchor="w", padx=10, pady=2)

        pw_row = ttk.Frame(opts, style="Panel.TFrame")
        pw_row.pack(fill="x", padx=10, pady=(2, 8))
        ttk.Label(pw_row, text="Signature Password:", style="Panel.TLabel").pack(side="left")
        self.password_var = tk.StringVar(value="")
        pw_entry = ttk.Entry(pw_row, textvariable=self.password_var, show="*", width=25)
        pw_entry.pack(side="left", padx=10)
        ttk.Label(
            pw_row, text="(kept in memory only, never saved to disk)", style="Dim.TLabel"
        ).pack(side="left")

        # --- Optional asset checker (only the tab-fallback image) ---
        assets_frame = ttk.Frame(self.root, style="Panel.TFrame")
        assets_frame.pack(fill="x", padx=pad, pady=5)
        ttk.Label(
            assets_frame, text="Optional fallback image (only used if tab isn't default-active)",
            style="Panel.TLabel"
        ).pack(anchor="w", padx=10, pady=(8, 0))
        self.asset_grid = ttk.Frame(assets_frame, style="Panel.TFrame")
        self.asset_grid.pack(fill="x", padx=10, pady=(2, 8))
        self.asset_labels = {}
        for name in IMAGE_ASSETS:
            lbl = ttk.Label(self.asset_grid, text=f"? {name}", style="Panel.TLabel")
            lbl.pack(anchor="w")
            self.asset_labels[name] = lbl
        ttk.Button(assets_frame, text="Re-check", command=self._refresh_asset_grid).pack(
            anchor="e", padx=10, pady=(0, 8)
        )

        # --- Controls ---
        controls = ttk.Frame(self.root, style="TFrame")
        controls.pack(fill="x", padx=pad, pady=5)
        self.start_btn = tk.Button(
            controls, text="START AUTOMATION", bg=GREEN, fg="black",
            font=("Segoe UI", 11, "bold"), command=self._start
        )
        self.start_btn.pack(side="left", padx=(0, 8), ipadx=10, ipady=6)
        self.abort_btn = tk.Button(
            controls, text="ABORT PROCESS", bg=RED, fg="white",
            font=("Segoe UI", 11, "bold"), command=self._abort, state="disabled"
        )
        self.abort_btn.pack(side="left", ipadx=10, ipady=6)

        tk.Button(
            controls, text="Test Nephro Focus", command=self._test_focus
        ).pack(side="left", padx=(8, 0), ipadx=6, ipady=4)
        self.status_label = ttk.Label(controls, text="Idle", style="TLabel")
        self.status_label.pack(side="right", padx=10)

        # --- Log console ---
        log_frame = ttk.Frame(self.root, style="TFrame")
        log_frame.pack(fill="both", expand=True, padx=pad, pady=(5, pad))
        self.log_box = tk.Text(
            log_frame, bg="#020617", fg=FG, font=FONT_MONO, wrap="word", state="disabled"
        )
        self.log_box.pack(fill="both", expand=True)
        self.log_box.tag_config("info", foreground=FG)
        self.log_box.tag_config("success", foreground=GREEN)
        self.log_box.tag_config("warn", foreground=YELLOW)
        self.log_box.tag_config("error", foreground=RED)

    def _add_slider(self, parent, label, var, lo, hi):
        row = ttk.Frame(parent, style="Panel.TFrame")
        row.pack(fill="x", padx=10, pady=4)
        ttk.Label(row, text=label, style="Panel.TLabel", width=20).pack(side="left")
        scale = ttk.Scale(row, from_=lo, to=hi, variable=var, orient="horizontal")
        scale.pack(side="left", fill="x", expand=True, padx=8)
        val_lbl = ttk.Label(row, text=f"{var.get():.1f}", style="Panel.TLabel", width=5)
        val_lbl.pack(side="left")

        def update_label(*_):
            val_lbl.config(text=f"{var.get():.1f}")

        var.trace_add("write", update_label)

    def _browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("Word documents", "*.docx")])
        if not path:
            return
        try:
            records = parse_docx(path)
        except Exception as exc:
            messagebox.showerror("Parse error", str(exc))
            self.file_label.config(text="No file loaded", foreground=FG_DIM)
            self.records = []
            return

        self.records = records
        self.docx_path = path
        self.file_label.config(
            text=f"{os.path.basename(path)}  ({len(records)} patients)", foreground=GREEN
        )
        self._log("info", f"Loaded {len(records)} patient records from {os.path.basename(path)}")

    def _refresh_asset_grid(self):
        results = self.engine.check_assets()
        for name, found in results.items():
            lbl = self.asset_labels[name]
            if found:
                lbl.config(text=f"\u2714 Found   {name}", foreground=GREEN)
            else:
                lbl.config(text=f"\u2716 Missing  {name}", foreground=RED)

    def _log(self, level, message):
        def append():
            timestamp = datetime.now().strftime("%H:%M:%S")
            self.log_box.config(state="normal")
            self.log_box.insert("end", f"[{timestamp}] {message}\n", level)
            self.log_box.see("end")
            self.log_box.config(state="disabled")

        self.root.after(0, append)

    def _start(self):
        if not self.records:
            messagebox.showwarning("No data", "Please load a .docx file first (0 patients).")
            return
        if not self.password_var.get():
            messagebox.showwarning("No password", "Please enter the signature password.")
            return

        if not self.tab_active_var.get():
            missing = [n for n, ok in self.engine.check_assets().items() if not ok]
            if missing:
                if not messagebox.askyesno(
                    "Missing fallback image",
                    f"{len(missing)} fallback image(s) missing:\n" + "\n".join(missing)
                    + "\n\nContinue anyway?",
                ):
                    return

        self.engine.patient_load_timeout = self.patient_load_timeout_var.get()
        self.engine.save_commit_timeout = self.save_commit_timeout_var.get()
        self.engine.assume_progress_tab_active = self.tab_active_var.get()
        self.engine.guided_mode = (self.mode_var.get() == "guided")
        self.engine.password = self.password_var.get()

        self.start_btn.config(state="disabled")
        self.abort_btn.config(state="normal")
        self.status_label.config(text="Running...")

        self.worker_thread = threading.Thread(
            target=self.engine.run,
            args=(self.records, self._on_done),
            daemon=True,
        )
        self.worker_thread.start()

    def _test_focus(self):
        def worker():
            try:
                self.engine._focus_nephro_window()
                self._log("success", "Focus test passed - Nephro is confirmed in the foreground.")
            except Exception as exc:
                self._log("error", f"Focus test failed: {exc}")

        threading.Thread(target=worker, daemon=True).start()

    def _abort(self):
        self.engine.request_abort()
        self.abort_btn.config(state="disabled")
        self.status_label.config(text="Aborting...")

    def _on_ready_to_sign(self, record):
        def show_popup():
            self.root.bell()
            win = tk.Toplevel(self.root)
            win.title("Ready to sign")
            win.configure(bg=BG_PANEL)
            win.attributes("-topmost", True)
            ttk.Label(
                win,
                text=(
                    f"Patient ID: {record.patient_id}\n\n"
                    "The verification dialog is open and the note has been pasted.\n"
                    "Click 'Continue' to let the tool type the password and sign,\n"
                    "or close this and abort if something looks wrong."
                ),
                style="Panel.TLabel",
                justify="left",
            ).pack(padx=20, pady=20)

            def confirm():
                win.destroy()
                self.engine.resume()

            tk.Button(
                win, text="Continue", bg=GREEN, fg="black",
                font=("Segoe UI", 10, "bold"), command=confirm
            ).pack(pady=(0, 15))

        self.root.after(0, show_popup)

    def _on_done(self, completed, total):
        def finish():
            self.start_btn.config(state="normal")
            self.abort_btn.config(state="disabled")
            self.status_label.config(text=f"Done: {completed}/{total}")

        self.root.after(0, finish)


if __name__ == "__main__":
    root = tk.Tk()
    app = MedvisionApp(root)
    root.mainloop()
