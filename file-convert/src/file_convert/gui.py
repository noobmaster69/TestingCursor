"""Tkinter GUI for file-convert."""

from __future__ import annotations

import os
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from file_convert import __version__
from file_convert.core import ConversionRequest, run_conversion, targets_for_source
from file_convert.formats import infer_format

# Friendly labels for target formats
TARGET_LABELS: dict[str, str] = {
    "html": "HTML (web page)",
    "csv": "CSV (spreadsheet)",
    "xlsx": "Excel (.xlsx)",
    "pdf": "PDF document",
    "png": "PNG image",
    "jpg": "JPEG image",
    "list": "List zip contents",
}


class FileConvertApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"File Convert v{__version__}")
        self.minsize(520, 480)
        self.geometry("620x560")

        self._input_path = tk.StringVar()
        self._output_path = tk.StringVar()
        self._force = tk.BooleanVar(value=False)
        self._target_values: list[str] = []
        self._busy = False
        self._last_output: Path | None = None

        self._build_ui()
        self._center_window()

    def _center_window(self) -> None:
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        x, y = (sw - w) // 2, (sh - h) // 2
        self.geometry(f"+{x}+{y}")

    def _build_ui(self) -> None:
        pad = {"padx": 12, "pady": 6}
        main = ttk.Frame(self, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            main,
            text="Convert files on your computer — nothing uploaded.",
            font=("Segoe UI", 10),
        ).pack(anchor=tk.W, pady=(0, 10))

        # Input
        in_frame = ttk.LabelFrame(main, text="1. Choose file", padding=8)
        in_frame.pack(fill=tk.X, **pad)
        ttk.Entry(in_frame, textvariable=self._input_path).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8)
        )
        ttk.Button(in_frame, text="Browse…", command=self._browse_input).pack(side=tk.RIGHT)

        # Target
        tgt_frame = ttk.LabelFrame(main, text="2. Convert to", padding=8)
        tgt_frame.pack(fill=tk.X, **pad)
        self._target_combo = ttk.Combobox(
            tgt_frame,
            state="readonly",
            width=40,
        )
        self._target_combo.pack(fill=tk.X)
        self._target_combo.bind("<<ComboboxSelected>>", self._on_target_changed)

        # Output
        out_frame = ttk.LabelFrame(main, text="3. Save location (optional)", padding=8)
        out_frame.pack(fill=tk.X, **pad)
        ttk.Label(out_frame, text="Leave empty to save next to the original file.").pack(
            anchor=tk.W, pady=(0, 4)
        )
        row = ttk.Frame(out_frame)
        row.pack(fill=tk.X)
        ttk.Entry(row, textvariable=self._output_path).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8)
        )
        ttk.Button(row, text="Browse…", command=self._browse_output).pack(side=tk.RIGHT)

        # Options
        self._opts_frame = ttk.LabelFrame(main, text="Options", padding=8)
        self._opts_frame.pack(fill=tk.X, **pad)
        self._opt_widgets: dict[str, tk.Widget] = {}

        self._sheet_var = tk.StringVar()
        self._theme_var = tk.StringVar(value="github")
        self._pages_var = tk.StringVar()
        self._dpi_var = tk.StringVar(value="150")
        self._quality_var = tk.StringVar(value="85")

        self._add_option("sheet", "Excel sheet name:", ttk.Entry(self._opts_frame, textvariable=self._sheet_var))
        self._add_option(
            "theme",
            "HTML theme:",
            ttk.Combobox(
                self._opts_frame,
                textvariable=self._theme_var,
                values=["github", "minimal"],
                state="readonly",
                width=20,
            ),
        )
        self._add_option("pages", "PDF pages (e.g. 1-3):", ttk.Entry(self._opts_frame, textvariable=self._pages_var))
        self._add_option("dpi", "PDF DPI:", ttk.Entry(self._opts_frame, width=8, textvariable=self._dpi_var))
        self._add_option("quality", "JPEG quality (1-100):", ttk.Entry(self._opts_frame, width=8, textvariable=self._quality_var))

        opts_row = ttk.Frame(main)
        opts_row.pack(fill=tk.X, **pad)
        ttk.Checkbutton(opts_row, text="Overwrite if file exists", variable=self._force).pack(
            side=tk.LEFT
        )

        # Actions
        btn_row = ttk.Frame(main)
        btn_row.pack(fill=tk.X, pady=12)
        self._convert_btn = ttk.Button(btn_row, text="Convert", command=self._on_convert)
        self._convert_btn.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_row, text="Open output folder", command=self._open_output_folder).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Button(btn_row, text="Open output file", command=self._open_output_file).pack(
            side=tk.LEFT, padx=4
        )

        # Log
        log_frame = ttk.LabelFrame(main, text="Status", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True, **pad)
        self._log = tk.Text(log_frame, height=10, wrap=tk.WORD, font=("Consolas", 9))
        scroll = ttk.Scrollbar(log_frame, command=self._log.yview)
        self._log.configure(yscrollcommand=scroll.set)
        self._log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self._log_insert("Pick a file to see available conversions.\n")

        # Trace input changes
        self._input_path.trace_add("write", lambda *_: self._refresh_targets())

        try:
            style = ttk.Style(self)
            if sys.platform == "win32":
                style.theme_use("vista")
        except tk.TclError:
            pass

    def _add_option(self, key: str, label: str, widget: tk.Widget) -> None:
        row = ttk.Frame(self._opts_frame)
        ttk.Label(row, text=label, width=22).pack(side=tk.LEFT)
        widget.pack(side=tk.LEFT, fill=tk.X, expand=True)
        row.pack(fill=tk.X, pady=2)
        self._opt_widgets[key] = row

    def _show_options_for(self, src: str | None, tgt: str | None) -> None:
        for row in self._opt_widgets.values():
            row.pack_forget()
        if not tgt:
            return
        if src == "xlsx" and tgt == "csv":
            self._opt_widgets["sheet"].pack(fill=tk.X, pady=2)
        if src == "md" and tgt == "html":
            self._opt_widgets["theme"].pack(fill=tk.X, pady=2)
        if src == "pdf" and tgt in ("png", "jpg"):
            self._opt_widgets["pages"].pack(fill=tk.X, pady=2)
            self._opt_widgets["dpi"].pack(fill=tk.X, pady=2)
        if tgt == "jpg" or (src == "pdf" and tgt == "jpg"):
            self._opt_widgets["quality"].pack(fill=tk.X, pady=2)

    def _browse_input(self) -> None:
        path = filedialog.askopenfilename(
            title="Select file to convert",
            filetypes=[
                ("All supported", "*.md *.csv *.xlsx *.pdf *.png *.jpg *.jpeg *.webp *.heic *.docx *.odt *.zip"),
                ("Documents", "*.md *.docx *.odt *.pdf"),
                ("Spreadsheets", "*.csv *.xlsx"),
                ("Images", "*.png *.jpg *.jpeg *.webp *.heic"),
                ("Archives", "*.zip"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self._input_path.set(path)
            self._refresh_targets()

    def _browse_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Save converted file as",
            defaultextension="",
            initialfile=Path(self._input_path.get()).stem if self._input_path.get() else "",
        )
        if path:
            self._output_path.set(path)

    def _refresh_targets(self) -> None:
        path = self._input_path.get().strip()
        if not path:
            self._target_combo["values"] = []
            self._target_format.set("")
            return
        src = infer_format(Path(path))
        targets = targets_for_source(src)
        labels = [f"{TARGET_LABELS.get(t, t.upper())} (.{t})" for t in targets]
        self._target_combo["values"] = labels
        self._target_values = targets
        if targets:
            self._target_combo.current(0)
            self._on_target_changed()
        else:
            self._log_insert(f"No conversions available for .{src or '?'}\n")

    def _on_target_changed(self, *_event) -> None:
        idx = self._target_combo.current()
        if idx < 0 or not hasattr(self, "_target_values"):
            return
        tgt = self._target_values[idx]
        src = infer_format(Path(self._input_path.get())) if self._input_path.get() else None
        self._show_options_for(src, tgt)

    def _selected_target(self) -> str | None:
        idx = self._target_combo.current()
        if idx < 0 or not hasattr(self, "_target_values"):
            return None
        return self._target_values[idx]

    def _build_options(self) -> dict:
        opts: dict = {"title": Path(self._input_path.get()).stem if self._input_path.get() else ""}
        if self._sheet_var.get().strip():
            opts["sheet"] = self._sheet_var.get().strip()
        if self._theme_var.get():
            opts["theme"] = self._theme_var.get()
        if self._pages_var.get().strip():
            opts["pages"] = self._pages_var.get().strip()
        if self._dpi_var.get().strip():
            try:
                opts["dpi"] = int(self._dpi_var.get())
            except ValueError:
                pass
        if self._quality_var.get().strip():
            try:
                opts["quality"] = int(self._quality_var.get())
            except ValueError:
                pass
        return opts

    def _log_insert(self, text: str) -> None:
        self._log.insert(tk.END, text)
        self._log.see(tk.END)

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        state = tk.DISABLED if busy else tk.NORMAL
        self._convert_btn.configure(state=state)

    def _on_convert(self) -> None:
        if self._busy:
            return
        src_path = self._input_path.get().strip()
        if not src_path:
            messagebox.showwarning("Missing file", "Please choose a file first.")
            return
        tgt = self._selected_target()
        if not tgt:
            messagebox.showwarning("Missing format", "Please choose a conversion type.")
            return

        out = self._output_path.get().strip() or None
        request = ConversionRequest(
            source=Path(src_path),
            target_format=tgt,
            output=Path(out) if out else None,
            force=self._force.get(),
            options=self._build_options(),
        )

        self._set_busy(True)
        self._log_insert(f"\nConverting → .{tgt} …\n")
        thread = threading.Thread(target=self._run_convert, args=(request,), daemon=True)
        thread.start()

    def _run_convert(self, request: ConversionRequest) -> None:
        response = run_conversion(request)
        self.after(0, lambda: self._on_convert_done(response))

    def _on_convert_done(self, response) -> None:
        self._set_busy(False)
        if response.ok:
            self._last_output = response.output_path
            self._log_insert(f"✓ {response.message}\n")
            if response.output_path and response.output_path.suffix:
                messagebox.showinfo("Done", response.message[:500])
            else:
                messagebox.showinfo("Done", "Conversion finished.")
        else:
            self._log_insert(f"✗ {response.message}\n")
            messagebox.showerror("Conversion failed", response.message)

    def _open_output_folder(self) -> None:
        if not self._last_output or not self._last_output.exists():
            messagebox.showinfo("Open folder", "Convert a file first.")
            return
        folder = self._last_output.parent if self._last_output.is_file() else self._last_output
        os.startfile(folder) if sys.platform == "win32" else subprocess.run(["xdg-open", str(folder)])

    def _open_output_file(self) -> None:
        if not self._last_output or not self._last_output.is_file():
            messagebox.showinfo("Open file", "No output file yet (or zip list has no file).")
            return
        os.startfile(self._last_output) if sys.platform == "win32" else subprocess.run(
            ["xdg-open", str(self._last_output)]
        )


def run() -> None:
    app = FileConvertApp()
    app.mainloop()


if __name__ == "__main__":
    run()
