"""Modernised Excel Cutter GUI with conversion utilities.

This module exposes the :class:`ExcelCutterApp` Tkinter application which allows
users to split Excel worksheets into multiple files, convert worksheets to
other formats, and build large CSV/JSON tables back into Excel workbooks.  It
is designed to be packaged with PyInstaller for distribution on Windows, macOS
and Linux.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Iterable

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Excel specific limits
MAX_EXCEL_ROWS = 1_048_576
DEFAULT_ROWS_PER_FILE = 50_000


def resource_path(relative: str) -> Path:
    """Resolve a resource path for development and PyInstaller builds."""

    base_path = getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)
    return Path(base_path, relative)


def chunk_dataframe(df: pd.DataFrame, rows_per_chunk: int) -> Iterable[pd.DataFrame]:
    """Yield DataFrame slices with ``rows_per_chunk`` rows."""

    if rows_per_chunk <= 0:
        raise ValueError("Rows per chunk must be greater than zero.")

    if df.empty:
        yield df
        return

    total_rows = len(df)
    for start in range(0, total_rows, rows_per_chunk):
        end = min(start + rows_per_chunk, total_rows)
        yield df.iloc[start:end]


def sanitise_sheet_name(name: str) -> str:
    """Return a filesystem friendly version of ``name``."""

    cleaned = [c if c.isalnum() or c in {"-", "_"} else "_" for c in name.strip()]
    candidate = "".join(cleaned).strip("_")
    return candidate or "sheet"


class ExcelCutterApp:
    """Tkinter application for slicing and converting Excel files."""

    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Excel Cutter")
        self.root.geometry("720x420")
        self.root.minsize(640, 360)

        try:
            icon = tk.PhotoImage(file=str(resource_path("knife.png")))
            self.root.iconphoto(True, icon)
        except Exception:
            # Fallback silently when the icon is not available.
            pass

        self.style = ttk.Style(self.root)
        # Use a modern looking theme if available.
        if "clam" in self.style.theme_names():
            self.style.theme_use("clam")
        self.style.configure("TFrame", padding=12)
        self.style.configure("TButton", padding=6)
        self.style.configure("TLabel", padding=(0, 2))

        self.status_var = tk.StringVar(value="Ready")

        # Split tab variables
        self.input_file_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.rows_var = tk.StringVar(value=str(DEFAULT_ROWS_PER_FILE))
        self.format_vars = {
            "xlsx": tk.BooleanVar(value=True),
            "csv": tk.BooleanVar(value=False),
            "ods": tk.BooleanVar(value=False),
            "json": tk.BooleanVar(value=False),
        }

        # Build tab variables
        self.table_file_var = tk.StringVar()
        self.output_xlsx_var = tk.StringVar()
        self.rows_per_sheet_var = tk.StringVar(value=str(MAX_EXCEL_ROWS))

        self._build_ui()

    # ------------------------------------------------------------------ UI
    def _build_ui(self) -> None:
        notebook = ttk.Notebook(self.root)
        notebook.grid(row=0, column=0, sticky="nsew")

        split_tab = ttk.Frame(notebook)
        assemble_tab = ttk.Frame(notebook)
        notebook.add(split_tab, text="Split Excel")
        notebook.add(assemble_tab, text="Build XLSX from CSV/JSON")

        self._build_split_tab(split_tab)
        self._build_assemble_tab(assemble_tab)

        status_bar = ttk.Label(self.root, textvariable=self.status_var, anchor="w")
        status_bar.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 12))

        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

    def _build_split_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(1, weight=1)

        ttk.Label(parent, text="Choose an Excel workbook and sheet to split.").grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 8)
        )

        ttk.Label(parent, text="Workbook:").grid(row=1, column=0, sticky="w")
        input_entry = ttk.Entry(parent, textvariable=self.input_file_var)
        input_entry.grid(row=1, column=1, sticky="ew", padx=(0, 8))
        ttk.Button(parent, text="Browse", command=self.select_input_file).grid(
            row=1, column=2, sticky="ew"
        )

        ttk.Label(parent, text="Sheet:").grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.sheet_box = ttk.Combobox(parent, textvariable=self.sheet_var, state="readonly")
        self.sheet_box.grid(row=2, column=1, sticky="ew", padx=(0, 8), pady=(8, 0))
        self.sheet_box.bind("<<ComboboxSelected>>", lambda _event: self._update_status("Ready"))

        ttk.Label(parent, text="Rows per file:").grid(row=3, column=0, sticky="w", pady=(8, 0))
        rows_spin = ttk.Spinbox(
            parent,
            from_=1,
            to=MAX_EXCEL_ROWS,
            textvariable=self.rows_var,
            increment=1000,
            width=10,
        )
        rows_spin.grid(row=3, column=1, sticky="w", pady=(8, 0))

        ttk.Label(parent, text="Output folder:").grid(row=4, column=0, sticky="w", pady=(8, 0))
        output_entry = ttk.Entry(parent, textvariable=self.output_dir_var)
        output_entry.grid(row=4, column=1, sticky="ew", padx=(0, 8), pady=(8, 0))
        ttk.Button(parent, text="Browse", command=self.select_output_folder).grid(
            row=4, column=2, sticky="ew", pady=(8, 0)
        )

        ttk.Label(parent, text="Export formats:").grid(row=5, column=0, sticky="nw", pady=(16, 0))
        format_frame = ttk.Frame(parent)
        format_frame.grid(row=5, column=1, columnspan=2, sticky="w", pady=(16, 0))
        for idx, (fmt, var) in enumerate(self.format_vars.items()):
            ttk.Checkbutton(format_frame, text=fmt.upper(), variable=var).grid(
                row=0, column=idx, padx=(0, 12)
            )

        ttk.Button(parent, text="Split", command=self.split_excel).grid(
            row=6, column=1, sticky="e", pady=(24, 0)
        )

    def _build_assemble_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(1, weight=1)

        ttk.Label(parent, text="Turn a CSV or JSON table into an Excel workbook.").grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 8)
        )

        ttk.Label(parent, text="Table file:").grid(row=1, column=0, sticky="w")
        table_entry = ttk.Entry(parent, textvariable=self.table_file_var)
        table_entry.grid(row=1, column=1, sticky="ew", padx=(0, 8))
        ttk.Button(parent, text="Browse", command=self.select_table_file).grid(
            row=1, column=2, sticky="ew"
        )

        ttk.Label(parent, text="Output XLSX:").grid(row=2, column=0, sticky="w", pady=(8, 0))
        output_entry = ttk.Entry(parent, textvariable=self.output_xlsx_var)
        output_entry.grid(row=2, column=1, sticky="ew", padx=(0, 8), pady=(8, 0))
        ttk.Button(parent, text="Save As", command=self.select_output_xlsx).grid(
            row=2, column=2, sticky="ew", pady=(8, 0)
        )

        ttk.Label(parent, text="Rows per sheet:").grid(row=3, column=0, sticky="w", pady=(8, 0))
        rows_spin = ttk.Spinbox(
            parent,
            from_=1,
            to=MAX_EXCEL_ROWS,
            textvariable=self.rows_per_sheet_var,
            increment=1000,
            width=10,
        )
        rows_spin.grid(row=3, column=1, sticky="w", pady=(8, 0))

        ttk.Button(parent, text="Build XLSX", command=self.build_xlsx).grid(
            row=4, column=1, sticky="e", pady=(24, 0)
        )

    # ------------------------------------------------------------ callbacks
    def _update_status(self, message: str) -> None:
        self.status_var.set(message)

    def select_input_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Select Excel workbook",
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if not file_path:
            return
        self.input_file_var.set(file_path)
        if not self.output_dir_var.get():
            self.output_dir_var.set(str(Path(file_path).parent))
        self._load_sheet_names(Path(file_path))

    def _load_sheet_names(self, path: Path) -> None:
        try:
            excel_file = pd.ExcelFile(path)
        except Exception as exc:
            messagebox.showerror("Unable to read workbook", str(exc))
            self.sheet_box["values"] = []
            self.sheet_var.set("")
            return

        self.sheet_box["values"] = excel_file.sheet_names
        if excel_file.sheet_names:
            self.sheet_var.set(excel_file.sheet_names[0])
        else:
            self.sheet_var.set("")
        self._update_status(f"Loaded {len(excel_file.sheet_names)} sheet(s)")

    def select_output_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            self.output_dir_var.set(folder)

    def split_excel(self) -> None:
        file_path = Path(self.input_file_var.get())
        output_dir = Path(self.output_dir_var.get())
        sheet_name = self.sheet_var.get()
        try:
            rows_per_file = int(self.rows_var.get())
        except ValueError:
            messagebox.showerror("Invalid value", "Rows per file must be a number.")
            return

        if not file_path.exists():
            messagebox.showerror("Missing file", "Please choose a workbook to split.")
            return
        if not sheet_name:
            messagebox.showerror("Missing sheet", "Please choose a worksheet to split.")
            return
        if rows_per_file <= 0:
            messagebox.showerror("Invalid rows", "Rows per file must be greater than zero.")
            return
        if not output_dir.exists():
            try:
                output_dir.mkdir(parents=True, exist_ok=True)
            except Exception as exc:
                messagebox.showerror("Output error", f"Cannot create folder: {exc}")
                return

        selected_formats = [fmt for fmt, var in self.format_vars.items() if var.get()]
        if not selected_formats:
            messagebox.showinfo(
                "No formats selected",
                "Please choose at least one format to export.",
            )
            return

        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)
        except Exception as exc:
            messagebox.showerror("Read error", f"Unable to read sheet: {exc}")
            return

        if df.empty:
            messagebox.showinfo("Empty sheet", "The selected sheet has no data to split.")
            return

        safe_sheet = sanitise_sheet_name(sheet_name)
        chunks_written = 0
        try:
            for idx, chunk in enumerate(chunk_dataframe(df, rows_per_file), start=1):
                if chunk.empty:
                    continue
                chunks_written += 1
                base_name = f"{file_path.stem}_{safe_sheet}_part{idx:03d}"
                if "xlsx" in selected_formats:
                    chunk.to_excel(output_dir / f"{base_name}.xlsx", index=False)
                if "csv" in selected_formats:
                    chunk.to_csv(output_dir / f"{base_name}.csv", index=False)
                if "json" in selected_formats:
                    chunk.to_json(
                        output_dir / f"{base_name}.json",
                        orient="records",
                        indent=2,
                        force_ascii=False,
                    )
                if "ods" in selected_formats:
                    chunk.to_excel(
                        output_dir / f"{base_name}.ods",
                        index=False,
                        engine="odf",
                    )
        except ImportError as exc:
            messagebox.showerror(
                "Missing dependency",
                f"{exc}. Install optional dependencies to export in all formats.",
            )
            return
        except Exception as exc:
            messagebox.showerror("Split error", str(exc))
            return

        self._update_status(f"Created {chunks_written} file(s) in {output_dir}.")
        messagebox.showinfo(
            "Split complete", f"Created {chunks_written} file(s) in {output_dir}."
        )

    def select_table_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Select table",
            filetypes=[
                ("Table files", "*.csv *.json"),
                ("CSV", "*.csv"),
                ("JSON", "*.json"),
                ("All files", "*.*"),
            ],
        )
        if not file_path:
            return
        self.table_file_var.set(file_path)
        if not self.output_xlsx_var.get():
            self.output_xlsx_var.set(str(Path(file_path).with_suffix(".xlsx")))

    def select_output_xlsx(self) -> None:
        file_path = filedialog.asksaveasfilename(
            title="Save Excel workbook",
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if file_path:
            self.output_xlsx_var.set(file_path)

    def build_xlsx(self) -> None:
        table_path = Path(self.table_file_var.get())
        output_path = Path(self.output_xlsx_var.get())
        try:
            rows_per_sheet = int(self.rows_per_sheet_var.get())
        except ValueError:
            messagebox.showerror("Invalid value", "Rows per sheet must be a number.")
            return

        if not table_path.exists():
            messagebox.showerror("Missing table", "Please choose a CSV or JSON file.")
            return
        if rows_per_sheet <= 0:
            messagebox.showerror(
                "Invalid rows", "Rows per sheet must be greater than zero."
            )
            return
        rows_per_sheet = min(rows_per_sheet, MAX_EXCEL_ROWS)
        if not output_path.parent.exists():
            try:
                output_path.parent.mkdir(parents=True, exist_ok=True)
            except Exception as exc:
                messagebox.showerror("Output error", f"Cannot create folder: {exc}")
                return

        try:
            if table_path.suffix.lower() == ".csv":
                df = pd.read_csv(table_path)
            elif table_path.suffix.lower() == ".json":
                df = pd.read_json(table_path)
            else:
                messagebox.showerror(
                    "Unsupported format", "Please choose a CSV or JSON table file."
                )
                return
        except Exception as exc:
            messagebox.showerror("Read error", f"Unable to read table: {exc}")
            return

        if df.empty:
            messagebox.showinfo("Empty table", "The selected file has no rows to export.")
            return

        sheet_count = 0
        try:
            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                for idx, chunk in enumerate(chunk_dataframe(df, rows_per_sheet), start=1):
                    if chunk.empty:
                        continue
                    sheet_name = f"Sheet{idx}"
                    chunk.to_excel(writer, sheet_name=sheet_name, index=False)
                    sheet_count += 1
        except Exception as exc:
            messagebox.showerror("Write error", f"Unable to write workbook: {exc}")
            return

        self._update_status(
            f"Created workbook with {sheet_count} sheet(s) at {output_path}."
        )
        messagebox.showinfo(
            "Build complete",
            f"Created workbook with {sheet_count} sheet(s) at:\n{output_path}",
        )

    # --------------------------------------------------------------- control
    def run(self) -> None:
        self.root.mainloop()


def launch_app() -> None:
    """Convenience entry point."""

    ExcelCutterApp().run()


if __name__ == "__main__":
    launch_app()
