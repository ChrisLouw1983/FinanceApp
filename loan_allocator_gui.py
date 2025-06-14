"""Loan allocation GUI using Tkinter and pandas.

This tool allows a user to select two Excel files:
1. Submission: contains loan information with instalments.
2. Collected: contains payments received.

Payments are allocated to the submission records by matching
ID NUMBER first, then EMPLOYEE NUMBER. Each record's PAID amount
is capped at the INSTALMENT AMOUNT. A DIFF column representing
the outstanding amount is added/updated.

Results are saved to an Excel file chosen by the user and a
summary is displayed.
"""

from __future__ import annotations

import os
import subprocess
import sys
import pandas as pd
from dataclasses import dataclass
from typing import Optional

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:  # fall back to regular tkinter if tkdnd is not available
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    DND_AVAILABLE = False
else:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk


@dataclass
class AllocationResult:
    records: int
    total_paid: float
    output_path: str


class LoanAllocatorGUI(tk.Tk):
    def __init__(self) -> None:
        if DND_AVAILABLE:
            tk.Tk.__init__(self)  # TkinterDnD uses mixin style
            self.__class__ = type(self.__class__.__name__, (TkinterDnD.Tk, self.__class__), {})
        else:
            super().__init__()
        self.title("Loan Allocator")
        self.geometry("600x300")
        self.resizable(False, False)
        self.submission_path: tk.StringVar = tk.StringVar()
        self.collected_path: tk.StringVar = tk.StringVar()
        self.summary_var: tk.StringVar = tk.StringVar()
        self.output_path: Optional[str] = None
        self.create_widgets()

    def create_widgets(self) -> None:
        padding = {"padx": 10, "pady": 5}

        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True, **padding)

        # Submission file widgets
        ttk.Label(frm, text="Submission file:").grid(row=0, column=0, sticky="w")
        sub_entry = ttk.Entry(frm, textvariable=self.submission_path, width=50)
        sub_entry.grid(row=0, column=1, sticky="we")
        ttk.Button(frm, text="Browse", command=self.browse_submission).grid(row=0, column=2)

        # Collected file widgets
        ttk.Label(frm, text="Collected file:").grid(row=1, column=0, sticky="w")
        coll_entry = ttk.Entry(frm, textvariable=self.collected_path, width=50)
        coll_entry.grid(row=1, column=1, sticky="we")
        ttk.Button(frm, text="Browse", command=self.browse_collected).grid(row=1, column=2)

        # Drag and drop support
        if DND_AVAILABLE:
            sub_entry.drop_target_register(DND_FILES)
            coll_entry.drop_target_register(DND_FILES)
            sub_entry.dnd_bind("<<Drop>>", lambda e: self.submission_path.set(e.data))
            coll_entry.dnd_bind("<<Drop>>", lambda e: self.collected_path.set(e.data))

        # Progress bar
        self.progress = ttk.Progressbar(frm, length=400, mode="determinate")
        self.progress.grid(row=2, column=0, columnspan=3, sticky="we", pady=(10, 0))

        # Process button
        ttk.Button(frm, text="Process", command=self.process_files).grid(row=3, column=0, columnspan=3, pady=(10, 0))

        # Summary label
        ttk.Label(frm, textvariable=self.summary_var, foreground="blue").grid(row=4, column=0, columnspan=3, sticky="we", pady=(10, 0))

        # Open folder button
        ttk.Button(frm, text="Open Output Folder", command=self.open_output_folder).grid(row=5, column=0, columnspan=3)

    # -- file dialogs -----------------------------------------------------
    def browse_submission(self) -> None:
        path = filedialog.askopenfilename(title="Select Submission file", filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.submission_path.set(path)

    def browse_collected(self) -> None:
        path = filedialog.askopenfilename(title="Select Collected file", filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.collected_path.set(path)

    # -- core processing --------------------------------------------------
    def process_files(self) -> None:
        sub_path = self.submission_path.get()
        coll_path = self.collected_path.get()
        if not sub_path or not coll_path:
            messagebox.showerror("Error", "Please select both Submission and Collected files.")
            return
        try:
            self.progress["value"] = 10
            self.update_idletasks()

            df_sub = pd.read_excel(sub_path)
            df_col = pd.read_excel(coll_path)
            self.progress["value"] = 30
            self.update_idletasks()
        except Exception as exc:  # catch file read errors
            messagebox.showerror("Error", f"Failed to read Excel files: {exc}")
            return

        try:
            result = self.allocate(df_sub, df_col)
            self.progress["value"] = 80
            self.update_idletasks()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return

        # Ask for output path
        save_path = filedialog.asksaveasfilename(title="Save output", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], initialfile="output.xlsx")
        if not save_path:
            return
        try:
            df_sub.to_excel(save_path, index=False)
            self.output_path = save_path
            self.progress["value"] = 100
            self.summary_var.set(f"Processed {result.records} records. Total paid: {result.total_paid:.2f}")
            messagebox.showinfo("Success", f"Output saved to {save_path}")
        except Exception as exc:
            messagebox.showerror("Error", f"Failed to save output: {exc}")

    def allocate(self, df_sub: pd.DataFrame, df_col: pd.DataFrame) -> AllocationResult:
        required_sub = {"ID NUMBER", "EMPLOYEE NUMBER", "INSTALMENT AMOUNT"}
        required_col = {"ID NUMBER", "EMPLOYEE NUMBER", "PAID"}

        if missing := required_sub - set(df_sub.columns):
            raise ValueError(f"Submission file missing columns: {', '.join(missing)}")
        if missing := required_col - set(df_col.columns):
            raise ValueError(f"Collected file missing columns: {', '.join(missing)}")

        if "PAID" not in df_sub.columns:
            df_sub["PAID"] = 0.0
        else:
            df_sub["PAID"] = df_sub["PAID"].fillna(0.0)

        df_sub["DIFF"] = df_sub["INSTALMENT AMOUNT"] - df_sub["PAID"]

        payments_by_id = df_col.groupby("ID NUMBER")["PAID"].sum().to_dict()
        payments_by_emp = df_col.groupby("EMPLOYEE NUMBER")["PAID"].sum().to_dict()

        total_paid = 0.0

        for idx in df_sub.index:
            row = df_sub.loc[idx]
            instalment = float(row["INSTALMENT AMOUNT"])
            paid = float(row["PAID"])
            id_number = row["ID NUMBER"]
            emp_number = row["EMPLOYEE NUMBER"]
            # allocate by ID
            available = payments_by_id.get(id_number, 0.0)
            if available > 0:
                needed = instalment - paid
                alloc = min(needed, available)
                paid += alloc
                payments_by_id[id_number] = available - alloc
            # allocate by employee number
            if paid < instalment:
                available = payments_by_emp.get(emp_number, 0.0)
                if available > 0:
                    needed = instalment - paid
                    alloc = min(needed, available)
                    paid += alloc
                    payments_by_emp[emp_number] = available - alloc
            paid = min(paid, instalment)
            df_sub.at[idx, "PAID"] = paid
            df_sub.at[idx, "DIFF"] = instalment - paid
            total_paid += paid

        return AllocationResult(records=len(df_sub), total_paid=total_paid, output_path="")

    def open_output_folder(self) -> None:
        if not self.output_path:
            messagebox.showinfo("Info", "No output file available yet.")
            return
        folder = os.path.dirname(self.output_path)
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.call(["open", folder])
            else:
                subprocess.call(["xdg-open", folder])
        except Exception as exc:
            messagebox.showerror("Error", f"Failed to open folder: {exc}")


if __name__ == "__main__":
    app = LoanAllocatorGUI()
    app.mainloop()
