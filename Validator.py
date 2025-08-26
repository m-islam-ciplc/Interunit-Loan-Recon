#!/usr/bin/env python3
"""
Interunit Loan Matcher GUI
--------------------------

This script provides a simple graphical user interface (GUI) to compare two
interunit loan spreadsheets.  Users can select the GeoTex and Steel Excel
files, process the data, and view a summary of matched records.  The GUI
leverages Tkinter for the interface and pandas for data processing.

Features
--------
* Two file selection buttons for the GeoTex and Steel spreadsheets.
* A processing button that parses the selected files, extracts audit
  information, determines which unit is the lender or borrower for each
  match ID, and calculates the corresponding debit and credit amounts.
* Results displayed in a scrollable table with columns for match ID,
  audit information (from each file), lender, borrower, lender debit amount,
  and borrower credit amount.

Usage
-----
Run this script with Python.  The GUI will appear allowing you to choose
the two Excel files and generate the comparison table.  If you wish to
convert the script into a standalone executable, you can use a tool such
as `pyinstaller`:

    pyinstaller --onefile --noconsole interunit_loan_gui.py

Ensure that the pandas package is installed in your Python environment:

    pip install pandas openpyxl
"""

import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import List, Dict, Any

import pandas as pd


def load_and_process(file_geo: str, file_steel: str) -> List[Dict[str, Any]]:
    """Load the provided Excel files, match records by Match ID, compare audit
    information, and compute debit/credit amounts.

    Parameters
    ----------
    file_geo: str
        Path to the GeoTex Excel file.
    file_steel: str
        Path to the Steel Excel file.

    Returns
    -------
    List[Dict[str, Any]]
        A list of dictionaries representing the matched records.  Each
        dictionary contains keys:
            'Match ID', 'GeoTex Audit Info', 'Steel Audit Info',
            'Lender', 'Borrower', 'Lender Debit amount', 'Borrower Credit amount'
    """
    # Read Excel files.  The data starts at row 9 (zero-based index 8),
    # which becomes the header row after skipping the first 8 rows.
    df_geo = pd.read_excel(file_geo, header=8)
    df_steel = pd.read_excel(file_steel, header=8)

    # Identify Match IDs present in both files.
    match_ids = set(df_geo['Match ID'].dropna()).intersection(
        set(df_steel['Match ID'].dropna())
    )

    results: List[Dict[str, Any]] = []

    for mid in sorted(match_ids):
        # Extract the first non-null audit info strings from each file
        geo_ai_series = df_geo.loc[df_geo['Match ID'] == mid, 'Audit Info'].dropna()
        steel_ai_series = df_steel.loc[df_steel['Match ID'] == mid, 'Audit Info'].dropna()

        geo_audit_str = geo_ai_series.iloc[0] if not geo_ai_series.empty else ""
        steel_audit_str = steel_ai_series.iloc[0] if not steel_ai_series.empty else ""

        # Determine debit and credit totals for each file
        geo_debit_sum = df_geo.loc[df_geo['Match ID'] == mid, 'Debit'].sum()
        geo_credit_sum = df_geo.loc[df_geo['Match ID'] == mid, 'Credit'].sum()
        steel_debit_sum = df_steel.loc[df_steel['Match ID'] == mid, 'Debit'].sum()
        steel_credit_sum = df_steel.loc[df_steel['Match ID'] == mid, 'Credit'].sum()

        # Identify lender and borrower based on where the debit is recorded
        if geo_debit_sum > 0:
            lender = "GeoTex"
            borrower = "Steel"
            lender_amount = geo_debit_sum
            borrower_amount = steel_credit_sum
        elif steel_debit_sum > 0:
            lender = "Steel"
            borrower = "GeoTex"
            lender_amount = steel_debit_sum
            borrower_amount = geo_credit_sum
        else:
            # Fallback in case no debit was recorded (should not happen with
            # correctly formatted files)
            lender = "Unknown"
            borrower = "Unknown"
            lender_amount = 0.0
            borrower_amount = 0.0

        results.append({
            'Match ID': mid,
            'GeoTex Audit Info': str(geo_audit_str),
            'Steel Audit Info': str(steel_audit_str),
            'Lender': lender,
            'Borrower': borrower,
            'Lender Debit amount': round(float(lender_amount), 2),
            'Borrower Credit amount': round(float(borrower_amount), 2),
        })

    return results


class InterunitLoanApp(tk.Tk):
    """Tkinter-based GUI application for interunit loan matching."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Interunit Loan Matcher")
        self.geometry("1000x600")

        # File paths
        self.geo_file_path: tk.StringVar = tk.StringVar()
        self.steel_file_path: tk.StringVar = tk.StringVar()

        # Build the UI
        self._build_widgets()

    def _build_widgets(self) -> None:
        """Construct and layout the widgets."""
        # Frame for file selectors
        frame = tk.Frame(self)
        frame.pack(pady=10, padx=10, fill=tk.X)

        tk.Label(frame, text="GeoTex Excel File:").grid(row=0, column=0, sticky=tk.W)
        tk.Entry(frame, textvariable=self.geo_file_path, width=80).grid(row=0, column=1, padx=5)
        tk.Button(frame, text="Browse", command=self._select_geo_file).grid(row=0, column=2, padx=5)

        tk.Label(frame, text="Steel Excel File:").grid(row=1, column=0, sticky=tk.W)
        tk.Entry(frame, textvariable=self.steel_file_path, width=80).grid(row=1, column=1, padx=5)
        tk.Button(frame, text="Browse", command=self._select_steel_file).grid(row=1, column=2, padx=5)

        # Process button
        tk.Button(frame, text="Process", command=self._process_files).grid(row=2, column=0, columnspan=3, pady=10)

        # Treeview for results
        self.tree = ttk.Treeview(
            self,
            columns=(
                "Match ID", "GeoTex Audit Info", "Steel Audit Info",
                "Lender", "Borrower", "Lender Debit amount",
                "Borrower Credit amount"
            ),
            show='headings',
        )

        # Define column headings and widths
        self.tree.heading("Match ID", text="Match ID")
        self.tree.heading("GeoTex Audit Info", text="GeoTex Audit Info")
        self.tree.heading("Steel Audit Info", text="Steel Audit Info")
        self.tree.heading("Lender", text="Lender")
        self.tree.heading("Borrower", text="Borrower")
        self.tree.heading("Lender Debit amount", text="Lender Debit amount")
        self.tree.heading("Borrower Credit amount", text="Borrower Credit amount")

        # Set column widths (adjust as needed)
        self.tree.column("Match ID", width=80)
        self.tree.column("GeoTex Audit Info", width=300)
        self.tree.column("Steel Audit Info", width=300)
        self.tree.column("Lender", width=80)
        self.tree.column("Borrower", width=80)
        self.tree.column("Lender Debit amount", width=150, anchor=tk.E)
        self.tree.column("Borrower Credit amount", width=150, anchor=tk.E)

        # Add vertical scrollbar
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    def _select_geo_file(self) -> None:
        """Open file dialog to select the GeoTex file."""
        filename = filedialog.askopenfilename(
            title="Select GeoTex Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if filename:
            self.geo_file_path.set(filename)

    def _select_steel_file(self) -> None:
        """Open file dialog to select the Steel file."""
        filename = filedialog.askopenfilename(
            title="Select Steel Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if filename:
            self.steel_file_path.set(filename)

    def _process_files(self) -> None:
        """Load the selected files, process them, and display the results."""
        geo_path = self.geo_file_path.get()
        steel_path = self.steel_file_path.get()

        if not geo_path or not steel_path:
            messagebox.showerror(
                "Missing File", "Please select both the GeoTex and Steel files before processing."
            )
            return

        try:
            results = load_and_process(geo_path, steel_path)
        except Exception as e:
            messagebox.showerror("Processing Error", f"An error occurred: {e}")
            return

        # Clear existing rows in the treeview
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Insert new rows
        for row in results:
            self.tree.insert(
                "",
                tk.END,
                values=(
                    row['Match ID'],
                    row['GeoTex Audit Info'],
                    row['Steel Audit Info'],
                    row['Lender'],
                    row['Borrower'],
                    f"{row['Lender Debit amount']:,}",
                    f"{row['Borrower Credit amount']:,}",
                ),
            )


def main() -> None:
    """Entry point for running the application."""
    app = InterunitLoanApp()
    app.mainloop()


if __name__ == "__main__":
    main()
