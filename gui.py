#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Minimal XML → Excel Converter — Tkinter GUI

This module provides a simple graphical interface for converting XML invoice files
(NFe-like structure) into a structured Excel spreadsheet (.xlsx).

Main features:
    • One button to select one or more XML files
    • One button to choose where to save the Excel file
    • Convert button with log display

Dependencies:
    - xmltodict
    - pandas
    - openpyxl
    - tkinter (standard library)

Usage:
    python gui_simple.py
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Dict, List, Optional

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
import xmltodict

# ---------------- Parsing helpers ----------------

def _get(d: Dict[str, Any], path: str, default: Any = None) -> Any:
    """
    Safely access a nested dictionary path using dot notation.

    Args:
        d (dict): Dictionary to traverse.
        path (str): Dot-separated path (e.g. "nfeProc.NFe.infNFe").
        default (Any): Value to return if path does not exist.

    Returns:
        Any: Value found at the given path, or default if missing.
    """
    cur = d
    for part in path.split("."):
        if not isinstance(cur, dict) or part not in cur:
            return default
        cur = cur[part]
    return cur


def parse_nfe_file(xml_path: Path) -> Optional[Dict[str, Any]]:
    """
    Parse a single NFe-like XML file into a normalized dictionary of fields.

    Args:
        xml_path (Path): Path to the XML file.

    Returns:
        dict | None: Dictionary with invoice data, or None if parsing fails.
            Keys include:
                - key, note_id, issuer_name, recipient_name
                - dest_street, dest_number, dest_district, dest_city,
                  dest_state, dest_zip, dest_country
                - file_name
    """
    try:
        with xml_path.open("rb") as f:
            data = xmltodict.parse(f)
    except Exception as e:
        logging.warning(f"Failed to parse XML: {xml_path.name} | {e}")
        return None

    inf = _get(data, "nfeProc.NFe.infNFe") or _get(data, "NFe.infNFe")
    if not isinstance(inf, dict):
        return None

    note_id = _get(inf, "@Id", "")
    key = note_id.replace("NFe", "") if note_id else ""

    issuer_name = _get(inf, "emit.xNome", "")
    recipient_name = _get(inf, "dest.xNome", "")

    end = _get(inf, "dest.enderDest", {}) or {}
    address = {
        "dest_street": end.get("xLgr", ""),
        "dest_number": end.get("nro", ""),
        "dest_district": end.get("xBairro", ""),
        "dest_city": end.get("xMun", ""),
        "dest_state": end.get("UF", ""),
        "dest_zip": end.get("CEP", ""),
        "dest_country": end.get("xPais", ""),
    }
    if not any(address.values()):
        end_emit = _get(inf, "emit.enderEmit", {}) or {}
        address = {
            "dest_street": end_emit.get("xLgr", ""),
            "dest_number": end_emit.get("nro", ""),
            "dest_district": end_emit.get("xBairro", ""),
            "dest_city": end_emit.get("xMun", ""),
            "dest_state": end_emit.get("UF", ""),
            "dest_zip": end_emit.get("CEP", ""),
            "dest_country": end_emit.get("xPais", ""),
        }

    return {
        "key": key,
        "note_id": note_id,
        "issuer_name": issuer_name,
        "recipient_name": recipient_name,
        **address,
        "file_name": xml_path.name,
    }

# ---------------- Minimal GUI ----------------

class App(ttk.Frame):
    """
    Tkinter GUI for converting XML invoice files into an Excel spreadsheet.

    Workflow:
        1. User selects one or more XML files via file dialog.
        2. User specifies the destination Excel file (.xlsx).
        3. On "Convert", the selected XMLs are parsed and results exported.

    Attributes:
        selected_files (List[Path]): List of chosen XML file paths.
        output_var (tk.StringVar): Holds the path to the Excel output file.
        listbox (tk.Listbox): Widget listing selected XML files.
        txt (tk.Text): Log output widget.
    """
    def __init__(self, master: tk.Tk) -> None:
        """
        Initialize the GUI application.

        Args:
            master (tk.Tk): Root Tkinter window.
        """
        super().__init__(master, padding=12)
        self.master.title("XML → Excel Converter — Minimal GUI")
        self.master.geometry("760x520")

        self.selected_files: List[Path] = []
        self.output_var = tk.StringVar(value=str(Path("output_file/Invoices.xlsx").resolve()))

        self._build_ui()

    def _build_ui(self) -> None:
        """
         - File selection list and buttons
        - Output file entry and save button
        - Convert button
        - Log area
        """
        grp_files = ttk.LabelFrame(self, text="Files")
        grp_files.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        ttk.Button(grp_files, text="Select XML files…", command=self._choose_files).grid(row=0, column=0, padx=6, pady=6, sticky="w")

        self.listbox = tk.Listbox(grp_files, height=8)
        self.listbox.grid(row=1, column=0, sticky="nsew", padx=6, pady=(0,6))
        yscroll = ttk.Scrollbar(grp_files, orient="vertical", command=self.listbox.yview)
        yscroll.grid(row=1, column=1, sticky="ns")
        self.listbox.configure(yscrollcommand=yscroll.set)

        grp_out = ttk.LabelFrame(self, text="Output")
        grp_out.grid(row=1, column=0, sticky="ew", padx=4, pady=4)
        ttk.Label(grp_out, text="Excel file (.xlsx):").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        ttk.Entry(grp_out, textvariable=self.output_var, width=64).grid(row=0, column=1, sticky="ew", padx=6, pady=6)
        ttk.Button(grp_out, text="Save as…", command=self._choose_output).grid(row=0, column=2, padx=6, pady=6)

        bar = ttk.Frame(self)
        bar.grid(row=2, column=0, sticky="ew", padx=4, pady=4)
        ttk.Button(bar, text="Convert", command=self._on_convert).grid(row=0, column=0, padx=6, pady=6)

        grp_log = ttk.LabelFrame(self, text="Log")
        grp_log.grid(row=3, column=0, sticky="nsew", padx=4, pady=4)
        self.txt = tk.Text(grp_log, height=12, wrap="word")
        self.txt.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        yscroll2 = ttk.Scrollbar(grp_log, orient="vertical", command=self.txt.yview)
        yscroll2.grid(row=0, column=1, sticky="ns")
        self.txt.configure(yscrollcommand=yscroll2.set)

        self.columnconfigure(0, weight=1)
        grp_files.columnconfigure(0, weight=1)
        grp_log.columnconfigure(0, weight=1)
        grp_log.rowconfigure(0, weight=1)

        self.pack(fill="both", expand=True)

    def _choose_files(self) -> None:
        """
        Open a dialog to select one or more XML files.
        Updates `selected_files` and refreshes the listbox.
        """
        file_paths = filedialog.askopenfilenames(
            title="Select XML files…",
            filetypes=[("XML files", "*.xml *.XML"), ("All files", "*.*")],
            initialdir=str(Path.cwd()),
        )
        if not file_paths:
            return
        self.selected_files = [Path(p) for p in file_paths]
        self._refresh_file_list()
        self._log(f"Selected {len(self.selected_files)} file(s).\n")

    def _choose_output(self) -> None:
        """
        Open a dialog to select the Excel output file.
        Updates `output_var`.
        """
        initial = Path(self.output_var.get()).parent if self.output_var.get() else Path.cwd()
        f = filedialog.asksaveasfilename(
            title="Save Excel as…",
            defaultextension=".xlsx",
            initialdir=str(initial),
            initialfile=Path(self.output_var.get()).name if self.output_var.get() else "Invoices.xlsx",
            filetypes=[("Excel file", ".xlsx")],
        )
        if f:
            self.output_var.set(f)

    def _on_convert(self) -> None:
        """
        Execute the conversion process:
        - Parse selected XML files
        - Build DataFrame
        - Save to Excel
        Logs progress and results in the log box.
        """
        if not self.selected_files:
            messagebox.showinfo("Select files", "Please select one or more XML files to convert.")
            return
        out_xlsx = Path(self.output_var.get())
        try:
            if not out_xlsx.parent.exists():
                out_xlsx.parent.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Output error", f"Failed to create output folder:\n{out_xlsx.parent}\n\n{e}")
            return

        self._log("Converting…\n")
        try:
            records: List[Dict[str, Any]] = []
            errors: List[str] = []
            for p in self.selected_files:
                rec = parse_nfe_file(p)
                if rec is None:
                    errors.append(p.name)
                else:
                    records.append(rec)

            if not records:
                self._log("No valid records produced. Nothing to save.\n")
                return

            cols = [
                "key", "note_id", "issuer_name", "recipient_name",
                "dest_street", "dest_number", "dest_district", "dest_city",
                "dest_state", "dest_zip", "dest_country", "file_name",
            ]
            df = pd.DataFrame(records)
            for c in cols:
                if c not in df.columns:
                    df[c] = ""
            df = df[cols].sort_values(by=["key", "file_name"], na_position="last")

            try:
                df.to_excel(out_xlsx, index=False)
            except ImportError:
                self._log("\n⚠ Missing dependency to write .xlsx. Install with:\n    pip install openpyxl\n")
                raise

            self._log(f"✔ Done! Saved: {out_xlsx}\n")
            if errors:
                self._log(f"Files ignored or with errors: {len(errors)}\n")
        except Exception as e:
            self._log(f"ERROR: {e}\n")

    def _refresh_file_list(self) -> None:
        """
        Refresh the listbox widget with the currently selected files.
        """
        self.listbox.delete(0, "end")
        for p in self.selected_files:
            self.listbox.insert("end", p.name)

    def _log(self, msg: str) -> None:
        """
        Append a message to the log text box and scroll to bottom.

        Args:
            msg (str): The message to display.
        """
        self.txt.insert("end", msg)
        self.txt.see("end")


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()

