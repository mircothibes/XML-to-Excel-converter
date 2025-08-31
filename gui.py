#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
XML → Excel Converter — Tkinter GUI

• Select an input folder with XML invoices (NFe-like, and easy to extend)
• Choose an output .xlsx path
• Click "Convert" to parse and export a spreadsheet

Dependencies: xmltodict, pandas, openpyxl

Run:
    python gui.py
"""
from __future__ import annotations

import logging
import threading
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Callable

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
import xmltodict

# --------------------------- Parsing utilities ---------------------------

def _get(d: Dict[str, Any], path: str, default: Any = None) -> Any:
    """Safe nested access: _get(obj, "a.b.c") or default if missing."""
    cur = d
    for part in path.split("."):
        if not isinstance(cur, dict) or part not in cur:
            return default
        cur = cur[part]
    return cur


def parse_nfe_file(xml_path: Path) -> Optional[Dict[str, Any]]:
    """Parse an NFe-like XML and return a normalized record dict, or None."""
    try:
        with xml_path.open("rb") as f:
            data = xmltodict.parse(f)
    except Exception as e:
        logging.warning(f"Failed to parse XML: {xml_path.name} | {e}")
        return None

    # Typical structure: nfeProc -> NFe -> infNFe, fallback to NFe -> infNFe
    inf = _get(data, "nfeProc.NFe.infNFe")
    if not isinstance(inf, dict):
        inf = _get(data, "NFe.infNFe")
    if not isinstance(inf, dict):
        logging.debug(f"Ignored (doesn't look like NFe): {xml_path.name}")
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

    if not any(address.values()):  # fallback to issuer address if destination missing
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


def iter_xml_files(root: Path) -> List[Path]:
    return sorted([p for p in root.glob("**/*.xml") if p.is_file()])


def parse_folder(
    input_dir: Path,
    on_progress: Optional[Callable[[int, int, str], None]] = None,
) -> Tuple[List[Dict[str, Any]], List[str]]:
    """
    Process all .xml files in the folder. Calls on_progress(cur, total, name).
    Returns (records, error_files).
    """
    records: List[Dict[str, Any]] = []
    errors: List[str] = []

    xmls = iter_xml_files(input_dir)
    total = len(xmls)
    if total == 0:
        logging.warning(f"No .xml files found in: {input_dir}")

    for i, xml_path in enumerate(xmls, start=1):
        if on_progress:
            on_progress(i, total, xml_path.name)
        rec = parse_nfe_file(xml_path)
        if rec is None:
            errors.append(xml_path.name)
        else:
            records.append(rec)

    return records, errors


# --------------------------- GUI application ---------------------------

class App(ttk.Frame):
    def __init__(self, master: tk.Tk) -> None:
        super().__init__(master, padding=12)
        self.master.title("XML → Excel Converter — Tkinter GUI")
        self.master.geometry("820x560")

        # State vars
        self.input_var = tk.StringVar(value=str(Path("input_folder").resolve()))
        self.output_var = tk.StringVar(value=str(Path("output_file/Invoices.xlsx").resolve()))
        self.verbose_var = tk.BooleanVar(value=False)
        self.selected_files: list[Path] = []

        self._build_ui()

    def _build_ui(self) -> None:
        # Paths group
        grp_paths = ttk.LabelFrame(self, text="Paths")
        grp_paths.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)

        ttk.Label(grp_paths, text="Input folder (optional):").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        ent_in = ttk.Entry(grp_paths, textvariable=self.input_var, width=72)
        ent_in.grid(row=0, column=1, sticky="ew", padx=6, pady=6)
        ttk.Button(grp_paths, text="Choose folder (optional)…", command=self._choose_input).grid(row=0, column=2, padx=6, pady=6)

        ttk.Label(grp_paths, text="Output Excel file (.xlsx):").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        ent_out = ttk.Entry(grp_paths, textvariable=self.output_var, width=72)
        ent_out.grid(row=1, column=1, sticky="ew", padx=6, pady=6)
        ttk.Button(grp_paths, text="Save as…", command=self._choose_output).grid(row=1, column=2, padx=6, pady=6)

        # Optional: select specific XML files (shows files in dialog)
        btns = ttk.Frame(grp_paths)
        btns.grid(row=2, column=0, columnspan=3, sticky="w", padx=6, pady=(0,6))
        ttk.Label(btns, text="Select the XML files you want to convert:").grid(row=0, column=0, padx=(0,10))
        ttk.Button(btns, text="Select XML files…", command=self._choose_files).grid(row=0, column=1, padx=(0,6))
        ttk.Button(btns, text="Clear list", command=self._clear_files).grid(row=0, column=2)

        # Preview of selected files (optional)
        grp_sel = ttk.LabelFrame(self, text="Selected files (optional)")
        grp_sel.grid(row=1, column=0, sticky="nsew", padx=4, pady=4)
        self.listbox = tk.Listbox(grp_sel, height=8)
        self.listbox.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        yscroll2 = ttk.Scrollbar(grp_sel, orient="vertical", command=self.listbox.yview)
        yscroll2.grid(row=0, column=1, sticky="ns")
        self.listbox.configure(yscrollcommand=yscroll2.set)

        # Options
        grp_opts = ttk.LabelFrame(self, text="Options")
        grp_opts.grid(row=2, column=0, sticky="ew", padx=4, pady=4)
        ttk.Checkbutton(grp_opts, text="Verbose logs", variable=self.verbose_var).grid(row=0, column=0, padx=6, pady=6)

        # Actions
        bar = ttk.Frame(self)
        bar.grid(row=3, column=0, sticky="ew", padx=4, pady=4)
        self.btn_run = ttk.Button(bar, text="Convert", command=self._on_run, style="Accent.TButton")
        self.btn_run.grid(row=0, column=0, padx=6, pady=4)
        self.prog = ttk.Progressbar(bar, mode="determinate", length=280)
        self.prog.grid(row=0, column=1, padx=12, pady=4)
        self.lbl_prog = ttk.Label(bar, text="Idle")
        self.lbl_prog.grid(row=0, column=2, padx=6, pady=4)

        # Log box
        grp_log = ttk.LabelFrame(self, text="Log")
        grp_log.grid(row=4, column=0, sticky="nsew", padx=4, pady=4)
        self.txt = tk.Text(grp_log, height=12, wrap="word")
        self.txt.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        yscroll = ttk.Scrollbar(grp_log, orient="vertical", command=self.txt.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        self.txt.configure(yscrollcommand=yscroll.set)

        # Layout weights
        self.columnconfigure(0, weight=1)
        for r in (0, 1, 4):
            self.rowconfigure(r, weight=1)
        grp_paths.columnconfigure(1, weight=1)
        grp_sel.columnconfigure(0, weight=1)
        grp_sel.rowconfigure(0, weight=1)
        grp_log.columnconfigure(0, weight=1)
        grp_log.rowconfigure(0, weight=1)

        self.pack(fill="both", expand=True)

    # --------------------- Handlers ---------------------
    def _choose_input(self) -> None:
        d = filedialog.askdirectory(initialdir=self.input_var.get())
        if d:
            self.input_var.set(d)
            # preview in log
            dir_path = Path(d)
            xmls = iter_xml_files(dir_path)
            self._log_clear()
            self._log(f"Selected folder: {dir_path}")
            self._log(f"Found {len(xmls)} XML file(s).")
            if xmls[:10]:
                self._log("Samples: - " + " - ".join(p.name for p in xmls[:10]) + "")
            if len(xmls) > 10:
                self._log(f"...and {len(xmls) - 10} more.")

    def _choose_files(self) -> None:
        file_paths = filedialog.askopenfilenames(
            title="Select XML files…",
            filetypes=[("XML files", "*.xml *.XML"), ("All files", "*.*")],
            initialdir=self.input_var.get() or str(Path.cwd()),
        )
        if file_paths:
            # accumulate unique files
            new_files = [Path(p) for p in file_paths]
            for p in new_files:
                if p not in self.selected_files:
                    self.selected_files.append(p)
            self._refresh_file_list()
            self._log(f"Added {len(new_files)} file(s). Total selected: {len(self.selected_files)}")

    def _clear_files(self) -> None:
        self.selected_files.clear()
        self._refresh_file_list()
        self._log("Cleared selected files list.")

    def _refresh_file_list(self) -> None:
        self.listbox.delete(0, "end")
        for p in self.selected_files:
            self.listbox.insert("end", p.name)

    def _choose_output(self) -> None:
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

    def _on_run(self) -> None:
        in_dir = Path(self.input_var.get())
        out_xlsx = Path(self.output_var.get())

        if not self.selected_files:
            # If user hasn't selected files yet, prompt now
            messagebox.showinfo("Select files", "Please select one or more XML files to convert.")
            self._choose_files()
            if not self.selected_files:
                return

        if not out_xlsx.parent.exists():
            try:
                out_xlsx.parent.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Output error", f"Failed to create output folder:{out_xlsx.parent}{e}")
                return

        self._lock_ui(True)
        self._log_clear()
        self._log(f"Output: {out_xlsx}")
        self._log(f"Using {len(self.selected_files)} selected file(s).")
        self._log("Starting…")

        files_copy = list(self.selected_files)
        t = threading.Thread(target=self._worker, args=(in_dir, out_xlsx, self.verbose_var.get(), files_copy), daemon=True)
        t.start()

    # --------------------- Worker thread ---------------------
    def _worker(self, in_dir: Path, out_xlsx: Path, verbose: bool, files: list[Path]) -> None:
        try:
            logging.getLogger().setLevel(logging.DEBUG if verbose else logging.INFO)

            # Determine source list: user-selected files (if any) or all .xml under folder
            # In file-selection mode we must have an explicit list
            xmls = [p for p in files if p.exists() and p.is_file()]
            total = len(xmls)
            if total == 0:
                self._log("No .xml files found.")
                self._set_progress(0, 1, "Idle")
                return

            def on_progress(cur: int, tot: int, name: str) -> None:
                self._set_progress(cur, tot, f"{cur}/{tot} — {name}")

            # Process files in deterministic order
            records: List[Dict[str, Any]] = []
            errors: List[str] = []
            for i, xml_path in enumerate(sorted(xmls), start=1):
                on_progress(i, total, xml_path.name)
                rec = parse_nfe_file(xml_path)
                if rec is None:
                    errors.append(xml_path.name)
                else:
                    records.append(rec)

            if not records:
                self._log("No valid records produced. Nothing to save.")
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
                self._log("⚠ Missing dependency to write .xlsx. Install with:pip install openpyxl")
                raise

            self._log(f"✔ Done! Saved: {out_xlsx}")
            if errors:
                self._log(f"Files ignored or with errors: {len(errors)}")
        except Exception as e:
            self._log(f"ERROR: {e}")
        finally:
            self._set_progress(1, 1, "Idle")
            self._lock_ui(False)
            self._set_progress(1, 1, "Idle")
            self._lock_ui(False)

    # --------------------- UI helpers ---------------------
    def _lock_ui(self, busy: bool) -> None:
        state = tk.DISABLED if busy else tk.NORMAL
        for child in self.winfo_children():
            self._set_state_recursive(child, state)
        # Keep log box always enabled for scrolling
        self.txt.configure(state=tk.NORMAL)

    def _set_state_recursive(self, widget: tk.Widget, state: str) -> None:
        try:
            widget.configure(state=state)
        except tk.TclError:
            pass
        if isinstance(widget, (ttk.Frame, ttk.LabelFrame)):
            for child in widget.winfo_children():
                self._set_state_recursive(child, state)

    def _set_progress(self, cur: int, tot: int, text: str) -> None:
        def _update():
            self.prog.configure(maximum=max(tot, 1))
            self.prog['value'] = min(cur, tot)
            self.lbl_prog.configure(text=text)
        self.after(0, _update)

    def _log(self, msg: str) -> None:
        def _append():
            self.txt.insert("end", msg)
            self.txt.see("end")
        self.after(0, _append)

    def _log_clear(self) -> None:
        self.txt.delete("1.0", "end")


if __name__ == "__main__":
    root = tk.Tk()
    # Try to use a nicer theme if available
    try:
        from tkinter import ttk
        root.call("source", "sun-valley.tcl")  # optional custom theme, ignore if missing
        root.call("set_theme", "light")
    except Exception:
        pass
    App(root)
    root.mainloop()
