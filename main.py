#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
NFe XML -> Excel Converter
- Reads all .xml files in a folder (recursively)
- Extracts key fields: id/key, issuer, recipient, and recipient address
- Outputs a clean .xlsx file
"""

from __future__ import annotations
import argparse
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import xmltodict


# --------------------------- Logging ---------------------------
def setup_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(level=level, format="%(levelname)s | %(message)s")


# --------------------------- XML helpers ---------------------------
def _get(d: Dict[str, Any], path: str, default: Any = None) -> Any:
    """
    Safe nested access:
    _get(obj, "a.b.c") -> obj["a"]["b"]["c"] or default if any key is missing.
    Supports '@attr' for XML attributes.
    """
    cur = d
    for part in path.split("."):
        if not isinstance(cur, dict) or part not in cur:
            return default
        cur = cur[part]
    return cur


def parse_nfe_file(xml_path: Path) -> Optional[Dict[str, Any]]:
    """
    Read an NFe XML file and return a dict with relevant fields.
    Returns None if the file doesn't look like a valid NFe payload.
    """
    try:
        with xml_path.open("rb") as f:
            data = xmltodict.parse(f)
    except Exception as e:
        logging.warning(f"Failed to parse XML: {xml_path.name} | {e}")
        return None

    # Typical structure: nfeProc -> NFe -> infNFe
    inf = _get(data, "nfeProc.NFe.infNFe")
    if not isinstance(inf, dict):
        # Some files may come without nfeProc (just NFe)
        inf = _get(data, "NFe.infNFe")
    if not isinstance(inf, dict):
        logging.debug(f"Ignored (doesn't look like NFe): {xml_path.name}")
        return None

    # Note ID (attribute @Id). Usually starts with "NFe" + key.
    note_id = _get(inf, "@Id", "")
    key = note_id.replace("NFe", "") if note_id else ""

    # Issuer / Recipient
    issuer_name = _get(inf, "emit.xNome", "")
    recipient_name = _get(inf, "dest.xNome", "")

    # Recipient address (when present)
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

    # Some NFes don't have enderDest (special operations). Fallback to enderEmit.
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

    # Final record
    return {
        "key": key,
        "note_id": note_id,
        "issuer_name": issuer_name,
        "recipient_name": recipient_name,
        **address,
        "file_name": xml_path.name,
    }


def parse_folder(input_dir: Path) -> Tuple[List[Dict[str, Any]], List[str]]:
    """
    Walk the input folder, process all .xml files.
    Returns (valid_records, error_file_names).
    """
    records: List[Dict[str, Any]] = []
    errors: List[str] = []

    xmls = sorted([p for p in input_dir.glob("**/*.xml") if p.is_file()])
    if not xmls:
        logging.warning(f"No .xml files found in: {input_dir}")

    for i, xml_path in enumerate(xmls, start=1):
        logging.debug(f"[{i}/{len(xmls)}] Reading {xml_path.name}")
        rec = parse_nfe_file(xml_path)
        if rec is None:
            errors.append(xml_path.name)
        else:
            records.append(rec)

    return records, errors


# --------------------------- CLI / Main ---------------------------
def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Convert NFe XML to Excel (.xlsx).")
    p.add_argument(
        "-i",
        "--input",
        type=Path,
        default=Path("NFs"),
        help="Folder containing .xml files (default: ./NFs)",
    )
    p.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("Invoices.xlsx"),
        help="Output .xlsx file path (default: ./Invoices.xlsx)",
    )
    p.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Enable verbose (debug) logs.",
    )
    return p


def main() -> None:
    args = build_arg_parser().parse_args()
    setup_logging(args.verbose)

    input_dir: Path = args.input
    output_xlsx: Path = args.output

    if not input_dir.exists() or not input_dir.is_dir():
        logging.error(f"Input folder does not exist or is not a directory: {input_dir}")
        raise SystemExit(1)

    records, errors = parse_folder(input_dir)

    if not records:
        logging.error("No valid records produced. Nothing to save.")
        if errors:
            logging.info(f"Files with error/ignored: {len(errors)}")
        raise SystemExit(2)

    # DataFrame sorted by key (when present)
    cols = [
        "key",
        "note_id",
        "issuer_name",
        "recipient_name",
        "dest_street",
        "dest_number",
        "dest_district",
        "dest_city",
        "dest_state",
        "dest_zip",
        "dest_country",
        "file_name",
    ]
    df = pd.DataFrame(records)
    # Ensure columns exist and are ordered
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    df = df[cols].sort_values(by=["key", "file_name"], na_position="last")

    # Save Excel
    try:
        df.to_excel(output_xlsx, index=False)
        logging.info(f"✔ Excel generated: {output_xlsx.resolve()}")
        logging.info(f"Records: {len(df)} | Files with error/ignored: {len(errors)}")
        if errors:
            logging.debug("Files with error: " + ", ".join(errors))
    except Exception as e:
        logging.error(f"Failed to save Excel: {e}")
        raise SystemExit(3)


if __name__ == "__main__":
    main()
