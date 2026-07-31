"""Bridges Streamlit's in-memory uploaded-file bytes to the filesystem-based
pipeline. pipeline.run() (and workbook_writer.save_workbook() beneath it)
require real Paths, not BytesIO, because save_workbook re-opens the
destination file as a zip twice to hand-patch XML (ignoredErrors, Slicer -
see docs/ARCHITECTURE.md). This module writes upload bytes to a scoped temp
directory, runs the existing pipeline against it, and reads the result back
into memory. It has no Streamlit import so it stays UI-agnostic and testable
with plain bytes.
"""
import tempfile
from pathlib import Path

from .config import COLUMNS_TO_KEEP, OUTSIDE_SHEET, WITHIN_SHEET
from .csv_loader import load_and_clean_rows
from .pipeline import run as run_pipeline
from .transform import build_trimmed_rows, split_by_reporting_period


class InvalidUpload(ValueError):
    """Raised for uploads that fail validation before reaching the pipeline."""


def validate_extension(filename: str) -> None:
    """Rejects filenames that don't end in .csv (case-insensitive)."""
    if not filename.lower().endswith(".csv"):
        raise InvalidUpload(
            f"'{filename}' is not a .csv file. Please upload the raw PO export as a .csv."
        )


def validate_columns(header: list) -> None:
    """Rejects a header missing any of the business columns the pipeline needs."""
    missing = [col for col in COLUMNS_TO_KEEP if col not in header]
    if missing:
        raise InvalidUpload(
            "This CSV is missing expected columns: "
            f"{', '.join(missing)}. Please upload the raw PO export, unmodified."
        )


def load_raw_preview(csv_bytes: bytes, limit: int = 20):
    """Writes csv_bytes to a temp file, loads+validates it, returns
    (header, rows[:limit]) exactly as uploaded, before any transformation.
    """
    with tempfile.TemporaryDirectory() as tmp_dir:
        src = Path(tmp_dir) / "upload.csv"
        src.write_bytes(csv_bytes)
        header, rows = load_and_clean_rows(src)
        validate_columns(header)
        return header, rows[:limit]


def process_upload(csv_bytes: bytes, preview_limit: int = 20) -> dict:
    """Runs the full pipeline against uploaded CSV bytes.

    Returns a dict with:
      - xlsx_bytes: the generated report as bytes, ready for download
      - counts: {sheet_name: row_count} from pipeline.run()
      - header: the trimmed output columns (reporting_month + business columns)
      - preview_sheets: {sheet_name: rows[:preview_limit]} for on-screen preview
    """
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_dir_path = Path(tmp_dir)
        src = tmp_dir_path / "upload.csv"
        dst = tmp_dir_path / "report.xlsx"
        src.write_bytes(csv_bytes)

        # Re-derive the within/outside split directly from the pure transform
        # functions for the on-screen preview, rather than reading the styled
        # .xlsx back with openpyxl - the workbook's styling/typing round-trip
        # is lossy and unnecessary just to show a preview table. This means
        # the CSV is read twice (once here, once inside run_pipeline below);
        # pipeline.run() has no pre-parsed-rows parameter, and adding one
        # purely for this convenience isn't worth changing a tested public
        # function's contract for typical monthly-export file sizes.
        raw_header, raw_rows = load_and_clean_rows(src)
        validate_columns(raw_header)
        header, rows = build_trimmed_rows(raw_header, raw_rows)
        sheets = split_by_reporting_period(header, rows)

        counts = run_pipeline(src=src, dst=dst)
        xlsx_bytes = dst.read_bytes()

    return {
        "xlsx_bytes": xlsx_bytes,
        "counts": counts,
        "header": header,
        "preview_sheets": {
            WITHIN_SHEET: sheets[WITHIN_SHEET][:preview_limit],
            OUTSIDE_SHEET: sheets[OUTSIDE_SHEET][:preview_limit],
        },
    }
