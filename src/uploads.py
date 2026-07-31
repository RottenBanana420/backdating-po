"""Bridges Streamlit's in-memory uploaded-file bytes to the filesystem-based
pipeline. pipeline.run() (and workbook_writer.save_workbook() beneath it)
require real Paths, not BytesIO, because save_workbook re-opens the
destination file as a zip twice to hand-patch XML (ignoredErrors, Slicer -
see docs/ARCHITECTURE.md). This module writes upload bytes to a scoped temp
directory, runs the existing pipeline against it, and reads the result back
into memory. It has no Streamlit import so it stays UI-agnostic and testable
with plain bytes.
"""
import logging
import tempfile
from collections.abc import Callable
from pathlib import Path

from .config import COLUMNS_TO_KEEP, OUTSIDE_SHEET, WITHIN_SHEET
from .csv_loader import load_and_clean_rows
from .pipeline import run as run_pipeline
from .transform import build_trimmed_rows, format_report_filename, split_by_reporting_period

logger = logging.getLogger(__name__)


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
        logger.warning("Upload rejected: missing column(s) %s", missing)
        raise InvalidUpload(
            "This CSV is missing expected columns: "
            f"{', '.join(missing)}. Please upload the raw PO export, unmodified."
        )
    logger.debug("Column validation passed: all %d required column(s) present", len(COLUMNS_TO_KEEP))


def load_raw_preview(csv_bytes: bytes, limit: int = 20):
    """Writes csv_bytes to a temp file, loads+validates it, returns
    (header, rows[:limit], dropped_row_count) exactly as uploaded, before any
    transformation. dropped_row_count is how many rows load_and_clean_rows
    couldn't recover (structurally misaligned) and excluded - surfaced here
    so the UI can warn the user at upload time, not just log it.
    """
    with tempfile.TemporaryDirectory() as tmp_dir:
        src = Path(tmp_dir) / "upload.csv"
        src.write_bytes(csv_bytes)
        header, rows, dropped_row_count = load_and_clean_rows(src)
        validate_columns(header)
        return header, rows[:limit], dropped_row_count


def process_upload(
    csv_bytes: bytes,
    preview_limit: int = 20,
    on_stage: Callable[[str], None] | None = None,
    on_progress: Callable[[int, int], None] | None = None,
) -> dict:
    """Runs the full pipeline against uploaded CSV bytes.

    `on_stage`, if given, is called with a short stage key ("reading",
    "preparing", "building_report") at each checkpoint below - it exists so
    a UI layer (src/app.py) can show live progress without this module
    importing Streamlit or owning any UI copy. Business-logic callers (the
    CLI, tests) simply omit it and get identical behavior to before.

    `on_progress`, if given, is called with (rows_done, rows_total) several
    times over the course of the "building_report" stage - that stage alone
    dominates total runtime on large files, so `on_stage` firing once at its
    start and then going quiet until it's done makes the UI look frozen for
    over a minute at 100k+ rows. Passed straight through to
    pipeline.run/build_workbook, which owns the actual reporting interval.

    Returns a dict with:
      - xlsx_bytes: the generated report as bytes, ready for download
      - filename: suggested download filename, "Backdating POs {start}-{end}.xlsx"
        with start/end derived from the data's earliest/latest reporting months
      - counts: {sheet_name: row_count} from pipeline.run()
      - header: the trimmed output columns (reporting_month + business columns)
      - preview_sheets: {sheet_name: rows[:preview_limit]} for on-screen preview
      - dropped_row_count: rows excluded because the CSV row was structurally
        misaligned (bad field count) and unrecoverable
      - skipped_row_count: rows excluded because receive_date/rcvtotalamt
        couldn't be parsed
      Both counts are 0 for a clean file; a non-zero value means the report
      doesn't cover every input row, so the caller (src/app.py) surfaces
      them to the user rather than letting the shortfall go unnoticed.
    """
    def _notify(stage: str) -> None:
        if on_stage is not None:
            on_stage(stage)

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_dir_path = Path(tmp_dir)
        src = tmp_dir_path / "upload.csv"
        dst = tmp_dir_path / "report.xlsx"
        src.write_bytes(csv_bytes)
        logger.debug("Staged upload at %s (%d bytes)", src, len(csv_bytes))

        _notify("reading")
        # Re-derive the within/outside split directly from the pure transform
        # functions for the on-screen preview, rather than reading the styled
        # .xlsx back with openpyxl - the workbook's styling/typing round-trip
        # is lossy and unnecessary just to show a preview table. This means
        # the CSV is read twice (once here, once inside run_pipeline below);
        # pipeline.run() has no pre-parsed-rows parameter, and adding one
        # purely for this convenience isn't worth changing a tested public
        # function's contract for typical monthly-export file sizes.
        raw_header, raw_rows, dropped_row_count = load_and_clean_rows(src)
        validate_columns(raw_header)

        _notify("preparing")
        header, rows, skipped_row_count, month_bounds = build_trimmed_rows(raw_header, raw_rows)
        sheets = split_by_reporting_period(header, rows)

        _notify("building_report")
        counts = run_pipeline(src=src, dst=dst, on_progress=on_progress)
        xlsx_bytes = dst.read_bytes()
        logger.info("Report ready: %d bytes, %d total row(s)", len(xlsx_bytes), sum(counts.values()))
        if dropped_row_count or skipped_row_count:
            logger.warning(
                "Rows excluded from report: %d (dropped_row_count=%d, skipped_row_count=%d)",
                dropped_row_count + skipped_row_count, dropped_row_count, skipped_row_count,
            )

    return {
        "xlsx_bytes": xlsx_bytes,
        "filename": format_report_filename(month_bounds),
        "counts": counts,
        "header": header,
        "preview_sheets": {
            WITHIN_SHEET: sheets[WITHIN_SHEET][:preview_limit],
            OUTSIDE_SHEET: sheets[OUTSIDE_SHEET][:preview_limit],
        },
        "dropped_row_count": dropped_row_count,
        "skipped_row_count": skipped_row_count,
    }
