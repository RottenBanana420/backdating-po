"""Orchestrates the full PO report pipeline: load -> transform -> split ->
build workbook -> save. This is the single function main.py calls.
"""
import logging
from pathlib import Path

from . import config
from .csv_loader import load_and_clean_rows
from .excel.workbook_builder import build_workbook
from .excel.workbook_writer import save_workbook
from .logging_config import log_stage
from .transform import build_trimmed_rows, format_report_filename, split_by_reporting_period

logger = logging.getLogger(__name__)


def run(
    src: Path = config.RAW_DATA_PATH,
    dst: Path | None = None,
    on_progress=None,
) -> dict:
    """`dst`, if omitted, is derived from the data itself once it's been read:
    data/output/"Backdating POs {start}-{end}.xlsx", with start/end taken
    from the earliest and latest reporting months found (see
    transform.format_report_filename). Pass an explicit `dst` to keep
    control over the exact output path/name instead.
    """
    logger.debug("Config: src=%s dst=%s encoding=%s", src, dst, config.ENCODING)

    if not src.exists():
        raise FileNotFoundError(
            f"Input file not found: {src}\n"
            "See README.md 'Running locally' for where to put your data, "
            "or pass --input data/sample/sample_raw_data.csv to try the demo dataset."
        )

    with log_stage(logger, f"full pipeline run ({src.name} -> {dst.name if dst else '<computed>'})"):
        raw_header, raw_rows, dropped_row_count = load_and_clean_rows(src)
        header, rows, skipped_row_count, month_bounds = build_trimmed_rows(raw_header, raw_rows)
        if dst is None:
            dst = config.OUTPUT_DIR / format_report_filename(month_bounds)
        sheets = split_by_reporting_period(header, rows)
        wb, text_column_letters_by_sheet, slicer_info_by_sheet = build_workbook(
            header, sheets, on_progress=on_progress
        )

        dst.parent.mkdir(parents=True, exist_ok=True)
        save_workbook(wb, text_column_letters_by_sheet, slicer_info_by_sheet, dst)

        logger.info("Written to: %s", dst)
        excluded_row_count = dropped_row_count + skipped_row_count
        if excluded_row_count:
            logger.warning(
                "Total rows excluded from this run: %d (see warnings above for detail)",
                excluded_row_count,
            )
        counts = {}
        for sheet_name, sheet_rows in sheets.items():
            logger.info("  %s: %d rows", sheet_name, len(sheet_rows))
            counts[sheet_name] = len(sheet_rows)

    return counts
