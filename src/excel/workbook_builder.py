"""Builds the styled openpyxl Workbook: one sheet per reporting-period
bucket, sorted and shaded by reporting_month, with tlc/receive_date
highlighted and a real Excel Table (backing the Slicer added later by
workbook_writer.save_workbook).
"""
import logging

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo

from ..config import HIGHLIGHT_COLUMNS, TEXT_COLUMNS
from ..logging_config import log_stage
from ..transform import parse_reporting_month, to_date_only
from .styles import (
    HEADER_ALIGNMENT,
    HEADER_BORDER,
    HEADER_FILL,
    HEADER_FONT,
    HEADER_HIGHLIGHT_FILL,
    HEADER_HIGHLIGHT_FONT,
    HIGHLIGHT_ALIGNMENT,
    HIGHLIGHT_BORDER,
    HIGHLIGHT_DATE_FORMAT,
    HIGHLIGHT_FILL,
    HIGHLIGHT_FONT,
    MAX_COLUMN_WIDTH,
    MIN_COLUMN_WIDTH,
    MONTH_FILL_COLORS,
)

logger = logging.getLogger(__name__)

TEXT_ALIGNMENT = Alignment(horizontal="left")


# log_stage is built on contextlib.contextmanager, which Python allows to
# double as a decorator - used that way here instead of a `with` block so
# this ~100-line function body doesn't need an extra indent level.
@log_stage(logger, "build styled workbook (shading, headers, tables)")
def build_workbook(header, sheets):
    # write_only streams each row straight to disk instead of materializing
    # every cell as a full-blown openpyxl Cell object in an in-memory grid -
    # at 100k+ rows the per-cell style assignments were ~76% of this
    # pipeline's runtime. The tradeoff is that a write-only worksheet is
    # append-only: cells can't be styled or read back after the fact, so
    # every style below is built into each WriteOnlyCell before it's
    # appended, and structural, sheet-level properties (column widths,
    # freeze_panes) must be set before the first row goes out.
    wb = Workbook(write_only=True)

    reporting_month_idx = header.index("reporting_month")
    highlight_indices = {header.index(c) for c in HIGHLIGHT_COLUMNS}
    # reporting_month / po / invoiceid must stay Text-formatted and
    # left-aligned so Excel doesn't auto-detect them and re-flag them.
    text_format_indices = {reporting_month_idx} | {header.index(c) for c in TEXT_COLUMNS}

    # sheet_name -> (text_column_letters, last_row), needed to re-patch the
    # "Number Stored as Text" ignoredErrors block after save (openpyxl does
    # not serialize it).
    text_column_letters_by_sheet = {}
    # sheet_name -> SheetSlicerTarget-ish info, needed to patch in a real
    # reporting_month Slicer after save (openpyxl can't create one itself -
    # see excel_slicers.py).
    slicer_info_by_sheet = {}

    for table_id, (sheet_name, rows) in enumerate(sheets.items(), start=1):
        logger.debug("Building sheet %r: %d row(s)", sheet_name, len(rows))
        ws = wb.create_sheet(sheet_name)

        data_rows = sorted(rows, key=lambda r: parse_reporting_month(r[reporting_month_idx]))
        data_rows = [
            tuple(
                to_date_only(value) if col_num in highlight_indices else value
                for col_num, value in enumerate(row)
            )
            for row in data_rows
        ]
        last_row = len(data_rows) + 1

        # Column widths (auto-sized to each column's longest value) and
        # freeze_panes both precede the cell data in the worksheet XML, so
        # write-only mode requires setting them before any row is appended.
        for col_idx, col_name in enumerate(header, start=1):
            col_letter = get_column_letter(col_idx)
            longest = max(
                [len(str(col_name))]
                + [len(str(row[col_idx - 1])) for row in data_rows if row[col_idx - 1] is not None]
            )
            ws.column_dimensions[col_letter].width = max(
                MIN_COLUMN_WIDTH, min(longest + 2, MAX_COLUMN_WIDTH)
            )
        ws.freeze_panes = "A2"

        # Header row: bold/centered/grey, with tlc/receive_date called out
        # in the accent color.
        header_cells = []
        for col_idx, col_name in enumerate(header, start=1):
            cell = WriteOnlyCell(ws, value=col_name)
            if col_idx - 1 in highlight_indices:
                cell.fill = HEADER_HIGHLIGHT_FILL
                cell.font = HEADER_HIGHLIGHT_FONT
                cell.border = HIGHLIGHT_BORDER
            else:
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.border = HEADER_BORDER
            cell.alignment = HEADER_ALIGNMENT
            header_cells.append(cell)
        ws.append(header_cells)

        # Data rows: shaded by reporting_month, alternating colors each time
        # the reporting_month changes; tlc/receive_date are then styled on
        # top of the shading so they stand out consistently in every row.
        color_idx = -1
        prev_month = None
        for row in data_rows:
            month = row[reporting_month_idx]
            if month != prev_month:
                color_idx = (color_idx + 1) % len(MONTH_FILL_COLORS)
                prev_month = month
            fill = PatternFill("solid", fgColor=MONTH_FILL_COLORS[color_idx])

            row_cells = []
            for col_num, value in enumerate(row):
                cell = WriteOnlyCell(ws, value=value)
                if col_num in highlight_indices:
                    cell.fill = HIGHLIGHT_FILL
                    cell.font = HIGHLIGHT_FONT
                    cell.border = HIGHLIGHT_BORDER
                    cell.number_format = HIGHLIGHT_DATE_FORMAT
                    cell.alignment = HIGHLIGHT_ALIGNMENT
                else:
                    cell.fill = fill
                    if col_num in text_format_indices:
                        cell.number_format = "@"
                        cell.alignment = TEXT_ALIGNMENT
                row_cells.append(cell)
            ws.append(row_cells)

        text_column_letters = [get_column_letter(header.index(c) + 1) for c in TEXT_COLUMNS]
        text_column_letters_by_sheet[ws.title] = (text_column_letters, last_row)

        # Real Excel Table backing the reporting_month Slicer added below.
        # No built-in style/banding, so it doesn't fight with the month
        # shading and header styling already applied above.
        table_ref = f"A1:{get_column_letter(len(header))}{last_row}"
        table = Table(
            id=table_id,
            displayName=f"Table{table_id}",
            ref=table_ref,
            tableColumns=[TableColumn(id=i, name=str(name)) for i, name in enumerate(header, start=1)],
        )
        table.tableStyleInfo = TableStyleInfo(
            showFirstColumn=False, showLastColumn=False, showRowStripes=False, showColumnStripes=False
        )
        ws.add_table(table)
        slicer_info_by_sheet[ws.title] = (table_id, reporting_month_idx + 1, len(header))

    return wb, text_column_letters_by_sheet, slicer_info_by_sheet
