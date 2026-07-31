"""Builds the styled openpyxl Workbook: one sheet per reporting-period
bucket, sorted and shaded by reporting_month, with tlc/receive_date
highlighted and a real Excel Table (backing the Slicer added later by
workbook_writer.save_workbook).
"""
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo

from ..config import HIGHLIGHT_COLUMNS, TEXT_COLUMNS
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
from ..transform import parse_reporting_month, to_date_only


def build_workbook(header, sheets):
    wb = Workbook()
    wb.remove(wb.active)

    reporting_month_idx = header.index("reporting_month")
    highlight_indices = {header.index(c) for c in HIGHLIGHT_COLUMNS}

    # sheet_name -> (text_column_letters, last_row), needed to re-patch the
    # "Number Stored as Text" ignoredErrors block after save (openpyxl does
    # not serialize it).
    text_column_letters_by_sheet = {}
    # sheet_name -> SheetSlicerTarget-ish info, needed to patch in a real
    # reporting_month Slicer after save (openpyxl can't create one itself -
    # see excel_slicers.py).
    slicer_info_by_sheet = {}

    for table_id, (sheet_name, rows) in enumerate(sheets.items(), start=1):
        ws = wb.create_sheet(sheet_name)
        ws.append(list(header))

        data_rows = sorted(rows, key=lambda r: parse_reporting_month(r[reporting_month_idx]))
        data_rows = [
            tuple(
                to_date_only(value) if col_num in highlight_indices else value
                for col_num, value in enumerate(row)
            )
            for row in data_rows
        ]
        for row in data_rows:
            ws.append(list(row))

        # reporting_month / po / invoiceid must stay Text-formatted and
        # left-aligned so Excel doesn't auto-detect them and re-flag them.
        reporting_month_letter = get_column_letter(reporting_month_idx + 1)
        for cell in ws[reporting_month_letter][1:]:
            cell.number_format = "@"
            cell.alignment = Alignment(horizontal="left")

        text_column_letters = []
        for col_name in TEXT_COLUMNS:
            col_idx = header.index(col_name) + 1
            col_letter = get_column_letter(col_idx)
            text_column_letters.append(col_letter)
            for cell in ws[col_letter][1:]:
                cell.number_format = "@"
                cell.alignment = Alignment(horizontal="left")
        text_column_letters_by_sheet[ws.title] = (text_column_letters, ws.max_row)

        # Shade each full row by its reporting_month, alternating colors
        # each time the reporting_month changes.
        color_idx = -1
        prev_month = None
        for row_num, row in enumerate(data_rows, start=2):
            month = row[reporting_month_idx]
            if month != prev_month:
                color_idx = (color_idx + 1) % len(MONTH_FILL_COLORS)
                prev_month = month
            fill = PatternFill("solid", fgColor=MONTH_FILL_COLORS[color_idx])
            for col_num in range(1, len(header) + 1):
                ws.cell(row=row_num, column=col_num).fill = fill

        # Style every header cell (bold, centered, grey fill, bottom border)
        # and auto-size each column to fit its longest value.
        for col_idx, col_name in enumerate(header, start=1):
            col_letter = get_column_letter(col_idx)
            header_cell = ws.cell(row=1, column=col_idx)
            header_cell.fill = HEADER_FILL
            header_cell.font = HEADER_FONT
            header_cell.border = HEADER_BORDER
            header_cell.alignment = HEADER_ALIGNMENT

            longest = max(
                [len(str(col_name))]
                + [len(str(row[col_idx - 1])) for row in data_rows if row[col_idx - 1] is not None]
            )
            ws.column_dimensions[col_letter].width = max(
                MIN_COLUMN_WIDTH, min(longest + 2, MAX_COLUMN_WIDTH)
            )

        # Make tlc / receive_date stand out on top of the row shading. This
        # runs after the generic header styling above so it overrides it,
        # replacing the inherited reporting_month row color with a fixed
        # fill so these two columns stand out consistently in every row.
        for col_name in HIGHLIGHT_COLUMNS:
            col_idx = header.index(col_name) + 1
            col_letter = get_column_letter(col_idx)

            header_cell = ws.cell(row=1, column=col_idx)
            header_cell.fill = HEADER_HIGHLIGHT_FILL
            header_cell.font = HEADER_HIGHLIGHT_FONT
            header_cell.border = HIGHLIGHT_BORDER

            for cell in ws[col_letter][1:]:
                cell.fill = HIGHLIGHT_FILL
                cell.font = HIGHLIGHT_FONT
                cell.border = HIGHLIGHT_BORDER
                cell.number_format = HIGHLIGHT_DATE_FORMAT
                cell.alignment = HIGHLIGHT_ALIGNMENT

        ws.freeze_panes = "A2"

        # Real Excel Table backing the reporting_month Slicer added below.
        # No built-in style/banding, so it doesn't fight with the month
        # shading and header styling already applied above.
        table_ref = f"A1:{get_column_letter(len(header))}{ws.max_row}"
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
