from openpyxl import load_workbook

from src.config import COLUMNS_TO_KEEP, HIGHLIGHT_COLUMNS, OUTSIDE_SHEET, WITHIN_SHEET
from src.excel.styles import HIGHLIGHT_BORDER, HIGHLIGHT_FILL, HIGHLIGHT_FONT, MONTH_FILL_COLORS
from src.excel.workbook_builder import build_workbook

HEADER = ["reporting_month", *COLUMNS_TO_KEEP]


def _row(reporting_month="January 2026", tlc="1/15/2026", receive_date="1/20/2026 00:00:00",
         vend="Acme Co", po="PO123", dept="IT"):
    return [reporting_month, vend, "1/1/2026", po, "1/10/2026", "INV1", tlc,
            receive_date, 100.00, dept, "TLC1", "Jane", "Doe"]


def _build_and_reload(tmp_path, sheets):
    """build_workbook returns a write-only Workbook, which streams cells
    straight to disk and can't be read back in memory - so tests save it and
    re-open it with load_workbook() to inspect the actual written content,
    the same round-trip save_workbook does in production."""
    wb, text_cols, slicer_info = build_workbook(HEADER, sheets)
    path = tmp_path / "out.xlsx"
    wb.save(path)
    return load_workbook(path), text_cols, slicer_info


def test_build_workbook_creates_one_sheet_per_bucket_with_correct_names(tmp_path):
    sheets = {WITHIN_SHEET: [_row()], OUTSIDE_SHEET: [_row()]}

    wb, _, _ = _build_and_reload(tmp_path, sheets)

    assert wb.sheetnames == [WITHIN_SHEET, OUTSIDE_SHEET]


def test_build_workbook_shades_rows_by_reporting_month(tmp_path):
    rows = [
        _row(reporting_month="January 2026", receive_date="1/20/2026 00:00:00"),
        _row(reporting_month="January 2026", receive_date="1/21/2026 00:00:00"),
        _row(reporting_month="February 2026", receive_date="2/5/2026 00:00:00"),
    ]

    wb, _, _ = _build_and_reload(tmp_path, {WITHIN_SHEET: rows})
    ws = wb[WITHIN_SHEET]

    # Rows are sorted by reporting_month before writing, so rows 2-3 are
    # both January (same shading color) and row 4 is February (next color
    # in the MONTH_FILL_COLORS cycle).
    jan_fill_a = ws.cell(row=2, column=1).fill.fgColor.rgb
    jan_fill_b = ws.cell(row=3, column=1).fill.fgColor.rgb
    feb_fill = ws.cell(row=4, column=1).fill.fgColor.rgb

    assert jan_fill_a == jan_fill_b == "00" + MONTH_FILL_COLORS[0]
    assert feb_fill == "00" + MONTH_FILL_COLORS[1]


def test_build_workbook_highlights_tlc_and_receive_date_columns(tmp_path):
    wb, _, _ = _build_and_reload(tmp_path, {WITHIN_SHEET: [_row()]})
    ws = wb[WITHIN_SHEET]

    for col_name in HIGHLIGHT_COLUMNS:
        col_idx = HEADER.index(col_name) + 1
        cell = ws.cell(row=2, column=col_idx)
        assert cell.fill.fgColor.rgb == HIGHLIGHT_FILL.fgColor.rgb
        assert cell.font.bold == HIGHLIGHT_FONT.bold
        assert cell.font.color.rgb == HIGHLIGHT_FONT.color.rgb
        assert cell.border.left.style == HIGHLIGHT_BORDER.left.style
        assert cell.border.left.color.rgb == HIGHLIGHT_BORDER.left.color.rgb
        assert cell.number_format == "mm/dd/yyyy"


def test_build_workbook_attaches_table_with_correct_headers(tmp_path):
    wb, _, slicer_info = _build_and_reload(tmp_path, {WITHIN_SHEET: [_row()]})
    ws = wb[WITHIN_SHEET]

    assert "Table1" in ws.tables
    table = ws.tables["Table1"]
    assert [col.name for col in table.tableColumns] == HEADER
    assert slicer_info[WITHIN_SHEET] == (1, HEADER.index("reporting_month") + 1, len(HEADER))


def test_build_workbook_handles_empty_bucket_without_error(tmp_path):
    wb, text_cols, slicer_info = _build_and_reload(
        tmp_path, {WITHIN_SHEET: [], OUTSIDE_SHEET: [_row()]}
    )
    ws = wb[WITHIN_SHEET]
    ws_outside = wb[OUTSIDE_SHEET]

    # Verify empty bucket is handled correctly
    assert ws.max_row == 1  # header row only, no data rows
    assert ws.cell(row=1, column=1).value == "reporting_month"
    assert "Table1" in ws.tables
    assert WITHIN_SHEET in text_cols
    assert WITHIN_SHEET in slicer_info

    # Verify sibling non-empty bucket still builds correctly
    assert ws_outside.max_row == 2  # header + 1 data row
    assert "Table2" in ws_outside.tables
    assert OUTSIDE_SHEET in text_cols
    assert OUTSIDE_SHEET in slicer_info
