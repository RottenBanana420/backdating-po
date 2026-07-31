from pathlib import Path

import pytest
from openpyxl import load_workbook

from src import config
from src.uploads import (
    InvalidUpload,
    load_raw_preview,
    process_upload,
    validate_columns,
    validate_extension,
)

SAMPLE = Path(__file__).parent.parent / "data" / "sample" / "sample_raw_data.csv"
WRONG_COLUMNS = Path(__file__).parent / "fixtures" / "wrong_columns.csv"
HEADER_ONLY = Path(__file__).parent / "fixtures" / "header_only.csv"


def test_validate_extension_rejects_non_csv_filename():
    with pytest.raises(InvalidUpload):
        validate_extension("data.txt")


def test_validate_extension_accepts_csv_case_insensitively():
    validate_extension("data.csv")
    validate_extension("DATA.CSV")


def test_validate_columns_rejects_missing_business_columns():
    header = ["vend", "reqdate", "po"]
    with pytest.raises(InvalidUpload, match="tlc"):
        validate_columns(header)


def test_validate_columns_accepts_full_header():
    header = config.COLUMNS_TO_KEEP + ["some_extra_column"]
    validate_columns(header)


def test_load_raw_preview_returns_header_and_limited_rows():
    csv_bytes = SAMPLE.read_bytes()

    header, rows = load_raw_preview(csv_bytes, limit=2)

    assert "vend" in header
    assert "reqhdr_id" in header  # raw header, not the trimmed business set
    assert len(rows) <= 2


def test_load_raw_preview_raises_invalid_upload_for_missing_columns():
    csv_bytes = WRONG_COLUMNS.read_bytes()

    with pytest.raises(InvalidUpload, match="tlc"):
        load_raw_preview(csv_bytes)


def test_load_raw_preview_raises_value_error_on_empty_bytes():
    with pytest.raises(ValueError, match="empty"):
        load_raw_preview(b"")


def test_process_upload_returns_xlsx_bytes_and_counts():
    csv_bytes = SAMPLE.read_bytes()

    result = process_upload(csv_bytes)

    assert result["xlsx_bytes"][:2] == b"PK"
    assert sum(result["counts"].values()) > 0
    assert set(result["preview_sheets"].keys()) == {
        config.WITHIN_SHEET,
        config.OUTSIDE_SHEET,
    }


def test_process_upload_xlsx_bytes_load_with_openpyxl(tmp_path):
    csv_bytes = SAMPLE.read_bytes()

    result = process_upload(csv_bytes)

    out = tmp_path / "out.xlsx"
    out.write_bytes(result["xlsx_bytes"])
    wb = load_workbook(out)
    assert set(wb.sheetnames) <= {config.WITHIN_SHEET, config.OUTSIDE_SHEET}


def test_process_upload_cleans_up_temp_directory():
    import tempfile

    captured_dirs = []
    real_tempdir = tempfile.TemporaryDirectory

    class TrackingTemporaryDirectory(real_tempdir):
        def __enter__(self):
            name = super().__enter__()
            captured_dirs.append(name)
            return name

    tempfile.TemporaryDirectory = TrackingTemporaryDirectory
    try:
        process_upload(SAMPLE.read_bytes())
    finally:
        tempfile.TemporaryDirectory = real_tempdir

    assert captured_dirs, "process_upload should use tempfile.TemporaryDirectory"
    assert not Path(captured_dirs[0]).exists()


def test_process_upload_propagates_pipeline_value_errors():
    bad_csv = (
        ",".join(config.COLUMNS_TO_KEEP) + "\n"
        + "Acme,1/1/2026,PO1,1/5/2026,INV1,1/10/2026,NOT-A-DATE,100.00,IT,TLC1,Jane,Doe\n"
    ).encode(config.ENCODING)

    with pytest.raises(ValueError):
        process_upload(bad_csv)
