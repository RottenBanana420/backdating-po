import logging
from pathlib import Path

import pytest

from src.csv_loader import load_and_clean_rows

FIXTURE = Path(__file__).parent / "fixtures" / "sample_raw_loader.csv"
HEADER_ONLY_FIXTURE = Path(__file__).parent / "fixtures" / "header_only.csv"
TOO_BROKEN_FIXTURE = Path(__file__).parent / "fixtures" / "too_broken.csv"


def test_load_and_clean_rows_repairs_split_rows():
    header, rows = load_and_clean_rows(FIXTURE)

    assert header == ["a", "b", "c", "d", "e"]
    assert len(rows) == 4
    assert all(len(row) == 5 for row in rows)

    # The row with the stray embedded \r in field c should have that
    # character stripped, not left dangling or dropped as misaligned.
    assert rows[2] == ["11", "12", "13XX", "14", "15"]


def test_load_and_clean_rows_handles_header_only_csv():
    header, rows = load_and_clean_rows(HEADER_ONLY_FIXTURE)

    assert header == ["a", "b", "c", "d", "e"]
    assert rows == []


def test_load_and_clean_rows_drops_rows_with_unrecoverable_field_count(caplog):
    caplog.set_level(logging.INFO)

    header, rows = load_and_clean_rows(TOO_BROKEN_FIXTURE)

    assert header == ["a", "b", "c", "d", "e"]
    assert rows == [["11", "12", "13", "14", "15"]]
    assert "Rows still misaligned after fix: 1" in caplog.text


def test_load_and_clean_rows_raises_cleanly_on_non_cp1252_bytes(tmp_path):
    bad_file = tmp_path / "bad_encoding.csv"
    # 0x81 is an unassigned cp1252 code point; open(..., encoding="cp1252")
    # raises UnicodeDecodeError on it rather than silently corrupting data.
    bad_file.write_bytes(b"a,b\r\n1,\x81\r\n")

    with pytest.raises(UnicodeDecodeError):
        load_and_clean_rows(bad_file)


def test_load_and_clean_rows_raises_clear_error_on_zero_byte_file(tmp_path):
    empty_file = tmp_path / "empty.csv"
    empty_file.write_bytes(b"")

    with pytest.raises(ValueError, match="empty"):
        load_and_clean_rows(empty_file)
