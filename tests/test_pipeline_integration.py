from pathlib import Path

import pytest
from openpyxl import load_workbook

from src.pipeline import run

FIXTURE = Path(__file__).parent.parent / "data" / "sample" / "sample_raw_data.csv"


def test_pipeline_end_to_end(tmp_path):
    dst = tmp_path / "out.xlsx"
    counts = run(src=FIXTURE, dst=dst)

    assert dst.exists()
    assert sum(counts.values()) > 0

    wb = load_workbook(dst)
    assert set(wb.sheetnames) <= {"POs Within Reporting Month", "POs Outside Reporting Month"}
    for name in wb.sheetnames:
        ws = wb[name]
        assert ws.cell(row=1, column=1).value == "reporting_month"


def test_pipeline_raises_clear_error_for_missing_input(tmp_path):
    missing = tmp_path / "does_not_exist.csv"
    with pytest.raises(FileNotFoundError, match="Input file not found"):
        run(src=missing, dst=tmp_path / "out.xlsx")
