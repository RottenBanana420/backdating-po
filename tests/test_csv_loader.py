from pathlib import Path

from backdating_po.csv_loader import load_and_clean_rows

FIXTURE = Path(__file__).parent / "fixtures" / "sample_raw_loader.csv"


def test_load_and_clean_rows_repairs_split_rows():
    header, rows = load_and_clean_rows(FIXTURE)

    assert header == ["a", "b", "c", "d", "e"]
    assert len(rows) == 4
    assert all(len(row) == 5 for row in rows)

    # The row with the stray embedded \r in field c should have that
    # character stripped, not left dangling or dropped as misaligned.
    assert rows[2] == ["11", "12", "13XX", "14", "15"]
