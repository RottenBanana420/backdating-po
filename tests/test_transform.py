import logging
from datetime import date, datetime

import pytest

from src.transform import (
    build_trimmed_rows,
    parse_date,
    split_by_reporting_period,
    to_date_only,
)

RAW_HEADER = [
    "vend", "reqdate", "po", "dlvdate", "invoiceid", "tlc",
    "receive_date", "rcvtotalamt", "dept", "rcv_tlc",
    "rcv_firstname", "rcv_lastname",
]


def _row(receive_date="1/15/2026 00:00:00", tlc="1/15/2026", rcvtotalamt="100.00"):
    return ["Acme Co", "1/1/2026", "PO123", "1/10/2026", "INV1", tlc,
            receive_date, rcvtotalamt, "IT", "TLC1", "Jane", "Doe"]


def test_build_trimmed_rows_adds_reporting_month():
    header, rows, skipped_row_count = build_trimmed_rows(RAW_HEADER, [_row()])

    assert header[0] == "reporting_month"
    assert rows[0][0] == "January 2026"
    assert isinstance(rows[0][header.index("rcvtotalamt")], float)
    assert skipped_row_count == 0


def test_parse_date_strips_time_component():
    assert parse_date("1/15/2026 00:00:00") == date(2026, 1, 15)


def test_split_by_reporting_period_buckets_correctly():
    header = ["reporting_month", *RAW_HEADER]

    def make_row(tlc, receive_date):
        return [None, *_row(receive_date=receive_date, tlc=tlc)]

    within = make_row(tlc="1/15/2026", receive_date="1/20/2026 00:00:00")
    outside = make_row(tlc="1/15/2026", receive_date="2/20/2026 00:00:00")
    dropped = make_row(tlc="1/15/2026", receive_date="1/15/2026 00:00:00")

    buckets = split_by_reporting_period(header, [within, outside, dropped])

    assert len(buckets["POs Within Reporting Month"]) == 1
    assert len(buckets["POs Outside Reporting Month"]) == 1
    assert buckets["POs Within Reporting Month"][0] == within
    assert buckets["POs Outside Reporting Month"][0] == outside


def test_build_trimmed_rows_handles_zero_rows():
    header, rows, skipped_row_count = build_trimmed_rows(RAW_HEADER, [])

    assert header[0] == "reporting_month"
    assert rows == []
    assert skipped_row_count == 0


def test_split_by_reporting_period_handles_zero_rows():
    header = ["reporting_month", *RAW_HEADER]

    buckets = split_by_reporting_period(header, [])

    assert buckets == {
        "POs Within Reporting Month": [],
        "POs Outside Reporting Month": [],
    }


def test_split_by_reporting_period_month_boundary_different_months():
    header = ["reporting_month", *RAW_HEADER]

    def make_row(tlc, receive_date):
        return [None, *_row(receive_date=receive_date, tlc=tlc)]

    # tlc on the last day of January, receive_date the very next day in
    # February - one day apart but different calendar months.
    row = make_row(tlc="1/31/2026", receive_date="2/1/2026 00:00:00")

    buckets = split_by_reporting_period(header, [row])

    assert buckets["POs Outside Reporting Month"] == [row]
    assert buckets["POs Within Reporting Month"] == []


def test_split_by_reporting_period_month_boundary_same_month():
    header = ["reporting_month", *RAW_HEADER]

    def make_row(tlc, receive_date):
        return [None, *_row(receive_date=receive_date, tlc=tlc)]

    # tlc on the 1st, receive_date on the 31st - 30 days apart but the
    # same calendar month.
    row = make_row(tlc="1/1/2026", receive_date="1/31/2026 00:00:00")

    buckets = split_by_reporting_period(header, [row])

    assert buckets["POs Within Reporting Month"] == [row]
    assert buckets["POs Outside Reporting Month"] == []


def test_parse_date_raises_on_malformed_input():
    with pytest.raises(ValueError):
        parse_date("not-a-date")


def test_build_trimmed_rows_skips_rows_with_malformed_receive_date():
    header, rows, skipped_row_count = build_trimmed_rows(RAW_HEADER, [_row(receive_date="not-a-date")])

    assert header[0] == "reporting_month"
    assert rows == []
    assert skipped_row_count == 1


def test_build_trimmed_rows_skips_rows_with_malformed_rcvtotalamt():
    _header, rows, skipped_row_count = build_trimmed_rows(RAW_HEADER, [_row(rcvtotalamt="not-a-number")])

    assert rows == []
    assert skipped_row_count == 1


def test_build_trimmed_rows_isolates_bad_rows_and_keeps_good_ones():
    good = _row()
    bad = _row(receive_date="not-a-date")

    _header, rows, skipped_row_count = build_trimmed_rows(RAW_HEADER, [good, bad])

    assert len(rows) == 1
    assert rows[0][0] == "January 2026"
    assert skipped_row_count == 1


def test_build_trimmed_rows_logs_warning_for_skipped_rows(caplog):
    caplog.set_level(logging.INFO)

    build_trimmed_rows(RAW_HEADER, [_row(receive_date="not-a-date")])

    assert "Rows skipped due to unparseable field values: 1" in caplog.text
    warnings = [r for r in caplog.records if r.levelno == logging.WARNING]
    assert len(warnings) == 1


def test_to_date_only_accepts_date_datetime_and_string_forms():
    assert to_date_only(date(2026, 1, 15)) == date(2026, 1, 15)
    assert to_date_only(datetime(2026, 1, 15, 13, 3, 7)) == date(2026, 1, 15)
    assert to_date_only("1/15/2026 00:00:00") == date(2026, 1, 15)
