from datetime import date

from backdating_po.transform import (
    build_trimmed_rows,
    parse_date,
    split_by_reporting_period,
)

RAW_HEADER = [
    "vend", "reqdate", "po", "dlvdate", "invoiceid", "tlc",
    "receive_date", "rcvtotalamt", "dept", "rcv_tlc",
    "rcv_firstname", "rcv_lastname",
]


def _row(receive_date="1/15/2026 00:00:00", tlc="1/15/2026"):
    return ["Acme Co", "1/1/2026", "PO123", "1/10/2026", "INV1", tlc,
            receive_date, "100.00", "IT", "TLC1", "Jane", "Doe"]


def test_build_trimmed_rows_adds_reporting_month():
    header, rows = build_trimmed_rows(RAW_HEADER, [_row()])

    assert header[0] == "reporting_month"
    assert rows[0][0] == "January 2026"
    assert isinstance(rows[0][header.index("rcvtotalamt")], float)


def test_parse_date_strips_time_component():
    assert parse_date("1/15/2026 00:00:00") == date(2026, 1, 15)


def test_split_by_reporting_period_buckets_correctly():
    header = ["reporting_month"] + RAW_HEADER

    def make_row(tlc, receive_date):
        return [None] + _row(receive_date=receive_date, tlc=tlc)

    within = make_row(tlc="1/15/2026", receive_date="1/20/2026 00:00:00")
    outside = make_row(tlc="1/15/2026", receive_date="2/20/2026 00:00:00")
    dropped = make_row(tlc="1/15/2026", receive_date="1/15/2026 00:00:00")

    buckets = split_by_reporting_period(header, [within, outside, dropped])

    assert len(buckets["POs Within Reporting Month"]) == 1
    assert len(buckets["POs Outside Reporting Month"]) == 1
    assert buckets["POs Within Reporting Month"][0] == within
    assert buckets["POs Outside Reporting Month"][0] == outside
