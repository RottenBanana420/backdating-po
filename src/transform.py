"""Row-level transforms: column trimming, date parsing, and the
within/outside reporting-period split that is the whole point of this
pipeline.
"""
import logging
from datetime import date, datetime

from .config import COLUMNS_TO_KEEP, OUTSIDE_SHEET, WITHIN_SHEET

logger = logging.getLogger(__name__)


def build_trimmed_rows(header, rows):
    output_columns = ["reporting_month"] + COLUMNS_TO_KEEP
    col_idx = {name: header.index(name) for name in COLUMNS_TO_KEEP}

    out_rows = []
    for row in rows:
        receive_date = datetime.strptime(row[col_idx["receive_date"]], "%m/%d/%Y %H:%M:%S")
        reporting_month = receive_date.strftime("%B %Y")
        new_row = [reporting_month]
        for col in COLUMNS_TO_KEEP:
            value = row[col_idx[col]]
            if col == "rcvtotalamt":
                value = float(value)
            new_row.append(value)
        out_rows.append(new_row)

    return output_columns, out_rows


def parse_date(value):
    date_part = value.split(" ")[0]
    return datetime.strptime(date_part, "%m/%d/%Y").date()


def split_by_reporting_period(header, rows):
    tlc_idx = header.index("tlc")
    receive_idx = header.index("receive_date")

    within_rows, outside_rows = [], []
    deleted_count = 0

    for row in rows:
        tlc_date = parse_date(row[tlc_idx])
        receive_date = parse_date(row[receive_idx])

        if tlc_date == receive_date:
            deleted_count += 1
            continue

        if (tlc_date.year, tlc_date.month) == (receive_date.year, receive_date.month):
            within_rows.append(row)
        else:
            outside_rows.append(row)

    logger.info("Deleted (tlc == receive_date): %d", deleted_count)
    logger.info("%s: %d rows", WITHIN_SHEET, len(within_rows))
    logger.info("%s: %d rows", OUTSIDE_SHEET, len(outside_rows))

    return {WITHIN_SHEET: within_rows, OUTSIDE_SHEET: outside_rows}


def parse_reporting_month(value):
    return datetime.strptime(value, "%B %Y")


def to_date_only(value):
    # tlc / receive_date come in as timestamp strings, e.g.
    # "5/26/2026 13:03:07:59" / "5/22/2026 00:00:00" - the time portion
    # isn't meaningful for the tlc-vs-receive_date comparison, so keep
    # only the date.
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    date_part = str(value).split(" ")[0]
    return datetime.strptime(date_part, "%m/%d/%Y").date()
