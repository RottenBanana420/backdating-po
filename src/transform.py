"""Row-level transforms: column trimming, date parsing, and the
within/outside reporting-period split that is the whole point of this
pipeline.
"""
import logging
from datetime import date, datetime

from .config import COLUMNS_TO_KEEP, OUTSIDE_SHEET, WITHIN_SHEET
from .logging_config import log_stage

logger = logging.getLogger(__name__)


def build_trimmed_rows(header, rows):
    with log_stage(logger, "trim rows to business columns and tag reporting_month"):
        logger.debug("Trimming %d row(s) down to %d business column(s)", len(rows), len(COLUMNS_TO_KEEP))
        output_columns = ["reporting_month", *COLUMNS_TO_KEEP]
        col_idx = {name: header.index(name) for name in COLUMNS_TO_KEEP}

        out_rows = []
        bad_rows = []
        # Tracked in this same pass (rather than rescanned from out_rows
        # afterwards) so the report's date-ranged output filename costs
        # nothing extra beyond two comparisons per already-visited row.
        min_month = None
        max_month = None
        for i, row in enumerate(rows, start=1):
            try:
                receive_date = datetime.strptime(row[col_idx["receive_date"]], "%m/%d/%Y %H:%M:%S")
                reporting_month = receive_date.strftime("%B %Y")
                new_row = [reporting_month]
                for col in COLUMNS_TO_KEEP:
                    value = row[col_idx[col]]
                    if col == "rcvtotalamt":
                        value = float(value)
                    new_row.append(value)
            except ValueError as exc:
                # An unparseable receive_date or rcvtotalamt is a data-quality
                # issue confined to this one row - quarantine it (skip, count,
                # keep going) the same way csv_loader.py isolates structurally
                # broken rows, rather than aborting a multi-thousand-row audit
                # over a single bad value.
                bad_rows.append((i, exc))
                continue
            out_rows.append(new_row)
            year_month = (receive_date.year, receive_date.month)
            if min_month is None or year_month < min_month:
                min_month = year_month
            if max_month is None or year_month > max_month:
                max_month = year_month

        if bad_rows:
            logger.warning("Rows skipped due to unparseable field values: %d (dropped)", len(bad_rows))
            for i, exc in bad_rows[:20]:
                logger.debug("  row %d: %s", i, exc)

    return output_columns, out_rows, len(bad_rows), (min_month, max_month)


def parse_date(value):
    date_part = value.split(" ")[0]
    return datetime.strptime(date_part, "%m/%d/%Y").date()


def split_by_reporting_period(header, rows):
    with log_stage(logger, "split rows by reporting period (within/outside)"):
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

        # Key pipeline metric: how the input was distributed across the
        # two report sheets, and how many rows were dropped as exact
        # tlc/receive_date matches (not backdated, so not report-worthy).
        logger.info("Deleted (tlc == receive_date): %d", deleted_count)
        logger.info("%s: %d rows", WITHIN_SHEET, len(within_rows))
        logger.info("%s: %d rows", OUTSIDE_SHEET, len(outside_rows))

    return {WITHIN_SHEET: within_rows, OUTSIDE_SHEET: outside_rows}


DEFAULT_REPORT_FILENAME = "po_reporting_periods.xlsx"


def _next_month(year_month):
    year, month = year_month
    return (year + 1, 1) if month == 12 else (year, month + 1)


def format_report_filename(month_bounds):
    """Builds the "Backdating POs {start}-{end}.xlsx" output filename from
    `month_bounds` (the (min_month, max_month) tuple `build_trimmed_rows`
    returns, each a (year, month) pair or None if there were no valid rows).

    `start` is the 1st of the earliest reporting month in the data; `end` is
    the 1st of the month *after* the latest reporting month - since this
    report runs monthly, a single-month file spans exactly that one month
    (e.g. reporting month July 2025 alone -> "7.1.25-8.1.25"), and a file
    covering several months' worth of data spans from the first to one past
    the last (e.g. July 2025 through June 2026 -> "7.1.25-7.1.26").
    """
    min_month, max_month = month_bounds
    if min_month is None or max_month is None:
        return DEFAULT_REPORT_FILENAME

    start_year, start_mo = min_month
    end_year, end_mo = _next_month(max_month)
    start = f"{start_mo}.1.{start_year % 100:02d}"
    end = f"{end_mo}.1.{end_year % 100:02d}"
    return f"Backdating POs {start}-{end}.xlsx"


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
