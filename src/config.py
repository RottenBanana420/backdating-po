"""Business/config constants: input columns, output paths, sheet names.

Paths default to the repo-relative data/ layout but can be overridden via
CLI flags in main.py (see pipeline.run()).
"""
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
RAW_DATA_PATH = ROOT / "data" / "raw" / "raw_data.csv"
OUTPUT_DIR = ROOT / "data" / "output"
# Fallback name used only when no valid row exists to derive a date range
# from (see transform.format_report_filename) - the normal case is a
# dynamically dated "Backdating POs {start}-{end}.xlsx" filename.
OUTPUT_PATH = OUTPUT_DIR / "po_reporting_periods.xlsx"
ENCODING = "cp1252"

COLUMNS_TO_KEEP = [
    "vend",
    "reqdate",
    "po",
    "dlvdate",
    "invoiceid",
    "tlc",
    "receive_date",
    "rcvtotalamt",
    "dept",
    "rcv_tlc",
    "rcv_firstname",
    "rcv_lastname",
]
TEXT_COLUMNS = ["po", "invoiceid"]
HIGHLIGHT_COLUMNS = ["tlc", "receive_date"]

WITHIN_SHEET = "POs Within Reporting Month"
OUTSIDE_SHEET = "POs Outside Reporting Month"

# Titles shown merged/centered above each sheet's table - distinct from the
# tab names above, which stay unchanged.
WITHIN_TITLE = "POs Received Within Reporting Month"
OUTSIDE_TITLE = "POs Received Outside Reporting Month"
