"""Loads and repairs the raw PO CSV export.

The source system occasionally emits rows with a stray bare CR or LF
embedded (unquoted) inside the vendoraccountid field. A naive line-based
CSV read tears such a row into two physical lines; this loader instead
splits only on the file's real record delimiter (\\r\\n), then strips any
stray \\r/\\n that survived within a line, so the row parses as one record
with the correct field count.
"""
import csv
import logging

from .config import ENCODING
from .logging_config import log_stage

logger = logging.getLogger(__name__)


def load_and_clean_rows(path):
    logger.debug("Reading CSV %s (encoding=%s)", path, ENCODING)
    with log_stage(logger, f"load and clean CSV: {path}"):
        with open(path, encoding=ENCODING, newline="") as f:
            raw = f.read()

        # Split into logical lines on the file's real record delimiter only.
        lines = [ln for ln in raw.split("\r\n") if ln.strip() != ""]

        header = None
        reader_rows = []
        bad_rows = []

        for i, line in enumerate(lines, start=1):
            # Strip any stray bare \r or \n that survived within the line
            # (these are the characters that caused rows to split early).
            cleaned = line.replace("\r", "").replace("\n", "")
            fields = next(csv.reader([cleaned]))
            if header is None:
                header = fields
                continue
            if len(fields) != len(header):
                bad_rows.append((i, len(fields), fields))
                continue
            reader_rows.append(fields)

        if header is None:
            raise ValueError(f"No header row found in {path} - the file appears to be empty.")

        logger.info("Header columns: %d", len(header))
        logger.info("Clean data rows: %d", len(reader_rows))
        if bad_rows:
            # A non-zero count means the source system emitted rows this
            # delimiter-repair pass couldn't fully recover - worth a
            # WARNING (not INFO) since those rows are dropped from the
            # report. The count is also returned to the caller (rather than
            # only logged) so it can reach the user - a WARNING log line
            # alone never reaches the Streamlit UI.
            logger.warning("Rows still misaligned after fix: %d (dropped)", len(bad_rows))
            for i, n, fields in bad_rows[:20]:
                logger.debug("  line %d: %d fields -> %s", i, n, fields)

    return header, reader_rows, len(bad_rows)
