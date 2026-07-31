"""Streamlit UI for the PO backdating-audit pipeline.

Wraps the same pipeline main.py runs from the command line: upload a raw PO
CSV export, preview it, run the pipeline, preview the styled report's two
sheets, then download the .xlsx. All business logic (temp-file handling,
validation, running the pipeline) lives in src/uploads.py - this script only
renders widgets and reacts to their state.
"""
import logging

import streamlit as st

from src import config
from src.logging_config import configure as configure_logging
from src.uploads import (
    InvalidUpload,
    load_raw_preview,
    process_upload,
    validate_extension,
)

configure_logging()
logger = logging.getLogger(__name__)

st.set_page_config(page_title="PO Backdating Audit", page_icon="\U0001F4CB")
st.title("PO Backdating Audit")
st.write(
    "Upload a raw PO-receipt CSV export to flag purchase orders received in "
    "a different calendar month than their transaction-log-confirmation "
    "(TLC) date. See "
    "[docs/ARCHITECTURE.md](https://github.com/RottenBanana420/backdating-po/blob/main/docs/ARCHITECTURE.md) "
    "for how the pipeline works."
)

for key in ("csv_bytes", "filename", "raw_preview", "result"):
    st.session_state.setdefault(key, None)
for key in ("upload_error", "process_error"):
    st.session_state.setdefault(key, None)


def _rows_to_records(header, rows):
    """Converts (header, rows) into the list-of-dicts shape st.dataframe expects."""
    return [dict(zip(header, row)) for row in rows]


uploaded_file = st.file_uploader(
    "Upload raw PO CSV export", type=["csv"], help="Drag and drop, or click to browse."
)

if uploaded_file is not None and (
    uploaded_file.name != st.session_state.filename
    or uploaded_file.size != st.session_state.get("file_size")
):
    st.session_state.filename = uploaded_file.name
    st.session_state.file_size = uploaded_file.size
    st.session_state.csv_bytes = uploaded_file.getvalue()
    st.session_state.result = None
    st.session_state.process_error = None
    st.session_state.raw_preview = None
    st.session_state.upload_error = None
    try:
        validate_extension(uploaded_file.name)
        header, rows = load_raw_preview(st.session_state.csv_bytes)
        st.session_state.raw_preview = (header, rows)
    except InvalidUpload as exc:
        st.session_state.upload_error = str(exc)
    except ValueError as exc:
        st.session_state.upload_error = f"Couldn't read this CSV: {exc}"
    except UnicodeDecodeError:
        st.session_state.upload_error = (
            "This file doesn't look like a cp1252-encoded CSV export. "
            "Please upload the raw export as-is, not a re-saved/re-encoded copy."
        )

if st.session_state.upload_error:
    st.error(st.session_state.upload_error)

if st.session_state.raw_preview:
    header, rows = st.session_state.raw_preview
    st.subheader("Raw data preview")
    st.dataframe(_rows_to_records(header, rows))

run_clicked = st.button(
    "Run pipeline", disabled=st.session_state.raw_preview is None
)

if run_clicked:
    st.session_state.result = None
    st.session_state.process_error = None
    with st.spinner("Processing upload..."):
        try:
            st.session_state.result = process_upload(st.session_state.csv_bytes)
        except InvalidUpload as exc:
            st.session_state.process_error = str(exc)
        except ValueError as exc:
            st.session_state.process_error = f"Couldn't process this file: {exc}"
        except FileNotFoundError:
            st.session_state.process_error = (
                "Internal error preparing the file for processing. Please try uploading again."
            )
        except Exception:
            logger.exception("Unexpected error processing uploaded file")
            st.session_state.process_error = (
                "Something went wrong processing this file. Please check it "
                "matches the expected PO export format, or see the README "
                "troubleshooting section."
            )

if st.session_state.process_error:
    st.error(st.session_state.process_error)

if st.session_state.result:
    result = st.session_state.result
    st.subheader("Processed result")

    within_tab, outside_tab = st.tabs([config.WITHIN_SHEET, config.OUTSIDE_SHEET])
    with within_tab:
        st.caption(f"{result['counts'][config.WITHIN_SHEET]} row(s)")
        st.dataframe(
            _rows_to_records(result["header"], result["preview_sheets"][config.WITHIN_SHEET])
        )
    with outside_tab:
        st.caption(f"{result['counts'][config.OUTSIDE_SHEET]} row(s)")
        st.dataframe(
            _rows_to_records(result["header"], result["preview_sheets"][config.OUTSIDE_SHEET])
        )

    st.download_button(
        "Download report (.xlsx)",
        data=result["xlsx_bytes"],
        file_name="po_reporting_periods.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
