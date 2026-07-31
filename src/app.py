"""Streamlit UI for the PO backdating-audit pipeline.

Wraps the same pipeline main.py runs from the command line: upload a raw PO
CSV export, preview it, run the pipeline, preview the styled report's two
sheets, then download the .xlsx. All business logic (temp-file handling,
validation, running the pipeline, and its own terminal logging) lives in
src/uploads.py and the modules below it - this script only renders widgets,
reacts to their state, and translates pipeline stage keys into plain-language
progress messages for non-technical users. Keep it that way: if a change
needs new business logic, it belongs in src/uploads.py, not here.
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

st.set_page_config(page_title="PO Backdating Audit", page_icon=":material/receipt_long:")

for key in ("csv_bytes", "filename", "raw_preview", "result"):
    st.session_state.setdefault(key, None)
for key in ("upload_error", "process_error"):
    st.session_state.setdefault(key, None)


# ---------------------------------------------------------------------------
# Reusable helpers
#
# Kept in this file (rather than a separate module) because src/app.py is
# still well under the size where splitting pays for itself - see
# .claude/skills/developing-with-streamlit/references/code-organization.md.
# ---------------------------------------------------------------------------

def _rows_to_records(header, rows):
    """Converts (header, rows) into the list-of-dicts shape st.dataframe expects."""
    return [dict(zip(header, row, strict=True)) for row in rows]


# Plain-language copy for each src.uploads.process_upload() checkpoint.
# uploads.py only knows technical stage keys ("reading", "preparing",
# "building_report") - it has no Streamlit import and shouldn't own UI copy.
# This dict is the one place that copy lives, reused for both the st.status
# label and its scrolling log line so the two never drift apart.
STAGE_MESSAGES = {
    "reading": "Reading and checking your file...",
    "preparing": "Sorting purchase orders into reporting periods...",
    "building_report": "Building your formatted Excel report...",
}

STEP_LABELS = ["1. Upload your file", "2. Run the audit", "3. Download your report"]


def _render_stage_tracker(current_index):
    """Shows a persistent 3-step tracker so the user always knows where they
    are, what's done, and what's left - independent of whichever step's
    detail section is currently expanded below it.
    """
    with st.container(horizontal=True, horizontal_alignment="distribute"):
        for i, label in enumerate(STEP_LABELS):
            if i < current_index:
                st.badge(label, icon=":material/check_circle:", color="green")
            elif i == current_index:
                st.badge(label, icon=":material/pending:", color="blue")
            else:
                st.badge(label, icon=":material/radio_button_unchecked:", color="gray")


def _render_technical_details(exc: Exception) -> None:
    """A collapsed-by-default expander holding the raw exception, so
    advanced users/developers can inspect the real failure without the
    default view leading with a wall of technical text.
    """
    with st.expander("Show technical details", icon=":material/code:"):
        st.code(f"{type(exc).__name__}: {exc}")


# ---------------------------------------------------------------------------
# Page intro
# ---------------------------------------------------------------------------

st.title("PO backdating audit")
st.write(
    "This tool checks your purchase orders and flags any that were "
    "received in a different month than they were confirmed in - so you "
    "can catch backdated POs before they hit your monthly report."
)
with st.expander("How this works", icon=":material/help:"):
    st.write(
        "1. Upload the raw PO-receipt export from your source system, "
        "unmodified, as a `.csv` file.\n"
        "2. The tool compares each PO's transaction-log-confirmation "
        "(TLC) date against its receive date.\n"
        "3. POs received in the same calendar month as their TLC date go "
        "on one report sheet; everything else goes on a second sheet for "
        "review.\n\n"
        "For the full technical write-up, see "
        "[docs/ARCHITECTURE.md](https://github.com/RottenBanana420/backdating-po/blob/main/docs/ARCHITECTURE.md)."
    )

_current_stage = 2 if st.session_state.result else (1 if st.session_state.raw_preview else 0)
_render_stage_tracker(_current_stage)


# ---------------------------------------------------------------------------
# Step 1 - Upload
# ---------------------------------------------------------------------------

with st.container(border=True):
    st.subheader("Step 1 - Upload your file", anchor=False)
    st.caption(
        "Use the raw PO-receipt export as-is - don't remove or rename "
        "columns before uploading, or the audit won't be able to find the "
        "data it needs."
    )

    uploaded_file = st.file_uploader(
        "Upload raw PO CSV export",
        type=["csv"],
        help="Drag and drop the file here, or click to browse for it.",
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
        logger.info("File selected: %s (%d bytes)", uploaded_file.name, uploaded_file.size)
        try:
            validate_extension(uploaded_file.name)
            header, rows, dropped_row_count = load_raw_preview(st.session_state.csv_bytes)
            st.session_state.raw_preview = (header, rows, dropped_row_count)
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
        st.error(st.session_state.upload_error, icon=":material/error:")

    if st.session_state.raw_preview:
        header, rows, dropped_row_count = st.session_state.raw_preview
        st.success(f"Loaded **{st.session_state.filename}** - looks good so far.", icon=":material/check_circle:")
        if dropped_row_count:
            st.warning(
                f"{dropped_row_count} row(s) in this file have data that couldn't be lined up into "
                "columns correctly and will be excluded from the audit. The rest of the file will "
                "still be processed normally.",
                icon=":material/warning:",
            )
        st.caption(f"Preview of the first {len(rows)} row(s), exactly as uploaded:")
        st.dataframe(_rows_to_records(header, rows), width="stretch")


# ---------------------------------------------------------------------------
# Step 2 - Run the audit
# ---------------------------------------------------------------------------

with st.container(border=True):
    st.subheader("Step 2 - Run the audit", anchor=False)
    st.caption(
        "This compares every PO's confirmation and receive dates and "
        "splits them into two report sheets. It usually takes a few "
        "seconds."
    )

    run_clicked = st.button(
        "Run the audit",
        icon=":material/play_arrow:",
        disabled=st.session_state.raw_preview is None,
        help="Enabled once a valid file is uploaded above.",
    )

    if run_clicked:
        st.session_state.result = None
        st.session_state.process_error = None
        st.session_state.process_error_exc = None
        logger.info("Audit started for %s", st.session_state.filename)

        with st.status("Running the audit...", expanded=True) as status:
            def _on_stage(stage: str) -> None:
                message = STAGE_MESSAGES.get(stage, stage)
                status.update(label=message)
                st.write(message)

            try:
                st.session_state.result = process_upload(
                    st.session_state.csv_bytes, on_stage=_on_stage
                )
            except InvalidUpload as exc:
                st.session_state.process_error = str(exc)
            except ValueError as exc:
                st.session_state.process_error = f"Couldn't process this file: {exc}"
            except FileNotFoundError:
                st.session_state.process_error = (
                    "Internal error preparing the file for processing. Please try uploading again."
                )
            except Exception as exc:
                logger.exception("Unexpected error processing uploaded file")
                st.session_state.process_error = (
                    "Something went wrong processing this file. Please check it "
                    "matches the expected PO export format, or see the README "
                    "troubleshooting section."
                )
                st.session_state.process_error_exc = exc

            if st.session_state.result:
                status.update(label="Audit complete", state="complete", expanded=False)
            else:
                status.update(label="Audit failed", state="error", expanded=False)

    if st.session_state.process_error:
        st.error(st.session_state.process_error, icon=":material/error:")
        if st.session_state.get("process_error_exc") is not None:
            _render_technical_details(st.session_state.process_error_exc)


# ---------------------------------------------------------------------------
# Step 3 - Results & download
# ---------------------------------------------------------------------------

if st.session_state.result:
    result = st.session_state.result

    with st.container(border=True):
        st.subheader("Step 3 - Download your report", anchor=False)
        st.success(
            f"Audit complete: **{sum(result['counts'].values())}** purchase order(s) reviewed.",
            icon=":material/task_alt:",
        )

        excluded_row_count = result["dropped_row_count"] + result["skipped_row_count"]
        if excluded_row_count:
            st.warning(
                f"**{excluded_row_count} row(s) were excluded** from this report and aren't "
                "reflected in the count above - "
                f"{result['dropped_row_count']} with data that couldn't be lined up into columns, "
                f"{result['skipped_row_count']} with a receive date or amount that couldn't be read. "
                "Check the original export for these rows before relying on this report as complete.",
                icon=":material/warning:",
            )

        within_tab, outside_tab = st.tabs([config.WITHIN_SHEET, config.OUTSIDE_SHEET])
        with within_tab:
            st.caption(f"{result['counts'][config.WITHIN_SHEET]} purchase order(s) - received the same month they were confirmed.")
            st.dataframe(
                _rows_to_records(result["header"], result["preview_sheets"][config.WITHIN_SHEET]),
                width="stretch",
            )
        with outside_tab:
            st.caption(f"{result['counts'][config.OUTSIDE_SHEET]} purchase order(s) - received a different month than they were confirmed.")
            st.dataframe(
                _rows_to_records(result["header"], result["preview_sheets"][config.OUTSIDE_SHEET]),
                width="stretch",
            )

        st.download_button(
            "Download report (.xlsx)",
            data=result["xlsx_bytes"],
            file_name="po_reporting_periods.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
            help="Downloads the full, formatted report - not just the preview rows shown above.",
        )
