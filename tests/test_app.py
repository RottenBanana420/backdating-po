"""Smoke test for src/app.py.

st.testing.v1.AppTest has no reliable way to simulate a file_uploader
interaction, so this only asserts the script loads without raising before
any upload happens. Upload/validation/processing logic is fully covered
against src/uploads.py directly in tests/test_uploads.py.

streamlit is an optional dependency (the "app" extra) so this module skips
cleanly via importorskip rather than failing collection - a plain "pip
install -e .[dev]" checkout must still be able to run the full test suite.
"""
import pytest

pytest.importorskip("streamlit")
from streamlit.testing.v1 import AppTest  # noqa: E402


def test_app_loads_without_exception():
    at = AppTest.from_file("src/app.py")
    at.run()
    assert at.exception == []
