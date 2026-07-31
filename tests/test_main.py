from pathlib import Path
from unittest.mock import patch

import main
from src import config

FIXTURE = Path(__file__).parent.parent / "data" / "sample" / "sample_raw_data.csv"


def test_main_with_no_args_launches_ui():
    with patch("main.launch_ui", return_value=0) as mock_launch_ui:
        assert main.main([]) == 0
    mock_launch_ui.assert_called_once()


def test_main_with_flags_runs_pipeline_not_ui(tmp_path):
    dst = tmp_path / "out.xlsx"
    with patch("main.launch_ui") as mock_launch_ui:
        exit_code = main.main(["--input", str(FIXTURE), "--output", str(dst)])
    mock_launch_ui.assert_not_called()
    assert exit_code == 0
    assert dst.exists()


def test_main_with_no_output_flag_derives_filename_from_data(tmp_path, monkeypatch):
    # sample_raw_data.csv's receive_date values span January-April 2026.
    monkeypatch.setattr(config, "OUTPUT_DIR", tmp_path)

    with patch("main.launch_ui") as mock_launch_ui:
        exit_code = main.main(["--input", str(FIXTURE)])

    mock_launch_ui.assert_not_called()
    assert exit_code == 0
    assert (tmp_path / "Backdating POs 1.1.26-5.1.26.xlsx").exists()


def test_launch_ui_reports_missing_streamlit():
    with patch("main.importlib.util.find_spec", return_value=None):
        assert main.launch_ui() == 1


def test_launch_ui_runs_streamlit_module_on_app_path():
    with (
        patch("main.importlib.util.find_spec", return_value=object()),
        patch("main.subprocess.run") as mock_run,
    ):
        mock_run.return_value.returncode = 0
        assert main.launch_ui() == 0
    mock_run.assert_called_once_with(
        [main.sys.executable, "-m", "streamlit", "run", str(main.APP_PATH)]
    )
