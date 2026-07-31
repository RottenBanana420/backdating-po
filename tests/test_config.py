import importlib
from pathlib import Path

from src import config as config_module

REPO_ROOT = Path(__file__).resolve().parent.parent


def test_root_is_the_actual_repo_root():
    assert config_module.ROOT == REPO_ROOT
    assert (config_module.ROOT / "pyproject.toml").is_file()
    assert (config_module.ROOT / "src").is_dir()


def test_root_unaffected_by_working_directory(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    reloaded = importlib.reload(config_module)
    assert reloaded.ROOT == REPO_ROOT


def test_raw_data_and_output_paths_are_under_repo_root():
    assert config_module.RAW_DATA_PATH == REPO_ROOT / "data" / "raw" / "raw_data.csv"
    assert config_module.OUTPUT_PATH == REPO_ROOT / "data" / "output" / "po_reporting_periods.xlsx"
