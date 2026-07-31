import logging
import time

import pytest

from src.logging_config import LOG_DATEFMT, LOG_FORMAT, configure, log_stage


def test_configure_calls_basicconfig_with_timestamped_leveled_format(monkeypatch):
    calls = []
    monkeypatch.setattr(logging, "basicConfig", lambda **kwargs: calls.append(kwargs))

    configure(level=logging.DEBUG)

    assert calls == [{"level": logging.DEBUG, "format": LOG_FORMAT, "datefmt": LOG_DATEFMT}]


def test_configure_defaults_to_info_level(monkeypatch):
    monkeypatch.delenv("LOG_LEVEL", raising=False)
    calls = []
    monkeypatch.setattr(logging, "basicConfig", lambda **kwargs: calls.append(kwargs))

    configure()

    assert calls == [{"level": logging.INFO, "format": LOG_FORMAT, "datefmt": LOG_DATEFMT}]


def test_configure_reads_level_from_log_level_env_var(monkeypatch):
    monkeypatch.setenv("LOG_LEVEL", "DEBUG")
    calls = []
    monkeypatch.setattr(logging, "basicConfig", lambda **kwargs: calls.append(kwargs))

    configure()

    assert calls == [{"level": logging.DEBUG, "format": LOG_FORMAT, "datefmt": LOG_DATEFMT}]


def test_configure_falls_back_to_info_for_invalid_log_level_env_var(monkeypatch):
    monkeypatch.setenv("LOG_LEVEL", "NOT_A_LEVEL")
    calls = []
    monkeypatch.setattr(logging, "basicConfig", lambda **kwargs: calls.append(kwargs))

    configure()

    assert calls == [{"level": logging.INFO, "format": LOG_FORMAT, "datefmt": LOG_DATEFMT}]


def test_log_stage_logs_start_and_finish(caplog):
    logger = logging.getLogger("test_log_stage")

    with caplog.at_level(logging.INFO, logger="test_log_stage"):
        with log_stage(logger, "example stage"):
            pass

    messages = [r.message for r in caplog.records]
    assert any(m.startswith("Starting: example stage") for m in messages)
    assert any(m.startswith("Finished: example stage") for m in messages)


def test_log_stage_logs_warning_and_reraises_on_failure(caplog):
    logger = logging.getLogger("test_log_stage_failure")

    with caplog.at_level(logging.INFO, logger="test_log_stage_failure"):
        with pytest.raises(ValueError, match="boom"):
            with log_stage(logger, "example stage"):
                raise ValueError("boom")

    warnings = [r for r in caplog.records if r.levelno == logging.WARNING]
    assert len(warnings) == 1
    assert "Failed: example stage" in warnings[0].message
    assert "ValueError: boom" in warnings[0].message
    # log_stage must not log the full traceback itself - that happens once,
    # at the top-level caller - so this record has no exc_info attached.
    assert warnings[0].exc_info is None


def test_log_stage_measures_elapsed_time(caplog):
    logger = logging.getLogger("test_log_stage_timing")

    with caplog.at_level(logging.INFO, logger="test_log_stage_timing"):
        with log_stage(logger, "slow stage"):
            time.sleep(0.01)

    finished = next(r.message for r in caplog.records if r.message.startswith("Finished"))
    assert "(0." in finished or "(1." in finished
