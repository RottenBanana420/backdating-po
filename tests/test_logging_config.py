import logging

from src.logging_config import configure


def test_configure_calls_basicconfig_with_message_only_format(monkeypatch):
    calls = []
    monkeypatch.setattr(logging, "basicConfig", lambda **kwargs: calls.append(kwargs))

    configure(level=logging.DEBUG)

    assert calls == [{"level": logging.DEBUG, "format": "%(message)s"}]


def test_configure_defaults_to_info_level(monkeypatch):
    calls = []
    monkeypatch.setattr(logging, "basicConfig", lambda **kwargs: calls.append(kwargs))

    configure()

    assert calls == [{"level": logging.INFO, "format": "%(message)s"}]
