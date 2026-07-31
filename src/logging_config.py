"""Central logging setup for the pipeline. Call configure() once, at the
CLI entry point, before any module-level logger emits output.
"""
import logging
import os
import time
from contextlib import contextmanager

# Timestamp + level + module name so a terminal log line is enough on its
# own to say when something happened, how serious it was, and where it came
# from - without that, tracing a failure back to a stage means re-running
# with print statements.
LOG_FORMAT = "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s"
LOG_DATEFMT = "%Y-%m-%d %H:%M:%S"

_UNSET = object()


def configure(level: int = _UNSET) -> None:
    """Configures root logging. If `level` isn't passed explicitly, it's
    read from the LOG_LEVEL env var (default INFO) so DEBUG-level detail
    can be turned on for a single run without a code change, e.g.
    `LOG_LEVEL=DEBUG python main.py --input ...`.
    """
    if level is _UNSET:
        level = logging.getLevelName(os.environ.get("LOG_LEVEL", "INFO").upper())
        if not isinstance(level, int):
            level = logging.INFO
    logging.basicConfig(level=level, format=LOG_FORMAT, datefmt=LOG_DATEFMT)


@contextmanager
def log_stage(logger: logging.Logger, name: str):
    """Times a pipeline stage and logs its start/finish, so call sites
    don't hand-roll timing + before/after log lines around every expensive
    step (CSV load, transform, workbook build, save).

    On failure this logs a one-line WARNING with elapsed time and the
    exception's type/message, then re-raises - it deliberately does NOT
    log the full traceback itself. log_stage calls nest (e.g. pipeline.run
    wraps a stage that itself wraps a stage); logging the full traceback at
    every level would print the same stack trace multiple times. The single
    full-traceback log happens once, at the top-level caller (app.py /
    main.py) via logger.exception.
    """
    logger.info("Starting: %s", name)
    start = time.perf_counter()
    try:
        yield
    except Exception as exc:
        elapsed = time.perf_counter() - start
        logger.warning(
            "Failed: %s after %.2fs (%s: %s)", name, elapsed, type(exc).__name__, exc
        )
        raise
    else:
        elapsed = time.perf_counter() - start
        logger.info("Finished: %s (%.2fs)", name, elapsed)
