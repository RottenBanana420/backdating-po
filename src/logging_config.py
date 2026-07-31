"""Central logging setup for the pipeline. Call configure() once, at the
CLI entry point, before any module-level logger emits output.
"""
import logging


def configure(level: int = logging.INFO) -> None:
    logging.basicConfig(level=level, format="%(message)s")
