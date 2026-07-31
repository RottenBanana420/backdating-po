# Changelog

All notable changes to this project are documented in this file, in the
style of [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [0.2.0] - 2026-07-30

Full repository overhaul: package restructure, tests, documentation, and
public-repo readiness.

### Added

- `src/backdating_po/` installable package, replacing the loose
  `scripts/build_report.py` monolith. Split into `config.py`, `styles.py`,
  `csv_loader.py`, `transform.py`, `workbook_builder.py`,
  `workbook_writer.py`, `excel_slicers.py` (moved unchanged), `pipeline.py`,
  and `logging_config.py` — one responsibility per module (see
  [ARCHITECTURE.md](ARCHITECTURE.md)).
- `pyproject.toml`, replacing `requirements.txt` as the single source of
  dependency truth. Install via `pip install -e ".[dev]"`.
- pytest suite (`tests/`): CSV row-repair coverage, date/period-split
  coverage, and an end-to-end pipeline integration test.
- `data/sample/sample_raw_data.csv` — a tracked, fully synthetic dataset
  for demos and tests, so the tool is runnable immediately after cloning
  without needing real PO data.
- `--input`/`--output` CLI flags on `main.py`.
- README, ARCHITECTURE, CONTRIBUTING, and this CHANGELOG.
- Apache 2.0 `LICENSE`.
- A purpose-built `.gitignore` covering real/generated data directories,
  Python artifacts, test/coverage output, editors, and OS files.

### Changed

- `main.py` now calls `backdating_po.pipeline.run()` directly instead of
  subprocessing `scripts/build_report.py` — same console output and exit
  codes, one less process hop, and it's what makes the pipeline
  importable/testable.
- Console output now goes through `logging` (`logging_config.py`) instead
  of bare `print()` calls.
- Missing-input-file errors now raise a clear `FileNotFoundError` with a
  pointer to the README, instead of a raw traceback from deep inside
  `csv_loader`.

### Removed

- `scripts/` directory (superseded by `src/backdating_po/`).
- `requirements.txt` (superseded by `pyproject.toml`).
- `reference/Backdating POs 7.1.25-7.1.26.xlsx` — a real Excel-authored
  workbook used during original development to reverse-engineer the
  Slicer XML schema. Contained real business data and wasn't needed at
  runtime; removed from the repo (see Security below).

### Security

- **Git history rewritten to remove real business data.** Earlier commits
  had added `data/raw/raw_data.csv`, `data/processed/raw_data_fixed.csv`,
  `data/output/po_reporting_periods.xlsx`,
  `data/output/raw_data_fixed.xlsx`, and the reference workbook above —
  all containing real vendor, employee, and PO financial data. History was
  rewritten with `git filter-repo` to strip these paths from every commit
  before this repository was made public. **Anyone with an existing local
  clone must re-clone**; the old commit hashes no longer exist upstream.
  Real data directories (`data/raw/`, `data/output/`, `reference/`) are
  now gitignored so this can't recur silently.

## [0.1.0] - 2026-07-30 (pre-refactor baseline)

The original four-stage script pipeline, for historical context:

- `scripts/01_fix_encoding_issues.py` — repaired CR/LF-torn CSV rows,
  writing a fixed CSV to disk.
- `scripts/02_build_report.py` — trimmed columns, derived
  `reporting_month`.
- `scripts/03_split_by_reporting_period.py` — bucketed rows into
  within/outside sheets.
- `scripts/04_format_report.py` — applied styling and wrote the final
  `.xlsx`.
- `main.py` ran all four stages via subprocess, each reading the previous
  stage's output file from disk.

This was later consolidated into a single `scripts/build_report.py`
(same logic, no intermediate files) immediately before the 0.2.0 refactor.
