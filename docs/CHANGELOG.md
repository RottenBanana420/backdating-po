# Changelog

All notable changes to this project are documented in this file, in the
style of [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [Unreleased]

### Added

- `.github/workflows/ci.yml`: runs `pytest` on Python 3.10/3.11/3.12 for
  every push/PR to `main`, with pip caching, a read-only `contents`
  permission, and `concurrency` cancellation of superseded runs.
- Unit test coverage for the previously-untested `src/excel/` subpackage
  (`workbook_builder.py`, `workbook_writer.py`, `excel_slicers.py`),
  `config.py`, and `logging_config.py`, plus edge-case coverage for
  `csv_loader.py`, `transform.py`, and `pipeline.py` (empty CSVs,
  unrecoverable rows, month-boundary dates, zero-row output, missing
  output directories). Suite grew from 6 to 37 tests.
- CI status badge in `README.md`.

### Changed

- `docs/CONTRIBUTING.md`'s testing-expectations section now covers all of
  `src/`, not just `csv_loader.py`/`transform.py`.

### Fixed

- `csv_loader.load_and_clean_rows` raised an unhandled `TypeError` on a
  genuinely zero-byte input CSV (`len(None)`, since no header row was ever
  found) instead of a clear error. Now raises `ValueError` with a message
  naming the empty file.

## [0.2.1] - 2026-07-30

### Changed

- Flattened the package layout: `src/backdating_po/` → `src/`, so `src/`
  itself is the importable package (`import src`) rather than wrapping a
  separate `backdating_po/` package directory. Trade-off: the installable
  distribution name (`backdating-po`, used for `pip install`) and the
  import name (`src`) now differ, which is unconventional for a package
  meant for wider distribution — acceptable here since this tool is not
  published to PyPI.
- Grouped the four openpyxl/Excel-output-specific modules (`styles.py`,
  `workbook_builder.py`, `workbook_writer.py`, `excel_slicers.py`) into an
  `excel/` subpackage; `config.py`, `csv_loader.py`, `transform.py`,
  `pipeline.py`, and `logging_config.py` stay at the top level of `src/`.
- Fixed `config.ROOT`, which had one extra `.parent` hop left over from
  the previous `src/backdating_po/config.py` depth — after the move to
  `src/config.py` (one level shallower), the stale calculation resolved
  the repo root one directory too high, breaking the default
  `data/raw/raw_data.csv` / `data/output/...` paths.

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
  workbook kept locally during original development to reverse-engineer
  the Slicer XML schema. Contained real business data; deleted from the
  local working tree once no longer needed (see Security below).

### Security

- **Audited full git history for real business data; none found.** Real
  files (`data/raw/raw_data.csv`, the reference workbook above,
  `data/output/*.xlsx`) were checked against every commit across the
  project's entire history (`git log --all --diff-filter=A --name-only`
  and a full tree walk over `git rev-list --all`). None were ever
  committed — the `.gitignore` from the very first commit already kept
  build/dependency artifacts out, and these data paths were simply never
  `git add`ed. No history rewrite was necessary. Real data directories
  (`data/raw/`, `data/output/`, `reference/`) are now explicitly
  gitignored (task 11) so this can't happen by accident going forward.

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
