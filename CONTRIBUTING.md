# Contributing

## Coding standards

- Follow PEP 8. Four-space indentation, no tabs.
- Add type hints where they clarify a function's contract (e.g. return
  types on public functions in `src/backdating_po/`); don't add them
  everywhere reflexively — a one-line helper with an obvious signature
  doesn't need them.
- Docstrings and comments explain **why**, not what — the code already
  says what it does. Reserve comments for non-obvious constraints (see
  `csv_loader.py` and `excel_slicers.py` for examples worth matching).
- One responsibility per module (see [ARCHITECTURE.md](ARCHITECTURE.md)).
  If a change doesn't fit cleanly into an existing module, that's a signal
  to add a new one rather than growing an existing file past its concern.
- No new runtime dependencies without a good reason — this project
  currently depends on `openpyxl` alone. `pytest` is dev-only.

## Branch strategy

- Feature branches off `main`, named descriptively (e.g.
  `fix-csv-encoding-detection`, `add-column-config`).
- No direct commits to `main` beyond trivial fixes (typos, formatting).
- Rebase or merge `main` into your branch before opening a PR if it's
  fallen behind.

## Commit message conventions

- Imperative mood, present tense: "Add X", "Fix Y", not "Added X" or
  "Fixes Y".
- One logical change per commit — matches the granularity already used in
  this repo's history (e.g. one commit per extracted module during the
  package refactor).
- No need for a conventional-commits prefix (`feat:`, `fix:`); a plain,
  specific imperative subject line is sufficient for a project this size.

## Pull request workflow

1. Describe what changed and why (not just what — the diff already shows
   what).
2. Link any related issue.
3. Ensure `pytest` passes locally before opening the PR.
4. Keep PRs scoped to one change; split unrelated changes into separate
   PRs.

## Testing expectations

- New or changed logic in `csv_loader.py` or `transform.py` needs unit
  test coverage (see `tests/test_csv_loader.py` and
  `tests/test_transform.py` for the existing style: small, synthetic
  fixtures, one behavior per test).
- Changes to `pipeline.py`'s stage wiring need
  `tests/test_pipeline_integration.py` updated to match.
- Never commit real business data as test fixtures — use synthetic values
  (see `data/sample/sample_raw_data.csv` and `tests/fixtures/`).
- Run the full suite before pushing:

  ```bash
  pytest
  ```
