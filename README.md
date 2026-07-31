# backdating-po

[![CI](https://github.com/RottenBanana420/backdating-po/actions/workflows/ci.yml/badge.svg)](https://github.com/RottenBanana420/backdating-po/actions/workflows/ci.yml)

Builds a styled Excel report that flags purchase orders (POs) received in a
different calendar month than the transaction-log-confirmation (TLC) date
on file — the discrepancy an auditor cares about when checking for
backdated POs.

## Features

- Repairs raw CSV export rows corrupted by stray CR/LF characters embedded
  in a single field (a defect in the upstream export, not this tool).
- Splits POs into two sheets — **within** vs. **outside** their reporting
  month — based on the TLC date vs. the receive date, dropping exact
  same-date rows (nothing to flag there).
- Produces a single `.xlsx` with per-month row shading, highlighted
  TLC/receive-date columns, and a working **Slicer** on `reporting_month`
  (openpyxl can't write slicers itself — see [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)
  for how this tool does it anyway).
- Runs end-to-end in memory: one CSV in, one `.xlsx` out, no intermediate
  files written to disk.

## Architecture

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for the full pipeline breakdown,
design decisions, and data flow. In short:

```
CSV file --> csv_loader --> transform --> workbook_builder --> workbook_writer --> .xlsx file
```

## Folder structure

```
backdating-po/
├── main.py                       # entry point: launches the web UI by
│                                  # default, or runs the CLI pipeline
│                                  # directly when given --input/--output
├── pyproject.toml                # package metadata + dependencies
├── requirements.txt               # Streamlit Community Cloud install shim (see Deployment)
├── LICENSE                       # Apache 2.0
├── .github/workflows/ci.yml      # CI: pytest on Python 3.10-3.12
├── src/
│   ├── app.py                    # Web UI (Streamlit), launched by main.py
│   ├── config.py                 # paths, column lists, sheet names
│   ├── csv_loader.py             # raw CSV read + row repair
│   ├── transform.py              # column trimming, date logic, period split
│   ├── pipeline.py               # orchestrates the stages below
│   ├── uploads.py                # bridges Web UI uploads to pipeline.run()
│   ├── logging_config.py         # console logging setup
│   └── excel/
│       ├── styles.py             # openpyxl fonts/fills/borders
│       ├── workbook_builder.py   # styled worksheet construction
│       ├── workbook_writer.py    # save + XML patching (ignoredErrors)
│       └── excel_slicers.py      # hand-authored OOXML Slicer injection
├── data/
│   ├── raw/                      # put your real raw_data.csv here (gitignored)
│   ├── sample/                   # synthetic demo dataset (tracked)
│   └── output/                   # generated report lands here (gitignored)
├── docs/
│   ├── ARCHITECTURE.md
│   ├── CONTRIBUTING.md
│   └── CHANGELOG.md
└── tests/                        # pytest suite
```

## Requirements

- Python 3.10+

## Installation

```bash
python3 -m venv venv
source venv/bin/activate       # Windows: venv\Scripts\activate
pip install -e ".[dev,app]"
```

The `app` extra installs Streamlit, needed for the [Web UI](#web-ui) that
`main.py` launches by default. Omit it (`pip install -e ".[dev]"`) if you
only ever intend to use [CLI mode](#cli-mode).

## Configuration

The set of columns kept in the report, the two sheet names, and the CSV
encoding (`cp1252`, matching the source export) are fixed business rules —
see `src/config.py`. Input/output file locations are the one
thing meant to vary, and are overridable via CLI flags (below) rather than
a config file, since there's nothing else to configure.

## Running locally

Install the `app` extra so the web UI is available, then run `main.py` with
no arguments — this is the default way to use the tool:

```bash
pip install -e ".[dev,app]"
python main.py
```

This launches the [Web UI](#web-ui) at `http://localhost:8501`.

### CLI mode

Passing `--input`/`--output` runs the pipeline directly instead, no browser
involved — useful for scripting or automation:

```bash
python main.py --input data/raw/raw_data.csv --output data/output/po_reporting_periods.xlsx
```

No real data on hand? Run the pipeline against the tracked synthetic sample
dataset instead:

```bash
python main.py --input data/sample/sample_raw_data.csv --output data/output/demo.xlsx
```

## Web UI

A Streamlit front-end (`src/app.py`, launched by `main.py`) wraps the same
pipeline used by CLI mode — see [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)
for how the underlying load → transform → split → build → save stages work;
the UI adds nothing to that pipeline, it only bridges an uploaded CSV's
in-memory bytes to it (see `src/uploads.py`).

1. **Upload** a raw PO-receipt `.csv` — drag-and-drop onto the upload box,
   or click it to browse. A preview of the first rows is shown immediately,
   exactly as uploaded.
2. Click **Run pipeline**. The app re-runs the same load/transform/split
   logic `main.py` uses, then previews both the "within" and "outside"
   reporting-period buckets.
3. Click **Download report (.xlsx)** to get the same styled, multi-sheet
   workbook `main.py` produces (row shading, highlighted TLC/receive-date
   columns, working `reporting_month` Slicer).

No real data on hand? Try `data/sample/sample_raw_data.csv`.

**Note:** the download is `.xlsx` only, not CSV. The processed output is
inherently a styled, two-sheet workbook with a Slicer — a flat CSV would
lose the within/outside split and all formatting that make the report
useful to an auditor, so `.xlsx` is the one, obvious download rather than
adding an unrequested CSV export alongside it.

## Deployment

The Web UI is deployed on [Streamlit Community Cloud](https://share.streamlit.io):

- **App URL:** https://backdating-po-lbkihwaovzjpqlrm8szlbe.streamlit.app
- **Source repo:** private (`RottenBanana420/backdating-po`) — the app's own
  sharing setting is public, so coworkers can open and use the URL directly
  without GitHub access or a Streamlit account. Repo privacy and app
  sharing are independent settings; either can be changed later (Settings →
  Sharing on the app, or the repo's visibility on GitHub) without breaking
  the deployment.
- **Deploy config:** repo `RottenBanana420/backdating-po`, branch `main`,
  main file path `src/app.py` (not `main.py` — Community Cloud runs the
  Streamlit script directly, and `main.py` only launches Streamlit via a
  `subprocess` call that doesn't work in that sandbox).
- **Redeploys automatically** on every push to `main`. No manual redeploy
  step needed.
- **Managing the app:** click **Manage app** (bottom-right corner while
  viewing it) for live logs, manual reboot, or resource usage — useful if
  it errors or you need to debug a deploy.

### `requirements.txt`

Re-added specifically for this deployment, after being deliberately removed
in [0.2.0](docs/CHANGELOG.md#020---2026-07-30) in favor of `pyproject.toml`
as the single source of dependency truth. `pyproject.toml` is still that
source of truth — `requirements.txt` is a one-line shim, not a duplicate
dependency list:

```
-e .[app]
```

Community Cloud only runs `pip install -r requirements.txt`; it does not
install the project itself from `pyproject.toml`. Without this file (or
with a plain `streamlit`/`openpyxl` list in it, which was the first attempt),
`streamlit` and `openpyxl` would be present but this repo's own `src`
package would not be — `src/app.py`'s `from src import config` then fails
with `ModuleNotFoundError` on Cloud, even though it works locally (where
`pip install -e ".[dev,app]"` already registered `src` as importable).
The editable install above installs the project (registering `src`) and its
`app` extra (`streamlit`) in one step. If `pyproject.toml`'s dependencies or
`app` extra ever change, no edit to `requirements.txt` is needed — a push
is enough to trigger Cloud's reinstall.

### Private-repo deploy gotcha

Deploying from a private repo needs GitHub's `repo` OAuth scope, which
Streamlit's default "Sign in with GitHub" link does **not** request (it only
asks for public-repo access). If the deploy form shows "This repository
does not exist" for a private repo you own, the fix is to find Streamlit's
separate private-repo authorization link (near the Repository field on the
deploy form, or under the workspace's GitHub connection settings) and accept
the GitHub prompt that lists **private repositories** under its Repositories
permission — not the general sign-in prompt, which only ever offers public
access no matter how many times it's re-accepted.

### Free-tier limits to expect

- Apps with no traffic for 12h go to sleep; the next visitor sees a
  "wake up" screen (~30-60s to restart). Not a bug if a coworker reports
  the app "not loading" after a quiet period.
- ~1GB memory, fractional shared CPU. Fine for normal CSV-sized PO exports;
  watch for a resource-limit error on unusually large files.
- Only **one** private-repo-backed app allowed at a time on the free tier.
- No custom domain on the free tier — the `*.streamlit.app` URL above is
  permanent unless the app is deleted/recreated.

## Testing

```bash
pytest
```

Covers CSV row repair and edge cases, date/period-split logic, `config.py`
path resolution, the styled-workbook builder (row shading, highlighted
columns, Excel Table), the writer's XML patching (`ignoredErrors`, Slicer
injection), the hand-authored Slicer XML itself, end-to-end pipeline
integration tests (including empty-result and missing-output-directory
cases) against the sample dataset, `main.py`'s UI-vs-CLI dispatch logic,
and the Web UI's upload-to-pipeline adapter (`src/uploads.py`) plus a
smoke test confirming `src/app.py` loads without error.

## Troubleshooting

- **`Error: Input file not found: <path>`** — the CSV you pointed `--input`
  at (or the default `data/raw/raw_data.csv`) doesn't exist. Either put
  your file there or point `--input` at the sample dataset to try the tool
  first.
- **`Rows still misaligned after fix: N` (N > 0) in the console output** —
  some rows in your CSV have a field count that still doesn't match the
  header after stray-character repair. These rows are logged (line number,
  field count) and dropped; they typically indicate a genuinely malformed
  export row rather than the CR/LF defect this tool repairs automatically.
- **`UnicodeDecodeError` on load** — the source CSV isn't `cp1252`-encoded.
  If your export uses a different encoding, change `ENCODING` in
  `src/config.py`.
- **`ValueError: No header row found...`** — the input CSV is empty (zero
  bytes). Point `--input` at a file that at least has a header row.
- **Excel shows an "Out of date" banner on the Slicer when you first open
  the report** — expected, one-time, harmless. Click **Update** once; see
  the docstring in `src/excel/excel_slicers.py` for why.

## Contributing

See [docs/CONTRIBUTING.md](docs/CONTRIBUTING.md).

## License

[Apache License 2.0](LICENSE).

## Changelog

See [docs/CHANGELOG.md](docs/CHANGELOG.md).

## Roadmap

- Configurable column list (currently hardcoded business rule in `config.py`).
- Optional CSV output alongside the styled `.xlsx`, for non-Excel consumers.
