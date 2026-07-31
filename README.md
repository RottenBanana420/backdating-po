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
├── main.py                       # CLI entry point
├── pyproject.toml                # package metadata + dependencies
├── LICENSE                       # Apache 2.0
├── .github/workflows/ci.yml      # CI: pytest on Python 3.10-3.12
├── src/
│   ├── config.py                 # paths, column lists, sheet names
│   ├── csv_loader.py             # raw CSV read + row repair
│   ├── transform.py              # column trimming, date logic, period split
│   ├── pipeline.py               # orchestrates the stages below
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
pip install -e ".[dev]"
```

## Configuration

The set of columns kept in the report, the two sheet names, and the CSV
encoding (`cp1252`, matching the source export) are fixed business rules —
see `src/config.py`. Input/output file locations are the one
thing meant to vary, and are overridable via CLI flags (below) rather than
a config file, since there's nothing else to configure.

## Running locally

Drop your raw export at `data/raw/raw_data.csv`, then:

```bash
python main.py
```

This writes `data/output/po_reporting_periods.xlsx`.

No real data on hand? Run the pipeline against the tracked synthetic sample
dataset instead:

```bash
python main.py --input data/sample/sample_raw_data.csv --output data/output/demo.xlsx
```

Full usage:

```bash
python main.py --input path/to/your.csv --output path/to/report.xlsx
```

## Testing

```bash
pytest
```

Covers CSV row repair and edge cases, date/period-split logic, `config.py`
path resolution, the styled-workbook builder (row shading, highlighted
columns, Excel Table), the writer's XML patching (`ignoredErrors`, Slicer
injection), the hand-authored Slicer XML itself, and end-to-end pipeline
integration tests (including empty-result and missing-output-directory
cases) against the sample dataset.

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
