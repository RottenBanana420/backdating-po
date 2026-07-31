# Architecture

## Overview

This is a single-purpose, single-pass ETL pipeline: one raw CSV in, one
styled `.xlsx` report out. There's no server, no database, no persistent
state between runs — every run reads the source CSV fresh and writes a
complete report. That simplicity is intentional; the module split below
exists to make each transformation step independently readable and
testable, not to support a larger system that doesn't exist.

## Design principles

- **One responsibility per module.** The original `build_report.py`
  mixed CSV repair, business transforms, Excel styling, Excel Table
  creation, and raw zip/XML patching in one 400-line file. Each of those
  is now its own module (see below), so a change to, say, row-shading
  colors touches `workbook_builder.py` only.
- **In-memory, single-pass.** Data flows through the pipeline as Python
  lists/dicts held in memory; nothing is written to disk until the final
  `.xlsx`. No intermediate CSV/XLSX files to keep in sync or clean up.
- **Preserve the escape hatches the original author already earned.** The
  CSV-repair logic and the hand-authored Slicer XML both exist because of
  specific, non-obvious real-world constraints (see below) — they're kept
  verbatim, not "simplified" away.

## Module responsibilities

```
src/
├── config.py            Business constants: input columns, output paths,
│                         CSV encoding, sheet names. Plain data, no logic.
├── csv_loader.py         Reads the raw CSV and repairs rows torn by stray
│                         CR/LF characters embedded in a field.
├── transform.py          Column trimming, reporting_month derivation, and
│                         the tlc-vs-receive_date within/outside/dropped
│                         split — the core business rule this tool exists
│                         to apply.
├── pipeline.py           Orchestrates the stages below in order.
├── logging_config.py     One-line logging.basicConfig() wrapper.
└── excel/                Everything openpyxl/Excel-output-specific, grouped
    │                     because these four modules change together and
    │                     share no callers outside this subpackage.
    ├── styles.py          openpyxl Font/Fill/Border/Alignment objects used
    │                     by workbook_builder. Split from config.py because
    │                     these are openpyxl-typed objects, not plain config.
    ├── workbook_builder.py   Builds the styled openpyxl Workbook: sheet per
    │                     bucket, month-based row shading, highlighted
    │                     tlc/receive_date columns, an Excel Table (backing
    │                     the Slicer added later).
    ├── workbook_writer.py    Saves the workbook, then patches in two things
    │                     openpyxl can't write itself (see below).
    └── excel_slicers.py      Hand-authored OOXML for a real Excel Slicer.
```

## Dependency flow

```
config.py ─┬─> csv_loader.py ─┐
           ├─> transform.py ──┤
           └─> excel/workbook_builder.py <── excel/styles.py
                                │
                                v
                excel/workbook_writer.py <── excel/excel_slicers.py
                                │
                                v
                          pipeline.py  (orchestrates all of the above)
                                │
                                v
                            main.py  (CLI: argument parsing, error handling)
```

Every module below `pipeline.py` is a pure function of its inputs (no
shared mutable state, no module reads another's output off disk) —
`pipeline.run()` is the only place that wires them together in sequence.

## Data flow

1. **Raw CSV row** (full source export schema, `cp1252`-encoded, `\r\n`-delimited) →
   `csv_loader.load_and_clean_rows` repairs any row torn by a stray CR/LF
   embedded in a field, returning `(header, rows)`. Raises `ValueError` if
   the file has no header row at all (i.e. is empty).
2. **Trimmed + dated row** → `transform.build_trimmed_rows` keeps only the
   12 business-relevant columns and prepends a `reporting_month` column
   derived from `receive_date` (`"%B %Y"`, e.g. `"January 2026"`).
3. **Bucketed sheet** → `transform.split_by_reporting_period` compares
   `tlc` and `receive_date` per row: same month → "within", different
   month → "outside", identical date → dropped (nothing to flag).
4. **Styled worksheet** → `workbook_builder.build_workbook` sorts each
   bucket by `reporting_month`, shades rows, highlights the two date
   columns, and attaches an Excel `Table` (required as the Slicer's data
   source).
5. **Patched `.xlsx` bytes** → `workbook_writer.save_workbook` saves via
   openpyxl, then re-opens the `.xlsx` as a zip and patches in XML openpyxl
   doesn't serialize: the "Number Stored as Text" `ignoredErrors` block
   (so Excel doesn't flag `po`/`invoiceid` as numbers-stored-as-text) and
   the `reporting_month` Slicer (via `excel_slicers.py`).

## Key design decisions

**Why hand-authored Slicer XML (`excel_slicers.py`)?** openpyxl's
maintainers have explicitly declined to add write support for Slicers —
they're an OOXML extension Microsoft added after ECMA-376 was ratified,
and openpyxl strips them from any file it re-saves
([issue #17](https://github.com/chronossc/openpyxl/issues/17),
[issue #30](https://github.com/chronossc/openpyxl/issues/30)). This module
builds the Slicer/SlicerCache/Drawing XML parts by hand, patching them
into the `.xlsx` zip after openpyxl saves — modeled on a real,
Excel-authored workbook containing working slicers (used during
development to reverse-engineer the exact XML schema, then deleted
locally — never committed to the repo — since it contained real business
data; see the audit note in [CHANGELOG.md](CHANGELOG.md)). Because the resulting slicer cache is
written directly rather than computed by a live Excel session, Excel shows
a one-time, harmless "Out of date" banner the first time the file opens;
clicking **Update** resolves it permanently.

**Why in-memory, single-pass, no intermediate files?** The pipeline
previously ran as four separate scripts, each reading the previous
script's output file from disk. That meant every schema change (e.g.
adding a column) had to be threaded through four file formats instead of
four function signatures, and left stale intermediate files on disk
between runs. Collapsing it to one in-memory pass removes both problems;
the module split in this refactor keeps the "one file, one concern"
readability of the four-script version without the disk I/O.

**Why does `main.py` call `pipeline.run()` directly instead of
subprocessing a script?** The original `main.py` used
`subprocess.run([sys.executable, "scripts/build_report.py"])` — a process
hop with no benefit once the pipeline is an importable package. A direct
function call is simpler, preserves the same console output and exit
codes, and is what makes the pipeline testable via `pipeline.run()` in
`tests/test_pipeline_integration.py`.
