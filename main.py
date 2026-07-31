"""CLI entry point for the PO reporting pipeline.

Usage:
    python main.py                                            # uses data/raw/raw_data.csv
    python main.py --input data/sample/sample_raw_data.csv     # run the demo dataset
    python main.py --input path/to.csv --output path/to.xlsx
"""
import argparse
import sys
from pathlib import Path

from backdating_po import config
from backdating_po.logging_config import configure
from backdating_po.pipeline import run


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=config.RAW_DATA_PATH, help="Path to the raw PO CSV.")
    parser.add_argument("--output", type=Path, default=config.OUTPUT_PATH, help="Path to write the report .xlsx to.")
    return parser.parse_args(argv)


def main(argv=None) -> int:
    configure()
    args = parse_args(argv)
    try:
        run(src=args.input, dst=args.output)
    except FileNotFoundError as e:
        print(f"\nError: {e}", file=sys.stderr)
        return 1
    print("\nPipeline completed successfully.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
