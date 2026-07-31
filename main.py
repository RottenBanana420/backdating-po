"""Entry point for the PO reporting pipeline.

Usage:
    python main.py                                            # launches the Streamlit web UI
    python main.py --input data/sample/sample_raw_data.csv     # run the pipeline directly (no UI)
    python main.py --input path/to.csv --output path/to.xlsx
"""
import argparse
import importlib.util
import subprocess
import sys
from pathlib import Path

from src import config
from src.logging_config import configure
from src.pipeline import run

APP_PATH = Path(__file__).resolve().parent / "src" / "app.py"


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=config.RAW_DATA_PATH, help="Path to the raw PO CSV.")
    parser.add_argument("--output", type=Path, default=config.OUTPUT_PATH, help="Path to write the report .xlsx to.")
    return parser.parse_args(argv)


def launch_ui() -> int:
    if importlib.util.find_spec("streamlit") is None:
        print(
            '\nError: streamlit is not installed. Install the "app" extra:\n'
            '  pip install -e ".[app]"\n',
            file=sys.stderr,
        )
        return 1
    result = subprocess.run([sys.executable, "-m", "streamlit", "run", str(APP_PATH)])
    return result.returncode


def main(argv=None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    if not argv:
        return launch_ui()

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
