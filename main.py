"""
Main entry point for the PO reporting pipeline.

Runs scripts/build_report.py, which reads data/raw/raw_data.csv and
writes the final report straight to data/output/po_reporting_periods.xlsx
in a single pass, with no intermediate files written to disk.
"""
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SCRIPT = ROOT / "scripts" / "build_report.py"


def main():
    result = subprocess.run([sys.executable, str(SCRIPT)])
    if result.returncode != 0:
        sys.exit(result.returncode)

    print("\nPipeline completed successfully.")


if __name__ == "__main__":
    main()
