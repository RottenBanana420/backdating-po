"""
Main entry point for the PO reporting pipeline.

Runs the pipeline stages in order:
  1. scripts/01_fix_encoding_issues.py       - repairs raw_data.csv
  2. scripts/02_build_report.py              - builds the xlsx report
  3. scripts/03_split_by_reporting_period.py - splits rows by tlc vs
                                                receive_date reporting period
  4. scripts/04_format_report.py             - sorts each sheet by
                                                reporting_month and color
                                                codes it for human review
"""
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SCRIPTS_DIR = ROOT / "scripts"

PIPELINE = [
    SCRIPTS_DIR / "01_fix_encoding_issues.py",
    SCRIPTS_DIR / "02_build_report.py",
    SCRIPTS_DIR / "03_split_by_reporting_period.py",
    SCRIPTS_DIR / "04_format_report.py",
]


def main():
    for script in PIPELINE:
        print(f"\n=== Running {script.name} ===")
        result = subprocess.run([sys.executable, str(script)])
        if result.returncode != 0:
            print(f"\n{script.name} failed with exit code {result.returncode}")
            sys.exit(result.returncode)

    print("\nPipeline completed successfully.")


if __name__ == "__main__":
    main()
