"""SmartSales Stage 5: run the complete reporting workflow with one command.

Usage:
    python smartsales_stage5.py
    python smartsales_stage5.py path/to/sales_data.csv

This script coordinates Stage 3 (validation + Excel report) and
Stage 4 (sales charts). Keep both stage scripts in this project's root.
"""

import argparse
import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_INPUT = PROJECT_ROOT / "../data" / "raw" / "sales_data_1500_with_invalid.csv"
STAGE3_SCRIPT = PROJECT_ROOT / "smartsales_stage3.py"
STAGE4_SCRIPT = PROJECT_ROOT / "smartsales_stage4.py"


def run_stage(stage_name, script_path, input_file):
    """Run one stage and stop the workflow if it fails."""
    print(f"\n{'=' * 60}")
    print(stage_name)
    print("=" * 60)

    result = subprocess.run(
        [sys.executable, str(script_path), str(input_file)],
        cwd=PROJECT_ROOT,
        check=False,
    )

    if result.returncode != 0:
        print(f"\nWorkflow stopped: {stage_name} failed.")
        return False

    return True


def main():
    """Run validation, Excel reporting, and chart generation in order."""
    parser = argparse.ArgumentParser(
        description="Generate the SmartSales Excel report and sales charts."
    )
    parser.add_argument(
        "csv_file",
        nargs="?",
        help="Optional path to a sales CSV file. Defaults to the project's sample data.",
    )
    args = parser.parse_args()

    if args.csv_file:
        input_file = Path(args.csv_file).expanduser()
        if not input_file.is_absolute():
            input_file = (Path.cwd() / input_file).resolve()
    else:
        input_file = DEFAULT_INPUT

    if not input_file.is_file():
        print(f"Error: CSV file not found: {input_file}")
        return 1

    if not STAGE3_SCRIPT.is_file():
        print(f"Error: Stage 3 script not found: {STAGE3_SCRIPT}")
        print("Keep smartsales_stage3.py beside smartsales_stage5.py.")
        return 1

    if not STAGE4_SCRIPT.is_file():
        print(f"Error: Stage 4 script not found: {STAGE4_SCRIPT}")
        print("Keep smartsales_stage4.py beside smartsales_stage5.py.")
        return 1

    print("SmartSales Stage 5: Automated Reporting Workflow")
    print(f"Input CSV: {input_file}")

    if not run_stage("Stage 3 — Validate data and create Excel report",
                     STAGE3_SCRIPT, input_file):
        return 1

    if not run_stage("Stage 4 — Generate sales charts",
                     STAGE4_SCRIPT, input_file):
        return 1

    report_path = PROJECT_ROOT / "reports" / "smartsales_report.xlsx"
    charts_folder = PROJECT_ROOT / "reports" / "charts"

    print(f"\n{'=' * 60}")
    print("WORKFLOW COMPLETED SUCCESSFULLY")
    print("=" * 60)
    print(f"Excel report: {report_path}")
    print(f"Charts folder: {charts_folder}")
    print("Charts:")
    print(f"  - {charts_folder / 'revenue_by_product.png'}")
    print(f"  - {charts_folder / 'revenue_by_category.png'}")
    print(f"  - {charts_folder / 'monthly_revenue.png'}")
    print("\nOnly valid sales records are included in the calculations.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
