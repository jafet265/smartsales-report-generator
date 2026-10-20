"""SmartSales Stage 5: run validation, reports, charts, and Stage 6 insights.

Usage:
    python smartsales_stage5.py
    python smartsales_stage5.py path/to/sales_data.csv

Keep Stage 3, Stage 4, and Stage 6 scripts beside this script.
"""

import argparse
import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_INPUT = PROJECT_ROOT / "../data" / "raw" / "sales_data_1500_with_invalid.csv"
STAGE3_SCRIPT = PROJECT_ROOT / "smartsales_stage3.py"
STAGE4_SCRIPT = PROJECT_ROOT / "smartsales_stage4.py"
STAGE6_SCRIPT = PROJECT_ROOT / "smartsales_stage6.py"


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
    parser = argparse.ArgumentParser(
        description="Generate the SmartSales Excel report, charts, and insights."
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

    required_scripts = [
        ("Stage 3", STAGE3_SCRIPT),
        ("Stage 4", STAGE4_SCRIPT),
        ("Stage 6", STAGE6_SCRIPT),
    ]
    for stage_name, script_path in required_scripts:
        if not script_path.is_file():
            print(f"Error: {stage_name} script not found: {script_path}")
            print(f"Keep {script_path.name} beside smartsales_stage5.py.")
            return 1

    print("SmartSales: Automated Sales Analysis Pipeline")
    print(f"Input CSV: {input_file}")

    stages = [
        ("Stage 3 — Validate data and create Excel report", STAGE3_SCRIPT),
        ("Stage 4 — Generate sales charts", STAGE4_SCRIPT),
        ("Stage 6 — Generate automated business insights", STAGE6_SCRIPT),
    ]
    for stage_name, script_path in stages:
        if not run_stage(stage_name, script_path, input_file):
            return 1

    reports_folder = PROJECT_ROOT / "reports"
    charts_folder = reports_folder / "charts"
    print(f"\n{'=' * 60}")
    print("SMARTSALES PIPELINE COMPLETED SUCCESSFULLY")
    print("=" * 60)
    print(f"Excel report: {reports_folder / 'smartsales_report.xlsx'}")
    print(f"Insights report: {reports_folder / 'sales_insights.txt'}")
    print(f"Charts folder: {charts_folder}")
    print(f"  - {charts_folder / 'revenue_by_product.png'}")
    print(f"  - {charts_folder / 'revenue_by_category.png'}")
    print(f"  - {charts_folder / 'monthly_revenue.png'}")
    print("\nOnly valid sales records are included in sales calculations.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
