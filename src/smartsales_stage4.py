"""SmartSales Stage 4: create sales charts from validated CSV data.

Usage:
    python smartsales_stage4.py
    python smartsales_stage4.py path/to/sales_data.csv

This script reuses validation functions from smartsales_stage3.py.
Charts are saved in reports/charts/.
"""

import sys
from collections import defaultdict
from decimal import Decimal
from pathlib import Path

import matplotlib

# Use a non-interactive backend so the script works in Codespaces.
matplotlib.use("Agg")

import matplotlib.pyplot as plt

from smartsales_stage3 import prepare_records, read_sales_file


DEFAULT_INPUT = Path("data/raw/sales_data_1500_with_invalid.csv")
CHARTS_FOLDER = Path("reports/charts")


def calculate_revenue(valid_records):
    """Calculate revenue totals by product, category, and month."""
    product_revenue = defaultdict(Decimal)
    category_revenue = defaultdict(Decimal)
    monthly_revenue = defaultdict(Decimal)

    for record in valid_records:
        revenue = record["quantity"] * record["unit_price"]

        product_revenue[record["product"]] += revenue
        category_revenue[record["category"]] += revenue
        month = record["order_date"][:7]  # YYYY-MM sorts chronologically.
        monthly_revenue[month] += revenue

    return product_revenue, category_revenue, monthly_revenue


def save_bar_chart(totals, title, x_label, output_path):
    """Save a bar chart with the largest total shown first."""
    sorted_items = sorted(totals.items(), key=lambda item: item[1], reverse=True)
    labels = [item[0] for item in sorted_items]
    values = [float(item[1]) for item in sorted_items]

    figure, axis = plt.subplots(figsize=(10, 6))
    axis.bar(labels, values)
    axis.set_title(title)
    axis.set_xlabel(x_label)
    axis.set_ylabel("Revenue ($)")
    axis.tick_params(axis="x", labelrotation=35)
    figure.tight_layout()
    figure.savefig(output_path, dpi=150)
    plt.close(figure)


def save_monthly_chart(totals, output_path):
    """Save a line chart showing revenue changes by month."""
    months = sorted(totals)
    values = [float(totals[month]) for month in months]

    figure, axis = plt.subplots(figsize=(10, 6))
    axis.plot(months, values, marker="o")
    axis.set_title("Monthly Revenue")
    axis.set_xlabel("Month")
    axis.set_ylabel("Revenue ($)")
    axis.tick_params(axis="x", labelrotation=35)
    axis.grid(True, axis="y", alpha=0.3)
    figure.tight_layout()
    figure.savefig(output_path, dpi=150)
    plt.close(figure)


def main():
    """Load, validate, summarize, and chart the sales data."""
    input_file = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_INPUT

    rows = read_sales_file(input_file)
    if rows is None:
        return 1

    valid_records, invalid_records = prepare_records(rows)
    if not valid_records:
        print("No valid sales records were found. No charts were created.")
        return 1

    product_revenue, category_revenue, monthly_revenue = calculate_revenue(
        valid_records
    )

    CHARTS_FOLDER.mkdir(parents=True, exist_ok=True)

    product_chart = CHARTS_FOLDER / "revenue_by_product.png"
    category_chart = CHARTS_FOLDER / "revenue_by_category.png"
    monthly_chart = CHARTS_FOLDER / "monthly_revenue.png"

    save_bar_chart(
        product_revenue,
        "Revenue by Product",
        "Product",
        product_chart,
    )
    save_bar_chart(
        category_revenue,
        "Revenue by Category",
        "Category",
        category_chart,
    )
    save_monthly_chart(monthly_revenue, monthly_chart)

    total_revenue = sum(
        (record["quantity"] * record["unit_price"] for record in valid_records),
        Decimal("0.00"),
    )

    print("SmartSales Stage 4 completed successfully.")
    print(f"Input file: {input_file}")
    print(f"Total records read: {len(rows)}")
    print(f"Valid records: {len(valid_records)}")
    print(f"Invalid records excluded: {len(invalid_records)}")
    print(f"Total revenue from valid records: ${total_revenue:,.2f}")
    print("\nCharts created:")
    print(f"- {product_chart}")
    print(f"- {category_chart}")
    print(f"- {monthly_chart}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
