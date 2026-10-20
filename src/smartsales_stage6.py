"""SmartSales Stage 6: generate automated business insights.

Usage:
    python smartsales_stage6.py
    python smartsales_stage6.py path/to/sales_data.csv

The script reuses Stage 3's validation functions and analyzes valid records only.
Output:
    reports/sales_insights.txt
"""

import argparse
import calendar
from collections import defaultdict
from datetime import datetime
from decimal import Decimal
from pathlib import Path

from smartsales_stage3 import read_sales_file, prepare_records, calculate_summary


PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_INPUT = PROJECT_ROOT / "data" / "raw" / "sales_data_1500_with_invalid.csv"
DEFAULT_OUTPUT = PROJECT_ROOT / "reports" / "sales_insights.txt"


def money(value):
    """Format a Decimal as US dollars."""
    return f"${value:,.2f}"


def analyze_sales(valid_records):
    """Summarize sales by product, category, and month."""
    revenue_by_product = defaultdict(lambda: Decimal("0.00"))
    units_by_product = defaultdict(int)
    revenue_by_category = defaultdict(lambda: Decimal("0.00"))
    units_by_category = defaultdict(int)
    revenue_by_month = defaultdict(lambda: Decimal("0.00"))
    units_by_month = defaultdict(int)

    total_revenue, total_units, _ = calculate_summary(valid_records)

    for record in valid_records:
        product = record["product"]
        category = record["category"]
        quantity = record["quantity"]
        revenue = quantity * record["unit_price"]
        month_key = record["order_date"][:7]

        revenue_by_product[product] += revenue
        units_by_product[product] += quantity
        revenue_by_category[category] += revenue
        units_by_category[category] += quantity
        revenue_by_month[month_key] += revenue
        units_by_month[month_key] += quantity

    return {
        "total_revenue": total_revenue,
        "total_units": total_units,
        "revenue_by_product": dict(revenue_by_product),
        "units_by_product": dict(units_by_product),
        "revenue_by_category": dict(revenue_by_category),
        "units_by_category": dict(units_by_category),
        "revenue_by_month": dict(revenue_by_month),
        "units_by_month": dict(units_by_month),
    }


def rank_items(mapping):
    """Return (key, value) pairs sorted by value descending, then name."""
    return sorted(mapping.items(), key=lambda item: (-item[1], item[0].lower()))


def month_label(month_key):
    """Convert YYYY-MM into a readable month label."""
    parsed = datetime.strptime(month_key, "%Y-%m")
    return f"{calendar.month_name[parsed.month]} {parsed.year}"


def build_insights(rows, valid_records, invalid_records, analysis):
    """Build a readable plain-text business insights report."""
    total_rows = len(rows)
    valid_count = len(valid_records)
    invalid_count = len(invalid_records)
    total_revenue = analysis["total_revenue"]
    total_units = analysis["total_units"]

    lines = [
        "SMARTSALES — AUTOMATED BUSINESS INSIGHTS",
        "=" * 48,
        "",
        "1. OVERALL PERFORMANCE",
        "-" * 48,
        f"CSV data rows analyzed: {total_rows:,}",
        f"Valid records: {valid_count:,}",
        f"Invalid records excluded: {invalid_count:,}",
        f"Valid-record rate: {(valid_count / total_rows * 100):.2f}%" if total_rows else "Valid-record rate: N/A",
        f"Total units sold: {total_units:,}",
        f"Total revenue: {money(total_revenue)}",
        (
            f"Average revenue per valid record: {money(total_revenue / valid_count)}"
            if valid_count else "Average revenue per valid record: N/A"
        ),
        "",
    ]

    if not valid_records:
        lines.extend([
            "No valid sales records were available for business analysis.",
            "Check the Invalid Records worksheet in the Excel report for validation errors.",
            "",
        ])
        return "\n".join(lines)

    products_by_revenue = rank_items(analysis["revenue_by_product"])
    products_by_units = rank_items(analysis["units_by_product"])
    categories_by_revenue = rank_items(analysis["revenue_by_category"])
    months_by_revenue = rank_items(analysis["revenue_by_month"])

    top_product, top_product_revenue = products_by_revenue[0]
    top_units_product, top_units = products_by_units[0]
    top_category, top_category_revenue = categories_by_revenue[0]
    top_month, top_month_revenue = months_by_revenue[0]

    lines.extend([
        "2. KEY FINDINGS",
        "-" * 48,
        f"Highest-revenue product: {top_product} ({money(top_product_revenue)})",
        f"Most units sold for a product: {top_units_product} ({top_units:,} units)",
        f"Highest-revenue category: {top_category} ({money(top_category_revenue)})",
        f"Highest-revenue month: {month_label(top_month)} ({money(top_month_revenue)})",
        "",
        "3. REVENUE BY PRODUCT",
        "-" * 48,
    ])

    for name, revenue in products_by_revenue:
        share = (revenue / total_revenue * 100) if total_revenue else Decimal("0")
        lines.append(f"{name}: {money(revenue)} ({share:.2f}% of total revenue)")

    lines.extend(["", "4. REVENUE BY CATEGORY", "-" * 48])
    for name, revenue in categories_by_revenue:
        share = (revenue / total_revenue * 100) if total_revenue else Decimal("0")
        lines.append(f"{name}: {money(revenue)} ({share:.2f}% of total revenue)")

    lines.extend(["", "5. MONTHLY REVENUE", "-" * 48])
    for month_key, revenue in sorted(analysis["revenue_by_month"].items()):
        lines.append(f"{month_label(month_key)}: {money(revenue)}")

    lines.extend(["", "6. SUGGESTED AREAS TO INVESTIGATE", "-" * 48])
    lines.append(
        f"Review stock availability and customer demand for '{top_units_product}', "
        f"the product with the highest unit sales."
    )
    lines.append(
        f"Investigate what contributed to '{top_category}' generating the most revenue; "
        "compare pricing, sales volume, and marketing activity."
    )
    lines.append(
        f"Review promotions, inventory, and sales activity around {month_label(top_month)}, "
        "the strongest month by revenue."
    )

    if invalid_count:
        lines.append(
            f"Review the {invalid_count:,} excluded records in the Excel report's "
            "'Invalid Records' worksheet before using the dataset for further decisions."
        )
    else:
        lines.append("No invalid records were found in this dataset.")

    lines.extend([
        "",
        "Note: These findings describe patterns in the supplied data. They do not establish "
        "why a product, category, or month performed as it did.",
        "",
    ])
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(
        description="Generate automated business insights from validated sales data."
    )
    parser.add_argument(
        "csv_file",
        nargs="?",
        help="Optional CSV path. Defaults to the project's sample dataset.",
    )
    parser.add_argument(
        "--output",
        help="Optional report output path. Defaults to reports/sales_insights.txt.",
    )
    args = parser.parse_args()

    input_file = Path(args.csv_file).expanduser() if args.csv_file else DEFAULT_INPUT
    if not input_file.is_absolute():
        input_file = (Path.cwd() / input_file).resolve()

    output_file = Path(args.output).expanduser() if args.output else DEFAULT_OUTPUT
    if not output_file.is_absolute():
        output_file = (Path.cwd() / output_file).resolve()

    rows = read_sales_file(input_file)
    if rows is None:
        return 1

    valid_records, invalid_records = prepare_records(rows)
    analysis = analyze_sales(valid_records)
    report = build_insights(rows, valid_records, invalid_records, analysis)

    try:
        output_file.parent.mkdir(parents=True, exist_ok=True)
        output_file.write_text(report + "\n", encoding="utf-8")
    except OSError as error:
        print(f"Error writing insights report: {error}")
        return 1

    print("SmartSales Stage 6: Automated Business Insights")
    print(f"Input CSV: {input_file}")
    print(f"Valid records: {len(valid_records):,}")
    print(f"Invalid records: {len(invalid_records):,}")
    print(f"Total revenue: {money(analysis['total_revenue'])}")
    print(f"Insights report: {output_file}")
    print("\nKey findings:")
    if valid_records:
        print(f"  Top-revenue product: {rank_items(analysis['revenue_by_product'])[0][0]}")
        print(f"  Top category: {rank_items(analysis['revenue_by_category'])[0][0]}")
        print(f"  Strongest month: {month_label(rank_items(analysis['revenue_by_month'])[0][0])}")
    else:
        print("  No valid sales records to analyze.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
