"""SmartSales: validate sales data and create an Excel report.

Usage:
    python smartsales_stage3.py path/to/sales_data.csv

If no CSV path is provided, the program looks for:
    data/raw/sales_data_1500_with_invalid.csv

The Excel report is saved to:
    reports/smartsales_report.xlsx
"""

import csv
import json
import logging
import sys
from collections import Counter
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException

from zipfile import BadZipFile

REQUIRED_COLUMNS = {
    "order_id",
    "order_date",
    "product",
    "category",
    "quantity",
    "unit_price",
}

VALID_CATEGORIES = {
    "Electronics",
    "Accessories",
    "Storage",
    "Furniture",
    "Office Supplies",
    "Stationery",
}

DEFAULT_INPUT = Path("data/raw/sales_data_1500_with_invalid.csv")
DEFAULT_OUTPUT = Path("reports/smartsales_report.xlsx")


CONFIG_FILE = Path("config.json")

def load_config(config_file=CONFIG_FILE):
    """Load report settings from JSON, or use built-in defaults."""
    settings = {
        "input_file": str(DEFAULT_INPUT),
        "output_file": str(DEFAULT_OUTPUT),
        "report_title": "SmartSales Report",
        "currency_symbol": "$",
        "top_products_count": 5,
    }

    config_path = Path(config_file)

    if not config_path.is_file():
        logging.info(
            "Configuration file not found; using built-in defaults."
        )
        return settings

    try:
        with config_path.open("r", encoding="utf-8") as file:
            user_settings = json.load(file)

        if not isinstance(user_settings, dict):
            raise ValueError("Configuration must be a JSON object.")

        string_settings = (
            "input_file",
            "output_file",
            "report_title",
            "currency_symbol",
        )

        for key in string_settings:
            if key in user_settings:
                value = user_settings[key]

                if not isinstance(value, str) or not value.strip():
                    raise ValueError(
                        f"'{key}' must be a non-empty string."
                    )

                settings[key] = value.strip()

        if "top_products_count" in user_settings:
            count = user_settings["top_products_count"]

            if not isinstance(count, int) or isinstance(count, bool) or count < 1:
                raise ValueError(
                    "'top_products_count' must be a positive integer."
                )

            settings["top_products_count"] = count

        logging.info("Configuration loaded from %s", config_path)
        return settings

    except (json.JSONDecodeError, OSError, ValueError) as error:
        raise ValueError(
            f"Unable to load configuration file "
            f"'{config_path}': {error}"
        ) from error


def read_sales_file(filename):
    """Read a CSV or Excel file and return rows with their original row numbers."""
    path = Path(filename)

    if not path.is_file():
        print(f"Error: Sales file not found: {path}")
        return None

    # Only CSV and modern Excel workbooks are supported.
    if path.suffix.lower() not in {".csv", ".xlsx"}:
        print("Error: Unsupported file type. Use a .csv or .xlsx file.")
        return None

    try:
        # Read CSV files.
        if path.suffix.lower() == ".csv":
            with path.open("r", encoding="utf-8-sig", newline="") as file:
                reader = csv.DictReader(file)

                if not reader.fieldnames:
                    print("Error: The CSV file has no header row.")
                    return None

                headers = [name.strip() for name in reader.fieldnames]
                reader.fieldnames = headers

                missing_columns = REQUIRED_COLUMNS - set(headers)
                if missing_columns:
                    print(
                        "Error: Missing required columns:",
                        ", ".join(sorted(missing_columns)),
                    )
                    return None

                rows = []

                for row_number, row in enumerate(reader, start=2):
                    if None in row:
                        print(
                            f"Error: CSV row {row_number} has more values "
                            "than expected."
                        )
                        return None

                    clean_row = {
                        key.strip(): (value or "").strip()
                        for key, value in row.items()
                    }
                    rows.append((row_number, clean_row))

                return rows

        # Read Excel files.
        workbook = load_workbook(
            path,
            read_only=True,
            data_only=True,
        )

        try:
            worksheet = workbook.active
            header_row = next(
                worksheet.iter_rows(values_only=True),
                None,
            )

            if not header_row:
                print("Error: The Excel file has no header row.")
                return None

            headers = [
                str(value).strip() if value is not None else ""
                for value in header_row
            ]

            if not any(headers):
                print("Error: The Excel file has no header row.")
                return None

            missing_columns = REQUIRED_COLUMNS - set(headers)
            if missing_columns:
                print(
                    "Error: Missing required columns:",
                    ", ".join(sorted(missing_columns)),
                )
                return None

            rows = []

            for row_number, values in enumerate(
                worksheet.iter_rows(min_row=2, values_only=True),
                start=2,
            ):
                # Ignore completely empty worksheet rows.
                if all(value is None for value in values):
                    continue

                # Reject data entered beyond the header columns.
                if any(value is not None for value in values[len(headers):]):
                    print(
                        f"Error: Excel row {row_number} has more values "
                        "than expected."
                    )
                    return None

                clean_row = {}

                for index, header in enumerate(headers):
                    value = values[index] if index < len(values) else None

                    # Convert Excel date cells into the expected date format.
                    if isinstance(value, (datetime, date)):
                        value = value.strftime("%Y-%m-%d")
                    elif value is None:
                        value = ""
                    else:
                        value = str(value).strip()

                    clean_row[header] = value

                rows.append((row_number, clean_row))

            return rows

        finally:
            workbook.close()

    except (OSError, csv.Error, UnicodeError, ValueError,
            BadZipFile, InvalidFileException) as error:
        print(f"Error reading sales file: {error}")
        return None


def find_duplicate_ids(rows):
    """Return non-empty order IDs that occur more than once."""
    counts = Counter()

    for _, row in rows:
        order_id = row.get("order_id", "").strip()
        if order_id:
            counts[order_id] += 1

    return {
        order_id
        for order_id, count in counts.items()
        if count > 1
    }


def validate_row(row, duplicate_ids):
    """Return validation errors and parsed quantity and price for a row."""
    errors = []

    order_id = row.get("order_id", "").strip()
    order_date = row.get("order_date", "").strip()
    product = row.get("product", "").strip()
    category = row.get("category", "").strip()
    quantity_text = row.get("quantity", "").strip()
    price_text = row.get("unit_price", "").strip()

    # Validate order ID.
    if not order_id:
        errors.append("Missing order_id")
    else:
        try:
            if int(order_id) <= 0:
                errors.append("order_id must be a positive integer")
        except ValueError:
            errors.append("order_id must be a positive integer")

        if order_id in duplicate_ids:
            errors.append(f"Duplicate order_id: {order_id}")

    # Validate date and require YYYY-MM-DD format.
    if not order_date:
        errors.append("Missing order_date")
    else:
        try:
            parsed_date = datetime.strptime(order_date, "%Y-%m-%d")
            if parsed_date.strftime("%Y-%m-%d") != order_date:
                errors.append("Date must use YYYY-MM-DD format")
        except ValueError:
            errors.append("Invalid date; expected YYYY-MM-DD")

    # Validate product and category.
    if not product:
        errors.append("Missing product")

    if not category:
        errors.append("Missing category")
    elif category not in VALID_CATEGORIES:
        errors.append(f"Unknown category: {category!r}")

    # Validate quantity.
    quantity = None
    if not quantity_text:
        errors.append("Missing quantity")
    else:
        try:
            quantity = int(quantity_text)
            if quantity <= 0:
                errors.append("Quantity must be a positive integer")
        except ValueError:
            errors.append("Quantity must be a positive integer")

    # Validate unit price.
    unit_price = None
    if not price_text:
        errors.append("Missing unit_price")
    else:
        try:
            unit_price = Decimal(price_text)
            if not unit_price.is_finite() or unit_price <= 0:
                errors.append("Unit price must be greater than zero")
        except InvalidOperation:
            errors.append("Unit price must be a valid number")

    return errors, quantity, unit_price


def prepare_records(rows):
    """Separate valid sales records from invalid rows and their errors."""
    duplicate_ids = find_duplicate_ids(rows)
    valid_records = []
    invalid_records = []

    for row_number, row in rows:
        errors, quantity, unit_price = validate_row(row, duplicate_ids)

        if errors:
            invalid_records.append({
                "csv_row": row_number,
                "order_id": row.get("order_id", ""),
                "order_date": row.get("order_date", ""),
                "product": row.get("product", ""),
                "category": row.get("category", ""),
                "quantity": row.get("quantity", ""),
                "unit_price": row.get("unit_price", ""),
                "errors": "; ".join(errors),
            })
        else:
            valid_records.append({
                "order_id": row["order_id"],
                "order_date": row["order_date"],
                "product": row["product"],
                "category": row["category"],
                "quantity": quantity,
                "unit_price": unit_price,
            })

    return valid_records, invalid_records


def calculate_summary(valid_records):
    """Calculate total revenue, units sold, and revenue per product."""
    total_revenue = Decimal("0.00")
    total_units = 0
    revenue_by_product = {}

    for record in valid_records:
        product = record["product"]
        revenue = record["quantity"] * record["unit_price"]

        total_revenue += revenue
        total_units += record["quantity"]
        revenue_by_product[product] = (
            revenue_by_product.get(product, Decimal("0.00")) + revenue
        )

    return total_revenue, total_units, revenue_by_product


def style_worksheet(worksheet):
    """Apply simple formatting to an Excel worksheet."""
    header_fill = PatternFill("solid", fgColor="1F4E78")

    for cell in worksheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    # Set each column width based on its content, with a sensible maximum.
    for column_cells in worksheet.columns:
        column_letter = get_column_letter(column_cells[0].column)
        longest_value = max(
            (len(str(cell.value)) if cell.value is not None else 0)
            for cell in column_cells
        )
        worksheet.column_dimensions[column_letter].width = min(longest_value + 2, 45)


def create_excel_report(
    rows,
    valid_records,
    invalid_records,
    output_file,
    report_title="SmartSales Report",
    currency_symbol="$",
):
    """Create an Excel workbook with summary, product revenue, and errors."""
    total_revenue, total_units, revenue_by_product = calculate_summary(valid_records)
    workbook = Workbook()


    # Worksheet 1: overall summary.
    summary = workbook.active
    summary.title = "Sales Summary"
    workbook.properties.title = report_title

    summary.append(["Metric", "Value"])
    summary.append(["Total records", len(rows)])
    summary.append(["Valid records", len(valid_records)])
    summary.append(["Invalid records", len(invalid_records)])
    summary.append(["Total units sold", total_units])
    summary.append(["Total revenue", float(total_revenue)])

    summary["B6"].number_format = f'"{currency_symbol}"#,##0.00'
    style_worksheet(summary)

    # Worksheet 2: revenue grouped by product.
    product_sheet = workbook.create_sheet("Product Revenue")
    product_sheet.append(["Product", "Revenue"])

    for product, revenue in sorted(
        revenue_by_product.items(), key=lambda item: item[1], reverse=True
    ):
        product_sheet.append([product, float(revenue)])

    for cell in product_sheet["B"][1:]:
        cell.number_format = f'"{currency_symbol}"#,##0.00'
    style_worksheet(product_sheet)

    # Worksheet 3: all invalid rows and their validation errors.
    errors_sheet = workbook.create_sheet("Invalid Records")
    error_columns = [
        "CSV Row", "Order ID", "Order Date", "Product", "Category",
        "Quantity", "Unit Price", "Validation Errors",
    ]
    errors_sheet.append(error_columns)

    for record in invalid_records:
        errors_sheet.append([
            record["csv_row"],
            record["order_id"],
            record["order_date"],
            record["product"],
            record["category"],
            record["quantity"],
            record["unit_price"],
            record["errors"],
        ])

    style_worksheet(errors_sheet)
    errors_sheet.column_dimensions["H"].width = 60

    # Make the error descriptions easier to read in Excel.
    for row in errors_sheet.iter_rows(min_row=2, min_col=8, max_col=8):
        row[0].alignment = Alignment(vertical="top", wrap_text=True)

    output_path = Path(output_file)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Save the Excel report.
    workbook.save(output_path)

    # Verify that the report was created successfully.
    if not output_path.is_file():
        raise OSError(f"Excel report was not created: {output_path}")

    if output_path.stat().st_size == 0:
        raise OSError(f"Excel report is empty: {output_path}")

    logging.info("Excel report created successfully: %s", output_path)

    return output_path, total_revenue, total_units



def main():
    """Run validation and create the Excel report."""

    # Configure logging for the terminal and GitHub Actions.
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )

    logging.info("SmartSales report generation started.")

    # Load the default settings from config.json.
    try:
        settings = load_config()
    except ValueError as error:
        logging.error("%s", error)
        return 1

    # Command-line arguments override configuration settings.
    input_file = (
        Path(sys.argv[1])
        if len(sys.argv) > 1
        else Path(settings["input_file"])
    )

    output_file = (
        Path(sys.argv[2])
        if len(sys.argv) > 2
        else Path(settings["output_file"])
    )

    logging.info("Input file: %s", input_file)
    logging.info("Output file: %s", output_file)

    # Read the sales data.
    rows = read_sales_file(input_file)

    if rows is None:
        logging.error("Unable to read the input file: %s", input_file)
        return 1

    try:
        # Validate records and separate valid from invalid data.
        valid_records, invalid_records = prepare_records(rows)

        # Generate and verify the Excel report.
        report_path, total_revenue, total_units = create_excel_report(
            rows,
            valid_records,
            invalid_records,
            output_file,
        )

        # Display the report summary.
        print("\nSmartSales - Excel Report Created")
        print("-" * 45)
        print(f"Input file:       {input_file}")
        print(f"Total records:    {len(rows)}")
        print(f"Valid records:    {len(valid_records)}")
        print(f"Invalid records:  {len(invalid_records)}")
        print(f"Total units sold: {total_units:,}")
        print(f"Total revenue:    ${total_revenue:,.2f}")
        print(f"Excel report:     {report_path}")
        print("\nThe report contains three worksheets:")
        print("  1. Sales Summary")
        print("  2. Product Revenue")
        print("  3. Invalid Records")
        print("Only valid records are included in sales calculations.")

        logging.info(
            "Processing completed. Total records: %d, valid: %d, invalid: %d",
            len(rows),
            len(valid_records),
            len(invalid_records),
        )
        logging.info("SmartSales report generation finished successfully.")

        return 0

    except Exception:
        logging.exception("Report generation failed.")
        return 1


if __name__ == "__main__":
    sys.exit(main())