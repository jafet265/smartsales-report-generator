import csv
from collections import Counter
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path


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
}


def read_sales_file(filename):
    """Read the CSV and return its rows, or stop with a clear error."""
    path = Path(filename)

    if not path.is_file():
        print(f"Error: File not found: {path}")
        return None

    try:
        with path.open("r", encoding="utf-8-sig", newline="") as file:
            reader = csv.DictReader(file)

            if not reader.fieldnames:
                print("Error: The CSV file has no header row.")
                return None

            # Allow accidental spaces around column names.
            reader.fieldnames = [name.strip() for name in reader.fieldnames]

            missing_columns = REQUIRED_COLUMNS - set(reader.fieldnames)
            if missing_columns:
                print("Error: Missing columns:", ", ".join(sorted(missing_columns)))
                return None

            # Keep the original CSV row number with each record.
            rows = []
            for row_number, row in enumerate(reader, start=2):
                rows.append((row_number, row))

            return rows

    except (OSError, csv.Error, UnicodeError) as error:
        print(f"Error reading CSV file: {error}")
        return None


def find_duplicate_ids(rows):
    """Return order IDs that appear more than once."""
    id_counts = Counter()

    for _, row in rows:
        order_id = (row.get("order_id") or "").strip()
        if order_id:
            id_counts[order_id] += 1

    return {
        order_id
        for order_id, count in id_counts.items()
        if count > 1
    }


def validate_row(row, duplicate_ids):
    """Return a list of errors for one row."""
    errors = []

    order_id = (row.get("order_id") or "").strip()
    order_date = (row.get("order_date") or "").strip()
    product = (row.get("product") or "").strip()
    category = (row.get("category") or "").strip()
    quantity_text = (row.get("quantity") or "").strip()
    price_text = (row.get("unit_price") or "").strip()

    # Check the order ID.
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

    # Check the date.
    if not order_date:
        errors.append("Missing order_date")
    else:
        try:
            parsed_date = datetime.strptime(order_date, "%Y-%m-%d")
            if parsed_date.strftime("%Y-%m-%d") != order_date:
                errors.append("Date must use YYYY-MM-DD format")
        except ValueError:
            errors.append("Invalid date; expected YYYY-MM-DD")

    # Check product and category.
    if not product:
        errors.append("Missing product")

    if not category:
        errors.append("Missing category")
    elif category not in VALID_CATEGORIES:
        errors.append(f"Unknown category: {category!r}")

    # Convert and validate quantity.
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

    # Convert and validate unit price.
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


def create_sales_summary(valid_records):
    """Calculate sales totals using valid records only."""
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


def main():
    filename = "data/raw/sales_data_1500_with_invalid.csv"
    rows = read_sales_file(filename)

    if rows is None:
        return

    duplicate_ids = find_duplicate_ids(rows)
    valid_records = []
    invalid_records = []

    # Validate each row and separate good data from bad data.
    for row_number, row in rows:
        errors, quantity, unit_price = validate_row(row, duplicate_ids)

        if errors:
            invalid_records.append((row_number, errors))
        else:
            valid_records.append({
                "order_id": row["order_id"].strip(),
                "order_date": row["order_date"].strip(),
                "product": row["product"].strip(),
                "category": row["category"].strip(),
                "quantity": quantity,
                "unit_price": unit_price,
            })

    total_revenue, total_units, revenue_by_product = create_sales_summary(
        valid_records
    )

    # Print the summary.
    print("\nSmartSales - Sales and Data Validation Report")
    print("-" * 55)
    print(f"Total records:   {len(rows)}")
    print(f"Valid records:   {len(valid_records)}")
    print(f"Invalid records: {len(invalid_records)}")

    print("\nRevenue by product (valid records only)")
    print("-" * 55)
    for product, revenue in sorted(
        revenue_by_product.items(),
        key=lambda item: item[1],
        reverse=True,
    ):
        print(f"{product:<20} ${revenue:,.2f}")

    print("-" * 55)
    print(f"Total units sold: {total_units:,}")
    print(f"Total revenue:    ${total_revenue:,.2f}")

    print("\nInvalid record details")
    print("-" * 55)
    if not invalid_records:
        print("No invalid records found.")
    else:
        for row_number, errors in invalid_records:
            print(f"CSV row {row_number}: {'; '.join(errors)}")

    print("\nNote: all occurrences of a duplicated ID are invalid.")
    print("Rows with any validation error are excluded from sales totals.")


if __name__ == "__main__":
    main()
