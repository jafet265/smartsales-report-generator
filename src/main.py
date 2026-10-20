import csv
from decimal import Decimal, InvalidOperation
from pathlib import Path


def calculate_revenue(csv_file):
    """Read sales data from a CSV file and calculate total revenue."""

    total_revenue = Decimal("0.00")
    valid_transactions = 0
    invalid_transactions = 0

    csv_path = Path(csv_file)

    # Check whether the CSV file exists.
    if not csv_path.is_file():
        print(f"Error: File not found: {csv_path}")
        return

    try:
        with csv_path.open(
            mode="r",
            encoding="utf-8-sig",
            newline=""
        ) as file:

            reader = csv.DictReader(file)

            required_columns = {
                "order_id",
                "order_date",
                "product",
                "category",
                "quantity",
                "unit_price",
            }

            # Validate CSV column names.
            if not reader.fieldnames or not required_columns.issubset(
                set(reader.fieldnames)
            ):
                print("Error: CSV file has missing required columns.")
                return

            print("SmartSale Revenue Report")
            print("-" * 60)

            for row_number, row in enumerate(reader, start=2):
                try:
                    # Validate required fields.
                    order_id = (row["order_id"] or "").strip()
                    order_date = (row["order_date"] or "").strip()
                    product = (row["product"] or "").strip()
                    category = (row["category"] or "").strip()
                    quantity_text = (row["quantity"] or "").strip()
                    price_text = (row["unit_price"] or "").strip()

                    if not all([
                        order_id,
                        order_date,
                        product,
                        category,
                        quantity_text,
                        price_text,
                    ]):
                        raise ValueError("Missing required field")

                    # Convert input values.
                    quantity = int(quantity_text)
                    unit_price = Decimal(price_text)

                    # Validate numeric values.
                    if quantity <= 0:
                        raise ValueError("Quantity must be positive")

                    if not unit_price.is_finite() or unit_price < 0:
                        raise ValueError("Invalid unit price")

                    # Calculate transaction revenue.
                    revenue = quantity * unit_price
                    total_revenue += revenue
                    valid_transactions += 1

                    print(
                        f"{order_id:<15}"
                        f"{product:<20}"
                        f"${revenue:>12,.2f}"
                    )

                except (
                    ValueError,
                    InvalidOperation,
                    TypeError,
                ) as error:
                    invalid_transactions += 1
                    print(
                        f"Skipping row {row_number}: {error}"
                    )

        # Display the final summary.
        print("-" * 60)
        print(f"Valid transactions:   {valid_transactions}")
        print(f"Invalid transactions: {invalid_transactions}")
        print(f"Total revenue:        ${total_revenue:,.2f}")

    except (OSError, csv.Error) as error:
        print(f"Error reading CSV file: {error}")


if __name__ == "__main__":
    calculate_revenue("data/raw/sales_data_1500_with_invalid.csv")
