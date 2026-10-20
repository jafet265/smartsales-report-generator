"""Interactive Streamlit dashboard for the SmartSales Report Generator.

Run from the repository root with:
    streamlit run app.py

The dashboard reuses Stage 3's validation rules, so invalid rows never affect
sales KPIs or charts. The supplied sample CSV is loaded by default; users may
upload another CSV file from the sidebar.
"""

from collections import defaultdict
from datetime import date
from decimal import Decimal
from pathlib import Path
import tempfile

import pandas as pd
import streamlit as st

from src.smartsales_stage3 import prepare_records, read_sales_file


PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_INPUT = PROJECT_ROOT / "data" / "raw" / "sales_data_1500_with_invalid.csv"

st.set_page_config(
    page_title="SmartSales Dashboard",
    page_icon="📊",
    layout="wide",
)


def money(value):
    """Format a numeric value as US dollars."""
    return f"${float(value):,.2f}"


def load_rows(uploaded_file):
    """Load rows through the project's shared Stage 3 CSV reader."""
    if uploaded_file is None:
        return read_sales_file(DEFAULT_INPUT), str(DEFAULT_INPUT)

    # Stage 3 accepts file paths. Save the uploaded bytes temporarily so that
    # the dashboard uses exactly the same header and row-reading behavior.
    temporary_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as temp_file:
            temp_file.write(uploaded_file.getvalue())
            temporary_path = Path(temp_file.name)
        rows = read_sales_file(temporary_path)
        return rows, uploaded_file.name
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


def make_dataframes(rows, valid_records, invalid_records):
    """Convert validated records into dataframes for filtering and display."""
    valid_rows = []
    for record in valid_records:
        valid_rows.append({
            **record,
            "order_date_parsed": pd.to_datetime(record["order_date"]).date(),
            "revenue": float(record["quantity"] * record["unit_price"]),
        })

    valid_df = pd.DataFrame(valid_rows)
    invalid_df = pd.DataFrame(invalid_records)
    all_rows_df = pd.DataFrame([
        {"CSV row": row_number, **row}
        for row_number, row in rows
    ])
    return valid_df, invalid_df, all_rows_df


def summarize(filtered_df):
    """Calculate the dashboard's metrics from the currently filtered rows."""
    if filtered_df.empty:
        return {
            "revenue": 0.0,
            "units": 0,
            "transactions": 0,
            "products": 0,
            "average": 0.0,
        }

    revenue = float(filtered_df["revenue"].sum())
    count = len(filtered_df)
    return {
        "revenue": revenue,
        "units": int(filtered_df["quantity"].sum()),
        "transactions": count,
        "products": int(filtered_df["product"].nunique()),
        "average": revenue / count if count else 0.0,
    }


def group_revenue(filtered_df, group_column):
    """Return a chart-ready dataframe sorted by revenue descending."""
    if filtered_df.empty:
        return pd.DataFrame(columns=[group_column, "revenue"])
    return (
        filtered_df.groupby(group_column, as_index=False)["revenue"]
        .sum()
        .sort_values("revenue", ascending=False)
    )


def main():
    st.title("📊 SmartSales Dashboard")
    st.caption(
        "Explore validated sales performance. Invalid records are excluded "
        "from all revenue and unit calculations."
    )

    with st.sidebar:
        st.header("Data source")
        uploaded_file = st.file_uploader(
            "Upload a sales CSV file",
            type=["csv"],
            help="The CSV must contain the six required SmartSales columns.",
        )

    rows, input_label = load_rows(uploaded_file)
    if rows is None:
        st.error("The sales file could not be loaded. Check the file and its column headers.")
        st.info("Required columns: order_id, order_date, product, category, quantity, unit_price")
        st.stop()

    valid_records, invalid_records = prepare_records(rows)
    valid_df, invalid_df, _ = make_dataframes(rows, valid_records, invalid_records)

    st.caption(f"**Data source:** `{input_label}` · **Rows checked:** {len(rows):,}")

    # Filtering is applied only to valid records; quality counts remain global.
    with st.sidebar:
        st.header("Filters")
        if valid_df.empty:
            selected_products = []
            selected_categories = []
            start_date = end_date = None
        else:
            products = sorted(valid_df["product"].unique().tolist())
            categories = sorted(valid_df["category"].unique().tolist())
            selected_products = st.multiselect("Products", products, default=products)
            selected_categories = st.multiselect("Categories", categories, default=categories)
            earliest = valid_df["order_date_parsed"].min()
            latest = valid_df["order_date_parsed"].max()
            selected_dates = st.date_input(
                "Order date range",
                value=(earliest, latest),
                min_value=earliest,
                max_value=latest,
            )
            if isinstance(selected_dates, (tuple, list)) and len(selected_dates) == 2:
                start_date, end_date = selected_dates
            elif isinstance(selected_dates, (tuple, list)) and len(selected_dates) == 1:
                start_date = end_date = selected_dates[0]
            else:
                start_date = end_date = selected_dates

    filtered_df = valid_df.copy()
    if not filtered_df.empty:
        filtered_df = filtered_df[
            filtered_df["product"].isin(selected_products)
            & filtered_df["category"].isin(selected_categories)
        ]
        if start_date is not None and end_date is not None:
            filtered_df = filtered_df[
                filtered_df["order_date_parsed"].between(start_date, end_date)
            ]

    metrics = summarize(filtered_df)
    first_row = st.columns(5)
    first_row[0].metric("Total revenue", money(metrics["revenue"]))
    first_row[1].metric("Units sold", f"{metrics['units']:,}")
    first_row[2].metric("Valid order lines", f"{metrics['transactions']:,}")
    first_row[3].metric("Distinct products", f"{metrics['products']:,}")
    first_row[4].metric("Avg. revenue / line", money(metrics["average"]))

    st.subheader("Data quality")
    quality_cols = st.columns(3)
    quality_cols[0].metric("Source rows", f"{len(rows):,}")
    quality_cols[1].metric("Valid rows", f"{len(valid_records):,}")
    quality_cols[2].metric("Invalid rows excluded", f"{len(invalid_records):,}")
    if invalid_records:
        st.warning(
            f"{len(invalid_records):,} invalid row(s) were excluded. "
            "Open the Data quality details section to review the reasons."
        )
    else:
        st.success("All source rows passed Stage 3 validation.")

    st.subheader("Sales performance")
    chart_left, chart_right = st.columns(2)
    with chart_left:
        st.markdown("**Revenue by product**")
        product_revenue = group_revenue(filtered_df, "product").set_index("product")
        if product_revenue.empty:
            st.info("No valid sales match the current filters.")
        else:
            st.bar_chart(product_revenue, y="revenue", x_label="Product", y_label="Revenue (USD)")

    with chart_right:
        st.markdown("**Revenue by category**")
        category_revenue = group_revenue(filtered_df, "category").set_index("category")
        if category_revenue.empty:
            st.info("No category revenue to display.")
        else:
            st.bar_chart(category_revenue, y="revenue", x_label="Category", y_label="Revenue (USD)")

    st.markdown("**Revenue over time**")
    if filtered_df.empty:
        st.info("No daily revenue to display.")
    else:
        daily_revenue = (
            filtered_df.groupby("order_date_parsed")["revenue"]
            .sum()
            .sort_index()
            .rename_axis("Order date")
        )
        st.line_chart(daily_revenue, y_label="Revenue (USD)")

    st.subheader("Business insights")
    if filtered_df.empty:
        st.info("Choose filters that include valid sales records to see business insights.")
    else:
        insight_cols = st.columns(2)
        product_revenue_values = group_revenue(filtered_df, "product")
        product_units = (
            filtered_df.groupby("product")["quantity"].sum()
            .sort_values(ascending=False)
        )
        category_revenue_values = group_revenue(filtered_df, "category")
        daily_values = (
            filtered_df.groupby("order_date_parsed")["revenue"].sum()
            .sort_values(ascending=False)
        )
        with insight_cols[0]:
            best_product = product_revenue_values.iloc[0]
            best_units_product = product_units.index[0]
            st.markdown(f"- **Highest-revenue product:** {best_product['product']} ({money(best_product['revenue'])})")
            st.markdown(f"- **Most units sold:** {best_units_product} ({int(product_units.iloc[0]):,} units)")
        with insight_cols[1]:
            best_category = category_revenue_values.iloc[0]
            best_day = daily_values.index[0]
            st.markdown(f"- **Highest-revenue category:** {best_category['category']} ({money(best_category['revenue'])})")
            st.markdown(f"- **Highest-revenue date:** {best_day:%B %d, %Y} ({money(daily_values.iloc[0])})")

    st.subheader("Detailed reports")
    product_report = group_revenue(filtered_df, "product")
    if not filtered_df.empty:
        units_report = filtered_df.groupby("product", as_index=False)["quantity"].sum()
        product_report = product_report.merge(units_report, on="product", how="left")
        product_report = product_report.rename(columns={"revenue": "Revenue (USD)", "quantity": "Units sold"})
        product_report["Revenue (USD)"] = product_report["Revenue (USD)"].round(2)
    st.markdown("**Product report**")
    st.dataframe(product_report, use_container_width=True, hide_index=True)

    with st.expander("Data quality details", expanded=bool(invalid_records)):
        if invalid_df.empty:
            st.success("No invalid rows were found.")
        else:
            st.dataframe(invalid_df, use_container_width=True, hide_index=True)

    with st.expander("Validated sales records"):
        display_columns = [
            "order_id", "order_date", "product", "category", "quantity", "unit_price", "revenue"
        ]
        if valid_df.empty:
            st.info("No valid records are available.")
        else:
            st.dataframe(
                filtered_df[display_columns].assign(
                    unit_price=lambda frame: frame["unit_price"].map(money),
                    revenue=lambda frame: frame["revenue"].map(money),
                ),
                use_container_width=True,
                hide_index=True,
            )

    st.caption("SmartSales Stage 7 · Built with Streamlit and the shared Stage 3 validation rules.")


if __name__ == "__main__":
    main()
