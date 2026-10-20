# SmartSales Report Generator

A Python automation project that validates sales data, calculates business metrics, identifies invalid records, and generates Excel reports. It also includes an interactive Streamlit dashboard for exploring sales performance.

## Project Overview

Small businesses often spend hours reviewing sales files, calculating revenue, checking data quality, and preparing reports manually.

SmartSales automates these tasks using Python. It reads a CSV file, validates each record, calculates sales metrics using valid records only, and generates an Excel workbook containing sales summaries and data-quality information.

## Features

### 1. CSV Input

- Reads sales data from CSV files.
- Checks that all required columns are present.
- Handles UTF-8 and UTF-8-with-BOM encoded files.
- Reports missing files, missing columns, and malformed CSV rows.

### 2. Data Validation

SmartSales checks sales records for problems including:

- Missing or invalid order IDs.
- Duplicate order IDs.
- Invalid or incorrectly formatted dates.
- Missing product names.
- Unsupported product categories.
- Quantities that are missing, non-integer, zero, or negative.
- Prices that are missing, non-numeric, zero, negative, or non-finite.
- CSV rows containing more values than expected.

Invalid records are excluded from financial calculations and included in the data-quality report with explanations.

### 3. Sales Calculations

For valid records, SmartSales calculates:

- Total revenue.
- Total units sold.
- Revenue by product.
- Revenue by product category in the dashboard.
- Revenue over time in the dashboard.
- Sales summaries and data-quality counts.

Revenue is calculated as:

`Revenue = Quantity × Unit Price`

Python's `Decimal` type is used in the reporting logic to help preserve monetary precision.

### 4. Excel Report

The report generator creates:

`reports/smartsales_report.xlsx`

The workbook contains three worksheets:

| Worksheet | Contents |
|---|---|
| Sales Summary | Total records, valid records, invalid records, units sold, and revenue |
| Product Revenue | Revenue grouped by product |
| Invalid Records | Invalid source records and their validation errors |

The report includes formatted headers, adjusted column widths, and readable error descriptions.

### 5. Interactive Dashboard

The Streamlit dashboard provides an interactive interface for exploring sales data.

Features include:

- KPI cards for key sales metrics.
- Product, category, and date filters.
- Revenue charts by product, category, and date.
- Insights into top-performing products and categories.
- Data-quality counts and invalid-record details.
- A table of validated sales records.
- CSV upload support for analyzing other sales files.

The dashboard reuses the shared CSV-reading and validation functions in `src/smartsales_stage3.py`.

## Project Structure

```text
smartsales-report-generator/
├── app.py
├── requirements.txt
├── README.md
├── src/
│   └── smartsales_stage3.py
├── tests/
│   └── test_smartsales.py
├── data/
│   └── raw/
│       └── sales_data_1500_with_invalid.csv
└── reports/
    └── smartsales_report.xlsx
```

The report and data directories may contain additional files. Generated reports should be treated as output artifacts rather than source code.

## Requirements

- Python 3.11 or later is recommended.
- pip
- Git (optional, but recommended for version control)

The project has been tested in a Python 3.14 environment.

## Installation

Clone the repository and enter its directory:

```bash
git clone <YOUR_GITHUB_REPOSITORY_URL>
cd smartsales-report-generator
```

Replace `<YOUR_GITHUB_REPOSITORY_URL>` with your actual repository URL.

Install the dependencies:

```bash
python -m pip install -r requirements.txt
```

## Run the Excel Report Generator

From the repository root, run the program using the default input file:

```bash
python src/smartsales_stage3.py
```

The default input file is:

```text
data/raw/sales_data_1500_with_invalid.csv
```

To process a different CSV file, supply its path:

```bash
python src/smartsales_stage3.py path/to/sales_data.csv
```

The generated Excel report is saved to:

```text
reports/smartsales_report.xlsx
```

The program prints a summary of the input records, valid and invalid records, units sold, revenue, and report location.

## Run the Streamlit Dashboard

Start the dashboard from the repository root:

```bash
streamlit run app.py
```

Open the local URL printed by Streamlit. In GitHub Codespaces, use the forwarded port or browser link provided by Codespaces.

Upload a CSV file through the dashboard to analyze another dataset without changing the Python source code.

## Configuration

SmartSales supports a JSON configuration file named `config.json` in the repository's root directory.

### Default configuration

```json
{
    "input_file": "data/raw/sales_data_1500_with_invalid.csv",
    "output_file": "reports/smartsales_report.xlsx"
}
```

- `input_file`: path to the sales data file. Supported formats are CSV and Excel (`.xlsx`).
- `output_file`: path where the generated Excel report will be saved.

### Run with configuration settings

```bash
python src/smartsales_stage3.py
```

When no command-line paths are provided, SmartSales uses the paths specified in `config.json`.

### Override configuration from the command line

```bash
python src/smartsales_stage3.py data/raw/another_sales_file.csv reports/custom_report.xlsx
```

Command-line arguments take precedence over the corresponding configuration settings.

### Default behavior

If `config.json` does not exist, SmartSales uses its built-in default paths. If the configuration file exists but contains invalid JSON or invalid path settings, the program reports an error.

### Run automated tests

```bash
python -m pytest -q
```

## Input CSV Format

The input file must contain these six columns:

```csv
order_id,order_date,product,category,quantity,unit_price
1001,2026-10-01,Laptop,Electronics,2,800.00
1002,2026-10-01,Mouse,Accessories,10,20.00
1003,2026-10-02,Keyboard,Accessories,5,50.00
```

### Column Definitions

| Column | Description |
|---|---|
| `order_id` | Unique order-line identifier |
| `order_date` | Transaction date in `YYYY-MM-DD` format |
| `product` | Product name |
| `category` | Product category |
| `quantity` | Positive integer quantity sold |
| `unit_price` | Positive, finite unit price in USD |

The current accepted categories are:

- Electronics
- Accessories
- Storage
- Furniture
- Office Supplies
- Stationery

For this project, revenue means quantity multiplied by unit price. Taxes, shipping, discounts, and refunds are outside the current scope.

## Run the Automated Tests

Run the test suite:

```bash
python -m pytest -v
```

The tests cover CSV reading, required columns, record validation, duplicate IDs, revenue calculations, invalid-record exclusion, Excel report creation, and malformed CSV rows.

The latest verified test run contained 25 passing tests.

## Current Scope and Future Improvements

The current implementation supports CSV input, Excel report generation, and an interactive Streamlit dashboard.

Potential future improvements include:

- Reading Excel `.xlsx` files as input.
- Generating PDF reports.
- Saving standalone chart images.
- Adding further validation and reporting options.
- Improving command-line configuration for input and output paths.

These are future enhancements, not claims about the current implementation.

## Learning Outcomes

This project demonstrates practical use of:

- Python functions and modules.
- CSV file processing.
- Input validation and exception handling.
- Decimal arithmetic for monetary calculations.
- Excel report generation with `openpyxl`.
- Data analysis and visualization with pandas and Streamlit.
- Automated testing with pytest.
- Git and GitHub version control.

## Project Status

The project currently includes a CSV-based Excel report generator, a Streamlit dashboard, and an automated test suite.

Development continues as the application is refined and additional reporting capabilities are introduced.