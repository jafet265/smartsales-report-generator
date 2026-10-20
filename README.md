## smartsales-report-generator
A Python automation tool that processes CSV and Excel sales data, detects invalid records, calculates key business metrics, and generates professional Excel reports, PDF summaries, and charts.

# Client: A small online store owner.

The owner sells products such as laptops, mice, keyboards, and monitors. Every week, the store exports its sales transactions into an Excel or CSV file.

The owner currently opens the file, calculates revenue manually, checks for errors, and creates a report. This process takes approximately 2–3 hours every week.

The owner wants you to automate this work using Python.
The client's request
I run a small online store and have a CSV or Excel file containing my sales transactions. I need a Python program that processes the file automatically, checks for data errors, calculates important sales metrics, and generates a professional report.

I want to run the program whenever I receive a new sales file without manually calculating anything.

The program should save the results in an output folder and clearly explain any problems found in the source data.

The task is to build a tool that fulfills this request.

# The input file your client provides
Client gives me this sample file, named sales_data.csv.
```bash
order_id,order_date,product,category,quantity,unit_price
1001,2026-10-01,Laptop,Electronics,2,800.00
1002,2026-10-01,Mouse,Accessories,10,20.00
1003,2026-10-02,Keyboard,Accessories,5,50.00
1004,2026-10-03,Monitor,Electronics,3,200.00
1005,2026-10-04,Laptop,Electronics,1,800.00
1006,2026-10-05,Mouse,Accessories,15,20.00
1007,2026-10-06,Monitor,Electronics,2,200.00
1008,2026-10-07,Keyboard,Accessories,4,50.00
1009,2026-10-08,Laptop,Electronics,3,800.00
1010,2026-10-09,Mouse,Accessories,8,20.00
```


Each row represents one order line. The file contains:
- order_id: Unique order-line identifier for this project.
- order_date: Date of the transaction.
- product: Product sold.
- category: Product category.
- quantity: Number of units sold.
- unit_price: Selling price per unit, in dollars.
For this initial project, assume that all sales are completed, prices are in USD, and revenue means quantity multiplied by unit price. Taxes, shipping, discounts, and refunds are outside the initial scope.

# What my Python program must do
Client requires six features.
Requirement 1 — Read the source file
The program must:
- Accept a CSV file or an Excel .xlsx file.
- Read its rows and column names.
- Display a useful error message if the file cannot be opened.
- Check whether all six required columns exist.

Requirement 2 — Validate the data
The program must detect:


| Problem          | Example            | Required behavior |
|------------------|--------------------|-------------------|
| Missing product  | Empty product name | Flag the row      |
| Invalid quantity | abc                | Flag the row      |
| Invalid price    | -50                | Flag the row      |
| ...              | ...                | ...               |

Business rule: Only valid records should contribute to the sales totals. Document which records were excluded and why.

Requirement 3 — Calculate sales metrics
For valid records, calculate the following:
1. Total revenue.
2. Total units sold.
3. Number of valid order lines.
4. Number of distinct products.
5. Average revenue per order line.
6. Revenue by product.
7. Revenue by category.
8. Revenue by date.

Requirement 4 — Identify the best-performing products
The program must determine:
- Which product generated the highest revenue?
- Which product sold the most units?
- Which category generated the highest revenue?
- Which date had the highest revenue?
These results should be included in the final report.

Requirement 5 — Generate output files
The program must create an output folder and generate three files.

A. sales_report.xlsx

An Excel workbook containing:
- Summary: the main sales metrics.
- Product Report: revenue and units sold by product.
- Category Report: revenue by category.
- Daily Report: revenue by date.
- Data Quality: invalid records and explanations.

B. sales_report.pdf

    A readable report showing the key metrics, best-performing products, and charts.

C. sales_chart.png

A bar chart showing revenue by product.

If there are invalid records, the report must disclose them.

Requirement 6 — Make the program reusable
The client does not want to edit the source code every week.

The program should support a command such as:
```bash
python report_generator.py sales_data.csv
```
It should also support Excel input:
```bash
python report_generator.py sales_data.xlsx
```

The program should generate the reports automatically and display a message such as:
```bash
Sales report generation completed.

Input file: sales_data.csv
Valid order lines: 10
Invalid order lines: 0
Total units sold: 53
Total revenue: $5,050.00
```
Output files:
- output/sales_report.xlsx
- output/sales_report.pdf
- output/sales_chart.png



Note: The example's unit count and revenue are illustrative placeholders. Your implementation should calculate the actual results from the supplied data rather than hard-code them.

## Stage 7 — Interactive Streamlit dashboard

Stage 7 adds an interactive dashboard for exploring validated sales data.

### Start the dashboard

From the repository root, install the project dependencies and run:

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

Streamlit will print a local URL in the terminal. Open that URL to use the dashboard. In GitHub Codespaces, use the forwarded port/browser link shown by Codespaces if prompted.

### Dashboard features

- KPI cards for total revenue, units sold, valid order lines, distinct products, and average revenue per valid line.
- Filters for product, category, and order-date range.
- Revenue charts by product, category, and date.
- Business insights for the top-revenue product, most-sold product, top-revenue category, and highest-revenue date.
- Data-quality counts and a table of invalid records with validation explanations.
- A table of validated sales records and a product summary.
- CSV upload support for analyzing a new sales file without changing the source code.

The dashboard reuses `src/smartsales_stage3.py` for reading and validation. Invalid records are excluded from the sales metrics and charts. The supplied sample CSV is loaded by default.
