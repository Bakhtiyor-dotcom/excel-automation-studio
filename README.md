# Excel Automation Tool

A Python command-line tool that merges CSV and Excel sales files, cleans invalid records, removes duplicate orders, and creates a formatted Excel workbook with business summaries.

## Features
- Combine multiple `.csv` and `.xlsx` files
- Normalize input column names (case, whitespace, spaces)
- Validate required fields, dates, quantities, and prices
- Deduplicate by `order_id` (first valid occurrence wins)
- Track rejected rows and duplicate order IDs
- Calculate revenue and units sold
- Export five Excel sheets: Overview, Cleaned Orders, By Product, By Customer, Rejected Rows
- Run automated tests with pytest

## Requirements
Python 3.10+ (Python 3.12 recommended).

## Installation (macOS / Linux)
```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

## Quick start
```bash
python main.py 'samples/*.csv' -o output/sales_report.xlsx
```

Expected console summary for bundled sample files:
```text
Valid orders: 6
Rejected rows: 3
Duplicate IDs removed: 1
Revenue: 708.00
```

Open `output/sales_report.xlsx` to inspect the results.

## Your own data
```bash
python main.py sales_january.xlsx sales_february.csv -o output/my_report.xlsx
```

Input columns (case-insensitive; spaces become underscores):

| Column | Meaning |
| --- | --- |
| order_id | Unique order identifier |
| date | Order date (e.g. 2026-01-15) |
| customer | Customer name |
| product | Product name |
| quantity | Positive whole number |
| unit_price | Nonnegative numeric price |

**Notes:** All input files must contain these columns. The tool treats `order_id` as unique, keeps the first valid row when duplicates appear, and does not convert currencies. CSV files are expected to be UTF-8 encoded. Use consistent date formats across files. This is a portfolio demo, not an accounting system.

## Tests
```bash
python -m pip install -r requirements-dev.txt
python -m pytest -q
```

## Technology
Python, pandas, openpyxl, pytest.

## Portfolio description
**Title:** Python Excel & CSV Automation — Data Cleaning and Sales Reports

**Description:** Built a Python command-line utility that merges spreadsheet files, validates and deduplicates sales records, and generates formatted Excel reports with product and customer summaries. Includes sample datasets, automated tests, and documentation.

## Future improvements
Configurable column mapping, configurable deduplication rules, currency support, GUI, and scheduled execution.


## Web dashboard (Streamlit)

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Open the local address shown in Terminal (usually http://localhost:8501). The dashboard starts with sample data; upload your own CSV/XLSX files from the sidebar. Review KPIs, charts, data tables and download a fresh Excel report. The existing CLI remains available.

**Portfolio note:** Screenshots should be captured from your own running app. This app is intended as a local portfolio demo; do not upload confidential customer files to public hosting without assessing privacy and security.

## Version 2 improvements

- Upload one or more CSV/XLSX files and map differently named columns to the six required fields.
- Recognizes common English and Russian column names automatically; manual selection is available.
- Prevents mapping the same source column to multiple required fields.
- Clearer metric cards in light and dark themes.
- Bundled sample data remains available for demos. Uploading files takes precedence over sample data.

The app expects order-level data. `quantity` must be a positive integer, `unit_price` must be nonnegative, and `order_id` must uniquely identify an order. Revenue is in the input file's currency; do not combine currencies without conversion.
