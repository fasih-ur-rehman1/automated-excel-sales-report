# Automated Excel Sales Report

An automated Python pipeline that transforms raw Excel sales data into a structured and professionally formatted sales report.

## Overview

This project demonstrates how Python can automate a repetitive business reporting workflow.

Instead of manually opening an Excel dataset, calculating summaries, creating report sheets, formatting cells, and adjusting the final workbook, the entire process is performed programmatically using **Pandas** and **OpenPyXL**.

## Business Problem

A business receives sales data in Excel and needs a clean management report containing:

* Overall sales and profit metrics
* Category-level sales and profit
* Region-level sales and profit
* Data-quality checks
* A professionally formatted Excel workbook

Performing these steps manually becomes repetitive and time-consuming when reports need to be generated regularly.

## Solution

The Python automation pipeline:

1. Loads the raw Excel dataset.
2. Inspects the dataset structure and data quality.
3. Checks for duplicates and missing values.
4. Checks for invalid negative values.
5. Calculates overall business metrics.
6. Generates category-level summaries.
7. Generates region-level summaries.
8. Exports the analysis to Excel.
9. Creates a report summary sheet.
10. Applies professional Excel formatting.
11. Automatically adjusts column widths.
12. Produces a reusable final report.

## Project Workflow

```text
Raw Excel Dataset
       ↓
     Pandas
       ↓
Data Inspection & Validation
       ↓
Business Analysis
       ↓
Category & Region Summaries
       ↓
ExcelWriter
       ↓
    OpenPyXL
       ↓
Formatting & Report Generation
       ↓
Final Excel Report
```

## Technologies Used

* Python
* Pandas
* OpenPyXL
* Excel
* Jupyter Notebook / Google Colab

## Key Features

### Data Analysis

* Dataset shape and column inspection
* Duplicate detection
* Missing-value detection
* Negative-value validation
* Sales aggregation
* Profit aggregation
* Average discount calculation
* Category analysis
* Regional analysis

### Excel Automation

* Multiple worksheet creation
* Automated data export
* Report summary generation
* Cell formatting
* Number formatting
* Percentage formatting
* Header styling
* Column width adjustment
* Merged report title
* Freeze panes
* Re-runnable report generation

## Output

The generated workbook contains three main sheets:

### Report Summary

Provides high-level business metrics including:

* Total Sales
* Total Profit
* Average Discount
* Total Orders
* Number of Categories
* Negative Profit Transactions

### Category Summary

Provides sales and profit totals grouped by product category.

### Region Summary

Provides sales and profit totals grouped by geographical region.

## Project Structure

```text
automated-excel-sales-report/
│
├── data/
│   └── Sample-sales-data-excel.xlsx
│
├── output/
│   └── Sales_Report.xlsx
│
├── src/
│   └── sales_report.py
│
├── docs/
│   └── project_report.pdf
│
└── README.md
```

## How to Run

Install the required libraries:
``` bash
pip install pandas openpyxl
```
Open the Jupyter/Colab notebook:
```bash
src/sales_report.ipynb
```
Run the cells in order to reproduce the analysis and generate the final Excel report in the output/ directory.

## Skills Demonstrated

This project demonstrates practical skills in:

* Python automation
* Excel automation
* Pandas data analysis
* OpenPyXL workbook manipulation
* Data validation
* Business reporting
* File and report generation
* Reusable automation workflows

## Why This Project Matters

The goal of this project is not simply to demonstrate Python syntax.

It demonstrates how Python can replace a repetitive manual business process with a reusable automated workflow.

The same approach can be adapted for recurring sales reports, financial reports, inventory reports, operational reports, and other Excel-based business workflows.
