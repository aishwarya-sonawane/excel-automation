# 🔧 Excel Automation Toolkit

A collection of Python scripts that automate the repetitive Excel tasks I used to do manually at Hero MotoCorp and SKF India. Each script replaces hours of copy-pasting with a single command.

![Python](https://img.shields.io/badge/Python-3.9+-blue)
![Pandas](https://img.shields.io/badge/Pandas-2.0+-green)
![License](https://img.shields.io/badge/License-MIT-green)

## The Problem

At Hero MotoCorp, I inherited a monthly pricing report that took 40+ hours to compile — pulling from 6 SAP extracts, manual copy-paste, rebuilding pivot tables every time. These scripts solve that class of problem.

## Scripts

### 1. `merge_csvs.py` — Multi-File Merger
Merges all CSV/Excel files in a folder into one clean dataset. Handles encoding issues, inconsistent columns, and duplicate rows.

```bash
# Generate sample data to try it
python merge_csvs.py --generate-sample

# Merge files
python merge_csvs.py --input ./raw_data --output merged.xlsx
```

### 2. `auto_pivot.py` — Pivot Table Generator
Reads any CSV/Excel and generates pivot tables with summary stats, top performers, and multi-sheet Excel output.

```bash
# Generate sample sales data
python auto_pivot.py --generate-sample

# Create pivot
python auto_pivot.py --input sample_sales.csv --rows Region Product --cols Quarter --values Revenue
```

### 3. `anomaly_detector.py` — Data Quality Checker
Scans any dataset for missing values, outliers, duplicates, type mismatches, and whitespace issues. Outputs a quality score and actionable report.

```bash
python anomaly_detector.py --input data.csv --output quality_report.xlsx
```

### 4. `format_report.py` — Report Formatter
Transforms raw data into a polished Excel report with KPI cards, conditional formatting, auto-adjusted widths, and pivot-ready aggregation.

```bash
python format_report.py --input raw_data.csv --title "Q3 Sales Report"
```

## Installation

```bash
pip install pandas openpyxl numpy
```

## Why I Built This

These scripts represent the kind of work I actually do. At every company I've worked at — Hero MotoCorp, SKF India, Sedris Vision, smart Europe — a significant chunk of time goes into preparing, cleaning, and formatting data in Excel. 

The manual version of these tasks isn't hard. It's just slow, error-prone, and nobody should have to do it every month. Automating them saved my team ~37 hours per month at Hero MotoCorp.

## Author

**Aishwarya Sonawane**  
[Portfolio](https://aishwaryasanjaysonawane-a11y.github.io/Portfolio/) · [LinkedIn](www.linkedin.com/in/aishwarya-s18)
