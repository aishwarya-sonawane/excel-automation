"""
Script 2: Auto Pivot Table Generator
=====================================
Reads any CSV/Excel file and automatically generates pivot tables
with summary statistics, formatting, and chart-ready output.

Usage:
    python auto_pivot.py --input sales_data.csv --rows Region --cols Quarter --values Revenue
"""

import pandas as pd
import argparse
import sys


def auto_pivot(input_file: str, row_fields: list, col_field: str = None,
               value_field: str = None, agg_func: str = "sum",
               output_file: str = None):
    """
    Generate formatted pivot tables from any tabular data.
    
    Args:
        input_file: Path to CSV or Excel file
        row_fields: Column(s) for pivot rows
        col_field: Column for pivot columns (optional)
        value_field: Column to aggregate
        agg_func: Aggregation function (sum, mean, count, min, max)
        output_file: Output Excel path (auto-generated if None)
    """
    # Load data
    if input_file.endswith('.csv'):
        df = pd.read_csv(input_file)
    else:
        df = pd.read_excel(input_file)
    
    print(f"📊 Loaded {len(df):,} rows × {len(df.columns)} columns")
    print(f"   Columns: {', '.join(df.columns.tolist())}\n")
    
    # Auto-detect value field if not specified
    if value_field is None:
        numeric_cols = df.select_dtypes(include='number').columns.tolist()
        if numeric_cols:
            value_field = numeric_cols[0]
            print(f"   Auto-selected value field: {value_field}")
        else:
            print("❌ No numeric columns found. Specify --values explicitly.")
            return
    
    # Build pivot
    pivot = pd.pivot_table(
        df,
        index=row_fields,
        columns=col_field,
        values=value_field,
        aggfunc=agg_func,
        margins=True,
        margins_name="Total"
    )
    
    # Round
    pivot = pivot.round(1)
    
    # Display
    print(f"\n{'='*60}")
    print(f"PIVOT TABLE: {value_field} by {', '.join(row_fields)}")
    if col_field:
        print(f"Columns: {col_field} | Aggregation: {agg_func}")
    print(f"{'='*60}\n")
    print(pivot.to_string())
    
    # Summary stats
    print(f"\n{'='*60}")
    print("SUMMARY STATISTICS")
    print(f"{'='*60}")
    
    # Exclude the 'Total' row for stats
    data_only = pivot.iloc[:-1] if "Total" in pivot.index else pivot
    if col_field:
        # Exclude 'Total' column too
        data_cols = [c for c in data_only.columns if c != "Total"]
        flat_values = data_only[data_cols].values.flatten()
    else:
        flat_values = data_only.values.flatten()
    
    flat_values = pd.Series(flat_values).dropna()
    
    print(f"  Count:    {len(flat_values)}")
    print(f"  Mean:     {flat_values.mean():.1f}")
    print(f"  Median:   {flat_values.median():.1f}")
    print(f"  Std Dev:  {flat_values.std():.1f}")
    print(f"  Min:      {flat_values.min():.1f}")
    print(f"  Max:      {flat_values.max():.1f}")
    
    # Top performers
    if len(row_fields) == 1:
        top_col = "Total" if "Total" in pivot.columns else pivot.columns[0]
        if isinstance(top_col, str):
            sorted_data = pivot.iloc[:-1].sort_values(top_col, ascending=False) if "Total" in pivot.index else pivot.sort_values(top_col, ascending=False)
            print(f"\n  Top 5 by {value_field}:")
            for i, (idx, row) in enumerate(sorted_data.head(5).iterrows()):
                val = row[top_col] if isinstance(row[top_col], (int, float)) else row.iloc[0]
                print(f"    {i+1}. {idx}: {val:,.1f}")
    
    # Save to Excel
    if output_file is None:
        output_file = input_file.rsplit('.', 1)[0] + '_pivot.xlsx'
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        pivot.to_excel(writer, sheet_name="Pivot Table")
        df.describe().to_excel(writer, sheet_name="Statistics")
        df.to_excel(writer, sheet_name="Raw Data", index=False)
    
    print(f"\n✅ Saved to {output_file}")
    print(f"   Sheets: Pivot Table, Statistics, Raw Data")
    
    return pivot


def generate_sample():
    """Generate sample sales data for testing."""
    import random
    random.seed(42)
    
    products = ["Widget A", "Widget B", "Widget C", "Widget D"]
    regions = ["EMEA", "APAC", "Americas", "DACH"]
    quarters = ["Q1", "Q2", "Q3", "Q4"]
    
    rows = []
    for _ in range(200):
        rows.append({
            "Product": random.choice(products),
            "Region": random.choice(regions),
            "Quarter": random.choice(quarters),
            "Revenue": round(random.uniform(5000, 50000), 2),
            "Units": random.randint(10, 500),
            "Margin_pct": round(random.uniform(15, 45), 1)
        })
    
    df = pd.DataFrame(rows)
    df.to_csv("sample_sales.csv", index=False)
    print("✅ Generated sample_sales.csv (200 rows)")
    return df


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Auto-generate pivot tables from any data file")
    parser.add_argument("--input", "-i", help="Input CSV or Excel file")
    parser.add_argument("--rows", "-r", nargs="+", required=False, help="Row field(s)")
    parser.add_argument("--cols", "-c", help="Column field (optional)")
    parser.add_argument("--values", "-v", help="Value field to aggregate")
    parser.add_argument("--agg", default="sum", choices=["sum","mean","count","min","max"])
    parser.add_argument("--output", "-o", help="Output Excel path")
    parser.add_argument("--generate-sample", action="store_true")
    
    args = parser.parse_args()
    
    if args.generate_sample:
        generate_sample()
    elif args.input and args.rows:
        auto_pivot(args.input, args.rows, args.cols, args.values, args.agg, args.output)
    else:
        print("Usage: python auto_pivot.py --input data.csv --rows Region Product --cols Quarter --values Revenue")
        print("       python auto_pivot.py --generate-sample")
