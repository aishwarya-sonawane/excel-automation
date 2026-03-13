"""
Script 1: Multi-CSV Merger & Cleaner
=====================================
Merges multiple CSV/Excel files from a folder into one clean dataset.
Handles inconsistent columns, duplicate headers, and encoding issues.

Usage:
    python merge_csvs.py --input ./raw_data --output merged_report.xlsx
"""

import pandas as pd
import os
import argparse
from pathlib import Path


def merge_files(input_dir: str, output_file: str, deduplicate: bool = True):
    """
    Merge all CSV/Excel files in a directory into a single clean dataset.
    
    Args:
        input_dir: Path to folder containing CSV/XLSX files
        output_file: Path for the merged output file
        deduplicate: Whether to remove duplicate rows
    """
    input_path = Path(input_dir)
    all_frames = []
    file_log = []
    
    supported_extensions = {'.csv', '.xlsx', '.xls', '.tsv'}
    files = [f for f in input_path.iterdir() if f.suffix.lower() in supported_extensions]
    
    if not files:
        print(f"❌ No supported files found in {input_dir}")
        return
    
    print(f"📁 Found {len(files)} files to merge\n")
    
    for file in sorted(files):
        try:
            if file.suffix.lower() == '.csv':
                # Try UTF-8 first, fall back to latin-1
                try:
                    df = pd.read_csv(file, encoding='utf-8')
                except UnicodeDecodeError:
                    df = pd.read_csv(file, encoding='latin-1')
            elif file.suffix.lower() == '.tsv':
                df = pd.read_csv(file, sep='\t')
            else:
                df = pd.read_excel(file)
            
            # Clean column names
            df.columns = (df.columns.str.strip()
                         .str.lower()
                         .str.replace(r'\s+', '_', regex=True)
                         .str.replace(r'[^\w]', '_', regex=True))
            
            # Remove completely empty rows
            df = df.dropna(how='all')
            
            # Add source file column
            df['_source_file'] = file.name
            
            all_frames.append(df)
            file_log.append({"file": file.name, "rows": len(df), "columns": len(df.columns), "status": "✅"})
            print(f"  ✅ {file.name}: {len(df)} rows, {len(df.columns)} cols")
            
        except Exception as e:
            file_log.append({"file": file.name, "rows": 0, "columns": 0, "status": f"❌ {str(e)[:50]}"})
            print(f"  ❌ {file.name}: {e}")
    
    if not all_frames:
        print("\n❌ No files could be read successfully.")
        return
    
    # Merge all dataframes
    merged = pd.concat(all_frames, ignore_index=True, sort=False)
    
    # Deduplicate
    initial_rows = len(merged)
    if deduplicate:
        # Exclude _source_file from dedup check
        data_cols = [c for c in merged.columns if c != '_source_file']
        merged = merged.drop_duplicates(subset=data_cols)
    
    dupes_removed = initial_rows - len(merged)
    
    # Save
    if output_file.endswith('.xlsx'):
        merged.to_excel(output_file, index=False, engine='openpyxl')
    else:
        merged.to_csv(output_file, index=False)
    
    # Summary
    print(f"\n{'='*50}")
    print(f"📊 MERGE SUMMARY")
    print(f"{'='*50}")
    print(f"  Files processed:  {len(file_log)}")
    print(f"  Successful:       {sum(1 for f in file_log if '✅' in f['status'])}")
    print(f"  Failed:           {sum(1 for f in file_log if '❌' in f['status'])}")
    print(f"  Total rows:       {len(merged):,}")
    print(f"  Duplicates removed: {dupes_removed:,}")
    print(f"  Output columns:   {len(merged.columns)}")
    print(f"  Saved to:         {output_file}")
    
    return merged


def generate_sample_data(output_dir: str = "./sample_data"):
    """Generate sample CSV files for testing."""
    os.makedirs(output_dir, exist_ok=True)
    
    # Sales Q1
    pd.DataFrame({
        "Date": ["2025-01-15", "2025-02-10", "2025-03-20"],
        "Product": ["Widget A", "Widget B", "Widget A"],
        "Region": ["EMEA", "APAC", "Americas"],
        "Revenue": [15000, 22000, 18500],
        "Units": [150, 180, 200]
    }).to_csv(f"{output_dir}/sales_q1.csv", index=False)
    
    # Sales Q2
    pd.DataFrame({
        "Date": ["2025-04-12", "2025-05-08", "2025-06-25"],
        "Product": ["Widget B", "Widget C", "Widget A"],
        "Region": ["EMEA", "EMEA", "APAC"],
        "Revenue": [25000, 12000, 30000],
        "Units": [200, 100, 280]
    }).to_csv(f"{output_dir}/sales_q2.csv", index=False)
    
    # Sales Q3 (slightly different columns to test robustness)
    pd.DataFrame({
        "Date": ["2025-07-05", "2025-08-18", "2025-09-30"],
        "Product": ["Widget A", "Widget C", "Widget B"],
        "Region": ["Americas", "APAC", "EMEA"],
        " Revenue ": [28000, 16000, 21000],  # note: extra spaces
        "Units": [250, 140, 175]
    }).to_csv(f"{output_dir}/sales_q3.csv", index=False)
    
    print(f"✅ Sample data generated in {output_dir}/")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Merge multiple CSV/Excel files into one clean dataset")
    parser.add_argument("--input", "-i", default="./sample_data", help="Input directory with CSV/XLSX files")
    parser.add_argument("--output", "-o", default="merged_report.xlsx", help="Output file path")
    parser.add_argument("--generate-sample", action="store_true", help="Generate sample data for testing")
    parser.add_argument("--no-dedup", action="store_true", help="Skip deduplication")
    
    args = parser.parse_args()
    
    if args.generate_sample:
        generate_sample_data()
    
    merge_files(args.input, args.output, deduplicate=not args.no_dedup)
