"""
Script 3: Data Anomaly Detector
================================
Scans any dataset for common data quality issues:
missing values, outliers, duplicates, format inconsistencies.
Generates a quality report with recommendations.

Usage:
    python anomaly_detector.py --input data.csv --output quality_report.xlsx
"""

import pandas as pd
import numpy as np
import argparse
from pathlib import Path


def detect_anomalies(input_file: str, output_file: str = None, z_threshold: float = 3.0):
    """
    Comprehensive data quality check on any tabular dataset.
    
    Args:
        input_file: Path to CSV or Excel file
        output_file: Path for quality report
        z_threshold: Z-score threshold for outlier detection
    """
    # Load
    if input_file.endswith('.csv'):
        df = pd.read_csv(input_file)
    else:
        df = pd.read_excel(input_file)
    
    total_rows = len(df)
    total_cols = len(df.columns)
    total_cells = total_rows * total_cols
    
    print(f"\n{'='*60}")
    print(f"📋 DATA QUALITY REPORT")
    print(f"{'='*60}")
    print(f"  File:     {input_file}")
    print(f"  Rows:     {total_rows:,}")
    print(f"  Columns:  {total_cols}")
    print(f"  Cells:    {total_cells:,}")
    
    issues = []
    
    # ---- 1. MISSING VALUES ----
    print(f"\n--- Missing Values ---")
    missing = df.isnull().sum()
    missing_pct = (missing / total_rows * 100).round(1)
    has_missing = missing[missing > 0]
    
    if len(has_missing) > 0:
        for col in has_missing.index:
            severity = "🔴 Critical" if missing_pct[col] > 20 else "🟡 Warning" if missing_pct[col] > 5 else "🟢 Minor"
            print(f"  {severity} {col}: {has_missing[col]} missing ({missing_pct[col]}%)")
            issues.append({
                "Type": "Missing Value",
                "Column": col,
                "Count": int(has_missing[col]),
                "Percentage": f"{missing_pct[col]}%",
                "Severity": severity.split()[0],
                "Recommendation": f"{'Drop column' if missing_pct[col] > 50 else 'Impute or investigate'}"
            })
    else:
        print("  ✅ No missing values found")
    
    # ---- 2. DUPLICATES ----
    print(f"\n--- Duplicate Rows ---")
    dupes = df.duplicated().sum()
    if dupes > 0:
        print(f"  🟡 {dupes} duplicate rows ({dupes/total_rows*100:.1f}%)")
        issues.append({
            "Type": "Duplicate Rows", "Column": "ALL",
            "Count": int(dupes), "Percentage": f"{dupes/total_rows*100:.1f}%",
            "Severity": "🟡", "Recommendation": "Review and remove duplicates"
        })
    else:
        print("  ✅ No duplicate rows")
    
    # ---- 3. OUTLIERS (numeric columns) ----
    print(f"\n--- Outliers (Z-score > {z_threshold}) ---")
    numeric_cols = df.select_dtypes(include='number').columns
    
    for col in numeric_cols:
        col_data = df[col].dropna()
        if len(col_data) < 10:
            continue
        
        z_scores = np.abs((col_data - col_data.mean()) / col_data.std())
        outlier_count = (z_scores > z_threshold).sum()
        
        if outlier_count > 0:
            outlier_values = col_data[z_scores > z_threshold]
            print(f"  🟡 {col}: {outlier_count} outliers (range: {outlier_values.min():.1f} to {outlier_values.max():.1f})")
            issues.append({
                "Type": "Outlier", "Column": col,
                "Count": int(outlier_count), "Percentage": f"{outlier_count/len(col_data)*100:.1f}%",
                "Severity": "🟡",
                "Recommendation": f"Values outside {col_data.mean():.1f} ± {z_threshold}×{col_data.std():.1f}"
            })
    
    if not any(i["Type"] == "Outlier" for i in issues):
        print("  ✅ No significant outliers detected")
    
    # ---- 4. DATA TYPE ISSUES ----
    print(f"\n--- Data Type Check ---")
    for col in df.columns:
        if df[col].dtype == 'object':
            # Check if it should be numeric
            try:
                numeric_count = pd.to_numeric(df[col], errors='coerce').notna().sum()
                if numeric_count > len(df[col]) * 0.7:
                    non_numeric = len(df[col]) - numeric_count
                    print(f"  🟡 {col}: looks numeric but has {non_numeric} non-numeric values")
                    issues.append({
                        "Type": "Type Mismatch", "Column": col,
                        "Count": int(non_numeric), "Percentage": f"{non_numeric/total_rows*100:.1f}%",
                        "Severity": "🟡",
                        "Recommendation": "Convert to numeric, investigate non-numeric entries"
                    })
            except:
                pass
            
            # Check for leading/trailing whitespace
            if df[col].dtype == 'object':
                stripped = df[col].dropna().str.strip()
                whitespace_issues = (df[col].dropna() != stripped).sum()
                if whitespace_issues > 0:
                    print(f"  🟢 {col}: {whitespace_issues} values have extra whitespace")
                    issues.append({
                        "Type": "Whitespace", "Column": col,
                        "Count": int(whitespace_issues), "Percentage": f"{whitespace_issues/total_rows*100:.1f}%",
                        "Severity": "🟢",
                        "Recommendation": "Strip whitespace with .str.strip()"
                    })
    
    # ---- 5. CARDINALITY CHECK ----
    print(f"\n--- Cardinality ---")
    for col in df.select_dtypes(include='object').columns:
        nunique = df[col].nunique()
        if nunique == total_rows:
            print(f"  ℹ️  {col}: all unique values (possible ID column)")
        elif nunique == 1:
            print(f"  🟡 {col}: only 1 unique value — consider dropping")
            issues.append({
                "Type": "Low Cardinality", "Column": col,
                "Count": 1, "Percentage": "100%",
                "Severity": "🟡", "Recommendation": "Constant column, consider removing"
            })
    
    # ---- SUMMARY ----
    critical = sum(1 for i in issues if "🔴" in i["Severity"])
    warnings = sum(1 for i in issues if "🟡" in i["Severity"])
    minor = sum(1 for i in issues if "🟢" in i["Severity"])
    
    print(f"\n{'='*60}")
    print(f"📊 SUMMARY")
    print(f"{'='*60}")
    print(f"  🔴 Critical:  {critical}")
    print(f"  🟡 Warnings:  {warnings}")
    print(f"  🟢 Minor:     {minor}")
    
    score = max(0, 100 - critical * 20 - warnings * 5 - minor * 1)
    grade = "A" if score >= 90 else "B" if score >= 75 else "C" if score >= 60 else "D" if score >= 40 else "F"
    print(f"\n  Data Quality Score: {score}/100 (Grade: {grade})")
    
    # Save report
    if output_file is None:
        output_file = str(Path(input_file).stem) + "_quality_report.xlsx"
    
    issues_df = pd.DataFrame(issues) if issues else pd.DataFrame(columns=["Type","Column","Count","Percentage","Severity","Recommendation"])
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        issues_df.to_excel(writer, sheet_name="Issues", index=False)
        
        # Column summary
        col_summary = pd.DataFrame({
            "Column": df.columns,
            "Type": df.dtypes.astype(str),
            "Non-Null": df.notna().sum(),
            "Null": df.isnull().sum(),
            "Null %": (df.isnull().sum() / total_rows * 100).round(1),
            "Unique": df.nunique(),
        })
        col_summary.to_excel(writer, sheet_name="Column Summary", index=False)
        
        # Numeric stats
        if len(numeric_cols) > 0:
            df[numeric_cols].describe().round(2).to_excel(writer, sheet_name="Numeric Stats")
    
    print(f"\n✅ Report saved to {output_file}")
    
    return issues_df


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Detect data quality issues in any dataset")
    parser.add_argument("--input", "-i", required=True, help="Input CSV or Excel file")
    parser.add_argument("--output", "-o", help="Output report path")
    parser.add_argument("--z-threshold", type=float, default=3.0, help="Z-score threshold for outliers")
    
    args = parser.parse_args()
    detect_anomalies(args.input, args.output, args.z_threshold)
