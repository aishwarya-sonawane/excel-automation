"""
Script 4: Auto Report Formatter
================================
Takes raw data and generates a formatted Excel report with:
- Summary sheet with KPIs
- Formatted data table with conditional formatting
- Auto-adjusted column widths
- Chart-ready data sheet

Usage:
    python format_report.py --input raw_data.csv --title "Q3 Sales Report"
"""

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, numbers
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
import argparse


def format_report(input_file: str, title: str = "Report", output_file: str = None):
    """
    Transform raw data into a polished, presentation-ready Excel report.
    """
    # Load
    if input_file.endswith('.csv'):
        df = pd.read_csv(input_file)
    else:
        df = pd.read_excel(input_file)
    
    if output_file is None:
        output_file = input_file.rsplit('.', 1)[0] + '_formatted.xlsx'
    
    wb = Workbook()
    
    # Styles
    BLUE = "2B579A"
    GOLD = "CFA64E"
    LIGHT_BLUE = "D6E4F0"
    LIGHT_GRAY = "F2F2F2"
    WHITE = "FFFFFF"
    
    title_font = Font(name="Calibri", size=16, bold=True, color=BLUE)
    header_font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    header_fill = PatternFill(start_color=BLUE, end_color=BLUE, fill_type="solid")
    alt_fill = PatternFill(start_color=LIGHT_GRAY, end_color=LIGHT_GRAY, fill_type="solid")
    kpi_font = Font(name="Calibri", size=24, bold=True, color=BLUE)
    kpi_label_font = Font(name="Calibri", size=9, color="888888")
    thin_border = Border(
        bottom=Side(style='thin', color='CCCCCC')
    )
    
    # ============================================================
    # SHEET 1: SUMMARY
    # ============================================================
    ws1 = wb.active
    ws1.title = "Summary"
    ws1.sheet_properties.tabColor = BLUE
    
    # Title
    ws1['A1'] = title
    ws1['A1'].font = title_font
    ws1['A2'] = f"Generated automatically · {len(df):,} records · {len(df.columns)} fields"
    ws1['A2'].font = Font(size=10, color="888888", italic=True)
    
    # KPIs
    numeric_cols = df.select_dtypes(include='number').columns.tolist()
    kpi_row = 5
    
    for i, col in enumerate(numeric_cols[:6]):  # Max 6 KPIs
        cell_col = get_column_letter(i * 3 + 1)
        
        ws1[f'{cell_col}{kpi_row}'] = col
        ws1[f'{cell_col}{kpi_row}'].font = kpi_label_font
        
        ws1[f'{cell_col}{kpi_row+1}'] = round(df[col].sum(), 1)
        ws1[f'{cell_col}{kpi_row+1}'].font = kpi_font
        ws1[f'{cell_col}{kpi_row+1}'].number_format = '#,##0.0'
        
        ws1[f'{cell_col}{kpi_row+2}'] = f"avg: {df[col].mean():.1f} · max: {df[col].max():.1f}"
        ws1[f'{cell_col}{kpi_row+2}'].font = Font(size=9, color="888888")
    
    # Column summary table
    summary_row = kpi_row + 5
    ws1[f'A{summary_row}'] = "Column Overview"
    ws1[f'A{summary_row}'].font = Font(size=12, bold=True, color=BLUE)
    
    summary_headers = ["Column", "Type", "Non-Null", "Unique", "Sample Value"]
    for i, h in enumerate(summary_headers):
        cell = ws1.cell(row=summary_row+1, column=i+1, value=h)
        cell.font = header_font
        cell.fill = header_fill
    
    for j, col in enumerate(df.columns):
        row = summary_row + 2 + j
        ws1.cell(row=row, column=1, value=col)
        ws1.cell(row=row, column=2, value=str(df[col].dtype))
        ws1.cell(row=row, column=3, value=int(df[col].notna().sum()))
        ws1.cell(row=row, column=4, value=int(df[col].nunique()))
        ws1.cell(row=row, column=5, value=str(df[col].dropna().iloc[0]) if not df[col].dropna().empty else "")
        
        if j % 2 == 1:
            for k in range(1, 6):
                ws1.cell(row=row, column=k).fill = alt_fill
    
    # ============================================================
    # SHEET 2: FORMATTED DATA
    # ============================================================
    ws2 = wb.create_sheet("Data")
    ws2.sheet_properties.tabColor = GOLD
    
    # Write data
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True)):
        for c_idx, value in enumerate(row, 1):
            cell = ws2.cell(row=r_idx + 1, column=c_idx, value=value)
            
            if r_idx == 0:  # Header
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal='center')
            else:
                cell.border = thin_border
                cell.font = Font(name="Calibri", size=10)
                
                # Alternate row shading
                if r_idx % 2 == 0:
                    cell.fill = alt_fill
                
                # Number formatting
                if isinstance(value, float):
                    cell.number_format = '#,##0.00'
                elif isinstance(value, int):
                    cell.number_format = '#,##0'
    
    # Auto-adjust column widths
    for col_idx in range(1, len(df.columns) + 1):
        max_length = max(
            len(str(ws2.cell(row=1, column=col_idx).value or "")),
            max((len(str(ws2.cell(row=r, column=col_idx).value or "")) for r in range(2, min(len(df)+2, 102))), default=0)
        )
        ws2.column_dimensions[get_column_letter(col_idx)].width = min(max_length + 4, 30)
    
    # Freeze header row
    ws2.freeze_panes = 'A2'
    # Auto-filter
    ws2.auto_filter.ref = f"A1:{get_column_letter(len(df.columns))}{len(df)+1}"
    
    # ============================================================
    # SHEET 3: PIVOT READY
    # ============================================================
    if len(numeric_cols) > 0:
        cat_cols = df.select_dtypes(include='object').columns.tolist()
        if cat_cols:
            ws3 = wb.create_sheet("Pivot Data")
            ws3.sheet_properties.tabColor = "4ADE80"
            
            pivot = df.groupby(cat_cols[0])[numeric_cols[0]].agg(['sum','mean','count']).round(1)
            pivot.columns = [f'{numeric_cols[0]}_{c}' for c in pivot.columns]
            pivot = pivot.sort_values(f'{numeric_cols[0]}_sum', ascending=False).reset_index()
            
            for r_idx, row in enumerate(dataframe_to_rows(pivot, index=False, header=True)):
                for c_idx, value in enumerate(row, 1):
                    cell = ws3.cell(row=r_idx + 1, column=c_idx, value=value)
                    if r_idx == 0:
                        cell.font = header_font
                        cell.fill = header_fill
    
    # Save
    wb.save(output_file)
    
    print(f"\n{'='*50}")
    print(f"✅ FORMATTED REPORT GENERATED")
    print(f"{'='*50}")
    print(f"  Title:   {title}")
    print(f"  Records: {len(df):,}")
    print(f"  Sheets:  Summary, Data, Pivot Data")
    print(f"  Output:  {output_file}")
    print(f"\n  Features applied:")
    print(f"    ✓ KPI cards on summary sheet")
    print(f"    ✓ Column headers with blue fill")
    print(f"    ✓ Alternating row shading")
    print(f"    ✓ Number formatting")
    print(f"    ✓ Auto-adjusted column widths")
    print(f"    ✓ Frozen header row")
    print(f"    ✓ Auto-filter enabled")
    print(f"    ✓ Pivot-ready aggregation sheet")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate formatted Excel reports from raw data")
    parser.add_argument("--input", "-i", required=True, help="Input CSV or Excel file")
    parser.add_argument("--title", "-t", default="Data Report", help="Report title")
    parser.add_argument("--output", "-o", help="Output Excel path")
    
    args = parser.parse_args()
    format_report(args.input, args.title, args.output)
