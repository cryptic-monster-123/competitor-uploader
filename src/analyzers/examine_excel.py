import pandas as pd
import sys

def examine_excel(file_path):
    """Examine an Excel file and print its structure."""
    print(f"Examining file: {file_path}")
    
    # Get sheet names
    excel_file = pd.ExcelFile(file_path)
    sheet_names = excel_file.sheet_names
    print(f"Sheets in file: {sheet_names}")
    
    # Examine each sheet
    for sheet in sheet_names:
        print(f"\nSheet: {sheet}")
        df = pd.read_excel(file_path, sheet_name=sheet)
        print(f"Shape: {df.shape}")
        print(f"Columns: {df.columns.tolist()}")
        
        # Check for non-null values in key columns
        print("\nNon-null counts for key columns:")
        for col in df.columns:
            non_null = df[col].notna().sum()
            print(f"  {col}: {non_null} non-null values ({non_null/len(df)*100:.1f}%)")
        
        # Print first few rows
        print("\nFirst 3 rows:")
        print(df.head(3))

if __name__ == "__main__":
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        file_path = "20240101/January2025 Store_Data.xlsx"
    
    examine_excel(file_path)
