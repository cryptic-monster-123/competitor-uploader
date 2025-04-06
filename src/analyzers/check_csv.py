import pandas as pd
import sys

def check_csv(file_path):
    """Check the CSV file for dash characters and print summary."""
    print(f"Checking file: {file_path}")
    
    # Read the CSV file
    df = pd.read_csv(file_path)
    
    # Print basic information
    print(f"Total rows: {len(df)}")
    print(f"Total columns: {len(df.columns)}")
    print(f"Columns: {df.columns.tolist()}")
    
    # Check for standalone dash characters (cells that contain only a dash)
    standalone_dash_columns = []
    for col in df.columns:
        if df[col].dtype == 'object':  # Only check string columns
            # Check for cells that contain only a dash
            has_standalone_dash = df[col].astype(str).str.match(r'^-+$').any()
            if has_standalone_dash:
                standalone_dash_columns.append(col)
                # Print the rows with standalone dashes
                dash_rows = df[df[col].astype(str).str.match(r'^-+$')]
                print(f"\nRows with standalone dash in column '{col}':")
                print(dash_rows[['POS Name', col, 'Source File']].head())
                
    if standalone_dash_columns:
        print(f"\nColumns with standalone dash characters: {standalone_dash_columns}")
    else:
        print("\nNo standalone dash characters found in any column!")
        
    # Check numeric columns for NaN values that might have been dash characters
    numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
    print(f"\nNumeric columns: {numeric_cols}")
    
    # Count NaN values in numeric columns
    nan_counts = df[numeric_cols].isna().sum()
    print("\nNaN counts in numeric columns:")
    for col, count in nan_counts.items():
        if count > 0:
            print(f"  {col}: {count} NaN values ({count/len(df)*100:.1f}%)")
    
    # Print first few rows
    print("\nSample of first 3 rows:")
    print(df.head(3))

if __name__ == "__main__":
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        # Find the most recent combined_data file
        import glob
        import os
        files = glob.glob("combined_data_*.csv")
        if files:
            file_path = max(files, key=os.path.getctime)
        else:
            print("No combined_data_*.csv files found")
            sys.exit(1)
    
    check_csv(file_path)
