import pandas as pd

# Read the CSV file
df = pd.read_csv('combined_data.csv')

# Print basic information
print(f"Total rows: {len(df)}")
print(f"Total columns: {len(df.columns)}")
print(f"Column names: {df.columns.tolist()}")

# Check for non-null values in key columns
print("\nNon-null values in key columns:")
for col in ['POS Name', 'Brand', 'Store Tagging', 'Territory', 'TSM']:
    print(f"  {col}: {df[col].notna().sum()} ({df[col].notna().sum() / len(df) * 100:.1f}%)")

# Check source files
print("\nRows per source file:")
source_counts = df['Source File'].value_counts()
for source, count in source_counts.items():
    print(f"  {source}: {count} rows")

# Check for numeric columns
print("\nNumeric columns statistics:")
numeric_cols = ['Retail Sales', 'Tonik Sales', 'HC Sales', 'Skyro Sales', 
                'Retailer Headcount', 'Tonik Headcount', 'HC Headcount', 'Skyro Headcount']
for col in numeric_cols:
    non_null = df[col].notna().sum()
    if non_null > 0:
        print(f"  {col}: {non_null} non-null values, Mean: {df[col].mean():.2f}")
    else:
        print(f"  {col}: {non_null} non-null values")
