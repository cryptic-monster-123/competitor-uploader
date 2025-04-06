import os
import pandas as pd
import argparse
from pathlib import Path

def concatenate_excel_files(folder_path, output_path, file_pattern=None, column_mapping=None):
    """
    Concatenate all Excel files in the specified folder with column mapping.
    
    Parameters:
    -----------
    folder_path : str
        Path to the folder containing Excel files
    output_path : str
        Path where the concatenated Excel file will be saved
    file_pattern : str, optional
        Pattern to match specific Excel files (e.g., '*.xlsx')
    column_mapping : dict, optional
        Dictionary mapping standard column names to possible variations in the files
    
    Returns:
    --------
    bool
        True if concatenation was successful, False otherwise
    """
    # Convert to Path object for better path handling
    folder = Path(folder_path)
    
    # Get all Excel files in the folder
    if file_pattern:
        excel_files = list(folder.glob(file_pattern))
    else:
        # Default to common Excel extensions, exclude temp files
        excel_files = []
        for pattern in ['*.xlsx', '*.xls', '*.xlsm']:
            excel_files.extend([f for f in folder.glob(pattern) if not f.name.startswith('~$')])
    
    if not excel_files:
        print(f"No Excel files found in {folder_path}")
        return False
    
    print(f"Found {len(excel_files)} Excel files")
    
    # Define standard column names if not provided
    if column_mapping is None:
        column_mapping = {
            'POS Name': ['POS Name', 'Store Name', 'STORE DATA JANUARY 2025', 'STORE DATA AS OF JANUARY 31'],
            'Retailer': ['Retailer', 'Retailer Name'],
            'Territory': ['Territory', 'Store Tagging Territory'],
            'TSM': ['TSM'],
            'Retail Sales': ['Retail Sales'],
            'Tonik Sales': ['Tonik Sales'],
            'NC Sales': ['NC Sales', 'HC Sales'],
            'Skyro Sales': ['Skyro Sales'],
            'Salmon': ['Salmon'],
            'In-House': ['In-House'],
            'Credit card': ['Credit card'],
            'Cash Sales': ['Cash Sales'],
            'Others': ['Others'],
            'Retailer Headcount': ['Retailer Headcount'],
            'Tonik Headcount': ['Tonik Headcount', 'Tonik Promoter Count'],
            'NC Headcount': ['NC Headcount', 'HC Promoter Count'],
            'Skyro Headcount': ['Skyro Headcount'],
            'Salmon Headcount': ['Salmon Headcount'],
            'Remarks': ['Remarks']
        }
    
    # Initialize an empty list to store DataFrames
    dfs = []
    
    # Read each Excel file and append to the list
    for file in excel_files:
        try:
            print(f"Reading {file.name}...")
            # Read the Excel file
            df = pd.read_excel(file)
            
            # Print original columns for debugging
            print(f"  - Original columns: {', '.join(str(col) for col in df.columns)}")
            
            # Create a new DataFrame with standardized columns
            new_df = pd.DataFrame()
            
            # Map columns based on the column_mapping dictionary
            for std_col, possible_cols in column_mapping.items():
                if isinstance(possible_cols, dict):
                    # Handle nested mappings (e.g., for categorized columns)
                    for sub_col, sub_possible_cols in possible_cols.items():
                        nested_col_name = f"{std_col} - {sub_col}"
                        # Find the first matching column in the DataFrame
                        matching_col = next((col for col in sub_possible_cols if col in df.columns), None)
                        
                        if matching_col:
                            new_df[nested_col_name] = df[matching_col]
                        else:
                            # If no matching column is found, add an empty column
                            new_df[nested_col_name] = None
                else:
                    # Handle regular mappings
                    # Find the first matching column in the DataFrame
                    matching_col = next((col for col in possible_cols if col in df.columns), None)
                    
                    if matching_col:
                        new_df[std_col] = df[matching_col]
                    else:
                        # If no matching column is found, add an empty column
                        new_df[std_col] = None
            
            # Add the standardized DataFrame to the list
            dfs.append(new_df)
            print(f"  - Standardized columns: {', '.join(new_df.columns)}")
            
        except Exception as e:
            print(f"Error reading {file.name}: {str(e)}")
    
    if not dfs:
        print("No valid Excel files could be read")
        return False
    
    # Concatenate all DataFrames
    try:
        combined_df = pd.concat(dfs, ignore_index=True)
        print(f"Combined data shape: {combined_df.shape}")
        
        # Create output directory if it doesn't exist
        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        
        # Save the concatenated DataFrame to Excel
        combined_df.to_excel(output_path, index=False)
        print(f"Concatenated data saved to {output_path}")
        return True
    except Exception as e:
        print(f"Error during concatenation or saving: {str(e)}")
        return False

def main():
    # Set up argument parser
    parser = argparse.ArgumentParser(description='Concatenate Excel files with column mapping.')
    parser.add_argument('folder', help='Folder containing Excel files')
    parser.add_argument('-o', '--output', default='combined_data.xlsx', 
                        help='Output file path (default: combined_data.xlsx in current directory)')
    parser.add_argument('-p', '--pattern', help='File pattern to match (e.g., "data_*.xlsx")')
    parser.add_argument('-m', '--mapping', help='JSON file containing column mapping')
    
    # Parse arguments
    args = parser.parse_args()
    
    # Load custom column mapping if provided
    column_mapping = None
    if args.mapping:
        try:
            with open(args.mapping, 'r') as f:
                import json
                column_mapping = json.load(f)
            print(f"Loaded custom column mapping from {args.mapping}")
        except Exception as e:
            print(f"Error loading column mapping: {str(e)}")
            return False
    
    # Call the concatenation function
    success = concatenate_excel_files(args.folder, args.output, args.pattern, column_mapping)
    
    if success:
        print("Concatenation completed successfully!")
    else:
        print("Concatenation failed.")

if __name__ == "__main__":
    main()
