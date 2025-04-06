import os
import pandas as pd
import argparse
from pathlib import Path

def concatenate_excel_files(folder_path, output_path, file_pattern=None):
    """
    Concatenate all Excel files in the specified folder that have the same column structure.
    
    Parameters:
    -----------
    folder_path : str
        Path to the folder containing Excel files
    output_path : str
        Path where the concatenated Excel file will be saved
    file_pattern : str, optional
        Pattern to match specific Excel files (e.g., '*.xlsx')
    
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
        # Default to common Excel extensions
        excel_files = list(folder.glob('*.xlsx')) + list(folder.glob('*.xls')) + list(folder.glob('*.xlsm'))
    
    if not excel_files:
        print(f"No Excel files found in {folder_path}")
        return False
    
    # Initialize an empty list to store DataFrames
    dfs = []
    
    # Read each Excel file and append to the list
    for file in excel_files:
        try:
            df = pd.read_excel(file)
            dfs.append(df)
        except Exception as e:
            print(f"Error reading {file.name}: {str(e)}")
    
    if not dfs:
        print("No valid Excel files could be read")
        return False
    
    # Concatenate all DataFrames
    try:
        combined_df = pd.concat(dfs, ignore_index=True)
        
        # Create output directory if it doesn't exist
        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        
        # Ensure the output file has a proper extension
        if not output_path.lower().endswith(('.xlsx', '.xls')):
            output_path = output_path + '.xlsx'
            output_file = Path(output_path)
        
        # Save the concatenated DataFrame to Excel
        combined_df.to_excel(output_path, index=False)
        return True
    except Exception as e:
        print(f"Error during concatenation or saving: {str(e)}")
        return False

def main():
    # Set up argument parser
    parser = argparse.ArgumentParser(description='Concatenate Excel files with the same column structure.')
    parser.add_argument('folder', help='Folder containing Excel files')
    parser.add_argument('-o', '--output', default='combined_data.xlsx', 
                        help='Output file path (default: combined_data.xlsx in current directory)')
    parser.add_argument('-p', '--pattern', help='File pattern to match (e.g., "data_*.xlsx")')
    
    # Parse arguments
    args = parser.parse_args()
    
    # Call the concatenation function
    success = concatenate_excel_files(args.folder, args.output, args.pattern)
    
    if success:
        print("Concatenation completed successfully!")
    else:
        print("Concatenation failed.")

if __name__ == "__main__":
    main()
