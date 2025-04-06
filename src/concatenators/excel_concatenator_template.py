import os
import pandas as pd
import argparse
from pathlib import Path
import re

def concatenate_excel_files_with_template(folder_path, output_path, template_path, file_pattern=None):
    """
    Concatenate Excel files using a template file for column structure.
    
    Parameters:
    -----------
    folder_path : str
        Path to the folder containing Excel files to concatenate
    output_path : str
        Path where the concatenated Excel file will be saved
    template_path : str
        Path to the template Excel file that defines the column structure
    file_pattern : str, optional
        Pattern to match specific Excel files (e.g., '*.xlsx')
    
    Returns:
    --------
    bool
        True if concatenation was successful, False otherwise
    """
    # Create a log file for detailed debugging
    import datetime
    log_file = f"concatenator_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    
    def log_message(message):
        print(message)
        with open(log_file, "a") as f:
            f.write(f"{message}\n")
    # Convert to Path objects for better path handling
    folder = Path(folder_path)
    template = Path(template_path)
    
    # Read the template file to get the column structure
    try:
        template_df = pd.read_excel(template)
        # Strip whitespace from column names to avoid matching issues
        template_columns = [col.strip() if isinstance(col, str) else col for col in template_df.columns.tolist()]
    except Exception as e:
        print(f"Error reading template file: {str(e)}")
        return False
    
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
    
    # Print details about found files
    print(f"Found {len(excel_files)} Excel files:")
    for file in excel_files:
        print(f"  - {file.name}")
    
    # Create a column mapping based on the template columns
    # This maps template column names to possible variations in the input files
    column_mapping = {
        'POS Name': ['POS Name', 'Store Name', 'Store name',
                    'STORE DATA', 'STORE DATA AO JANUARY', 'Store', 'Store Location', 'Location', 'LOCATION',
                    'Branch', 'BRANCH', 'Branch Name', 'BRANCH NAME', 'STORE', 'STORE NAME', 'STORE DATA JANUARY'],
        'Retailer': ['Retailer', 'Retailer Name', 'Brand', 'Store Brand', 'Company', 'RETAILER', 'RETAILER NAME',
                    'BRAND', 'COMPANY', 'STORE BRAND', 'Store Company', 'STORE COMPANY'],
        'Territory': ['Territory', 'Store Tagging Territory', 'Region', 'Area', 'Location', 'Store Region',
                    'TERRITORY', 'REGION', 'AREA', 'STORE REGION', 'Store Territory', 'STORE TERRITORY'],
        'TSM': ['TSM', 'Manager', 'Store Manager', 'Territory Sales Manager', 'Sales Manager', 'MANAGER',
                'STORE MANAGER', 'TERRITORY SALES MANAGER', 'SALES MANAGER', 'Store TSM', 'STORE TSM'],
        'Retail Sales': ['Retail Sales', 'Sales', 'Total Sales', 'Retailer Sales', 'Store Sales', 'Gross Sales',
                        'RETAIL SALES', 'SALES', 'TOTAL SALES', 'RETAILER SALES', 'STORE SALES', 'GROSS SALES', 'Total Revenue'],
        'Tonik Sales': ['Tonik Sales', 'Brand A', 'Tonik', 'TONIK', 'Tonik Brand Sales', 'TONIK SALES',
                        'BRAND A', 'TONIK BRAND SALES', 'Brand A Sales', 'BRAND A SALES'],
        'HC Sales': ['HC Sales', 'NC Sales', 'Brand B', 'HC', 'HC Brand', 'HC Brand Sales', 'HC SALES',
                    'NC SALES', 'BRAND B', 'HC BRAND', 'HC BRAND SALES', 'Brand B Sales', 'BRAND B SALES'],
        'Skyro Sales': ['Skyro Sales', 'Brand C', 'Skyro', 'SKYRO', 'Skyro Brand Sales', 'SKYRO SALES',
                        'BRAND C', 'SKYRO BRAND SALES', 'Brand C Sales', 'BRAND C SALES'],
        'Salmon': ['Salmon', 'Brand D', 'Salmon Sales', 'SALMON', 'Salmon Brand', 'Salmon Brand Sales',
                    'SALMON SALES', 'BRAND D', 'SALMON BRAND', 'SALMON BRAND SALES', 'Brand D Sales', 'BRAND D SALES'],
        'In-House': ['In-House', 'Store Credit', 'In House', 'In House Sales', 'Store Credit Sales',
                    'IN-HOUSE', 'STORE CREDIT', 'IN HOUSE', 'IN HOUSE SALES', 'STORE CREDIT SALES'],
        'Credit card': ['Credit card', 'Credit Card Sales', 'Credit Card', 'CC Sales', 'Card Sales',
                        'CREDIT CARD', 'CREDIT CARD SALES', 'CC SALES', 'CARD SALES'],
        'Cash Sales': ['Cash Sales', 'Cash', 'Cash Payment', 'Cash Transactions',
                        'CASH SALES', 'CASH', 'CASH PAYMENT', 'CASH TRANSACTIONS'],
        'Others': ['Others', 'Other Payment', 'Other', 'Miscellaneous', 'Other Sales',
                    'OTHER PAYMENT', 'OTHER', 'MISCELLANEOUS', 'OTHER SALES'],
        'Retailer Headcount': ['Retailer Headcount', 'Store Staff', 'Staff Count', 'Employee Count', 'Store Employees',
                                'RETAILER HEADCOUNT', 'STORE STAFF', 'STAFF COUNT', 'EMPLOYEE COUNT', 'STORE EMPLOYEES'],
        'Tonik Headcount': ['Tonik Headcount', 'Tonik Promoter Count', 'Brand A Staff', 'Tonik Staff', 'Tonik Employees',
                            'TONIK HEADCOUNT', 'TONIK PROMOTER COUNT', 'BRAND A STAFF', 'TONIK STAFF', 'TONIK EMPLOYEES'],
        'HC Headcount': ['HC Headcount', 'NC Headcount', 'HC Promoter Count', 'Brand B Staff', 'HC Staff', 'HC Employees',
                        'HC HEADCOUNT', 'NC HEADCOUNT', 'HC PROMOTER COUNT', 'BRAND B STAFF', 'HC STAFF', 'HC EMPLOYEES'],
        'Skyro Headcount': ['Skyro Headcount', 'Brand C Staff', 'Skyro Staff', 'Skyro Employees', 'Skyro Promoters',
                            'SKYRO HEADCOUNT', 'BRAND C STAFF', 'SKYRO STAFF', 'SKYRO EMPLOYEES', 'SKYRO PROMOTERS'],
        'Salmon Headcount': ['Salmon Headcount', 'Brand D Staff', 'Salmon Staff', 'Salmon Employees', 'Salmon Promoters',
                            'SALMON HEADCOUNT', 'BRAND D STAFF', 'SALMON STAFF', 'SALMON EMPLOYEES', 'SALMON PROMOTERS'],
        'Remarks': ['Remarks', 'Comments', 'Notes', 'Store Headcount', 'Additional Info', 'Additional Information',
                    'REMARKS', 'COMMENTS', 'NOTES', 'ADDITIONAL INFO', 'ADDITIONAL INFORMATION']
    }
    
    # Add variations with whitespace to the mapping
    whitespace_mapping = {}
    for key, values in column_mapping.items():
        # Add stripped versions and handle case sensitivity
        expanded_values = []
        for v in values:
            if isinstance(v, str):
                expanded_values.append(v.strip())
                expanded_values.append(v.strip().lower())
                expanded_values.append(v.strip().upper())
                
        whitespace_mapping[key.strip()] = values + expanded_values
    
    column_mapping = whitespace_mapping
    
    # Initialize an empty list to store DataFrames
    dfs = []
    
    # Define a function to clean store names for better matching
    def clean_store_name(name):
        if not isinstance(name, str):
            return name
        
        # Convert to lowercase for better matching
        name = name.lower().strip()
        
        # Remove common prefixes/suffixes and standardize format
        name = re.sub(r'^(store|branch|location|pos|site|outlet|mall)[:]\s*', '', name)
        name = re.sub(r'\s+', ' ', name)  # Normalize whitespace
        
        # Remove common store suffixes and noise words
        name = re.sub(r'\b(store|branch|outlet|mall|shop|location)\b', '', name)
        name = re.sub(r'\bsm\b', '', name)  # Remove shopping mall abbreviation
        
        # Remove other common noise patterns
        name = re.sub(r'[-_]', ' ', name)  # Replace hyphens and underscores with spaces
        name = re.sub(r'[()]', '', name)   # Remove parentheses
        name = re.sub(r'\s+', ' ', name)   # Normalize whitespace again after removals
        
        # Special case replacements
        name = name.replace('saint', 'st')
        name = name.replace('avenue', 'ave')
        name = name.replace('road', 'rd')
        
        # Final cleanup
        name = name.strip()
        
        return name
    
    # First, read all files to determine which ones have non-null data for each column
    file_data = []
    for file in excel_files:
        try:
            # Try to read all sheets in the Excel file
            print(f"Reading file: {file.name}")
            excel = pd.ExcelFile(file)
            sheet_names = excel.sheet_names
            print(f"  Found {len(sheet_names)} sheets: {', '.join(sheet_names)}")
            
            # Read each sheet and append to file_data
            for sheet_name in sheet_names:
                try:
                    print(f"  Processing sheet: {sheet_name}")
                    # Read the Excel file with all columns as strings
                    df = pd.read_excel(file, sheet_name=sheet_name, dtype=str)
                    
                    if df.empty:
                        print(f"  Sheet {sheet_name} is empty, skipping")
                        continue
                        
                    # Print the shape of the DataFrame
                    print(f"  Sheet shape: {df.shape} - {len(df.columns)} columns, {len(df)} rows")
                    
                    # Clean the data: strip whitespace and convert "-" to NaN
                    for col in df.columns:
                        # Strip whitespace
                        df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
                        # Convert "-", "N/A", "None", or empty strings to NaN
                        df[col] = df[col].replace(["-", "N/A", "None", ""], pd.NA)
                    
                    # Store the DataFrame and its non-null column info
                    non_null_cols = {col: df[col].notna().sum() for col in df.columns if df[col].notna().sum() > 0}
                    
                    # Count the number of sales-related columns with data
                    sales_columns_count = sum(1 for col in non_null_cols.keys()
                                            if any(sales_term.lower() in str(col).lower()
                                                   for sales_term in ['sales', 'retail', 'tonik', 'hc', 'skyro', 'salmon']))
                    
                    print(f"  Found {sales_columns_count} sales-related columns")
                    file_data.append((f"{file.name} - {sheet_name}", df, non_null_cols, sales_columns_count))
                except Exception as e:
                    print(f"  Error reading sheet {sheet_name}: {str(e)}")
        except Exception as e:
            print(f"Error reading {file.name}: {str(e)}")
    
    if not file_data:
        print("No valid Excel files could be read")
        return False
    
    # Sort files to prioritize those with more sales data, then by overall data completeness
    file_data.sort(key=lambda x: (x[3], sum(x[2].values())), reverse=True)
    
    # Initialize a DataFrame with the template columns using object dtype for consistency
    combined_df = pd.DataFrame(columns=template_columns, dtype=object)
    
    # Process each file in order
    for file_name, df, _, _ in file_data:
        # Map columns from the input file to the template columns
        for template_col in template_columns:
            if template_col in column_mapping:
                # Find the first matching column in the DataFrame
                # Also check for columns with whitespace
                df_cols_stripped = [col.strip() if isinstance(col, str) else col for col in df.columns]
                matching_cols = []  # We'll try to find all matching columns, not just the first one
                
                # First try exact match
                for possible_col in column_mapping[template_col]:
                    if possible_col in df.columns:
                        matching_cols.append(possible_col)
                
                # If no exact match, try stripped match
                if not matching_cols:
                    for i, col in enumerate(df_cols_stripped):
                        if col in column_mapping[template_col]:
                            matching_cols.append(df.columns[i])
                
                # If we found multiple matching columns, prioritize those with more non-null values
                if matching_cols:
                    if len(matching_cols) > 1:
                        matching_cols.sort(key=lambda col: df[col].notna().sum(), reverse=True)
                    
                    matching_col = matching_cols[0]
                    
                    # Only copy non-null values to avoid overwriting good data with NaNs
                    non_null_mask = df[matching_col].notna()
                    if non_null_mask.any():  # Check if column has any non-null values
                        # If the column doesn't exist in combined_df yet, create it
                        if template_col not in combined_df:
                            combined_df[template_col] = None
                        
                        # Special handling for sales columns to ensure they're properly captured
                        is_sales_column = any(sales_term in template_col.lower() for sales_term in ['sales', 'retail', 'tonik', 'hc', 'skyro', 'salmon'])
                        
                        # Only update rows that have non-null values in this file
                        for idx in range(len(df)):
                            if non_null_mask.iloc[idx]:
                                # Check if this is potentially a store name
                                is_store_name = template_col.lower() in ['pos name', 'store name', 'store', 'location']
                                
                                # Extend combined_df if needed
                                while idx >= len(combined_df):
                                    combined_df = pd.concat([combined_df, pd.DataFrame({template_col: [None]})], ignore_index=True)
                                
                                # Update the value
                                value = df.at[idx, matching_col]
                                
                                # For sales columns, ensure we're not overwriting a non-null value with a zero
                                if is_sales_column and (value == "0" or value == 0) and pd.notna(combined_df.at[idx, template_col]):
                                    continue
                                
                                # For string values, strip whitespace
                                if isinstance(value, str):
                                    value = value.strip()
                                    # Convert "-", "N/A", "None" to NaN
                                    if value in ["-", "N/A", "None", ""]:
                                        value = pd.NA
                                    # Clean store names
                                    elif is_store_name and len(value) > 0:
                                        value = clean_store_name(value)
                                
                                # Don't overwrite existing values with NaN
                                if pd.isna(value) and pd.notna(combined_df.at[idx, template_col]):
                                    continue
                                    
                                combined_df.at[idx, template_col] = value
    
    # Collect all unique store names from all files
    all_stores = set()
    store_data = {}
    
    log_message("\nExtracting store data from all files...")
    
    # Expanded list of possible store name column identifiers - more comprehensive
    store_col_identifiers = [
        'POS Name', 'Store Name', 'STORE DATA', 'Store', 'Location', 'Branch', 'BRANCH', 'LOCATION',
        'STORE', 'Name', 'BRANCH NAME', 'NAME', 'SITE', 'Site', 'OUTLET', 'Outlet', 'ADDRESS', 'Address',
        'AREA', 'Area', 'STORE DATA JANUARY', 'STORE NAME', 'BRANCH NAME'
    ]
    
    # Words that indicate a header row rather than actual store names
    header_indicators = [
        'pos name', 'store name', 'store', 'location', 'name', 'branch', 'branch name',
        'address', 'area', 'territory', 'region', 'outlet', 'site', 'no.', 'number', 'id',
        'code', 'status', 'remarks', 'comment'
    ]
    
    # Scan each file for store names
    for file_name, df, _, _ in file_data:
        log_message(f"Extracting store data from {file_name}")
        
        # First attempt: look for columns that are likely to contain store names
        store_cols = []
        for col in df.columns:
            if any(pos_col.lower() in str(col).lower() for pos_col in store_col_identifiers):
                store_cols.append(col)
                
        if not store_cols:
            # No obvious store column found, try heuristics
            # If a column has a lot of unique string values, it might be store names
            for col in df.columns:
                if df[col].dtype == 'object':
                    unique_values = df[col].dropna().unique()
                    if len(unique_values) > 5 and len(unique_values) < len(df) * 0.9:
                        # Likely a categorical column like store names
                        store_cols.append(col)
        
        if store_cols:
            log_message(f"  Found potential store name columns: {store_cols}")
        else:
            log_message(f"  No store name columns identified, checking all columns")
            store_cols = df.columns.tolist()
            
        # Process each potential store column
        for col in store_cols:
            for idx, value in enumerate(df[col]):
                if pd.notna(value) and str(value).strip() != '':
                    store_name = str(value).strip()
                    
                    # Skip rows that are likely headers or not actual store names
                    if any(indicator.lower() in store_name.lower() for indicator in header_indicators):
                        continue
                    
                    # Skip if too short or likely not a store name (all numbers)
                    if len(store_name) <= 3 or store_name.isdigit():
                        continue
                        
                    # Clean up store names using the clean_store_name function
                    store_name = clean_store_name(store_name)
                    
                    # Skip if the cleaned name is too short
                    if len(store_name) <= 2:
                        continue
                    
                    # Add to our collection
                    all_stores.add(store_name)
                    if store_name not in store_data:
                        store_data[store_name] = {}
                        log_message(f"    Found new store: {store_name} in column {col}")
                    
                    # Collect all data for this store from this file
                    for df_col in df.columns:
                        if pd.notna(df.at[idx, df_col]) and str(df.at[idx, df_col]).strip() != '' and str(df.at[idx, df_col]).strip() != '0':
                            # If we already have data for this column but it's "0" or empty, replace it
                            if df_col in store_data[store_name]:
                                existing_value = store_data[store_name][df_col]
                                if str(existing_value).strip() in ['0', '']:
                                    store_data[store_name][df_col] = df.at[idx, df_col]
                            else:
                                store_data[store_name][df_col] = df.at[idx, df_col]
    
    log_message(f"Found {len(all_stores)} unique stores across all files")
    # Log store names for debugging
    log_message("All stores found:")
    for store in sorted(all_stores):
        log_message(f"  - {store}")
    
    # Create a new DataFrame with all unique stores
    new_combined_df = pd.DataFrame(columns=template_columns, dtype=object)
    
    # Add all stores to the new DataFrame with correct column types
    for store_name in all_stores:
        new_row = pd.Series(index=template_columns, dtype=object)  # Use object dtype to avoid type warnings
        new_row['POS Name'] = store_name
        new_combined_df = pd.concat([new_combined_df, pd.DataFrame([new_row])], ignore_index=True)
    
    # Map data from all files to the new DataFrame
    for store_name, data in store_data.items():
        # Find the index of this store in the new DataFrame
        store_idx = new_combined_df[new_combined_df['POS Name'] == store_name].index
        if len(store_idx) > 0:
            store_idx = store_idx[0]
            # Map data from all columns
            for df_col, value in data.items():
                for template_col in template_columns:
                    if template_col in column_mapping:
                        if df_col in column_mapping[template_col] or df_col.strip() in column_mapping[template_col]:
                            if pd.isna(new_combined_df.at[store_idx, template_col]) or new_combined_df.at[store_idx, template_col] == '' or new_combined_df.at[store_idx, template_col] == '0':
                                new_combined_df.at[store_idx, template_col] = value
    
    # Replace the combined_df with the new one
    combined_df = new_combined_df
    
    # Check for stores with missing sales data
    # Get the sales columns directly from the template columns to ensure matching
    sales_columns = []
    for col in template_columns:
        if any(sales_term.lower() in col.lower() for sales_term in ['sales', 'retail', 'tonik', 'hc', 'skyro', 'salmon']):
            sales_columns.append(col)
    
    if not sales_columns:
        # If no sales columns found in template, use default list
        sales_columns = ['Retail Sales', 'Retailer Sales', 'Tonik Sales', 'HC Sales', 'Skyro Sales', 'Salmon']
    
    log_message(f"\nDetected sales columns: {', '.join(sales_columns)}")
    
    stores_missing_data = []
    for idx, row in combined_df.iterrows():
        store_name = row['POS Name']
        missing_sales = True
        for col in sales_columns:
            if col in combined_df.columns and pd.notna(row[col]) and row[col] != '' and row[col] != '0':
                missing_sales = False
                break
        if missing_sales:
            stores_missing_data.append(store_name)
    
    log_message(f"\nStores missing sales data ({len(stores_missing_data)}):")
    for store in sorted(stores_missing_data):
        log_message(f"  - {store}")
    
    # Expand the list of stores that need special handling
    stores_to_fix = stores_missing_data.copy()
    
    # Look for potential duplicate stores (similar names) before processing missing data
    log_message("\nAnalyzing possible store duplicates:")
    store_name_map = {}  # Map of variations to canonical store names
    store_list = list(set(row['POS Name'] for _, row in combined_df.iterrows() if pd.notna(row['POS Name'])))
    
    for i in range(len(store_list)):
        for j in range(i+1, len(store_list)):
            name1 = store_list[i].lower()
            name2 = store_list[j].lower()
            
            # Check for similarity
            similarity = False
            similarity_reason = ""
            if name1 in name2 or name2 in name1:
                similarity = True
                similarity_reason = "one name contains the other"
            elif len(name1) > 5 and len(name2) > 5:
                # Check if first 5 characters match
                if name1[:5] == name2[:5]:
                    similarity = True
                    similarity_reason = "first 5 characters match"
                # Check if removing common words (Mall, Branch, etc.) makes them similar
                common_words = ['mall', 'branch', 'store', 'shop', 'outlet']
                name1_clean = ' '.join(word for word in name1.split() if word.lower() not in common_words)
                name2_clean = ' '.join(word for word in name2.split() if word.lower() not in common_words)
                if name1_clean in name2_clean or name2_clean in name1_clean:
                    similarity = True
                    similarity_reason = "similar after removing common words"
            
            if similarity:
                # Use the longer name as canonical or the one with more data
                idx1 = combined_df[combined_df['POS Name'] == store_list[i]].index[0]
                idx2 = combined_df[combined_df['POS Name'] == store_list[j]].index[0]
                
                # Count non-null values in each row
                non_null_count1 = combined_df.iloc[idx1].notna().sum()
                non_null_count2 = combined_df.iloc[idx2].notna().sum()
                
                # Prefer the one with more data, or the longer name if equal
                if non_null_count1 >= non_null_count2:
                    canonical = store_list[i]
                    variation = store_list[j]
                else:
                    canonical = store_list[j]
                    variation = store_list[i]
                
                store_name_map[variation] = canonical
                log_message(f"  Detected duplicate stores: '{variation}' -> '{canonical}' (reason: {similarity_reason})")
    
    # Consolidate duplicate stores
    if store_name_map:
        for variation, canonical in store_name_map.items():
            var_idx = combined_df[combined_df['POS Name'] == variation].index
            canon_idx = combined_df[combined_df['POS Name'] == canonical].index
            
            if len(var_idx) > 0 and len(canon_idx) > 0:
                var_idx = var_idx[0]
                canon_idx = canon_idx[0]
                
                # Merge data from variation to canonical
                for col in combined_df.columns:
                    if col != 'POS Name' and pd.isna(combined_df.at[canon_idx, col]) and pd.notna(combined_df.at[var_idx, col]):
                        combined_df.at[canon_idx, col] = combined_df.at[var_idx, col]
                
                # Mark the row for deletion
                combined_df.at[var_idx, 'POS Name'] = f"DUPLICATE_{variation}"
        
        # Remove duplicated stores
        combined_df = combined_df[~combined_df['POS Name'].str.startswith('DUPLICATE_', na=False)]
        
        # Update stores_to_fix to reference canonical names
        stores_to_fix = [store_name_map.get(name, name) for name in stores_to_fix]
        stores_to_fix = list(set(stores_to_fix))  # Remove any duplicates
    
    # Process all stores with missing data
    log_message("\nAttempting to fix stores with missing data:")
    stores_fixed = 0
    for store_name in stores_to_fix:
        log_message(f"\n  Processing store: {store_name}")
        found_match = False
        store_fixed = False
        
        # First try exact match
        for file_name, df, _, _ in file_data:
            for col in df.columns:
                # Try to find the store by exact name (case insensitive)
                exact_matches = df[df[col].astype(str).str.lower() == store_name.lower()]
                if not exact_matches.empty:
                    log_message(f"    Found exact match in {file_name}, column {col} ({len(exact_matches)} rows)")
                    found_match = True
                    # Find the index in combined_df
                    store_idx = combined_df[combined_df['POS Name'] == store_name].index
                    if len(store_idx) > 0:
                        store_idx = store_idx[0]
                        # Copy all non-null values
                        data_added = []
                        for idx, row in exact_matches.iterrows():
                            for df_col in df.columns:
                                for template_col in template_columns:
                                    if template_col in column_mapping:
                                        # Check for column mapping match
                                        col_matches = False
                                        for possible_col in column_mapping[template_col]:
                                            if isinstance(possible_col, str) and (
                                                df_col.lower() == possible_col.lower() or
                                                df_col.strip().lower() == possible_col.strip().lower()):
                                                col_matches = True
                                                break
                                                
                                        if col_matches:
                                            if pd.notna(row[df_col]) and str(row[df_col]).strip() not in ['', '0', '-', 'N/A', 'None']:
                                                old_value = combined_df.at[store_idx, template_col]
                                                combined_df.at[store_idx, template_col] = row[df_col]
                                                data_added.append(f"{template_col}: {old_value} -> {row[df_col]}")
                                                if template_col in sales_columns:
                                                    store_fixed = True
                        
                        if data_added:
                            log_message(f"    Added data: {', '.join(data_added[:5])}{' and more...' if len(data_added) > 5 else ''}")
                        else:
                            log_message(f"    No valid data found to add")
        
        # If exact match didn't work, try partial match with improved matching logic
        store_idx = combined_df[combined_df['POS Name'] == store_name].index
        if len(store_idx) > 0:
            store_idx = store_idx[0]
            # Check if we still have missing sales data
            missing_sales = True
            for col in sales_columns:
                if col in combined_df.columns and pd.notna(combined_df.at[store_idx, col]) and str(combined_df.at[store_idx, col]) not in ['', '0', '-', 'N/A', 'None']:
                    missing_sales = False
                    break
            
            if missing_sales:
                log_message(f"    Exact match didn't find sales data, trying partial matching")
                # Multiple matching strategies
                store_name_variants = [
                    store_name.lower(),
                    # Remove common words
                    ' '.join(word for word in store_name.lower().split() if word not in ['mall', 'branch', 'store', 'shop', 'outlet']),
                    # Try extracting the main name part
                    store_name.split()[0] if len(store_name.split()) > 0 else ""
                ]
                
                log_message(f"    Trying with variants: {store_name_variants}")
                
                for file_name, df, _, _ in file_data:
                    for col in df.columns:
                        found_match = False
                        partial_match_strategy = ""
                        
                        # Try different matching strategies in order
                        for variant in store_name_variants:
                            if not variant or len(variant) < 3:
                                continue
                                
                            # Try partial match (case insensitive)
                            partial_matches = df[df[col].astype(str).str.lower().str.contains(variant, na=False)]
                            
                            if not partial_matches.empty:
                                found_match = True
                                partial_match_strategy = f"variant '{variant}'"
                                log_message(f"      Found partial match using variant '{variant}' in {file_name}, column {col} ({len(partial_matches)} matches)")
                                break
                        
                        # If no match found with above strategies, try matching parts of the store name
                        if not found_match:
                            # Try matching parts of the store name
                            store_parts = [part for part in store_name.split() if len(part) > 3]
                            if store_parts:
                                log_message(f"      Trying with store name parts: {', '.join(store_parts)}")
                            for part in store_parts:
                                partial_matches = df[df[col].astype(str).str.lower().str.contains(part.lower(), na=False)]
                                if not partial_matches.empty:
                                    found_match = True
                                    partial_match_strategy = f"part '{part}'"
                                    log_message(f"      Found partial match using part '{part}' in {file_name}, column {col} ({len(partial_matches)} matches)")
                                    break
                        
                        # Apply matches found with any strategy
                        if found_match and not partial_matches.empty:
                            # Prioritize rows with more non-null values
                            non_null_counts = partial_matches.notna().sum(axis=1)
                            partial_matches = partial_matches.iloc[non_null_counts.argsort(kind='stable')[::-1]]
                            
                            # Copy all non-null values from best matching row
                            data_added = []
                            partial_store_fixed = False
                            
                            for idx, row in partial_matches.iterrows():
                                if partial_store_fixed:  # If we already fixed with one row, stop
                                    break
                                    
                                log_message(f"      Processing match: {row.get(col, 'unknown')}")
                                
                                for df_col in df.columns:
                                    for template_col in template_columns:
                                        if template_col in column_mapping:
                                            # Check for column mapping match
                                            col_matches = False
                                            for possible_col in column_mapping[template_col]:
                                                if isinstance(possible_col, str) and (
                                                    df_col.lower() == possible_col.lower() or
                                                    df_col.strip().lower() == possible_col.strip().lower()):
                                                    col_matches = True
                                                    break
                                                    
                                            if col_matches:
                                                if pd.notna(row[df_col]) and str(row[df_col]).strip() not in ['', '0', '-', 'N/A', 'None']:
                                                    old_value = combined_df.at[store_idx, template_col]
                                                    combined_df.at[store_idx, template_col] = row[df_col]
                                                    data_added.append(f"{template_col}: {old_value} -> {row[df_col]}")
                                                    if template_col in sales_columns:
                                                        partial_store_fixed = True
                                                        store_fixed = True
                            
                            if data_added:
                                log_message(f"      Added data using {partial_match_strategy}: {', '.join(data_added[:5])}{' and more...' if len(data_added) > 5 else ''}")
                                if partial_store_fixed:
                                    log_message(f"      Successfully added sales data for store")
                            else:
                                log_message(f"      No valid data found to add from partial match")
    
    # Check for any remaining stores with missing sales data
    stores_still_missing = []
    for idx, row in combined_df.iterrows():
        store_name = row['POS Name']
        missing_sales = True
        for col in sales_columns:  # Using the dynamically identified sales columns
            if col in combined_df.columns and pd.notna(row[col]) and str(row[col]) not in ['', '0', '-', 'N/A', 'None']:
                missing_sales = False
                break
        if missing_sales:
            stores_still_missing.append(store_name)
    
    log_message(f"\nStores still missing sales data after all processing: {len(stores_still_missing)}")
    for store in sorted(stores_still_missing):
        log_message(f"  - {store}")
    
    # Potential new issue: Check if some stores were completely missed from input files
    log_message("\nChecking for potentially missed stores...")
    all_possible_stores = set()
    for file_name, df, _, _ in file_data:
        for col in df.columns:
            for value in df[col].dropna():
                if isinstance(value, str) and len(value.strip()) > 3 and not value.isdigit():
                    cleaned = clean_store_name(value.strip())
                    if len(cleaned) > 3:
                        all_possible_stores.add(cleaned)
    
    # Check if all stores we found are in the final DataFrame
    final_stores = set(row['POS Name'] for _, row in combined_df.iterrows() if pd.notna(row['POS Name']))
    potentially_missed_stores = all_possible_stores - final_stores
    
    if potentially_missed_stores:
        log_message(f"\nPotentially missed stores (not in final output): {len(potentially_missed_stores)}")
        for store in sorted(potentially_missed_stores)[:20]:  # Limit to first 20 to avoid excessive output
            log_message(f"  - {store}")
        if len(potentially_missed_stores) > 20:
            log_message(f"  ... and {len(potentially_missed_stores) - 20} more")
            
    # Summary of processing results
    log_message("\n===== SUMMARY =====")
    log_message(f"Total stores found across all files: {len(all_stores)}")
    log_message(f"Stores in final output: {len(final_stores)}")
    log_message(f"Stores with missing sales data initially: {len(stores_missing_data)}")
    log_message(f"Stores fixed during processing: {stores_fixed}")
    log_message(f"Stores still missing sales data: {len(stores_still_missing)}")
    log_message(f"Potentially missed stores: {len(potentially_missed_stores)}")
    
    try:
        # Create output directory if it doesn't exist
        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        
        # Ensure the output file has a proper extension
        if not output_path.lower().endswith(('.xlsx', '.xls')):
            output_path = output_path + '.xlsx'
            output_file = Path(output_path)
        
        # Save the concatenated DataFrame to Excel with all columns as strings
        combined_df.to_excel(output_path, index=False)
        log_message(f"\nSaved final output with {len(combined_df)} rows to {output_path}")
        return True
    except Exception as e:
        print(f"Error during concatenation or saving: {str(e)}")
        return False

def main():
    # Set up argument parser
    parser = argparse.ArgumentParser(description='Concatenate Excel files using a template file for column structure.')
    parser.add_argument('folder', help='Folder containing Excel files to concatenate')
    parser.add_argument('-o', '--output', default='combined_data.xlsx', 
                        help='Output file path (default: combined_data.xlsx in current directory)')
    parser.add_argument('-t', '--template', required=True,
                        help='Template Excel file that defines the column structure')
    parser.add_argument('-p', '--pattern', help='File pattern to match (e.g., "data_*.xlsx")')
    
    # Parse arguments
    args = parser.parse_args()
    
    # Call the concatenation function
    success = concatenate_excel_files_with_template(args.folder, args.output, args.template, args.pattern)
    
    if success:
        print("Concatenation completed successfully!")
    else:
        print("Concatenation failed.")

if __name__ == "__main__":
    main()
