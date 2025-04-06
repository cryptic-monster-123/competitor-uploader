import os
import pandas as pd
import argparse
from pathlib import Path
import re
import logging
from datetime import datetime
from fuzzywuzzy import process

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(f"logs/concatenator_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
    ]
)
logger = logging.getLogger(__name__)

# Key columns that should be prioritized in matching
KEY_COLUMNS = ['pos name', 'brand', 'store name', 'location', 'address']

# Special column mappings for columns that might have different names but represent the same data
SPECIAL_MAPPINGS = {
    'hc headcount': ['hc promoter count', 'hc headcount', 'headcount hc'],
    'tonik headcount': ['tonik promoter count', 'tonik headcount', 'headcount tonik'],
    'skyro headcount': ['skyro promoter count', 'skyro headcount', 'headcount skyro'],
    'salmon headcount': ['salmon promoter count', 'salmon headcount', 'headcount salmon'],
    'retailer headcount': ['retailer promoter count', 'retailer headcount', 'headcount retailer'],
    'brand': ['brand', 'retailer', 'retailer name']
}

# Characters that should be treated as empty/null values
NULL_VALUE_CHARS = ['-', '—', '–', '−', '⁃', '‐', '‑', '‒', '–', '—', '―', '⁓', '⁻', '₋', '−']

def normalize_column_name(col_name):
    """Normalize column names for better matching."""
    if not isinstance(col_name, str):
        return ""
    
    # Convert to lowercase and strip whitespace
    normalized = col_name.lower().strip()
    
    # Remove special characters
    normalized = re.sub(r'[^\w\s]', ' ', normalized)
    
    # Replace multiple spaces with single space
    normalized = re.sub(r'\s+', ' ', normalized)
    
    # Common abbreviations and variations
    replacements = {
        'pos ': 'point of sale ',
        'hc': 'headcount',
        'jan': 'january',
        'feb': 'february',
        'mar': 'march',
        'apr': 'april',
        'jun': 'june',
        'jul': 'july',
        'aug': 'august',
        'sep': 'september',
        'oct': 'october',
        'nov': 'november',
        'dec': 'december'
    }
    
    for abbr, full in replacements.items():
        normalized = normalized.replace(abbr, full)
    
    return normalized

def get_best_column_match(template_col, file_cols, file_name, min_score=70):
    """Find the best match for a template column in the file columns."""
    # Normalize the template column for matching
    norm_template_col = normalize_column_name(template_col)
    
    # Normalize all file columns
    norm_file_cols = [normalize_column_name(col) for col in file_cols]
    
    # Check for special mappings first (for headcount columns)
    for special_key, alternatives in SPECIAL_MAPPINGS.items():
        if special_key == template_col.lower():
            logger.info(f"Checking special mappings for '{template_col}'")
            # Try to find any of the alternative names in the file columns
            for alt in alternatives:
                norm_alt = normalize_column_name(alt)
                for i, norm_col in enumerate(norm_file_cols):
                    # Check for exact match with any alternative
                    if norm_col == norm_alt or norm_col.startswith(norm_alt) or norm_alt.startswith(norm_col):
                        logger.info(f"Special mapping match found for '{template_col}' -> '{file_cols[i]}' (alternative: '{alt}')")
                        return file_cols[i], 100
    
    # Check for exact matches
    for i, norm_col in enumerate(norm_file_cols):
        if norm_col == norm_template_col:
            logger.info(f"Exact match found for '{template_col}' -> '{file_cols[i]}'")
            return file_cols[i], 100
    
    # If no exact match, use fuzzy matching
    match, score = process.extractOne(norm_template_col, norm_file_cols)
    
    # Adjust threshold for key columns and special columns
    threshold = min_score
    if any(key in norm_template_col for key in KEY_COLUMNS):
        threshold = max(60, threshold - 10)  # Lower threshold for key columns
    
    # Lower threshold for headcount columns
    if 'headcount' in norm_template_col or 'promoter' in norm_template_col:
        threshold = max(60, threshold - 5)  # Even lower threshold for headcount columns
    
    if score >= threshold:
        matched_col = file_cols[norm_file_cols.index(match)]
        logger.info(f"Matched '{template_col}' to '{matched_col}' (score: {score})")
        return matched_col, score
    else:
        logger.warning(f"No good match found for '{template_col}' in {file_name} (best: '{match}' with score {score})")
        return None, score

def select_best_sheet(excel_file, sales_columns=['sales', 'headcount']):
    """Select the sheet with the most non-null values in sales columns."""
    excel = pd.ExcelFile(excel_file)
    sheet_names = excel.sheet_names
    
    if len(sheet_names) == 1:
        return sheet_names[0]
    
    logger.info(f"File has multiple sheets: {sheet_names}")
    
    # Track the best sheet and its score
    best_sheet = sheet_names[0]
    best_score = 0
    
    for sheet in sheet_names:
        df = pd.read_excel(excel_file, sheet_name=sheet)
        
        # Count non-null values in columns that might contain sales data
        sales_score = 0
        for col in df.columns:
            if any(sales_term in col.lower() for sales_term in sales_columns):
                non_null_count = df[col].notna().sum()
                sales_score += non_null_count
        
        logger.info(f"Sheet '{sheet}' has {sales_score} non-null values in sales columns")
        
        if sales_score > best_score:
            best_score = sales_score
            best_sheet = sheet
    
    logger.info(f"Selected sheet '{best_sheet}' with {best_score} non-null values in sales columns")
    return best_sheet

def concatenate_excel_to_csv(folder_path, output_path, template_path, file_pattern=None, exclude_columns=None, period=None):
    """Concatenate Excel files using a template file for column structure and save as CSV."""
    # Convert to Path objects
    folder = Path(folder_path)
    template = Path(template_path)
    
    # Add timestamp to output filename if not already present
    if '{timestamp}' in output_path:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = output_path.replace('{timestamp}', timestamp)
    
    # Default columns to exclude
    if exclude_columns is None:
        exclude_columns = []
    
    logger.info(f"Starting concatenation process")
    logger.info(f"Template file: {template}")
    logger.info(f"Folder path: {folder}")
    logger.info(f"Output path: {output_path}")
    logger.info(f"Excluding columns: {exclude_columns}")

    # Read template columns
    try:
        template_df = pd.read_excel(template)
        # Create mappings for column names and dtypes
        template_columns = template_df.columns.tolist()
        template_col_mapping = {col.strip().lower(): col for col in template_columns}
        template_dtypes = {col: dtype for col, dtype in zip(template_columns, template_df.dtypes)}
        
        logger.info(f"Template columns: {template_columns}")
    except Exception as e:
        logger.error(f"Error reading template file: {str(e)}")
        return False

    # Get all Excel files
    excel_files = []
    for pattern in ['*.xlsx', '*.xls', '*.xlsm']:
        excel_files.extend([f for f in folder.glob(pattern) if not f.name.startswith('~$')])

    if not excel_files:
        logger.error(f"No Excel files found in {folder_path}")
        return False
    
    logger.info(f"Found {len(excel_files)} Excel files to process")

    # Process files
    dfs = []
    for file in excel_files:
        logger.info(f"Processing file: {file.name}")
        try:
            # Select the best sheet based on sales data
            best_sheet = select_best_sheet(file)
            
            # Read the Excel file using the best sheet
            df = pd.read_excel(file, sheet_name=best_sheet)
            # Strip whitespace from column names
            df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]
            logger.info(f"File columns from sheet '{best_sheet}': {df.columns.tolist()}")
            
            # Create fuzzy mapping for current file's columns
            file_cols = df.columns.tolist()
            file_col_mapping = {}
            
            # Find best matches for template columns
            for template_col in template_col_mapping.keys():
                matched_col, score = get_best_column_match(
                    template_col, 
                    file_cols, 
                    file.name
                )
                if matched_col:
                    file_col_mapping[template_col] = matched_col

            # Standardize columns with data validation
            new_df = pd.DataFrame()
            required_match_rate = 0.6  # Require at least 60% of columns matched
            matched_count = 0
            
            # Process key columns first to ensure they're properly handled
            for template_col, original_col in template_col_mapping.items():
                # Skip excluded columns
                if original_col in exclude_columns:
                    logger.info(f"Skipping excluded column: {original_col}")
                    continue
                    
                if any(key in template_col.lower() for key in KEY_COLUMNS):
                    src_col = file_col_mapping.get(template_col)
                    if src_col is not None:
                        # Convert to string and handle dash characters
                        temp_series = df[src_col].fillna('').astype(str).str.strip()
                        
                        # Replace all null value characters with NaN
                        for null_char in NULL_VALUE_CHARS:
                            temp_series = temp_series.replace(null_char, pd.NA)
                        
                        # Also replace any string that contains only these characters
                        temp_series = temp_series.replace(r'^[\s\-–—]+$', pd.NA, regex=True)
                        
                        new_df[original_col] = temp_series
                        matched_count += 1
                    else:
                        new_df[original_col] = pd.NA
            
            # Process remaining columns
            for template_col, original_col in template_col_mapping.items():
                # Skip excluded columns and already processed key columns
                if original_col in exclude_columns or any(key in template_col.lower() for key in KEY_COLUMNS):
                    continue
                    
                src_col = file_col_mapping.get(template_col)
                if src_col is not None:
                    # For numeric columns, handle special cases
                    if ('sales' in original_col.lower() or 
                        'headcount' in original_col.lower() or 
                        'count' in original_col.lower() or
                        'amount' in original_col.lower()):
                        
                        # Convert to string first to handle special characters
                        temp_series = df[src_col].astype(str).str.strip()
                        
                        # Replace all null value characters with NaN
                        for null_char in NULL_VALUE_CHARS:
                            temp_series = temp_series.replace(null_char, pd.NA)
                        
                        # Also replace any string that contains only these characters
                        temp_series = temp_series.replace(r'^[\s\-–—]+$', pd.NA, regex=True)
                        
                        # Log the conversion for debugging
                        non_numeric_count = sum(~pd.to_numeric(temp_series, errors='coerce').notna() & temp_series.notna())
                        if non_numeric_count > 0:
                            logger.info(f"Column '{original_col}' from '{file.name}': {non_numeric_count} non-numeric values converted to NaN")
                        
                        # Convert to numeric, coercing errors to NaN
                        new_df[original_col] = pd.to_numeric(temp_series, errors='coerce')
                    else:
                        # For non-numeric columns, still handle dash characters
                        temp_series = df[src_col].fillna('').astype(str).str.strip()
                        
                        # Replace all null value characters with NaN
                        for null_char in NULL_VALUE_CHARS:
                            temp_series = temp_series.replace(null_char, pd.NA)
                        
                        # Also replace any string that contains only these characters
                        temp_series = temp_series.replace(r'^[\s\-–—]+$', pd.NA, regex=True)
                        
                        new_df[original_col] = temp_series
                    
                    matched_count += 1
                else:
                    new_df[original_col] = pd.NA  # Use NA instead of empty strings

            # Skip files with too many unmatched columns
            match_rate = matched_count / len(template_col_mapping)
            logger.info(f"Match rate for {file.name}: {match_rate:.1%} ({matched_count}/{len(template_col_mapping)} columns)")
            
            if match_rate < required_match_rate:
                logger.warning(f"Skipping {file.name} - only {match_rate:.1%} columns matched")
                continue
            
            # Add file source information
            new_df['sourceFile'] = file.name
            
            # Append to list of dataframes
            dfs.append(new_df)
            logger.info(f"Successfully processed {file.name}")
            
        except Exception as e:
            logger.error(f"Error processing {file.name}: {str(e)}")

    if not dfs:
        logger.error("No valid data found")
        return False

    # Concatenate and save as CSV
    try:
        combined_df = pd.concat(dfs, ignore_index=True)
        logger.info(f"Combined data shape: {combined_df.shape}")
        logger.info(f"Attempting to save to: {output_path}")
        
        # Ensure output directory exists and file is writable
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
            
        if os.path.exists(output_path):
            os.remove(output_path)  # Remove existing file if any
        
        # Replace "Brand" and "Retailer Name" with "Retailer" in column names
        if "Brand" in combined_df.columns:
            combined_df = combined_df.rename(columns={"Brand": "Retailer"})
        if "Retailer Name" in combined_df.columns:
            combined_df = combined_df.rename(columns={"Retailer Name": "Retailer"})
        
        # Add Period column if specified
        if period:
            combined_df["datePeriod"] = period
        
        # Replace spaces with underscores in all column names

        combined_df.columns = [col.replace(" ", "_") for col in combined_df.columns]
            
        combined_df.to_csv(output_path, index=False)
        logger.info(f"Successfully saved concatenated data to {output_path}")
        return True
    except Exception as e:
        logger.error(f"Error saving CSV: {str(e)}")
        return False

def main():
    parser = argparse.ArgumentParser(description='Concatenate Excel files to CSV using template columns')
    parser.add_argument('folder', help='Folder containing Excel files')
    parser.add_argument('-o', '--output', default='combined_data_{timestamp}.csv', help='Output CSV file path')
    parser.add_argument('-t', '--template', required=True, help='Template Excel file with column structure')
    parser.add_argument('-e', '--exclude', nargs='+', default=['M POS Status'], help='Columns to exclude from output')
    parser.add_argument('-p', '--period', help='Add a Period column with the specified value')
    
    args = parser.parse_args()
    
    # Add timestamp to output filename if not already present
    if '{timestamp}' in args.output:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = args.output.replace('{timestamp}', timestamp)
    else:
        output_path = args.output
    
    success = concatenate_excel_to_csv(
        args.folder,
        output_path,
        args.template,
        exclude_columns=args.exclude,
        period=args.period
    )
    
    if success:
        logger.info("Process completed successfully")
    else:
        logger.error("Process failed")

if __name__ == "__main__":
    main()
