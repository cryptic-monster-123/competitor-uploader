# Excel Concatenator

A Python script to concatenate multiple Excel files into a single Excel file. Three versions are provided:

1. **excel_concatenator.py**: Basic version that concatenates Excel files with identical column structures
2. **excel_concatenator_enhanced.py**: Advanced version that maps different column names to a standard set
3. **excel_concatenator_template.py**: Template-based version that uses an Excel file as a template for the output structure

## Project Structure

```
competitor-uploader/
├── README.md
├── requirements.txt
├── run_concatenator.bat
├── run_csv_concatenator.bat
├── src/
│   ├── __init__.py
│   ├── concatenators/
│   │   ├── __init__.py
│   │   ├── excel_concatenator.py
│   │   ├── excel_concatenator_enhanced.py
│   │   ├── excel_concatenator_template.py
│   │   └── excel_to_csv_concatenator.py
│   ├── analyzers/
│   │   ├── __init__.py
│   │   ├── analyze_csv.py
│   │   ├── check_csv.py
│   │   └── examine_excel.py
│   ├── config/
│   │   ├── __init__.py
│   │   └── column_mapping_example.json
│   └── utils/
│       └── __init__.py
├── data/
│   ├── 2024-01/
│   │   └── [January 2024 Excel files]
│   ├── 2025-02/
│   │   └── [February 2025 Excel files]
│   └── test/
│       └── [test data files]
├── templates/
│   ├── uploader_template.csv
│   └── uploader_template.xlsx
├── output/
│   ├── archive/
│   ├── csv/
│   │   └── [combined CSV files]
│   └── excel/
│       └── [output Excel files]
└── logs/
    └── [log files]
```

## Quick Start

1. Install the required packages:
   ```
   pip install -r requirements.txt
   ```

2. Generate sample data for testing:
   ```
   python src/concatenators/create_sample_data.py
   ```
   This creates sample Excel files in the `test_data` directory.

3. Run the basic concatenator (for files with identical columns):
   ```
   python src/concatenators/excel_concatenator.py data/test -o output/excel/combined_result.xlsx
   ```

4. Run the enhanced concatenator (for files with different column names):
   ```
   python src/concatenators/excel_concatenator_enhanced.py data/test -o output/excel/combined_result_enhanced.xlsx
   ```

5. Run the template-based concatenator (to match a specific column structure):
   ```
   python src/concatenators/excel_concatenator_template.py data/test -o output/excel/combined_result_template.xlsx -t templates/uploader_template.xlsx
   ```

6. Alternatively, use the interactive scripts:
   - Windows: Double-click `run_concatenator.bat`
   - Linux/macOS: Run `./run_concatenator.sh` (make it executable first with `chmod +x run_concatenator.sh`)

## Requirements

- Python 3.6+
- pandas
- openpyxl

## Installation

1. Ensure you have Python installed on your system.
2. Install the required packages:

```
pip install pandas openpyxl
```

## Using as a Python Package

The project is structured as a proper Python package with `__init__.py` files, allowing you to import and use the modules in your own Python scripts:

```python
# Import specific modules
from src.concatenators import excel_concatenator
from src.config import load_column_mapping

# Use the concatenator programmatically
mapping = load_column_mapping()
excel_concatenator.concatenate_excel_files('data/2024-01', 'output/excel/result.xlsx')
```

This package structure makes it easier to:
- Import specific modules or functions
- Reuse code across different scripts
- Extend the functionality with new modules
- Maintain a clean separation of concerns

## Usage

### Basic Usage

```
python src/concatenators/excel_concatenator.py path/to/excel/files
```

This will:
1. Find all Excel files (*.xlsx, *.xls, *.xlsm) in the specified folder
2. Concatenate them into a single Excel file
3. Save the result as `output/excel/combined_data.xlsx` in the current directory

### Advanced Options

```
python src/concatenators/excel_concatenator.py path/to/excel/files -o output/excel/output.xlsx -p "data_*.xlsx"
```

Parameters:
- `folder`: Path to the folder containing Excel files (required)
- `-o, --output`: Path where the output file will be saved (default: output/excel/combined_data.xlsx)
- `-p, --pattern`: File pattern to match specific Excel files (e.g., "data_*.xlsx")

## Examples

Combine all Excel files in the "data/2024-01" folder:
```
python src/concatenators/excel_concatenator.py data/2024-01
```

Combine only files matching a pattern and save to a specific location:
```
python src/concatenators/excel_concatenator.py data/2025-02 -o output/excel/combined_report.xlsx -p "monthly_*.xlsx"
```

## How It Works

### Basic Version (excel_concatenator.py)

The basic script:
1. Scans the specified folder for Excel files (or files matching the provided pattern)
2. Reads each Excel file into a pandas DataFrame
3. Concatenates all DataFrames vertically (row-wise)
4. Saves the combined data to a new Excel file

### Enhanced Version (excel_concatenator_enhanced.py)

The enhanced script:
1. Scans the specified folder for Excel files (excluding temporary files)
2. Reads each Excel file and maps columns to a standard set of column names
3. Creates a new DataFrame with standardized columns for each file
4. Concatenates all standardized DataFrames vertically
5. Saves the combined data to a new Excel file

The enhanced version is particularly useful when:
- Excel files have different column names for the same data
- Some files have extra columns that you don't need
- You want to ensure a consistent output format

### Template-Based Version (excel_concatenator_template.py)

The template-based script:
1. Reads a template Excel file to determine the desired column structure
2. Scans the specified folder for Excel files (excluding temporary files)
3. Maps columns from each input file to match the template structure
4. Creates a new DataFrame with only the columns specified in the template
5. Concatenates all DataFrames vertically
6. Saves the combined data to a new Excel file with the exact column structure of the template

The template-based version is ideal when:
- You need the output to match a specific format exactly
- You want to select only certain columns from the input files
- You have an existing Excel file that defines the desired structure

## Notes

### Basic Version
- All Excel files must have the same column structure (same column names)
- The script preserves all data and column names
- Files are processed in the order they are found in the directory

### Enhanced Version
- Works with Excel files that have different column names
- Maps columns to a standard set of column names
- Supports nested column mappings for more complex data organization
- Handles missing columns by filling with NULL values
- Excludes temporary Excel files (those starting with ~$)
- Provides more detailed logging of the column mapping process

### Template-Based Version
- Uses an existing Excel file as a template for the output structure
- Only includes columns that are present in the template
- Handles whitespace and variations in column names
- Ideal for creating reports with a specific format
- Can be used with an empty Excel file that just defines the column headers

## Custom Column Mapping

### JSON Mapping (Enhanced Version)

The enhanced version supports custom column mappings via a JSON file:

```
python src/concatenators/excel_concatenator_enhanced.py data -o output/excel/output.xlsx -m src/config/column_mapping_example.json
```

A sample column mapping file (`src/config/column_mapping_example.json`) is provided, which demonstrates:
- Simple column mappings (one-to-many)
- Nested column mappings for categorized data
- How to handle variations in column names across different files

Example mapping structure:
```json
{
    "Output Column Name": ["Possible Input Column 1", "Possible Input Column 2"],
    "Categorized Data": {
        "Category 1": ["Input Column A", "Alternative Name A"],
        "Category 2": ["Input Column B", "Alternative Name B"]
    }
}
```

This will create columns named "Output Column Name" and "Categorized Data - Category 1", "Categorized Data - Category 2", etc.

### Excel Template (Template-Based Version)

The template-based version uses an Excel file as a template:

```
python src/concatenators/excel_concatenator_template.py data -o output/excel/output.xlsx -t templates/uploader_template.xlsx
```

The template file should:
- Contain the desired column headers
- Have the exact structure you want for the output
- Can be an empty file with just the headers (no data rows required)

This approach is simpler than creating a JSON mapping file and ensures the output matches your desired format exactly.

Both versions:
- Create the output directory if it doesn't exist
- Support file pattern matching to select specific files
