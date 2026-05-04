#!/bin/bash
echo "Excel Concatenator"
echo "================="
echo

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "Error: Python is not installed."
    echo "Please install Python from https://www.python.org/downloads/"
    echo
    exit 1
fi

# Check if required packages are installed
echo "Checking required packages..."
if ! python3 -c "import pandas" &> /dev/null; then
    echo "Installing pandas and openpyxl..."
    pip3 install -r requirements.txt
    if [ $? -ne 0 ]; then
        echo "Error installing required packages."
        exit 1
    fi
fi

echo
echo "This script will automatically concatenate all Excel files in the data/2024-01 folder."
echo

# Set fixed parameters
SCRIPT_NAME="src/concatenators/excel_concatenator_template.py"
FOLDER_PATH="$(pwd)/data/2024-01"
OUTPUT_PATH="output/excel/combined_data.xlsx"
TEMPLATE_PATH="$(pwd)/templates/uploader_template.xlsx"

echo "Using template-based version"
echo "Using folder: $FOLDER_PATH"
echo "Using template: $TEMPLATE_PATH"
echo "Output file: $OUTPUT_PATH"

# Validate folder exists
if [ ! -d "$FOLDER_PATH" ]; then
    echo "Error: Folder does not exist."
    exit 1
fi

# Validate template exists
if [ ! -f "$TEMPLATE_PATH" ]; then
    echo "Error: Template file does not exist."
    exit 1
fi

# Create output directory if it doesn't exist
mkdir -p "$(dirname "$OUTPUT_PATH")"

echo
echo "Processing with $SCRIPT_NAME..."
echo

# Run the Python script with the provided parameters
python3 $SCRIPT_NAME "$FOLDER_PATH" -o "$OUTPUT_PATH" -t "$TEMPLATE_PATH"

echo
if [ $? -eq 0 ]; then
    echo "Concatenation completed successfully!"
else
    echo "Concatenation failed."
fi
