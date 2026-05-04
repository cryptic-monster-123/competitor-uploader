#!/bin/bash
echo "Running Excel to CSV Concatenator..."

# Create output directory if it doesn't exist
mkdir -p output/csv
mkdir -p logs

python3 src/concatenators/excel_to_csv_concatenator.py "$(pwd)/data/2025-03" -t "$(pwd)/templates/uploader_template.xlsx" -o "output/csv/combined_data_$(date +%Y%m%d_%H%M%S).csv" -e "M POS Status" -p "March 2025"
echo
echo "If successful, the combined data is saved with a timestamp in the filename"
