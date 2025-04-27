@echo off
echo Running Excel to CSV Concatenator...
python src\concatenators\excel_to_csv_concatenator.py "C:\Users\matth_9lb83h2\Desktop\TONIK\Coding Projects\competitor-uploader\data\2025-03" -t "C:\Users\matth_9lb83h2\Desktop\TONIK\Coding Projects\competitor-uploader\templates\uploader_template.xlsx" -o "output\csv\combined_data_{timestamp}.csv" -e "M POS Status" -p "March 2025"
echo.
echo If successful, the combined data is saved with a timestamp in the filename
pause
