@echo off
echo Excel Concatenator
echo =================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: Python is not installed or not in PATH.
    echo Please install Python from https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

REM Check if required packages are installed
echo Checking required packages...
pip show pandas >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing pandas and openpyxl...
    pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo Error installing required packages.
        pause
        exit /b 1
    )
)

echo.
echo This script will automatically concatenate all Excel files in the data\2024-01 folder.
echo.

REM Set fixed parameters
set SCRIPT_NAME=src\concatenators\excel_concatenator_template.py
set FOLDER_PATH=C:\Users\matth_9lb83h2\Desktop\TONIK\Coding Projects\competitor-uploader\data\2024-01
set OUTPUT_PATH=output\excel\combined_data.xlsx
set TEMPLATE_PATH=C:\Users\matth_9lb83h2\Desktop\TONIK\Coding Projects\competitor-uploader\output\excel\Output.xlsx

echo Using template-based version
echo Using folder: %FOLDER_PATH%
echo Using template: %TEMPLATE_PATH%
echo Output file: %OUTPUT_PATH%

REM Validate folder exists
if not exist "%FOLDER_PATH%" (
    echo Error: Folder does not exist.
    pause
    exit /b 1
)

REM Validate template exists
if not exist "%TEMPLATE_PATH%" (
    echo Error: Template file does not exist.
    pause
    exit /b 1
)

echo.
echo Processing with %SCRIPT_NAME%...
echo.

REM Run the Python script with the provided parameters
python %SCRIPT_NAME% "%FOLDER_PATH%" -o "%OUTPUT_PATH%" -t "%TEMPLATE_PATH%"

echo.
if %errorlevel% equ 0 (
    echo Concatenation completed successfully!
) else (
    echo Concatenation failed.
)

pause
