"""
Data Analysis Modules

This package contains modules for analyzing CSV and Excel data:
- analyze_csv.py: Tools for analyzing CSV files
- check_csv.py: Validation and checking tools for CSV files
- examine_excel.py: Tools for examining Excel file structure and content
"""

from . import analyze_csv
from . import check_csv
from . import examine_excel

__all__ = [
    'analyze_csv',
    'check_csv',
    'examine_excel',
]
