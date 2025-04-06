"""
Excel Concatenator Modules

This package contains modules for concatenating Excel files with different approaches:
- Basic concatenator for files with identical column structures
- Enhanced concatenator with column mapping capabilities
- Template-based concatenator that uses an Excel template for output structure
- Excel to CSV concatenator for converting and combining Excel files to CSV
"""

from . import excel_concatenator
from . import excel_concatenator_enhanced
from . import excel_concatenator_template
from . import excel_to_csv_concatenator

__all__ = [
    'excel_concatenator',
    'excel_concatenator_enhanced',
    'excel_concatenator_template',
    'excel_to_csv_concatenator',
]
