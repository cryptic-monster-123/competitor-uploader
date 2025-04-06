"""
Configuration Module

This package contains configuration files and utilities for the Excel Concatenator:
- column_mapping_example.json: Example mapping configuration for column standardization
"""

import os
import json

def load_column_mapping(mapping_file='column_mapping_example.json'):
    """
    Load column mapping configuration from a JSON file.
    
    Args:
        mapping_file (str): Name of the mapping file in the config directory
        
    Returns:
        dict: Column mapping configuration
    """
    current_dir = os.path.dirname(os.path.abspath(__file__))
    mapping_path = os.path.join(current_dir, mapping_file)
    
    try:
        with open(mapping_path, 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        raise FileNotFoundError(f"Column mapping file not found: {mapping_path}")
    except json.JSONDecodeError:
        raise ValueError(f"Invalid JSON in column mapping file: {mapping_path}")

__all__ = ['load_column_mapping']
