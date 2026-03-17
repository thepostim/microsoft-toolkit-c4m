# excel_to_csv_converter.py

import pandas as pd
import os
import re

def convert_excel_to_csv(excel_file, sheet_name, output_dir):
    """
    Converts a specified sheet in an Excel file to CSV format.

    Parameters:
        excel_file (str): Path to the Excel file.
        sheet_name (str): Name of the sheet to convert.
        output_dir (str): Directory to save the CSV file.

    Raises:
        FileNotFoundError: If the Excel file does not exist.
        ValueError: If the specified sheet does not exist in the Excel file.
    """
    
    # Check if the Excel file exists
    if not os.path.isfile(excel_file):
        raise FileNotFoundError(f"The file '{excel_file}' does not exist.")
    
    # Get available sheet names to validate
    try:
        excel_data = pd.ExcelFile(excel_file)
        available_sheets = excel_data.sheet_names
    except Exception as e:
        raise Exception(f"Error reading the Excel file: {e}")

    # Check if the specified sheet exists
    if sheet_name not in available_sheets:
        raise ValueError(f"The sheet '{sheet_name}' does not exist in the Excel file.")
    
    # Read the specified sheet
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
    except Exception as e:
        raise Exception(f"Error reading sheet '{sheet_name}': {e}")

    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)
    
    # Sanitize sheet name for filename
    safe_sheet_name = re.sub(r'[<>:"/\\|?*]', '_', sheet_name)
    
    # Prepare output CSV file path
    output_file = os.path.join(output_dir, f"{safe_sheet_name}.csv")

    # Save the DataFrame to a CSV file
    try:
        df.to_csv(output_file, index=False)
        print(f"Successfully converted '{sheet_name}' to '{output_file}'")
    except Exception as e:
        raise Exception(f"Error writing to CSV file: {e}")

# TODO: Add support for multiple sheets conversion
# TODO: Implement a command-line interface for convenience
# TODO: Add logging instead of print statements for better error tracking
