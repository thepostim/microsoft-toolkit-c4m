import os
import sys
import argparse
import pandas as pd

def convert_excel_to_csv(input_file, output_file):
    """Convert Excel file to CSV using pandas"""
    df = pd.read_excel(input_file)
    df.to_csv(output_file, index=False)

def main():
    # Create an argument parser for command line inputs
    parser = argparse.ArgumentParser(description='Convert Excel sheets to CSV files.')
    parser.add_argument('input_file', type=str, help='Path to the input Excel file.')
    parser.add_argument('output_file', type=str, help='Path to save the output CSV file.')
    
    # Parse the arguments
    args = parser.parse_args()

    # Check if the input file exists
    if not os.path.isfile(args.input_file):
        print(f"Error: The file '{args.input_file}' does not exist.")
        sys.exit(1)

    # Attempt to convert the Excel file to CSV
    try:
        convert_excel_to_csv(args.input_file, args.output_file)
        print(f"Successfully converted '{args.input_file}' to '{args.output_file}'")
    except Exception as e:
        print(f"An error occurred during conversion: {e}")
        sys.exit(1)

if __name__ == '__main__':
    main()
