import os
import logging
import pandas as pd

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def ensure_directory_exists(directory):
    """Ensure that the specified directory exists; create it if it does not."""
    if not directory:  # Handle empty string case
        return
    if not os.path.exists(directory):
        try:
            os.makedirs(directory)
            logging.info(f"Created directory: {directory}")
        except Exception as e:
            logging.error(f"Failed to create directory {directory}: {e}")
            raise

def read_excel_file(file_path):
    """Read an Excel file and return a DataFrame.
    
    Raises:
        FileNotFoundError: If the file does not exist.
        ValueError: If the file format is incorrect.
    """
    if not os.path.isfile(file_path):
        logging.error(f"File not found: {file_path}")
        raise FileNotFoundError(f"The file {file_path} does not exist.")
    
    try:
        df = pd.read_excel(file_path)
        logging.info(f"Successfully read Excel file: {file_path}")
        return df
    except ValueError as e:
        logging.error(f"Error reading the Excel file {file_path}: {e}")
        raise

def save_to_csv(data_frame, output_path):
    """Save the DataFrame to a CSV file.
    
    Raises:
        Exception: If saving the CSV fails.
    """
    try:
        data_frame.to_csv(output_path, index=False)
        logging.info(f"Successfully saved CSV file: {output_path}")
    except Exception as e:
        logging.error(f"Failed to save CSV file {output_path}: {e}")
        raise

def convert_excel_to_csv(excel_path, csv_path):
    """Convert an Excel file to a CSV file.
    
    Args:
        excel_path (str): Path to the Excel file.
        csv_path (str): Path where the CSV file will be saved.
    
    TODO: Add functionality to specify sheet names.
    """
    csv_dir = os.path.dirname(csv_path)
    if csv_dir:  # Only create directory if dirname is not empty
        ensure_directory_exists(csv_dir)
    df = read_excel_file(excel_path)
    save_to_csv(df, csv_path)

# Example usage (commented out)
# if __name__ == "__main__":
#     convert_excel_to_csv('example.xlsx', 'output/example.csv')
