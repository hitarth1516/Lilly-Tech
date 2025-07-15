import pandas as pd
import os
import sys
import logging
import csv # Import csv for quoting
from datetime import datetime

# Configure logging
timestmp = datetime.now().strftime("%Y%m%d_%H%M%S")
log_filename = f'./trim_spaces_{timestmp}.log'
logging.basicConfig(filename=log_filename, filemode='w', level=logging.INFO)

def trim_csv_spaces(file_path):
    """
    Reads a CSV file, trims spaces from all string columns,
    and overwrites the original file.
    """
    try:
        if not os.path.exists(file_path):
            logging.error(f"Error: File not found: {file_path}")
            print(f"Error: File not found: {file_path}")
            sys.exit(1)

        # Read the CSV file using '|' as delimiter and no quoting
        # Assuming Metric_Config_File.csv is pipe-delimited
        df = pd.read_csv(file_path, sep='|', quoting=csv.QUOTE_NONE)
        logging.info(f"Successfully read CSV: {file_path}")

        # Trim spaces from all string columns
        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].astype(str).str.strip()
        logging.info("Trimmed spaces from string columns.")

        # Overwrite the original CSV file with trimmed data
        df.to_csv(file_path, sep='|', index=False, quoting=csv.QUOTE_NONE)
        logging.info(f"Successfully wrote trimmed data back to: {file_path}")
        print(f"Successfully trimmed spaces in {file_path}")

    except Exception as e:
        logging.error(f"Failed to trim spaces in {file_path}: {e}")
        print(f"Failed to trim spaces in {file_path}: {e}")
        sys.exit(1)
    finally:
        logging.shutdown()

if __name__ == "__main__":
    # Updated default path to be relative to the repository root and reflect deeper nesting
    # This default will be used if the environment variable is not set by GitHub Actions
    metric_config_path = os.environ.get('METRIC_CONFIG_PATH', 'INPUT/03_METADATA_FILES/METRIC_SEGMENT_ASSET/CONFIG_FILE/METRIC_CONFIG/Metric_Config_File.csv')
    if not metric_config_path:
        logging.error("METRIC_CONFIG_PATH environment variable not set.")
        print("Error: METRIC_CONFIG_PATH environment variable not set.")
        sys.exit(1)
    trim_csv_spaces(metric_config_path)

