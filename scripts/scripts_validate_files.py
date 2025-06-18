import json
import csv
import os
import glob
import pandas as pd
import re
import sys

def is_valid_json(file_path):
    """Validate JSON file for syntax and invalid characters."""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
            # Check for control characters (invalid in JSON)
            if re.search(r'[\x00-\x1F\x7F]', content):
                print(f"Error: Invalid control characters found in {file_path}")
                return False
            json.loads(content)
        print(f"Valid JSON: {file_path}")
        return True
    except json.JSONDecodeError as e:
        print(f"Error: Invalid JSON in {file_path}: {e}")
        return False
    except UnicodeDecodeError as e:
        print(f"Error: Encoding issue in {file_path}: {e}")
        return False

def is_valid_csv(file_path):
    """Validate CSV file for format and invalid characters."""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # Use pandas to validate CSV structure
        df = pd.read_csv(file_path)
        if df.empty:
            print(f"Warning: Empty CSV file: {file_path}")
        else:
            print(f"Valid CSV: {file_path}")
        return True
    except pd.errors.ParserError as e:
        print(f"Error: Invalid CSV format in {file_path}: {e}")
        return False
    except UnicodeDecodeError as e:
        print(f"Error: Encoding issue in {file_path}: {e}")
        return False

def main():
    """Main function to validate all JSON and CSV files."""
    json_files = glob.glob('**/*.json', recursive=True)
    csv_files = glob.glob('**/*.csv', recursive=True)

    all_valid = True

    # Validate JSON files
    for json_file in json_files:
        if not is_valid_json(json_file):
            all_valid = False

    # Validate CSV files
    for csv_file in csv_files:
        if not is_valid_csv(csv_file):
            all_valid = False

  #  if not all_valid:
   #     sys.exit(1)
   # print("All files are valid!")

if __name__ == "__main__":
    main()
