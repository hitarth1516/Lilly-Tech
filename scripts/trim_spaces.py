import pandas as pd
import glob
import os
import sys

def trim_csv_spaces(file_path):
    """Trim leading and trailing spaces from CSV columns."""
    try:
        # Specify the delimiter as '|' to match the file format
        df = pd.read_csv(file_path, sep='|')
        # Trim spaces from object (string) columns
        df = df.apply(lambda x: x.str.strip() if x.dtype == 'object' else x)
        # Write back to the same file
        df.to_csv(file_path, sep='|', index=False)
        print(f"Successfully trimmed spaces in {file_path}")
    except FileNotFoundError:
        print(f"Error: File not found: {file_path}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error: Failed to trim spaces in {file_path}: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    directory = "INPUT/03_METADATA_FILES/METRIC_SEGMENT_ASSET/CONFIG_FILE/METRIC_CONFIG"
    csv_files = glob.glob(os.path.join(directory, "**/*.csv"), recursive=True)

    if not csv_files:
        print(f"Error: No CSV files found in {directory} or its subdirectories", file=sys.stderr)
        sys.exit(1)

    for file_path in csv_files:
        trim_csv_spaces(file_path)
