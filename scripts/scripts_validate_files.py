import json
import csv
import os
import pandas as pd
import re
import sys
import logging
from openpyxl import Workbook
from datetime import datetime

# Configure logging
timestmp = datetime.now().strftime("%Y%m%d_%H%M%S")
log_filename = f'./validate_files_{timestmp}.log'
logging.basicConfig(filename=log_filename, filemode='w', level=logging.INFO)

def is_valid_csv(file_path):
    """Validate CSV file for format and invalid characters."""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # Use pandas to validate CSV structure with | delimiter
        df = pd.read_csv(file_path, sep='|', quoting=csv.QUOTE_NONE)
        if df.empty:
            print(f"Warning: Empty CSV file: {file_path}")
            logging.warning(f"Empty CSV file: {file_path}")
            return False
        else:
            print(f"Valid CSV: {file_path}")
            logging.info(f"Valid CSV: {file_path}")
            return True
    except pd.errors.ParserError as e:
        print(f"Error: Invalid CSV format in {file_path}: {e}")
        logging.error(f"Invalid CSV format in {file_path}: {e}")
        return False
    except UnicodeDecodeError as e:
        print(f"Error: Encoding issue in {file_path}: {e}")
        logging.error(f"Encoding issue in {file_path}: {e}")
        return False

def check_duplicates(metric_config_df, master_config_df):
    """Check for duplicates in master and metric config files."""
    status1, status2 = 'SUCCESS', 'SUCCESS'
    master_config_active_query_df = master_config_df[master_config_df['TYPE'] == 'QUERY']
    metric_config_active_entry_df = metric_config_df[['PARAMETER_VALUE', 'CONFIG_TYPE', 'DESCRIPTION']]

    filtered_df = master_config_active_query_df[
        ~master_config_active_query_df['Identifier_Value1'].isin(['NPP_OBU_MASTER_DERIVED', 'NPP_OBU_MASTER_PASSTHROUGH'])
    ]
    duplicates_df1 = pd.DataFrame() if not filtered_df['Identifier_Value1'].duplicated().any() else \
        pd.DataFrame(filtered_df[filtered_df['Identifier_Value1'].duplicated(keep=False)]['Identifier_Value1'].unique(), columns=['Duplicate IdentifierValue1 Values'])
    if not duplicates_df1.empty:
        status1 = 'FAILED'

    filtered_df_obu = master_config_active_query_df[
        master_config_active_query_df['Identifier_Value1'].isin(['NPP_OBU_MASTER_DERIVED', 'NPP_OBU_MASTER_PASSTHROUGH'])
    ]
    if filtered_df_obu.duplicated(subset=['Identifier_Value1', 'Identifier_Value2']).any():
        duplicates_obu = filtered_df_obu[filtered_df_obu.duplicated(subset=['Identifier_Value1', 'Identifier_Value2'], keep=False)][['Identifier_Value1', 'Identifier_Value2']].drop_duplicates()
        duplicates_df1 = pd.concat([duplicates_df1, duplicates_obu['Identifier_Value1']], ignore_index=True)
        status1 = 'FAILED'

    duplicates_df2 = pd.DataFrame() if not metric_config_active_entry_df.duplicated().any() else \
        metric_config_active_entry_df[metric_config_active_entry_df.duplicated(keep=False)].copy()
    if not duplicates_df2.empty:
        status2 = 'FAILED'

    return {
        'MasterConfig_DuplicateCheck': (status1, duplicates_df1),
        'MetricConfig_DuplicateCheck': (status2, duplicates_df2)
    }

def validate_master_files(master_config_df):
    """Validate presence of expected master files."""
    status = 'SUCCESS'
    master_file_list = ['customer', 'product', 'rep_alignment', 'intermediate_output', 'npp_output', 'pp_output', 'seg_output']
    active_files = master_config_df[(master_config_df['TYPE'] == 'INPUT_FILE') | (master_config_df['TYPE'] == 'OUTPUT_FILE')]['Identifier_Name1'].dropna().unique()
    missing_files = [f for f in master_file_list if f.lower() not in [file.lower() for file in active_files]]
    missing_files_df = pd.DataFrame(missing_files, columns=['Master Input Missing Files']) if missing_files else pd.DataFrame()
    if missing_files:
        status = 'FAILED'
    return {'MissingFilesCheck': (status, missing_files_df)}

def extract_table_names(sql_query):
    """Extract table names from SQL queries."""
    table_name_pattern = re.compile(r'\bFROM\s+(\w+)|\bJOIN\s+(\w+)', re.IGNORECASE)
    matches = table_name_pattern.findall(sql_query)
    return list(set(name for match in matches for name in match if name))

def validate_table_names(master_config_df):
    """Validate tables used in SQL queries."""
    status = 'SUCCESS'
    active_df = master_config_df.copy()
    active_df['Extracted_Tables'] = active_df['Identifier_Value2'].apply(lambda x: extract_table_names(x) if pd.notnull(x) else [])
    extracted_tables = [t.lower() for sublist in active_df['Extracted_Tables'] for t in sublist if t.lower() != 'temp']
    unique_values = [v.lower() for v in master_config_df['File_Name'].dropna().unique()]
    missing_tables = [t for t in extracted_tables if t not in unique_values]
    missing_tables_df = pd.DataFrame(missing_tables, columns=['SQL INPUT MISSING TABLES']) if missing_tables else pd.DataFrame()
    if missing_tables:
        status = 'FAILED'
    return {'MissingSQLTablesCheck': (status, missing_tables_df)}

def print_inactive_files(master_config_df):
    """List inactive files in master config."""
    status = 'SUCCESS' if len(master_config_df[master_config_df["Status"] == 'InActive']['File_Name'].dropna().unique()) == 0 else 'USER_EVALUATE'
    inactive_files_df = pd.DataFrame(master_config_df[master_config_df["Status"] == 'InActive']['File_Name'].dropna().unique(), columns=['INACTIVE INPUT FILE LIST'])
    return {'InactiveInputFiles-MasterConfig': (status, inactive_files_df)}

def print_inactive_metrics(master_config_df):
    """List inactive metrics in master config."""
    status = 'SUCCESS' if len(master_config_df[master_config_df["Status"] == 'InActive']['Identifier_Name2'].dropna().unique()) == 0 else 'USER_EVALUATE'
    inactive_metrics_df = pd.DataFrame(master_config_df[master_config_df["Status"] == 'InActive']['Identifier_Name2'].dropna().unique(), columns=['INACTIVE INPUT FILE LIST'])
    return {'InactiveScripts-MasterConfig': (status, inactive_metrics_df)}

def validate_sql_template_strings(master_config_df):
    """Validate template strings against dynamic statements."""
    status = 'SUCCESS'
    failure_df = pd.DataFrame(columns=['IdentifierValue', 'Number of Template Strings', 'Number of Dynamic Statements / Trigger Names'])
    for index, row in master_config_df.iterrows():
        if row['Identifier_Name1'] in ('NPP', 'PP', 'SEG'):
            template_pattern = re.compile(r"#template_string_\d+")
            unique_templates = set(template_pattern.findall(row['Identifier_Value2'] or ''))
            num_templates = len(unique_templates)
            num_statements = len([s for s in (row['DynamicStatements'] or '').split('||') if s.strip()]) if row['DynamicStatements'] else 0
            if num_templates != num_statements:
                status = 'FAILED'
                failure_df = pd.concat([failure_df, pd.DataFrame([{
                    'IdentifierValue': row['Identifier_Value1'],
                    'Number of Template Strings': num_templates,
                    'Number of Dynamic Statements / Trigger Names': num_statements
                }])], ignore_index=True)
    return {'TemplateString-DynamicStatement': (status, failure_df)}

def validate_quotes_in_sql(master_config_df):
    """Validate double quotes in SQL queries."""
    status = 'SUCCESS'
    failure_df = pd.DataFrame(columns=['IdentifierName', 'IdentifierValue'])
    for index, row in master_config_df.iterrows():
        if row['Identifier_Name1'] in ('NPP', 'PP', 'SEG'):
            if row['Identifier_Value2'] and (row['Identifier_Value2'].startswith('"') or row['Identifier_Value2'].endswith('"')):
                status = 'FAILED'
                failure_df = pd.concat([failure_df, pd.DataFrame([{
                    'IdentifierName': row['Identifier_Name2'],
                    'IdentifierValue': row['Identifier_Value1']
                }])], ignore_index=True)
    return {'DoubleQuoteSQLScripts': (status, failure_df)}

def validate_quotes_in_metric_config(metric_config_df):
    """Validate double quotes in metric config values."""
    status = 'SUCCESS'
    failure_df = pd.DataFrame(columns=['ParameterValue', 'ConfigType'])
    for index, row in metric_config_df.iterrows():
        if row['PARAMETER_TYPE'] == 'DYNAMIC_METRIC_ID':
            if row['CONFIG_VALUE'] and (row['CONFIG_VALUE'].startswith('"') or row['CONFIG_VALUE'].endswith('"')):
                status = 'FAILED'
                failure_df = pd.concat([failure_df, pd.DataFrame([{
                    'ParameterValue': row['PARAMETER_VALUE'],
                    'ConfigType': row['CONFIG_TYPE']
                }])], ignore_index=True)
    return {'DoubleQuoteMetricNames': (status, failure_df)}

def validate_active_metrics(metric_config_df, env_data):
    """Validate active metrics against environment data (simplified)."""
    status = 'SUCCESS'
    # Assuming env_data has an 'active_metrics' key; adjust if different
    env_metrics = env_data.get('active_metrics', [])
    active_metrics = metric_config_df[metric_config_df['ACTIVE'] == 'Y']['DESCRIPTION'].dropna().unique()
    discrepancies = [m for m in active_metrics if m not in env_metrics]
    discrepancy_df = pd.DataFrame(discrepancies, columns=['Metrics-InMetricAsset-NotinENV']) if discrepancies else pd.DataFrame()
    if discrepancies:
        status = 'FAILED'
    return {'Metrics-InMetricAsset-NotinENV': (status, discrepancy_df)}

def print_inactive_metrics_list(metric_config_df):
    """List inactive metrics in metric config."""
    status = 'SUCCESS' if len(metric_config_df[metric_config_df["ACTIVE"] == 'N']['DESCRIPTION'].dropna().unique()) == 0 else 'USER_EVALUATE'
    inactive_metrics_df = pd.DataFrame(metric_config_df[metric_config_df["ACTIVE"] == 'N']['DESCRIPTION'].dropna().unique(), columns=['INACTIVE METRICS LIST'])
    return {'InactiveMetrics-MetricConfig': (status, inactive_metrics_df)}

def validate_activeinsights_metrics(master_config_df, metric_config_df, env_data):
    """Validate metrics in active insights (simplified)."""
    status1, status2 = 'SUCCESS', 'SUCCESS'
    # Simplified logic; adjust based on actual rules_df and active_rules_extract_df
    active_master_df = pd.DataFrame()
    for index, row in master_config_df.iterrows():
        if row['Identifier_Name1'] in ('NPP', 'PP', 'SEG') and pd.notna(row['DynamicStatements']):
            statements = row['DynamicStatements'].split('||')
            names = [s.split('#')[0] for s in statements]
            active_master_df = pd.concat([active_master_df, pd.DataFrame({
                'Identifier_Value1': row['Identifier_Value1'],
                'NameBeforeHash': [n + '#' for n in names]
            })], ignore_index=True)

    merged_df = pd.merge(active_master_df, metric_config_df, left_on=['Identifier_Value1', 'NameBeforeHash'],
                        right_on=['CONFIG_TYPE', 'PARAMETER_VALUE'], how='inner')
    metric_asset_active = merged_df[merged_df['ACTIVE'] == 'Y']['DESCRIPTION']
    env_active = env_data.get('active_metrics', [])
    non_matching = [m for m in metric_asset_active if m not in env_active]
    non_matching_df = pd.DataFrame(non_matching, columns=['Metrics-InOC-NotActiveMetrics']) if non_matching else pd.DataFrame()
    if non_matching:
        status1 = 'FAILED'

    result_df = pd.DataFrame(columns=['suggestionName', 'colName'])  # Placeholder; adjust with actual data
    if not result_df.empty:
        status2 = 'FAILED'

    return {
        'Metrics-InOC-NotActiveMetrics': (status1, non_matching_df),
        'Metrics-InOC-NotMetricAsset': (status2, result_df)
    }

def split_dynamic_statements_templateid(row):
    """Split dynamic statements to extract template IDs."""
    split_df = pd.DataFrame(columns=['IdentifierValue', 'TemplateString'])
    if pd.isna(row['DynamicStatements']):
        return split_df
    statements = row['DynamicStatements'].split('||')
    for part in statements:
        if '#template_id' in part:
            metric_name_match = re.search(r'(\w+)#', part)
            if metric_name_match:
                template_string = metric_name_match.group(1)
                split_df = pd.concat([split_df, pd.DataFrame([{
                    'IdentifierValue': row['Identifier_Value1'],
                    'TemplateString': template_string
                }])], ignore_index=True)
    return split_df

def validate_dynamic_statements_vaetemplateids(master_config_df, metric_config_df):
    """Validate VAE template IDs in dynamic statements."""
    status = 'SUCCESS'
    failure_df = pd.DataFrame(columns=['IdentifierValue', 'TemplateString'])
    master_config_active_query_df = master_config_df[master_config_df['TYPE'] == 'QUERY']
    master_config_validation_df = master_config_active_query_df[['Identifier_Value1', 'DynamicStatements']]
    metric_config_validation_df = metric_config_df[metric_config_df['PARAMETER_TYPE'] == 'DYNAMIC_METRIC_ID'][['PARAMETER_VALUE', 'CONFIG_TYPE', 'CONFIG_VALUE']].drop_duplicates()

    expanded_rows = []
    for _, row in master_config_validation_df.iterrows():
        split_df = split_dynamic_statements_templateid(row)
        expanded_rows.append(split_df)
    expanded_df = pd.concat(expanded_rows, ignore_index=True) if expanded_rows else pd.DataFrame(columns=['IdentifierValue', 'TemplateString'])
    master_config_expanded_df = expanded_df[['IdentifierValue', 'TemplateString']].drop_duplicates()

    for index, row in master_config_expanded_df.iterrows():
        count = 0
        for index1, row1 in metric_config_validation_df.iterrows():
            if (row['IdentifierValue'] == row1['CONFIG_TYPE']) and (row['TemplateString'] + '#' == row1['PARAMETER_VALUE']):
                if '"template_id"' in row1['CONFIG_VALUE']:
                    count += 1
        if count == 0:
            status = 'FAILED'
            failure_df = pd.concat([failure_df, pd.DataFrame([{
                'IdentifierValue': row['IdentifierValue'],
                'TemplateString': row['TemplateString'] + '#'
            }])], ignore_index=True)

    return {'MASTER VS METRIC TEMPLATEID': (status, failure_df.drop_duplicates())}

def flag_double_quote(master_config_df, metric_config_df):
    """Flag rows with double or single quotes at start/end."""
    status1, status2 = 'SUCCESS', 'SUCCESS'
    failure_df1 = pd.DataFrame(columns=['ID Number'])
    failure_df2 = pd.DataFrame(columns=['ROW DESCRIPTION NAME'])

    master_check_df = master_config_df
    metric_check_df = metric_config_df

    mask1 = master_check_df.applymap(lambda x: isinstance(x, str) and ('""' in x or (x.startswith('"') or x.endswith('"'))))
    mask2 = metric_check_df.applymap(lambda x: isinstance(x, str) and ('""' in x or (x.startswith('"') or x.endswith('"'))))

    master_check_df['Flag'] = mask1.any(axis=1)
    metric_check_df['Flag'] = mask2.any(axis=1)

    for index, row in master_check_df.iterrows():
        if row['Flag']:
            status1 = 'FAILED'
            failure_df1 = pd.concat([failure_df1, pd.DataFrame([{'ID Number': row.get('ID', 'N/A')}])], ignore_index=True)
    for index, row in metric_check_df.iterrows():
        if row['Flag']:
            status2 = 'FAILED'
            failure_df2 = pd.concat([failure_df2, pd.DataFrame([{'ROW DESCRIPTION NAME': row.get('DESCRIPTION', 'N/A')}])], ignore_index=True)

    return {
        'MASTER CONFIG DOUBLEQUOTES ROW': (status1, failure_df1),
        'METRIC CONFIG DOUBLEQUOTES ROW': (status2, failure_df2)
    }

def generate_report(results, output_path):
    """Generate an Excel report with validation results."""
    wb = Workbook()
    for sheet, (status, df) in results.items():
        ws = wb.create_sheet(title=sheet) if sheet != 'Validation Report' else wb.active
        ws.append(["File" if sheet == 'Validation Report' else "Check", "Status", "Details"])
        if sheet == 'Validation Report':
            for file_path, (status, details) in results.get('initial_validation', {}).items():
                ws.append([os.path.basename(file_path), status, details])
        else:
            ws.append([f"{sheet} Result", status, "See details below"])
            if not df.empty:
                for column in df.columns:
                    ws.append([column] + df[column].tolist())
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    logging.info(f"Validation report saved to {output_path}")
    print(f"Validation report saved to {output_path}")

def main():
    """Main function to validate specific CSV files and perform additional checks."""
    try:
        # Get file paths from environment variables
        metric_config_path = os.environ.get('METRIC_CONFIG_PATH', '/work/METRIC_SEGMENT_ASSET/CONFIG_FILE/METRIC_CONFIG.File.csv')
        master_config_path = os.environ.get('MASTER_CONFIG_PATH', '/work/METRIC_SEGMENT_ASSET/CONFIG_FILE/MASTER_CONFIG.File.csv')
        output_path = os.environ.get('OUTPUT_PATH', '/work/output/dqm_report.xlsx')

        # Read data for additional validations
        print("reading metric_config_df")
        metric_config_df = pd.read_csv(metric_config_path, sep='|', quoting=csv.QUOTE_NONE)
        print("reading master_config_df")
        master_config_df = pd.read_csv(master_config_path, sep=',', quoting=csv.QUOTE_MINIMAL)
        env_data = {}  # Placeholder; adjust with actual env data source

        # Run additional validations
        results.update(check_duplicates(metric_config_df, master_config_df))
        results.update(validate_master_files(master_config_df))
        results.update(validate_table_names(master_config_df))
        results.update(print_inactive_files(master_config_df))
        results.update(print_inactive_metrics(master_config_df))
        results.update(validate_sql_template_strings(master_config_df))
        results.update(validate_quotes_in_sql(master_config_df))
        results.update(validate_quotes_in_metric_config(metric_config_df))
        results.update(validate_active_metrics(metric_config_df, env_data))
        results.update(print_inactive_metrics_list(metric_config_df))
        results.update(validate_activeinsights_metrics(master_config_df, metric_config_df, env_data))
        results.update(validate_dynamic_statements_vaetemplateids(master_config_df, metric_config_df))
        results.update(flag_double_quote(master_config_df, metric_config_df))

        # Generate report
        generate_report(results, output_path)

        if not all_valid:
            sys.exit(1)
        print("All validated files and checks are valid!")

    except Exception as e:
        logging.error(f"Execution failed: {e}")
        print(f"Execution failed: {e}")
        sys.exit(1)
    finally:
        logging.shutdown()

if __name__ == "__main__":
    main()
