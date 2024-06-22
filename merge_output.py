# Import Libraries
import argparse
import os
from openpyxl import load_workbook, Workbook
import pandas as pd
from pandas import ExcelFile

# Functions
def parse_arguments():
    parser = argparse.ArgumentParser(description='Run the prediction model on input Excel file and save the output.')
    parser.add_argument('output_file', type=str, help='Output Excel file name')
    parser.add_argument('all_file', type=str, help='All Projects Output Excel file name')
    return parser.parse_args()

# read output file
def read_excel_file(file_name):
    file_path = 'output/{}.xlsx'.format(file_name)
    return pd.ExcelFile(file_path)

# merge output file with all projects file
def merge_output(output_file_name,all_file_name):
    output_folder_path = 'output'
    output_file = '{}.xlsx'.format(output_file_name)
    all_projects_file = '{}.xlsx'.format(all_file_name)
    os.makedirs(output_folder_path, exist_ok=True)

    all_projects_path = os.path.join(output_folder_path, all_projects_file)

    # Retrieve the ExcelFile object and then parse the first sheet into a DataFrame
    output_excel_file = read_excel_file(output_file_name)
    output_df = pd.read_excel(output_excel_file, sheet_name=0)

    # check if all projects file is already exist yet
    if os.path.exists(all_projects_path):
        wb = load_workbook(all_projects_path)
        if 'Sheet' in wb.sheetnames:
          del wb['Sheet']
    else:
        # If the Excel file doesn't exist, create a new workbook
        wb = Workbook()

    base_sheet_name = os.path.splitext(output_file_name)[0]
    sheet_name = base_sheet_name
    sheet_index = 1
    # Find a unique sheet name
    while sheet_name in wb.sheetnames:
        sheet_index += 1
        sheet_name = f"{base_sheet_name} ({sheet_index})"

    # Remove the default 'Sheet' if it exists and has no data
    if 'Sheet' in wb.sheetnames:
        sheet = wb['Sheet']
        if not any(sheet.iter_rows(min_row=2, max_row=2, min_col=1, max_col=1)):
            del wb['Sheet']

    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(title=sheet_name)

    for r in dataframe_to_rows(output_df, index=False, header=True):
        ws.append(r)

    wb.save(all_projects_path)
    print(f"Data from {output_file} merged into {all_projects_file}")

def dataframe_to_rows(df, index=True, header=True):
    import itertools
    rows = itertools.chain(
        ([df.columns.values.tolist()] if header else []),
        df.itertuples(index=index, name=None)
    )
    return rows

def main():
    args = parse_arguments()
    output_file_name = os.path.splitext(args.output_file)[0]
    all_file_name = os.path.splitext(args.all_file)[0]

    merge_output(output_file_name,all_file_name)

if __name__ == "__main__":
   main()