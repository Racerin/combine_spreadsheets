# Code that 'inspired' me: https://stackoverflow.com/a/73540079

import logging
import os
import configparser
import json

import pandas as pd
import openpyxl

# Logging
logging.basicConfig(
    # filename='example.log', 
    encoding='utf-8', level=logging.INFO
    )

logging.info("Copying excel sheets from multiple files to one file")

# CONFIGURATION
config = dict()
with open("config.json", "r") as jsonfile:
    config = json.load(jsonfile)
    logging.info("Read 'config.json' successful.")

# CONFIGURATIONS
quiet = config['DEFAULT'].get('quiet', False)
excel_files_dir = config['DEFAULT'].get('excel_files_dir', os.getcwd())
output_excel_file_name = config['DEFAULT'].get('output_excel_file_name', 'Combined/combined_file.xlsx')

# FUNCTIONS
def ensure_str_unique(input_str:str, str_list:list[str]=[]):
    """ 
    If 'input_str' is inside of list, update it's name and try again to add it.
    Return the accepted str.
    'str_list' automatically updates with each tested 'input_str'
    """
    n_str_list = len(str_list)
    for _ in range(n_str_list + 1):
        if input_str in str_list:
            # Append to 'input_str' name and try again
            input_str = input_str + "X"
        else:
            str_list.append(input_str)
            break
    else:
        raise OverflowError("Code Error: Str options exceeded.")
    return input_str

def is_excel_file(file_name):
    return any((file_name.endswith(ext) for ext in ('.xls','.xlsx')))

_df_rename_rename_redundants_dict = dict()
def _df_rename_rename_redundants():
    """ 
    There is a method in panda called 'df.rename'.
    It requires that the data be structured as {old_field_name:new_field_name}
    for each field to rename.
    The configuration variable "Combine_Fields" has a different structure and 
    must be restructured for the .rename method.
      """
    global _df_rename_rename_redundants_dict
    for k,v_list in config["Combine_Fields"].items():
        for val in v_list:
            _df_rename_rename_redundants_dict.update({val:k})
_df_rename_rename_redundants()
assert len(_df_rename_rename_redundants_dict) > 0, "This should not be empty."

# Check-out each excel file and then combine
frames = []
files = os.listdir(excel_files_dir)
with pd.ExcelWriter(output_excel_file_name,mode='w') as writer:
    for file in files:
        try:
            # Is it an excel file?
            if is_excel_file(file):
                file_path = os.path.join(excel_files_dir, file)
                excel_file = pd.ExcelFile(file_path)
                sheets = excel_file.sheet_names
                # loop through sheets inside an Excel file
                for sheet in sheets:
                    logging.info(f"File: {file} | Sheet: {sheet}")
                    df = excel_file.parse(sheet_name=sheet)
                    # Ignore empty dataframes
                    if df.empty: continue
                    # Rename redundants Excel Fields
                    df.rename(columns=_df_rename_rename_redundants_dict)
                    # Add to be concatenated
                    frames.append(df)
        except PermissionError:
            logging.warning(f"The file '{file}' cannot be opened.")
    # Create one sheet containing all spreadsheet data.
    result = pd.concat(frames)
    result.to_excel(writer, index=False)

logging.info("Reading files complete. Done.")
if not quiet: dali=input()
