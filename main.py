# Code that 'inspired' me: https://stackoverflow.com/a/73540079

import logging
import os
import json

import pandas as pd
# import openpyxl

from lib import *

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
rename_columns = config['DEFAULT'].get('rename_columns', False)
reduce_field_whitespace : bool = config['DEFAULT'].get('reduce_field_whitespace', False)
assert all((isinstance(var, bool) for var in [quiet, rename_columns, reduce_field_whitespace,])), "One of the flags are not a bool"

# FUNCTIONS


_df_rename__rename_redundants_dict = dict()
def _df_rename__rename_redundants():
    """ 
    There is a method in panda called 'df.rename'.
    It requires that the data be structured as {old_field_name:new_field_name}
    for each field to rename.
    The configuration variable "Combine_Fields" has a different structure and 
    must be restructured for the .rename method.
    Ref: https://stackoverflow.com/a/11354850
      """
    global _df_rename__rename_redundants_dict
    for k,v_list in config["Combine_Fields"].items():
        for val in v_list:
            _df_rename__rename_redundants_dict.update({val:k})
_df_rename__rename_redundants()
# assert len(_df_rename__rename_redundants_dict) > 0, "This should not be empty."
logging.debug(_df_rename__rename_redundants_dict)

def reduce_field_whitespace_custom(str1):
    """ Wrapper for 'dict__old_to_unique_new_field' """
    return dict__old_to_unique_new_field(str1, strip=True, titleize=True)

def main():
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
                        # Remove whitespace in all columns
                        if reduce_field_whitespace:
                            # df.rename(reduce_whitespace_to_one_space, axis="columns", inplace=True)
                            columns = df.columns.to_list()
                            new_column_names_dict = reduce_field_whitespace_custom(columns)
                            df.rename(columns=new_column_names_dict, inplace=True)
                        if rename_columns:
                            df.rename(columns=_df_rename__rename_redundants_dict, inplace=True)
                        # Add to be concatenated
                        frames.append(df)
            except PermissionError:
                logging.warning(f"The file '{file}' cannot be opened.")
        # Create one sheet containing all spreadsheet data.
        result = pd.concat(frames)
        result.to_excel(writer, index=False)

    logging.info("Reading files complete. Press ENTER to continue.")
    if not quiet: dali=input()

if __name__ == '__main__':
    main()