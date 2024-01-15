# Code that 'inspired' me: https://stackoverflow.com/a/73540079

import logging
import os
import configparser

import pandas as pd
import openpyxl

logging.basicConfig(
    # filename='example.log', 
    encoding='utf-8', level=logging.INFO
    )

logging.info("Copying excel sheets from multiple files to one file")

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

# Setup Configuration reader
config = configparser.ConfigParser()
config.read('config.ini')

# Configuration inputs
quiet = config['DEFAULT'].get('quiet', False)
if isinstance(quiet, str):
    quiet = 'true'==quiet
cwd = os.getcwd(); # os.path.abspath('')
excel_files_dir = config['DEFAULT'].get('excel_files_dir', cwd)
output_excel_file_name = config['DEFAULT'].get('output_excel_file_name', 'Combined/combined_file.xlsx')

# Panda files setup
df_total = pd.DataFrame()
df_total.to_excel(output_excel_file_name) #create a new file
workbook=openpyxl.load_workbook(output_excel_file_name)
ss_sheet = workbook['Sheet1']
ss_sheet.title = 'TempExcelSheetForDeleting'
workbook.save(output_excel_file_name)

files = os.listdir(excel_files_dir)

# Check-out each excel file and then combine
""" 
for file in files:
    try:
        # Is it an excel file?
        if any((file.endswith(ext) for ext in ('.xls','.xlsx'))):
            file_path = os.path.join(excel_files_dir, file)
            excel_file = pd.ExcelFile(file_path)
            sheets = excel_file.sheet_names
            # loop through sheets inside an Excel file
            for sheet in sheets:
                str_sheet = str(sheet)
                logging.info(f"{file}, {sheet}")
                df = excel_file.parse(sheet_name=sheet)
                with pd.ExcelWriter(output_excel_file_name,mode='a') as writer:  
                    sheet_name = compare_str_list(str_sheet)
                    if sheet_name != str_sheet:
                        logging.info(f"The sheet named '{str_sheet}' changed to '{sheet_name}'.")
                    df.to_excel(writer, sheet_name=str(sheet_name), index=False)
                #df.to_excel(output_excel_file_name, sheet_name=f"{sheet}")
    except PermissionError:
        logging.warning(f"The file '{file}' cannot be opened.") 
"""
frames = []
with pd.ExcelWriter(output_excel_file_name,mode='a') as writer:
    for file in files:
        try:
            # Is it an excel file?
            if any((file.endswith(ext) for ext in ('.xls','.xlsx'))):
                file_path = os.path.join(excel_files_dir, file)
                excel_file = pd.ExcelFile(file_path)
                sheets = excel_file.sheet_names
                # loop through sheets inside an Excel file
                for sheet in sheets:
                    str_sheet = str(sheet)
                    logging.info(f"{file}, {sheet}")
                    df = excel_file.parse(sheet_name=sheet)
                    frames.append(df)
        except PermissionError:
            logging.warning(f"The file '{file}' cannot be opened.") 
    result = pd.concat(frames)
    # df.to_excel(writer, sheet_name='sheet1', index=False)
    result.to_excel(writer, sheet_name='sheet1', index=False)


workbook=openpyxl.load_workbook(output_excel_file_name)
std=workbook["TempExcelSheetForDeleting"]
workbook.remove(std)
workbook.save(output_excel_file_name)
logging.info("Loaded, press ENTER to end")
if not quiet: dali=input()
#df_total.to_excel(output_excel_file_name)
logging.info("Done")
if not quiet: dali=input()
