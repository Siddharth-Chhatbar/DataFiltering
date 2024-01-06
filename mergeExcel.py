# Program to merge multiple excel files into one

import pandas as pd
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
import sqlite3
def jsonToExcel(files_json):
    # Convert json to excel
    for f in files_json:
        data = pd.read_json(f, lines=True)
        # Save the file as xlsx
        data.to_excel(f[:-4] + 'xlsx', index=None, header=True)
    
    input("Files saved as xlsx. Press any key to continue...")
    main()

def csvToExcel(files_csv):
    # Convert csv to excel
    for f in files_csv:
        data = pd.read_csv(f)
        data.to_excel(f[:-3] + 'xlsx', index=None, header=True)
    input("Files saved as xlsx. Press any key to continue...")
    main()


def mergeExcel(files_xls):
    # Create an empty list to store the dataframes
    df_list = []

    # Loop through the files and append data from all sheets to the list
    for f in files_xls:
        xls = pd.ExcelFile(f)
        sheet_names = xls.sheet_names  # Get all sheet names in the Excel file

        # Loop through each sheet in the file
        for sheet_name in sheet_names:
            data = pd.read_excel(f, sheet_name=sheet_name)
            df_list.append(data)

    # Concatenate all dataframes in the list
    big_df = pd.concat(df_list)

    # Export to excel
    big_df.to_excel('output.xlsx', index=False)

    input("File saved as output.xlsx. Press any key to continue...")
    main()


def browseFiles():
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilenames(parent=root, title='Choose a file')

def main():

    print("Menu:\n1. Merge Excel files\n2. Convert CSV to Excel\n3. JSON to Excel\n4. Excel to SQLite3\n0. Exit")
    choice = int(input("Enter your choice: "))
    if choice == 1:
        files = browseFiles()
        files_xls = [f for f in files if f[-4:] == 'xlsx']
        if len(files_xls) > 0:
            mergeExcel(files_xls)
    elif choice == 2:
        files = browseFiles()
        files_csv = [f for f in files if f[-3:] == 'csv']
        if len(files_csv) > 0:
            csvToExcel(files_csv)
    elif choice == 3:
        files = browseFiles()
        files_json = [f for f in files if f[-4:] == 'json']
        if len(files_json) > 0:
            jsonToExcel(files_json)
    elif choice == 4:
        db_name = 'output-' + datetime.now().strftime("%Y%m%d-%H%M%S") + '.db'
        db = sqlite3.connect(db_name)
        files = browseFiles()
        files_xls = [f for f in files if f[-4:] == 'xlsx']
        if len(files_xls) > 0:
            for f in files_xls:
                dfs = pd.ExcelFile(f)
                sheet_names = dfs.sheet_names
                for table in sheet_names:
                    df = pd.read_excel(f, sheet_name=table)
                    df.to_sql(table, db)
        input("File saved as " + db_name + ". Press any key to continue...")
        main()
    else:
        exit(0)


if __name__ == "__main__":
    main()
