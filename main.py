import pandas as pd
import glob
import openpyxl

filepaths = glob.glob("invoices/*.xlsx")  # Use glob to read all xlsx files

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)