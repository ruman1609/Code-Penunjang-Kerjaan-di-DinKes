import pandas as pd
import glob

files = glob.glob("./source/*.xlsx")

for file in files:
    xls = pd.ExcelFile(file)
    sheetsName = xls.sheet_names
    print(f"{file}: {sheetsName}")