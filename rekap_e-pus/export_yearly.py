import pandas as pd
import shutil
import os
from tqdm import tqdm
from glob import glob

excel_files = glob("./excels/*.xlsx")

if (os.path.isdir("./results")):
    shutil.rmtree("./results")
os.mkdir("./results")

combined_df = pd.DataFrame()

def group_by_sort_df(df: pd.DataFrame):
    return df.groupby("id", as_index=False).agg({
        "name": "first",
        "man": "sum",
        "woman": "sum",
        "total": "sum"
    }).sort_values(by="total", ascending=False)

for file in tqdm(excel_files, desc="Processing Excel Files"):
    excels_df = pd.read_excel(file, sheet_name=None, skiprows=5)
    
    df = pd.concat(excels_df.values(), ignore_index=True)
    df = df.drop(columns=[df.columns[0]])
    df = df.rename(columns={
        df.columns[0]: "id",
        df.columns[1]: "name",
        df.columns[2]: "man",
        df.columns[3]: "woman",
        df.columns[4]: "total"
    })
    df = group_by_sort_df(df)

    combined_df = pd.concat([combined_df, df], ignore_index=True)

    df.to_excel(f"./results/{file.rsplit('/', 1)[1].split('.')[0].strip()}.xlsx", index=False)


combined_df = group_by_sort_df(combined_df)
combined_df.to_excel("./results/Laporan - Kota Samarinda.xlsx", index=False)