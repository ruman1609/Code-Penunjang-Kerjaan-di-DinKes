import pandas as pd
import glob
import traceback
import shutil
import os
from tqdm import tqdm

files = glob.glob("./source/*.xlsx")

columns_not_salin = ["NAMA", "RUMAH SAKIT/PUSKESMAS", "KECAMATAN", "SUDAH TERLAYANI/BELUM", "JENIS PELAYANAN", "KETERANGAN"]

df_bersalin = pd.DataFrame(columns=["KECAMATAN", "NAMA IBU MELAHIRKAN", "TANGGAL MELAHIRKAN", "LAHIR HIDUP/MENINGGAL", "PENYEBAB KEMATIAN", "KETERANGAN"])
df_hipertensi = pd.DataFrame(columns=columns_not_salin)
df_diabetes = pd.DataFrame(columns=columns_not_salin)
df_odgj = pd.DataFrame(columns=columns_not_salin)

list_bersalin = []
list_hipertensi = []
list_diabetes = []
list_odgj = []

for file in tqdm(files, desc="Processing Excel Files"):
    try:
        xls = pd.ExcelFile(file)
        current_bersalin = pd.read_excel(xls, sheet_name="IBU BERSALIN", dtype={"TANGGAL MELAHIRKAN": str})
        current_hipertensi = pd.read_excel(xls, sheet_name="HIPERTENSI")
        current_diabetes = pd.read_excel(xls, sheet_name="DIABETES MELLITUS")
        current_odgj = pd.read_excel(xls, sheet_name="ODGJ")

        def clean_df(df):
            df.columns = df.columns.str.strip().str.upper()
            return df.map(lambda x: str(x).strip().upper() if pd.notnull(x) else "")

        current_bersalin = clean_df(current_bersalin)
        current_hipertensi = clean_df(current_hipertensi)
        current_diabetes = clean_df(current_diabetes)
        current_odgj = clean_df(current_odgj)
        
        # print(f"{file}:\n{current_bersalin.keys()}\n{current_hipertensi.keys()}\n{current_diabetes.keys()}\n{current_odgj.keys()}\n====\n\n")
        # print(f"{file}: {current_bersalin.shape}, {current_hipertensi.shape}, {current_diabetes.shape}, {current_odgj.shape}")
        # continue

        current_hipertensi = current_hipertensi.drop(columns=["NO", "NO."], errors="ignore")
        current_diabetes = current_diabetes.drop(columns=["NO", "NO."], errors="ignore")
        current_odgj = current_odgj.drop(columns=["NO", "NO."], errors="ignore")

        def map_not_salin(current: pd.DataFrame):
            return pd.DataFrame({
                "NAMA": current["NAMA PASIEN"],
                "RUMAH SAKIT/PUSKESMAS": current["PUSKESMAS"],
                "KECAMATAN": current["KECAMATAN"]
            })
        
        list_bersalin.append(current_bersalin[["KECAMATAN", "NAMA IBU MELAHIRKAN", "TANGGAL MELAHIRKAN", "LAHIR HIDUP/MENINGGAL", "PENYEBAB KEMATIAN", "KETERANGAN"]])
        list_diabetes.append(map_not_salin(current_diabetes))
        list_hipertensi.append(map_not_salin(current_hipertensi))
        list_odgj.append(map_not_salin(current_odgj))
    except:
        print("Error on " + file)
        traceback.print_exc()
        exit()


df_bersalin = pd.concat(list_bersalin, ignore_index=True)
df_hipertensi = pd.concat(list_hipertensi, ignore_index=True)
df_diabetes = pd.concat(list_diabetes, ignore_index=True)
df_odgj = pd.concat(list_odgj, ignore_index=True)

if os.path.isdir("./results"):
    shutil.rmtree("./results")
os.mkdir("./results")

df_bersalin.to_excel("./results/bersalin.xlsx", index=True)
df_hipertensi.to_excel("./results/hipertensi.xlsx", index=True)
df_diabetes.to_excel("./results/diabetes.xlsx", index=True)
df_odgj.to_excel("./results/odgj.xlsx", index=True)

print("Done!")