import glob
import pdfplumber
import numpy as np
import dateparser
from datetime import datetime
import pandas as pd
import shutil
import os
from tqdm import tqdm
import traceback


listPdfFiles = glob.glob("./pdf/*.pdf")

header = ["No. Urut", "Nama", "Tgl Lahir", "NIP", "NPWP", "Status Pegawai - Golongan", "Penghasilan - TPP Beban Kerja", "Penghasilan - TPP Tempat Bertugas", "Penghasilan - TPP Kondisi Kerja", "Penghasilan - TPP Kelangkaan Profesi", "Penghasilan - TPP Prestasi Kerja", "Penghasilan - TJ. Kesehatan", "Penghasilan - TJ. Kecelakaan", "Penghasilan - TJ. Kematian", "Penghasilan - TJ. PPh 21", "Penghasilan - Jumlah Bruto", "Potongan - IWP Kesehatan", "Potongan - IWP Kecelakaan", "Potongan - IWP Kematian", "Potongan - PPh 21", "Potongan - Iuran Tapera", "Potongan - IWP 8%", "Potongan - IWP 1%", "Potongan - Zakat", "Potongan - Bulog", "Potongan - Jumlah Potongan", "Jumlah Dibayarkan", "TTD", "Masa Kerja", "No. Rek"]

parse_settings = {
    'RELATIVE_BASE': datetime(2024, 1, 1), 
    'PREFER_DATES_FROM': 'past'
}

if (os.path.isdir("result")):
    shutil.rmtree("result")
os.mkdir("result")

def handlePdf(pdf : pdfplumber.PDF, pdfFilePath: str):
    pdf_title = pdfFilePath.rsplit("/", 1)[1].split(".")[0]
    all_data = []
    for page_number, page in enumerate(tqdm(pdf.pages, desc="Processing " + pdf_title)):
        extracted = page.extract_table()

        if extracted is None:
            continue
        
        for data_index, data in enumerate(extracted):
            if page_number == 0 and data_index < 2:
                continue
            
            if data is None:
                continue

            if data[0] is None:
                continue
            
            if data[0].lower() == "jumlah":
                continue

            if data[0].lower() == header[0].lower():
                continue

            if '' in data:
                for item_index, to_be_merged_item in enumerate(data):
                    if to_be_merged_item is None or to_be_merged_item == "":
                        continue

                    all_data[-1][item_index] = f"{str(all_data[-1][item_index]).strip()}\n{str(to_be_merged_item).strip()}"
                continue

            all_data.append(data)
    
    flatten_extracted = []
    for data_index, data in enumerate(all_data):
        result = []
        for item in data:
            splitted = item.split("\n")
            result.extend(splitted)

        must_be_date_index = 2
        while True:
            if any(char.isdigit() for char in result[must_be_date_index]) and dateparser.parse(result[must_be_date_index], languages=['id'], settings=parse_settings) != None:
                break
            else:
                result[must_be_date_index - 1] = result[must_be_date_index - 1] + " " + result[must_be_date_index]
                result.pop(must_be_date_index)

        if (len(result) != len(header)):
            print("RESULT NOT SAME", result, len(result))
        
        flatten_extracted.append(result)
    
    flatten_extracted = np.array(flatten_extracted)

    df = pd.DataFrame(flatten_extracted, columns=header)

    df.to_excel("result/" + pdf_title + ".xlsx", index=False)

for _, pdfFilePath in enumerate(tqdm(listPdfFiles, "Current PDF")):
    with pdfplumber.open(pdfFilePath) as pdf:
        try:
            handlePdf(pdf, pdfFilePath)
        except:
            print("Error on " + pdfFilePath)
            traceback.print_exc()
            exit()