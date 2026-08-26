import glob
import re
import os
import pandas as pd
import shutil
from tqdm import tqdm

def parse_age_group(age_str):
    if pd.isna(age_str):
        return "Tidak Diketahui"
    
    thn_match = re.search(r'(\d+)\s*Thn', str(age_str), re.IGNORECASE)
    thn = int(thn_match.group(1)) if thn_match else 0
    
    if thn < 1:
        return "< 1 tahun"
    elif 1 <= thn <= 4:
        return "1-4"
    elif 5 <= thn <= 9:
        return "5-9"
    elif 10 <= thn <= 14:
        return "10-14"
    elif 15 <= thn <= 19:
        return "15-19"
    else:
        return ">= 20"

def process_excel_files(file_pattern="*.xlsx", output_filename="Rekap ISPA Anak.xlsx"):
    if os.path.exists('results'):
        shutil.rmtree('results')
        os.mkdir('results')
    else:
        os.mkdir('results')

    file_list = glob.glob(file_pattern)
    file_list = [f for f in file_list if not os.path.basename(f).startswith("Rekap_")]
    
    if not file_list:
        print(f"Tidak ditemukan file Excel dengan pola '{file_pattern}'.")
        return

    all_data = []
    
    for file_path in tqdm(file_list, desc="Processing Excel Files"):
        try:
            df = pd.read_excel(file_path, skiprows=11)
            
            if 'Umur' in df.columns:
                if 'Kode Diagnosa' in df.columns:
                    df['Kode_Penyakit'] = df['Kode Diagnosa'].astype(str).apply(lambda x: x.split('|')[0].strip())
                else:
                    df['Kode_Penyakit'] = "N/A"
                
                df['Kelompok_Umur'] = df['Umur'].apply(parse_age_group)
                all_data.append(df)
            else:
                print(f"Peringatan: Kolom 'Umur' tidak ditemukan pada {file_path}")
        except Exception as e:
            print(f"Gagal membaca file {file_path}: {e}")
            
    if not all_data:
        print("Tidak ada data valid yang dapat diproses.")
        return
        
    combined_df = pd.concat(all_data, ignore_index=True)
    
    age_categories = ["< 1 tahun", "1-4", "5-9", "10-14", "15-19", ">= 20"]
    
    # --- Sheet 1: Detail Per Penyakit ---
    pivot_counts = combined_df.groupby(['Kode_Penyakit', 'Kelompok_Umur']).size().unstack(fill_value=0)
    pivot_counts = pivot_counts.reindex(columns=age_categories, fill_value=0)
    
    rows_detail = []
    for kode, row in pivot_counts.iterrows():
        total = row.sum()
        for cat in age_categories:
            cnt = row[cat]
            pct = round((cnt / total * 100), 2) if total > 0 else 0.0
            rows_detail.append({
                'Kode Penyakit': kode,
                'Rentang Umur': cat,
                'Jumlah': cnt,
                'Total': total,
                'Persentase (%)': pct
            })
            
    rekap_df = pd.DataFrame(rows_detail)
    
    # --- Sheet 2: Total Keseluruhan (Gabungan 3 Penyakit) ---
    total_by_age = combined_df.groupby('Kelompok_Umur').size().reindex(age_categories, fill_value=0)
    grand_total = total_by_age.sum()
    
    rows_total = []
    for cat in age_categories:
        cnt = total_by_age[cat]
        pct = round((cnt / grand_total * 100), 2) if grand_total > 0 else 0.0
        rows_total.append({
            'Rentang Umur': cat,
            'Jumlah': cnt,
            'Total Keseluruhan': grand_total,
            'Persentase (%)': pct
        })
        
    total_df = pd.DataFrame(rows_total)
    
    # Simpan ke Excel
    output_path = os.path.join('results', output_filename)
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        rekap_df.to_excel(writer, sheet_name='Rekapitulasi Per Penyakit', index=False)
        total_df.to_excel(writer, sheet_name='Total Gabungan Penyakit', index=False)
        
        # Matrix summary optional
        pivot_counts['Total'] = pivot_counts.sum(axis=1)
        pivot_counts.to_excel(writer, sheet_name='Ringkasan Matriks')

    print(f"Berhasil! Hasil rekapitulasi telah disimpan ke: {output_path}")

if __name__ == "__main__":
    process_excel_files("excels/*.xlsx")
