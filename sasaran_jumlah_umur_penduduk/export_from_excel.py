import pandas as pd
import os
import numpy as np

if os.path.isfile("result.xlsx"):
    os.remove("result.xlsx")

WILAYAH_KEY = "WILAYAH"
PUSKESMAS_KEY = "PUSKESMAS"

source = pd.read_excel("./source.xlsx", sheet_name="SOURCE")

source_man = pd.concat([source[WILAYAH_KEY].copy(), source.filter(regex="^L").copy()], axis=1)
source_woman = pd.concat([source[WILAYAH_KEY].copy(), source.filter(regex="^P").copy()], axis=1)

source_man.columns = source_man.columns.str.replace("^L", "", regex=True).map(str.strip).dropna()
source_woman.columns = source_woman.columns.str.replace("^P", "", regex=True).map(str.strip).dropna()

results_man = source[[WILAYAH_KEY]].copy()
results_woman = source[[WILAYAH_KEY]].copy()
results_union = source[[WILAYAH_KEY]].copy()

results_man[WILAYAH_KEY] = results_man[WILAYAH_KEY].str.strip().str.upper()
results_woman[WILAYAH_KEY] = results_woman[WILAYAH_KEY].str.strip().str.upper()
results_union[WILAYAH_KEY] = results_union[WILAYAH_KEY].str.strip().str.upper()

pusdatin_man = {
    "0": 6696,
    "0-1": 13361,
    "0-2": 19953,
    "0-4": 32292,
    "1-4": 26296,
    "1-2": 13257,
    "2-5": 26177,
    "5-6": 13054,
    "6": 6508,
    "7": 6487,
    "8": 6477,
    "9": 6476,
    "10": 6489,
    "11": 6522,
    "12": 6590,
    "7-17": 73856,
    "7-15": 59530,
    "7+": 394549,
    "1-12": 78391,
    "10-21": 84189,
    "10-17": 54416,
    "13": 6695,
    "15": 6968,
    "13-15": 20489,
    "16": 7102,
    "16-18": 21652,
    "13-18": 42141,
    "0-18": 127228,
    "19-64": 284593,
    "0-14": 98608,
    "15-64": 313213,
    "65+": 28774,
    "21-26": 44755,
    "18-30": 95124,
    "31-45": 100731,
    "46-59": 76928,
    "18-59": 272783,
    "15+": 341987,
    "18+": 320693,
    "15-39": 177274,
    "40+": 164713,
    "15-59": 294077,
    "30-50": 95941,
    "15-49": 95941,
    "15-49W": 0,
    "30-50": 0,
    "30-69": 0,
    "45-59": 83405,
    "60+": 47910,
    "50-59": 52045,
    "60-69": 32459,
    "70+": 15451
}

pusdatin_woman = {
    "0": 6422,
    "0-1": 12820,
    "0-2": 19151,
    "0-4": 31677,
    "1-4": 25255,
    "1-2": 12729,
    "2-5": 25114,
    "5-6": 12472,
    "6": 6215,
    "7": 6195,
    "8": 6196,
    "9": 6215,
    "10": 6247,
    "11": 6299,
    "12": 6365,
    "7-17": 70677,
    "7-15": 57132,
    "7+": 384737,
    "1-12": 75244,
    "10-21": 79810,
    "10-17": 52071,
    "13": 6445,
    "15": 6634,
    "13-15": 19615,
    "16": 6731,
    "16-18": 20421,
    "13-18": 40036,
    "0-18": 121702,
    "19-64": 277959,
    "0-14": 94647,
    "15-64": 305014,
    "65+": 29225,
    "21-26": 42393,
    "18-30": 91168,
    "31-45": 98016,
    "46-59": 76375,
    "18-59": 265559,
    "15+": 334239,
    "18+": 314060,
    "15-39": 171107,
    "40+": 163132,
    "15-59": 285738,
    "15-49": 92636,
    "15-49W": 233694,
    "30-50": 135074,
    "30-69": 213625,
    "45-59": 82633,
    "60+": 48501,
    "50-59": 52044,
    "60-69": 32362,
    "70+": 16139
}

puskesmas_data = {
    # PALARAN
    "RAWA MAKMUR": "PALARAN",
    "HANDIL BAKTI": "PALARAN",
    "SIMPANG PASIR": "PALARAN",
    # BUKUAN
    "BUKUAN": "BUKUAN",
    # BANTUAS
    "BANTUAS": "BANTUAS",
    # SAMARINDA KOTA
    "PELABUHAN": "SAMARINDA KOTA",
    "KARANG MUMUS": "SAMARINDA KOTA",
    "PASAR PAGI": "SAMARINDA KOTA",
    "SUNGAI PINANG LUAR": "SAMARINDA KOTA",
    "BUGIS": "SAMARINDA KOTA",
    # BAQA
    "SUNGAI KELEDANG": "BAQA",
    "BAQA": "BAQA",
    "GUNUNG PANJANG": "BAQA",
    # MANGKUPALAS
    "MANGKUPALAS": "MANGKUPALAS",
    "MASJID": "MANGKUPALAS",
    "KAMPUNG TENUN": "MANGKUPALAS",
    # AIR PUTIH
    "AIR PUTIH": "AIR PUTIH",
    "BUKIT PINANG": "AIR PUTIH",
    # JUANDA
    "AIR HITAM": "JUANDA",
    "GUNUNG KELUA": "JUANDA",
    # SEGIRI
    "SIDODADI": "SEGIRI",
    "DADIMULYA": "SEGIRI",
    # PASUNDAN
    "TELUK LERONG ILIR": "PASUNDAN",
    "JAWA": "PASUNDAN",
    # SIDOMULYO
    "SELILI": "SIDOMULYO",
    "SUNGAI DAMA": "SIDOMULYO",
    "SIDOMULYO": "SIDOMULYO",
    "SIDODAMAI": "SIDOMULYO",
    "PELITA": "SIDOMULYO",
    # SUNGAI SIRING
    "SUNGAI SIRING": "SUNGAI SIRING",
    "TANAH MERAH": "SUNGAI SIRING",
    "BUDAYA PAMPANG": "SUNGAI SIRING",
    # SEMPAJA
    "SEMPAJA SELATAN": "SEMPAJA",
    "SEMPAJA BARAT": "SEMPAJA",
    # BENGKURING
    "SEMPAJA UTARA": "BENGKURING",
    "SEMPAJA TIMUR": "BENGKURING",
    # LEMPAKE
    "LEMPAKE": "LEMPAKE",
    # LOA BAKUNG
    "LOA BAKUNG": "LOA BAKUNG",
    "LOA BUAH": "LOA BAKUNG",
    # KARANG ASAM
    "KARANG ASAM ULU": "KARANG ASAM",
    "KARANG ASAM ILIR": "KARANG ASAM",
    # LOK BAHU
    "LOK BAHU": "LOK BAHU",
    # WONOREJO
    "TELUK LERONG ULU": "WONOREJO",
    "KARANG ANYAR": "WONOREJO",
    # SAMBUTAN
    "SAMBUTAN": "SAMBUTAN",
    # SUNGAI KAPIH
    "SUNGAI KAPIH": "SUNGAI KAPIH",
    # MAKROMAN
    "MAKROMAN": "MAKROMAN",
    "SINDANG SARI": "MAKROMAN",
    "PULAU ATAS": "MAKROMAN",
    # TRAUMA CENTER
    "SENGKOTEK": "TRAUMA CENTER",
    "SIMPANG TIGA": "TRAUMA CENTER",
    "TANI AMAN": "TRAUMA CENTER",
    # HARAPAN BARU
    "HARAPAN BARU": "HARAPAN BARU",
    "RAPAK DALAM": "HARAPAN BARU",
    # REMAJA
    "TEMINDUNG PERMAI": "REMAJA",
    "GUNUNG LINGAI": "REMAJA",
    "BANDARA": "REMAJA",
    # TEMINDUNG
    "SUNGAI PINANG DALAM": "TEMINDUNG",
    "MUGIREJO": "TEMINDUNG"
}

results_man[PUSKESMAS_KEY] = results_man[WILAYAH_KEY].map(puskesmas_data)
results_woman[PUSKESMAS_KEY] = results_woman[WILAYAH_KEY].map(puskesmas_data)
results_union[PUSKESMAS_KEY] = results_union[WILAYAH_KEY].map(puskesmas_data)

def process_df(results_base: pd.DataFrame, source_df: pd.DataFrame, pusdatin: dict) -> pd.DataFrame:
    new_cols = {}
    
    def add_calc(key, start, end=None):
        if end is None:
            val = source_df[key].copy()
        elif end == "":
            val = source_df.loc[:, start:].sum(axis=1)
        else:
            val = source_df.loc[:, start:end].sum(axis=1)
            
        # new_cols[key] = val

        proportion = (val / val.sum()) * pusdatin[key]
        # new_cols[f"{key}_Pr"] = proportion.apply(np.floor)
        new_cols[key] = proportion.apply(np.floor)

    add_calc("0", "0")
    add_calc("0-1", "0", "1")
    add_calc("0-2", "0", "2")
    add_calc("0-4", "0", "4")
    add_calc("1-4", "1", "4")
    add_calc("1-2", "1", "2")
    add_calc("2-5", "2", "5")
    add_calc("5-6", "5", "6")
    add_calc("6", "6")
    add_calc("7", "7")
    add_calc("8", "8")
    add_calc("9", "9")
    add_calc("10", "10")
    add_calc("11", "11")
    add_calc("12", "12")
    add_calc("7-17", "7", "17")
    add_calc("7-15", "7", "15")
    add_calc("7+", "7", "")
    add_calc("1-12", "1", "12")
    add_calc("10-21", "10", "21")
    add_calc("10-17", "10", "17")
    add_calc("13", "13")
    add_calc("15", "15")
    add_calc("13-15", "13", "15")
    add_calc("16", "16")
    add_calc("16-18", "16", "18")
    add_calc("13-18", "13", "18")
    add_calc("0-18", "0", "18")
    add_calc("19-64", "19", "64")
    add_calc("0-14", "0", "14")
    add_calc("15-64", "15", "64")
    add_calc("65+", "65", "")
    add_calc("21-26", "21", "26")
    add_calc("18-30", "18", "30")
    add_calc("31-45", "31", "45")
    add_calc("46-59", "46", "59")
    add_calc("18-59", "18", "59")
    add_calc("15+", "15", "")
    add_calc("18+", "18", "")
    add_calc("15-39", "15", "39")
    add_calc("40+", "40", "")
    add_calc("15-59", "15", "59")
    add_calc("15-49", "15", "49")
    
    
    val_15_49w = source_df.loc[:, "15":"49"].sum(axis=1)
    # new_cols["15-49W"] = val_15_49w

    proportion = (val_15_49w / val_15_49w.sum()) * pusdatin["15-49W"]
    new_cols["15-49W"] = proportion.round(0)

    add_calc("30-50", "30", "50")
    add_calc("30-69", "30", "69")
    add_calc("45-59", "45", "59")
    add_calc("60+", "60", "")
    add_calc("50-59", "50", "59")
    add_calc("60-69", "60", "69")
    add_calc("70+", "70", "")

    new_df = pd.DataFrame(new_cols)
    return pd.concat([results_base, new_df], axis=1)


results_man = process_df(results_man, source_man, pusdatin_man)
results_woman = process_df(results_woman, source_woman, pusdatin_woman)

def group_sum_puskesmas(results: pd.DataFrame) -> pd.DataFrame:
    def add_total(group: pd.DataFrame, puskesmas_name: str):
        total_row = group.sum(numeric_only=True).to_frame().T
        total_row[WILAYAH_KEY] = "TOTAL"
        
        combined = pd.concat([group, total_row], ignore_index=True)
        combined[PUSKESMAS_KEY] = puskesmas_name
        return combined
    
    return results.groupby(PUSKESMAS_KEY, group_keys=False, sort=False).apply(
        lambda g: add_total(g, g.name), 
        include_groups=False
    )


results_man = group_sum_puskesmas(results_man)
results_woman = group_sum_puskesmas(results_woman)


idx_results_man = results_man.set_index([PUSKESMAS_KEY, WILAYAH_KEY])
idx_results_woman = results_woman.set_index([PUSKESMAS_KEY, WILAYAH_KEY])
results_union = idx_results_man.add(idx_results_woman, fill_value=0)

results = pd.concat([idx_results_man, idx_results_woman, results_union], axis=1, keys=["L", "P", "TOTAL"])
results = results.sort_index(axis=1, level=1).reset_index()

results.to_excel("result.xlsx")
print("DONE!")