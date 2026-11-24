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
    "0": 6663,
    "0-1": 13263,
    "0-2": 19842,
    "0-4": 32806,
    "1-4": 26143,
    "1-2": 13179,
    "2-5": 26179,
    "5-6": 13241,
    "6": 6605,
    "7": 6575,
    "8": 6538,
    "9": 6497,
    "10": 6462,
    "11": 6450,
    "12": 6481,
    "7-17": 73123,
    "7-15": 59086,
    "7+": 397395,
    "1-12": 78387,
    "10-21": 82940,
    "10-17": 53513,
    "13": 6565,
    "15": 6827,
    "13-15": 20083,
    "16": 6955,
    "16-18": 21237,
    "13-18": 41320,
    "0-18": 126370,
    "19-64": 286518,
    "0-14": 98306,
    "15-64": 314582,
    "65+": 30554,
    "21-26": 44851,
    "18-30": 95153,
    "31-45": 100463,
    "46-59": 78313,
    "18-59": 273929,
    "15+": 345136,
    "18+": 324272,
    "15-39": 176594,
    "40+": 168542,
    "15-59": 294793,
    "15-49": 95777,
    "15-49W": 0,
    "30-50": 0,
    "30-69": 0,
    "45-59": 84824,
    "60+": 50343,
    "50-59": 53175,
    "60-69": 33925,
    "70+": 16418,
}

pusdatin_woman = {
    "0": 6390,
    "0-1": 12725,
    "0-2": 19042,
    "0-4": 31497,
    "1-4": 25107,
    "1-2": 12652,
    "2-5": 25108,
    "5-6": 12641,
    "6": 6305,
    "7": 6279,
    "8": 6254,
    "9": 6231,
    "10": 6214,
    "11": 6216,
    "12": 6250,
    "7-17": 70091,
    "7-15": 56717,
    "7+": 388595,
    "1-12": 75192,
    "10-21": 78849,
    "10-17": 51327,
    "13": 6321,
    "15": 6532,
    "13-15": 19273,
    "16": 6640,
    "16-18": 20179,
    "13-18": 39452,
    "0-18": 121034,
    "19-64": 280253,
    "0-14": 94323,
    "15-64": 306964,
    "65+": 31446,
    "21-26": 42236,
    "18-30": 90867,
    "31-45": 98312,
    "46-59": 77777,
    "18-59": 266956,
    "15+": 338410,
    "18+": 318504,
    "15-39": 170684,
    "40+": 167726,
    "15-59": 286862,
    "15-49": 92598,
    "15-49W": 233597,
    "30-50": 135667,
    "30-69": 217214,
    "45-59": 84068,
    "60+": 51548,
    "50-59": 53265,
    "60-69": 34211,
    "70+": 17337,
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

column_orders = [
    "0", "0-1", "0-2", "0-4", "1-4", "1-2", "2-5", "5-6", "6", "7", "8", "9", 
    "10", "11", "12", "7-17", "7-15", "7+", "1-12", "10-21", "10-17", "13", 
    "15", "13-15", "16", "16-18", "13-18", "0-18", "19-64", "0-14", "15-64", 
    "65+", "21-26", "18-30", "31-45", "46-59", "18-59", "15+", "18+", "15-39", 
    "40+", "15-59", "15-49", "15-49W", "30-50", "30-69", "45-59", "60+", 
    "50-59", "60-69", "70+"
]

results_man[PUSKESMAS_KEY] = results_man[WILAYAH_KEY].map(puskesmas_data)
results_woman[PUSKESMAS_KEY] = results_woman[WILAYAH_KEY].map(puskesmas_data)
results_union[PUSKESMAS_KEY] = results_union[WILAYAH_KEY].map(puskesmas_data)

def process_df(results_base: pd.DataFrame, source_df: pd.DataFrame, pusdatin: dict) -> pd.DataFrame:
    new_cols = {}
    
    def add_calc(key, start, end=None, mode="proportional"):
        if end is None:
            val = source_df[key].copy()
        elif end == "":
            val = source_df.loc[:, start:].sum(axis=1)
        else:
            val = source_df.loc[:, start:end].sum(axis=1)

        target = float(pusdatin[key])
        n_rows = len(val)
        if n_rows == 0:
            raise ValueError(f"source_df has no rows for key='{key}'.")

        if mode == "equal":
            weights = pd.Series(np.ones(n_rows, dtype=float), index=val.index)
        elif mode == "proportional":
            weights = val.astype(float).clip(lower=0.0)
        elif mode == "inverse_proportional":
            # Guard against zero/non-positive values so 1/val does not produce inf or NaN
            weights = pd.Series(
                np.where(val > 0, 1.0 / val.where(val > 0), 0.0),
                index=val.index,
                dtype=float,
            )
        else:
            raise ValueError(
                f"Unsupported mode='{mode}'. Expected 'equal', 'proportional', or 'inverse_proportional'."
            )

        total_weight = weights.sum()
        if total_weight == 0:
            raise ValueError(
                f"Total weight is zero for mode='{mode}' and key='{key}'; cannot compute distribution."
            )

        proportion = (weights / total_weight) * target
        rounded_val = proportion.round(0).astype(int)

        # Adjust the largest item so the total matches the rounded target closely/exactly
        diff = int(round(target)) - int(rounded_val.sum())
        if diff != 0:
            rounded_val.iloc[int(rounded_val.argmax())] += diff

        new_cols[key] = rounded_val


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
    rounded_val = proportion.round(0).astype(int)
    diff = pusdatin["15-49W"] - rounded_val.sum()
    if diff != 0:
        rounded_val.iloc[rounded_val.argmax()] += diff
    new_cols["15-49W"] = rounded_val

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

new_column_order = []
for age in column_orders:
    for gender in ["L", "P", "TOTAL"]:
        # Only add if the column actually exists (handles cases like 15-49W which might only be in P)
        if (gender, age) in results.columns:
            new_column_order.append((gender, age))

results = results.reindex(columns=new_column_order)
results = results.reorder_levels([1, 0])
results = results.reset_index()

results.to_excel("result.xlsx")
print("DONE!")