import os
import pandas as pd
from tqdm import tqdm
from collections import defaultdict

# پارامترهایی که باید در شیت‌های جدا ذخیره شوند
parameters = [
    "tmax_m", "tmax_max", "tmax_min", "tmin_m", "tmin_min", "tmin_max",
    "ntmin_0", "rrr24", "umin", "umax", "umm", "sshn", "radglo24", "tm_m",
    "twetm_m", "tsoilm_m", "tsoilmin_min", "umax_m", "umin_m",
    "evt_s", "evt_min", "evt_max"
]

START_YEAR = 1950
END_YEAR = 2024

def process_files(input_folder, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    all_files = [
        os.path.join(input_folder, f)
        for f in os.listdir(input_folder)
        if f.endswith(".xlsx") and not f.startswith("~$")
    ]

    station_data = defaultdict(list)

    for file in tqdm(all_files, desc="📂 در حال خواندن فایل‌ها"):
        try:
            df = pd.read_excel(file, engine="openpyxl", header=1)
        except Exception as e:
            print(f"❌ خطا در خواندن {file}: {e}")
            continue

        # بررسی وجود ستون‌های اصلی
        required_columns = ["station_name", "data"]
        missing_cols = [col for col in required_columns if col not in df.columns]
        if missing_cols:
            print(f"⚠️ فایل {file} ستون‌های لازم را ندارد. ستون‌های گمشده: {missing_cols}")
            continue

        # تبدیل و فیلتر بازه سال
        df['data'] = pd.to_datetime(df['data'], errors='coerce')
        df['year'] = df['data'].dt.year
        df = df[df['year'].between(START_YEAR, END_YEAR)]

        # دسته‌بندی داده‌ها بر اساس نام ایستگاه
        for station in df['station_name'].dropna().unique():
            station_df = df[df['station_name'] == station].copy()
            station_data[station].append(station_df)

    # نوشتن خروجی‌ها
    for station_name, dfs in station_data.items():
        full_df = pd.concat(dfs, ignore_index=True)
        full_df = full_df.sort_values(by=['year', 'data'])

        safe_station_name = station_name.replace(" ", "_").replace("/", "_")
        output_path = os.path.join(output_folder, f"{safe_station_name}_{START_YEAR}-{END_YEAR}.xlsx")

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for param in parameters:
                if param not in full_df.columns:
                    continue

                param_df = full_df[['year', 'data', param]].copy()
                param_df = param_df.dropna(subset=[param])
                param_df.columns = ['year', 'date', 'value']

                param_df.to_excel(writer, sheet_name=param[:31], index=False)

        print(f"✅ ساخته شد: {output_path}")

if __name__ == "__main__":
    input_folder = "converted_xlsx"      # مسیر فایل‌های ورودی
    output_folder = "output_stations"    # مسیر فایل‌های خروجی

    process_files(input_folder, output_folder)
