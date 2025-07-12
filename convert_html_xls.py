import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def is_html_xls(file_path):
    """بررسی اینکه فایل xls در واقع HTML هست یا نه"""
    try:
        with open(file_path, 'rb') as f:
            head = f.read(512).lower()
            return b'<html' in head or b'<table' in head
    except Exception as e:
        print(f"⚠️ خطا در بررسی فرمت فایل {file_path}: {e}")
        return False

def convert_html_xls_to_xlsx(input_folder, output_folder):
    os.makedirs(output_folder, exist_ok=True)

    files = [f for f in os.listdir(input_folder) if f.lower().endswith('.xls') and not f.startswith('~$')]

    if not files:
        messagebox.showwarning("هیچ فایلی پیدا نشد", "📂 هیچ فایل .xls در پوشه انتخاب‌شده نیست.")
        return

    success_count = 0
    for file in files:
        file_path = os.path.join(input_folder, file)
        if is_html_xls(file_path):
            try:
                tables = pd.read_html(file_path)
                df = tables[0]
                output_filename = file.replace('.xls', '.xlsx')
                output_path = os.path.join(output_folder, output_filename)
                df.to_excel(output_path, index=False)
                success_count += 1
                print(f"✅ تبدیل شد: {file} → {output_filename}")
            except Exception as e:
                print(f"❌ خطا در تبدیل {file}: {e}")
        else:
            print(f"⏩ رد شد (فرمت HTML نبود): {file}")

    messagebox.showinfo("✅ عملیات انجام شد", f"{success_count} فایل با موفقیت تبدیل شدند.")

def select_folders_and_convert():
    input_folder = filedialog.askdirectory(title="📥 انتخاب پوشه فایل‌های HTML xls")
    if not input_folder:
        return

    output_folder = filedialog.askdirectory(title="📤 انتخاب پوشه برای ذخیره فایل‌های xlsx")
    if not output_folder:
        return

    convert_html_xls_to_xlsx(input_folder, output_folder)

# رابط گرافیکی ساده برای اجرای راحت‌تر
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    select_folders_and_convert()
