import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar
import threading
from generate_by_station_param import process_files  # حواست باشه اسم فایل همینه

def start_processing():
    messagebox.showinfo("📂 انتخاب پوشه", "پوشه‌ای که فایل‌های تبدیل‌شده‌ی xlsx داخلشه رو انتخاب کن.")

    input_folder = filedialog.askdirectory(title="📥 انتخاب پوشه فایل‌های ورودی")
    if not input_folder:
        return

    output_folder = filedialog.askdirectory(title="📤 انتخاب پوشه فایل‌های خروجی")
    if not output_folder:
        return

    status_label.config(text="⏳ در حال پردازش فایل‌ها...")
    progress_bar.start()

    def run():
        try:
            process_files(input_folder, output_folder)
            messagebox.showinfo("✅ موفقیت", "همه‌ی فایل‌ها با موفقیت تولید شدند.")
        except Exception as e:
            messagebox.showerror("❌ خطا", str(e))
        finally:
            status_label.config(text="آماده")
            progress_bar.stop()

    threading.Thread(target=run).start()

# ساخت رابط گرافیکی
app = tk.Tk()
app.title("📊 پردازش داده‌های هواشناسی به تفکیک ایستگاه و پارامتر")
app.geometry("600x300")
app.resizable(False, False)

tk.Label(app, text="📊 ابزار تبدیل داده‌های هواشناسی", font=("B Nazanin", 16, "bold")).pack(pady=20)

tk.Button(
    app,
    text="🚀 انتخاب پوشه‌ها و شروع پردازش",
    command=start_processing,
    bg="black", fg="white",
    font=("Arial", 13),
    padx=20, pady=10
).pack()

progress_bar = Progressbar(app, orient="horizontal", length=400, mode="indeterminate")
progress_bar.pack(pady=30)

status_label = tk.Label(app, text="آماده", font=("Arial", 11))
status_label.pack()

app.mainloop()
