import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import re
import time
import datetime as dt
import jdatetime
import requests
import pandas as pd
from bs4 import BeautifulSoup

# --- کلاس TGJUGoldFetcher: منطق استخراج داده (بدون تغییر) ---
class TGJUGoldFetcher:
    HEADERS = {"User-Agent": "Mozilla/5.0"}

    def __init__(self, status_callback=None):
        self.status_callback = status_callback
        self.stop_flag = False

    def _update_status(self, message):
        if self.status_callback:
            self.status_callback(message)

    def stop(self):
        self.stop_flag = True

    def fetch_data(self, base_url, start_jalali_str, end_jalali_str, output_filepath):
        self.stop_flag = False
        try:
            start_jalali_date = jdatetime.date.fromisoformat(start_jalali_str)
            end_jalali_date = jdatetime.date.fromisoformat(end_jalali_str)

            if start_jalali_date > end_jalali_date:
                raise ValueError("تاریخ شروع باید قبل از یا برابر با تاریخ پایان باشد.")

            start_gregorian_date = start_jalali_date.togregorian()
            end_gregorian_date = end_jalali_date.togregorian()

            self._update_status(f"در حال جمع‌آوری داده‌ها از {start_jalali_str} تا {end_jalali_str}...")
            all_data = []
            page = 1
            reached_start_date_in_history = False 

            while not self.stop_flag:
                url = f"{base_url}?p={page}"
                self._update_status(f"در حال دریافت صفحه {page}...")
                
                try:
                    html = requests.get(url, headers=self.HEADERS, timeout=15).text
                except requests.exceptions.RequestException as e:
                    raise IOError(f"خطا در ارتباط شبکه: {e}. لطفا اتصال اینترنت و آدرس URL را بررسی کنید.")

                soup = BeautifulSoup(html, "html.parser")
                rows = soup.findAll("tr")
                
                if not rows and page > 1: break
                page_processed_any_data_row = False
                
                for r in rows:
                    if self.stop_flag: break
                    cols = [c.get_text(strip=True).replace(',', '') for c in r.findAll("td")]
                    
                    if len(cols) >= 6 and re.match(r"^\d{4}-\d{2}-\d{2}$", cols[0]):
                        page_processed_any_data_row = True
                        gdate = dt.date.fromisoformat(cols[0])

                        if gdate < start_gregorian_date:
                            reached_start_date_in_history = True
                            break
                        if gdate > end_gregorian_date: continue

                        try:
                            high = int(cols[1])
                            low = int(cols[2])
                            avg = (high + low) // 2
                            all_data.append([gdate, high, low, avg])
                        except (ValueError, IndexError):
                            self._update_status(f"هشدار: داده نامعتبر در تاریخ {gdate}. نادیده گرفته شد.")
                            continue
                
                if reached_start_date_in_history or (not page_processed_any_data_row and page > 1): break
                page += 1
                time.sleep(0.5)

            if self.stop_flag:
                self._update_status("عملیات توسط کاربر متوقف شد.")
                return False

            if not all_data:
                self._update_status("هیچ داده‌ای برای بازه تاریخ مشخص شده یافت نشد.")
                return False

            df = pd.DataFrame(all_data, columns=["GregorianDate", "High", "Low", "Average"])
            df = df.drop_duplicates(subset=["GregorianDate"]).sort_values("GregorianDate")
            df = df[(df["GregorianDate"] >= start_gregorian_date) & (df["GregorianDate"] <= end_gregorian_date)]
            
            if df.empty:
                self._update_status("هیچ داده‌ای پس از فیلتر نهایی در بازه تاریخ یافت نشد.")
                return False

            df["PersianDate"] = [jdatetime.date.fromgregorian(date=d) for d in df["GregorianDate"]]
            df['Previous_Average'] = df['Average'].shift(1)
            
            def get_trend(row):
                if pd.isna(row['Previous_Average']): return "---"
                if row['Average'] > row['Previous_Average']: return "صعودی"
                if row['Average'] < row['Previous_Average']: return "نزولی"
                return "بدون تغییر"

            df["Trend"] = df.apply(get_trend, axis=1)
            df = df[["PersianDate", "Low", "High", "Average", "Trend"]]
            df.columns = ["تاریخ", "حداقل", "حداکثر", "میانگین", "روند"]

            df.to_excel(output_filepath, index=False)
            self._update_status(f"عملیات با موفقیت انجام شد. فایل در: {output_filepath} ذخیره شد.")
            return True

        except (ValueError, IOError, Exception) as e:
            self._update_status(f"خطا: {e}")
            return False

# --- کلاس GoldApp: رابط کاربری استاندارد Tkinter ---
class GoldApp:
    def __init__(self, master):
        self.master = master
        master.title("استخراج قیمت طلا از TGJU")
        master.geometry("500x350")
        master.resizable(False, False)

        self.fetcher = TGJUGoldFetcher(self.update_status)
        self.current_thread = None

        self._create_widgets()

    def _create_widgets(self):
        main_frame = ttk.Frame(self.master, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="آدرس صفحه (URL):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.url_entry = ttk.Entry(main_frame)
        self.url_entry.grid(row=0, column=1, columnspan=2, padx=5, pady=5, sticky="ew")
        self.url_entry.insert(0, "https://english.tgju.org/profile/sekee")

        ttk.Label(main_frame, text="تاریخ شروع (مثال: 1403-01-01):").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.start_date_entry = ttk.Entry(main_frame)
        self.start_date_entry.grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky="ew")
        self.start_date_entry.insert(0, jdatetime.date(1403, 1, 1).isoformat())

        ttk.Label(main_frame, text="تاریخ پایان (مثال: 1404-05-31):").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.end_date_entry = ttk.Entry(main_frame)
        self.end_date_entry.grid(row=2, column=1, columnspan=2, padx=5, pady=5, sticky="ew")
        self.end_date_entry.insert(0, jdatetime.date.today().isoformat()) 

        ttk.Label(main_frame, text="مسیر ذخیره فایل اکسل:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.output_path_entry = ttk.Entry(main_frame)
        self.output_path_entry.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        self.output_path_entry.insert(0, "gold_prices.xlsx")

        self.browse_button = ttk.Button(main_frame, text="مرور", command=self.browse_output_path)
        self.browse_button.grid(row=3, column=2, padx=5, pady=5)

        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, pady=10)

        self.start_button = ttk.Button(button_frame, text="شروع استخراج", command=self.start_fetching)
        self.start_button.pack(side=tk.LEFT, padx=5)

        self.stop_button = ttk.Button(button_frame, text="توقف", command=self.stop_fetching, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)

        self.status_label = ttk.Label(main_frame, text="آماده به کار", wraplength=480, justify="right", foreground="blue")
        self.status_label.grid(row=5, column=0, columnspan=3, padx=5, pady=10, sticky="ew")

        main_frame.grid_columnconfigure(1, weight=1)

    def update_status(self, message):
        self.master.after(0, lambda: self.status_label.config(text=message))

    def browse_output_path(self):
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=self.output_path_entry.get()
        )
        if filepath:
            self.output_path_entry.delete(0, tk.END)
            self.output_path_entry.insert(0, filepath)

    def start_fetching(self):
        base_url = self.url_entry.get().strip()
        start_date_str = self.start_date_entry.get().strip()
        end_date_str = self.end_date_entry.get().strip()
        output_filepath = self.output_path_entry.get().strip()

        if not all([base_url, start_date_str, end_date_str, output_filepath]):
            messagebox.showerror("خطا", "لطفا تمام فیلدها را پر کنید.")
            return

        if not re.match(r"^\d{4}-\d{2}-\d{2}$", start_date_str) or \
           not re.match(r"^\d{4}-\d{2}-\d{2}$", end_date_str):
            messagebox.showerror("خطا", "لطفا تاریخ را با فرمت صحیح YYYY-MM-DD وارد کنید.")
            return

        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.update_status("شروع عملیات استخراج...")

        self.current_thread = threading.Thread(target=self._run_fetching_thread, args=(base_url, start_date_str, end_date_str, output_filepath))
        self.current_thread.start()

    def _run_fetching_thread(self, base_url, start_date_str, end_date_str, output_filepath):
        success = self.fetcher.fetch_data(base_url, start_date_str, end_date_str, output_filepath)
        
        self.master.after(0, self.reset_buttons)
        
        if success:
            self.master.after(0, lambda: messagebox.showinfo("پایان عملیات", "داده‌ها با موفقیت استخراج شدند."))
        elif not self.fetcher.stop_flag:
            self.master.after(0, lambda: messagebox.showerror("خطا", "عملیات با خطا مواجه شد."))

    def reset_buttons(self):
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)

    def stop_fetching(self):
        if self.current_thread and self.current_thread.is_alive():
            self.fetcher.stop()
            self.stop_button.config(state=tk.DISABLED)
            self.update_status("درخواست توقف ارسال شد...")

if __name__ == "__main__":
    try:
        requests.get("https://www.google.com", timeout=5)
    except requests.exceptions.ConnectionError:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("خطای اتصال", "لطفاً اتصال اینترنت خود را بررسی کنید.")
        exit()

    root = tk.Tk()
    app = GoldApp(root)
    root.mainloop()
