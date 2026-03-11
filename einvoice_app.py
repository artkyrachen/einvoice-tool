"""
財政部電子發票 — 統一編號查詢工具
Windows GUI 版本（tkinter）
"""
import re
import io
import time
import base64
import threading
import queue
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
from PIL import Image, ImageFilter, ImageEnhance
import pytesseract
import pandas as pd
import openpyxl
from webdriver_manager.chrome import ChromeDriverManager

TARGET_URL = "https://www.einvoice.nat.gov.tw/portal/btc/btc604w/search"
MAX_RETRY  = 5
WAIT_SEC   = 8
_selenium_ready = True


def create_driver(headless=True):
    opts = Options()
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1280,900")
    opts.add_argument("--lang=zh-TW")
    if headless:
        opts.add_argument("--headless=new")
    try:
        drv = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()), options=opts)
    except Exception:
        drv = webdriver.Chrome(options=opts)
    return drv


def ocr_captcha(driver):
    selectors = [
        "img[src*='captcha']", "img[src*='kaptcha']",
        "img[id*='captcha']",  "img[id*='code']",
        ".captcha img", "#captchaImg", "#imgCaptcha",
    ]
    img_el = None
    for sel in selectors:
        try:
            img_el = driver.find_element(By.CSS_SELECTOR, sel)
            break
        except Exception:
            continue
    if img_el is None:
        raise RuntimeError("找不到驗證碼圖片")

    src = img_el.get_attribute("src")
    if src and src.startswith("data:image"):
        b64 = src.split(",", 1)[1]
        img = Image.open(io.BytesIO(base64.b64decode(b64)))
    else:
        img = Image.open(io.BytesIO(img_el.screenshot_as_png))

    img = img.convert("L")
    img = img.resize((img.width * 3, img.height * 3), Image.LANCZOS)
    img = ImageEnhance.Contrast(img).enhance(2.5)
    img = img.point(lambda p: 255 if p > 140 else 0)
    img = img.filter(ImageFilter.MedianFilter(size=3))

    cfg = ("--oem 3 --psm 8 "
           "-c tessedit_char_whitelist="
           "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789abcdefghijklmnopqrstuvwxyz")
    text = pytesseract.image_to_string(img, config=cfg).strip()
    return re.sub(r"[^A-Za-z0-9]", "", text)


def _map_field(result, label, value):
    mapping = {
        "公司名稱": "公司名稱", "名稱": "公司名稱",
        "負責人": "負責人", "組織": "組織類型",
        "組織類型": "組織類型", "使用統一": "使用發票",
        "使用發票": "使用發票", "業別": "業別", "地址": "地址",
    }
    for key, col in mapping.items():
        if key in label:
            result[col] = value
            return


def query_one(driver, ban_id, log_fn=None):
    ban_id = ban_id.strip()
    result = {
        "統一編號": ban_id, "公司名稱": "", "負責人": "",
        "組織類型": "", "使用發票": "", "業別": "", "地址": "",
        "狀態": "", "查詢時間": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "備註": "",
    }

    def log(msg):
        if log_fn:
            log_fn(msg)

    for attempt in range(1, MAX_RETRY + 1):
        try:
            driver.get(TARGET_URL)
            wait = WebDriverWait(driver, WAIT_SEC)

            # 填統一編號
            input_sels = [
                "input[name='banId']", "input[id='banId']",
                "input[name*='ban']", "input[placeholder*='統一編號']",
                "input[placeholder*='編號']",
            ]
            ban_input = None
            for sel in input_sels:
                try:
                    ban_input = wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, sel)))
                    break
                except Exception:
                    continue
            if ban_input is None:
                result["備註"] = "找不到統一編號輸入欄位"
                return result

            ban_input.clear()
            ban_input.send_keys(ban_id)

            # OCR 驗證碼
            captcha_text = ocr_captcha(driver)
            log(f"  [{attempt}/{MAX_RETRY}] OCR：'{captcha_text}'")

            captcha_in_sels = [
                "input[name*='captcha']", "input[id*='captcha']",
                "input[name*='code']",    "input[id*='code']",
                "input[placeholder*='驗證']",
            ]
            captcha_input = None
            for sel in captcha_in_sels:
                try:
                    captcha_input = driver.find_element(By.CSS_SELECTOR, sel)
                    break
                except Exception:
                    continue
            if captcha_input is None:
                result["備註"] = "找不到驗證碼輸入欄位"
                return result

            captcha_input.clear()
            captcha_input.send_keys(captcha_text)

            # 送出
            submitted = False
            for sel in ["button[type='submit']", "input[type='submit']"]:
                try:
                    driver.find_element(By.CSS_SELECTOR, sel).click()
                    submitted = True
                    break
                except Exception:
                    pass
            if not submitted:
                for xp in ["//button[contains(text(),'查詢')]",
                            "//input[@value='查詢']"]:
                    try:
                        driver.find_element(By.XPATH, xp).click()
                        submitted = True
                        break
                    except Exception:
                        pass
            if not submitted:
                from selenium.webdriver.common.keys import Keys
                captcha_input.send_keys(Keys.RETURN)

            time.sleep(2)

            page = driver.page_source
            if any(k in page for k in ["驗證碼錯誤", "驗證碼不正確",
                                        "請重新輸入", "圖形驗證碼"]):
                log("  驗證碼錯誤，重試...")
                continue

            # 解析 table
            tables = driver.find_elements(By.TAG_NAME, "table")
            for table in tables:
                for row in table.find_elements(By.TAG_NAME, "tr"):
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if len(cells) >= 2:
                        _map_field(result, cells[0].text.strip(),
                                   cells[1].text.strip())
            if result["公司名稱"]:
                result["狀態"] = "成功"
                return result

            if any(k in page for k in ["查無資料", "查無此筆", "無相符"]):
                result["狀態"] = "查無資料"
                return result

            result["狀態"] = "解析失敗"
            result["備註"] = "頁面無法解析"
            return result

        except Exception as e:
            log(f"  嘗試 {attempt} 錯誤：{e}")
            if attempt == MAX_RETRY:
                result["狀態"] = "錯誤"
                result["備註"] = str(e)
    return result


# ════════════════════════════════════════════════════════
#  GUI
# ════════════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("統一編號查詢工具 — 財政部電子發票平台")
        self.geometry("900x680")
        self.resizable(True, True)
        self.configure(bg="#F0F4F8")

        self._results = []
        self._running = False
        self._queue  = queue.Queue()

        self._build_ui()
        self._poll_queue()

    # ── UI 建立 ────────────────────────────────────────
    def _build_ui(self):
        # ── 標題列 ──
        header = tk.Frame(self, bg="#1A365D", height=56)
        header.pack(fill="x")
        tk.Label(header, text="📄  統一編號查詢工具",
                 font=("Microsoft JhengHei", 16, "bold"),
                 bg="#1A365D", fg="white").pack(side="left", padx=20, pady=12)
        tk.Label(header, text="財政部電子發票整合服務平台",
                 font=("Microsoft JhengHei", 9),
                 bg="#1A365D", fg="#90CDF4").pack(side="left", padx=4, pady=16)

        # ── 主體 ──
        body = tk.Frame(self, bg="#F0F4F8")
        body.pack(fill="both", expand=True, padx=20, pady=16)

        # 左側輸入區
        left = tk.LabelFrame(body, text="輸入設定",
                              font=("Microsoft JhengHei", 10, "bold"),
                              bg="#F0F4F8", fg="#2D3748", bd=1,
                              relief="groove")
        left.pack(side="left", fill="both", expand=False,
                  padx=(0, 12), ipadx=8, ipady=8)
        left.configure(width=280)
        left.pack_propagate(False)

        # 單筆
        tk.Label(left, text="單筆查詢（統一編號）：",
                 font=("Microsoft JhengHei", 9),
                 bg="#F0F4F8", fg="#4A5568").pack(anchor="w", padx=8, pady=(12,2))
        single_frame = tk.Frame(left, bg="#F0F4F8")
        single_frame.pack(fill="x", padx=8)
        self.single_var = tk.StringVar()
        tk.Entry(single_frame, textvariable=self.single_var,
                 font=("Consolas", 11), width=12).pack(side="left")
        tk.Button(single_frame, text="查詢", bg="#2B6CB0", fg="white",
                  font=("Microsoft JhengHei", 9, "bold"),
                  relief="flat", padx=10,
                  command=self._run_single).pack(side="left", padx=(6, 0))

        ttk.Separator(left, orient="horizontal").pack(fill="x",
                                                       padx=8, pady=14)

        # 批次
        tk.Label(left, text="批次查詢（每行一個編號）：",
                 font=("Microsoft JhengHei", 9),
                 bg="#F0F4F8", fg="#4A5568").pack(anchor="w", padx=8, pady=(0,4))
        self.batch_text = scrolledtext.ScrolledText(
            left, width=28, height=10,
            font=("Consolas", 10), wrap="none")
        self.batch_text.pack(padx=8)

        btn_row = tk.Frame(left, bg="#F0F4F8")
        btn_row.pack(fill="x", padx=8, pady=8)
        tk.Button(btn_row, text="📂 匯入 CSV/Excel",
                  bg="#EDF2F7", fg="#2D3748",
                  font=("Microsoft JhengHei", 9),
                  relief="flat", command=self._import_file
                  ).pack(side="left", fill="x", expand=True, padx=(0,4))
        tk.Button(btn_row, text="▶ 批次執行",
                  bg="#276749", fg="white",
                  font=("Microsoft JhengHei", 9, "bold"),
                  relief="flat", command=self._run_batch
                  ).pack(side="left", fill="x", expand=True)

        # Headless 選項
        self.headless_var = tk.BooleanVar(value=True)
        tk.Checkbutton(left, text="背景執行（不顯示瀏覽器）",
                       variable=self.headless_var,
                       font=("Microsoft JhengHei", 9),
                       bg="#F0F4F8", activebackground="#F0F4F8"
                       ).pack(anchor="w", padx=8, pady=(4,0))

        tk.Button(left, text="⏹ 停止", bg="#C53030", fg="white",
                  font=("Microsoft JhengHei", 9, "bold"),
                  relief="flat", command=self._stop
                  ).pack(fill="x", padx=8, pady=(8,4))

        # 右側結果區
        right = tk.Frame(body, bg="#F0F4F8")
        right.pack(side="left", fill="both", expand=True)

        # 結果 Treeview
        tree_frame = tk.LabelFrame(right, text="查詢結果",
                                   font=("Microsoft JhengHei", 10, "bold"),
                                   bg="#F0F4F8", fg="#2D3748",
                                   bd=1, relief="groove")
        tree_frame.pack(fill="both", expand=True)

        cols = ("統一編號", "公司名稱", "負責人", "業別", "地址", "狀態")
        self.tree = ttk.Treeview(tree_frame, columns=cols,
                                  show="headings", height=12)
        col_w = {"統一編號": 90, "公司名稱": 160, "負責人": 70,
                 "業別": 80, "地址": 180, "狀態": 60}
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=col_w[c], minwidth=50)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical",
                             command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True, padx=4, pady=4)

        # 狀態顏色 tag
        self.tree.tag_configure("成功",    background="#F0FFF4")
        self.tree.tag_configure("查無資料", background="#FFFBEB")
        self.tree.tag_configure("錯誤",    background="#FFF5F5")
        self.tree.tag_configure("解析失敗", background="#FFF5F5")

        # 輸出按鈕列
        out_row = tk.Frame(right, bg="#F0F4F8")
        out_row.pack(fill="x", pady=(8, 0))
        tk.Button(out_row, text="💾 匯出 Excel",
                  bg="#2B6CB0", fg="white",
                  font=("Microsoft JhengHei", 9, "bold"),
                  relief="flat", command=lambda: self._export("xlsx")
                  ).pack(side="left", padx=(0, 6))
        tk.Button(out_row, text="💾 匯出 CSV",
                  bg="#2B6CB0", fg="white",
                  font=("Microsoft JhengHei", 9, "bold"),
                  relief="flat", command=lambda: self._export("csv")
                  ).pack(side="left", padx=(0, 6))
        tk.Button(out_row, text="🗑 清除結果",
                  bg="#718096", fg="white",
                  font=("Microsoft JhengHei", 9),
                  relief="flat", command=self._clear
                  ).pack(side="left")

        # 日誌
        log_frame = tk.LabelFrame(right, text="執行紀錄",
                                   font=("Microsoft JhengHei", 9),
                                   bg="#F0F4F8", fg="#4A5568",
                                   bd=1, relief="groove")
        log_frame.pack(fill="x", pady=(8, 0))
        self.log_box = scrolledtext.ScrolledText(
            log_frame, height=5, state="disabled",
            font=("Consolas", 9), bg="#1A202C", fg="#A0AEC0",
            insertbackground="white", wrap="word")
        self.log_box.pack(fill="x", padx=4, pady=4)

        # 進度條
        self.progress = ttk.Progressbar(right, mode="indeterminate")
        self.progress.pack(fill="x", pady=(6, 0))

    # ── 日誌 ────────────────────────────────────────────
    def _log(self, msg):
        self._queue.put(("log", msg))

    def _poll_queue(self):
        try:
            while True:
                kind, data = self._queue.get_nowait()
                if kind == "log":
                    self.log_box.configure(state="normal")
                    self.log_box.insert("end",
                        f"[{datetime.now().strftime('%H:%M:%S')}] {data}\n")
                    self.log_box.see("end")
                    self.log_box.configure(state="disabled")
                elif kind == "row":
                    r = data
                    tag = r.get("狀態", "")
                    self.tree.insert("", "end", values=(
                        r["統一編號"], r["公司名稱"], r["負責人"],
                        r["業別"], r["地址"], r["狀態"]), tags=(tag,))
                    self._results.append(r)
                elif kind == "done":
                    self.progress.stop()
                    self._running = False
                    self._log(f"✅ 完成，共 {len(self._results)} 筆")
                elif kind == "error":
                    self.progress.stop()
                    self._running = False
                    messagebox.showerror("錯誤", data)
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    # ── 匯入檔案 ────────────────────────────────────────
    def _import_file(self):
        path = filedialog.askopenfilename(
            title="選擇清單檔案",
            filetypes=[("CSV/Excel", "*.csv *.xlsx *.xls"), ("所有檔案", "*.*")])
        if not path:
            return
        try:
            if path.endswith(".csv"):
                import pandas as _pd
                df = _pd.read_csv(path, dtype=str)
            else:
                import pandas as _pd
                df = _pd.read_excel(path, dtype=str)
            col = None
            for c in df.columns:
                if "編號" in c or "ban" in c.lower():
                    col = c
                    break
            if col is None:
                col = df.columns[0]
            ids = df[col].dropna().tolist()
            self.batch_text.delete("1.0", "end")
            self.batch_text.insert("end", "\n".join(ids))
            self._log(f"已匯入 {len(ids)} 筆（欄位：{col}）")
        except Exception as e:
            messagebox.showerror("匯入失敗", str(e))

    # ── 執行查詢 ────────────────────────────────────────
    def _get_ban_list_from_textbox(self):
        raw = self.batch_text.get("1.0", "end").strip()
        ids = [x.strip() for x in re.split(r"[\n,，\s]+", raw) if x.strip()]
        return ids

    def _run_single(self):
        ban = self.single_var.get().strip()
        if not ban:
            messagebox.showwarning("提示", "請輸入統一編號")
            return
        self._run([ban])

    def _run_batch(self):
        ids = self._get_ban_list_from_textbox()
        if not ids:
            messagebox.showwarning("提示", "請輸入統一編號清單")
            return
        self._run(ids)

    def _run(self, ban_list):
        if self._running:
            messagebox.showwarning("提示", "查詢進行中，請稍候")
            return
        self._running = True
        self.progress.start(10)
        headless = self.headless_var.get()
        t = threading.Thread(target=self._worker,
                             args=(ban_list, headless), daemon=True)
        t.start()

    def _worker(self, ban_list, headless):
        try:
            drv = create_driver(headless=headless)
            try:
                for i, ban in enumerate(ban_list, 1):
                    if not self._running:
                        break
                    self._queue.put(("log",
                        f"[{i}/{len(ban_list)}] 查詢：{ban}"))
                    r = query_one(drv, ban, log_fn=self._log)
                    self._queue.put(("row", r))
            finally:
                drv.quit()
        except Exception as e:
            self._queue.put(("error", str(e)))
            return
        self._queue.put(("done", None))

    def _stop(self):
        self._running = False
        self._log("⏹ 使用者停止查詢")

    # ── 匯出 ────────────────────────────────────────────
    def _export(self, fmt):
        if not self._results:
            messagebox.showwarning("提示", "尚無查詢結果")
            return
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        if fmt == "xlsx":
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=f"einvoice_{ts}.xlsx",
                filetypes=[("Excel", "*.xlsx")])
        else:
            path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                initialfile=f"einvoice_{ts}.csv",
                filetypes=[("CSV", "*.csv")])
        if not path:
            return
        try:
            import pandas as _pd
            df = _pd.DataFrame(self._results)
            if fmt == "xlsx":
                import openpyxl as _ox
                with _pd.ExcelWriter(path, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name="查詢結果")
                    ws = writer.sheets["查詢結果"]
                    for col_cells in ws.columns:
                        length = max(
                            len(str(c.value or "")) for c in col_cells)
                        ws.column_dimensions[
                            col_cells[0].column_letter].width = min(
                                length + 4, 50)
            else:
                df.to_csv(path, index=False, encoding="utf-8-sig")
            self._log(f"✅ 已匯出：{path}")
            messagebox.showinfo("完成", f"已儲存至：\n{path}")
        except Exception as e:
            messagebox.showerror("匯出失敗", str(e))

    def _clear(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._results.clear()
        self._log("已清除結果")


# ════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = App()
    app.mainloop()
