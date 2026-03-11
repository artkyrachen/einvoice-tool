# 統一編號查詢工具 — 使用說明

## 一、首次使用（打包成 .exe）

### 前置需求
1. **Python 3.9+**　https://www.python.org/downloads/
   - 安裝時勾選「Add Python to PATH」
2. **Google Chrome**（需與 ChromeDriver 版本相符）
3. **Tesseract OCR**　https://github.com/UB-Mannheim/tesseract/wiki
   - 安裝路徑建議：`C:\Program Files\Tesseract-OCR\`
   - 安裝完畢後將路徑加入系統 PATH

### 打包步驟
1. 雙擊執行 **`打包exe.bat`**
2. 等待約 1–3 分鐘
3. 完成後執行檔位於 `dist\統一編號查詢工具.exe`
4. 可將 `.exe` 複製到任意位置使用（目標機器需安裝 Tesseract + Chrome）

---

## 二、工具操作說明

### 單筆查詢
1. 在左上方「統一編號」欄位輸入 8 碼編號
2. 點擊「查詢」按鈕

### 批次查詢
- **手動輸入**：在文字區域每行貼一個統一編號
- **匯入檔案**：點「📂 匯入 CSV/Excel」，選擇含統一編號欄位的檔案
  - CSV：欄位名稱含「編號」或 `ban`（或第一欄）
  - Excel：同上規則
- 點「▶ 批次執行」開始查詢

### 匯出結果
- 「💾 匯出 Excel」→ 儲存為 `.xlsx`（自動調整欄寬）
- 「💾 匯出 CSV」→ 儲存為 `.csv`（UTF-8 BOM，Excel 可直接開啟）

---

## 三、常見問題

| 問題 | 解決方式 |
|------|----------|
| 驗證碼一直失敗 | 取消勾選「背景執行」觀察瀏覽器狀況；或調整 `MAX_RETRY` |
| 找不到輸入欄位 | 網站改版，請回報 selector 讓開發者更新 |
| Tesseract 找不到 | 在 `einvoice_app.py` 頂部加入：`pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"` |
| Chrome 版本不符 | `webdriver-manager` 會自動下載對應版本；或手動安裝 ChromeDriver |

---

## 四、輸出欄位說明

| 欄位 | 說明 |
|------|------|
| 統一編號 | 查詢輸入值 |
| 公司名稱 | 登記名稱 |
| 負責人 | 負責人姓名 |
| 組織類型 | 公司/行號/其他 |
| 使用發票 | 是否使用統一發票 |
| 業別 | 行業別 |
| 地址 | 登記地址 |
| 狀態 | 成功 / 查無資料 / 錯誤 / 解析失敗 |
| 查詢時間 | 執行時間戳記 |
| 備註 | 錯誤訊息（如有） |
