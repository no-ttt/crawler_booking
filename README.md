# Booking.com 飯店價格爬蟲

這是一個使用 Python 和 Selenium 開發的網路爬蟲，專門用於抓取 [Booking.com](https://www.booking.com/) 網站上的飯店房型與價格資訊。

專案包含兩種抓取模式：
1.  **標準模式 (`booking.py`)**: 以匿名訪客身份進行爬取。
2.  **登入模式 (`booking_login.py`)**: 使用預先儲存的 Cookie 模擬會員登入狀態，可能可以抓取到「Genius 會員價」。

抓取完成後，程式會將結果整理並匯出成一個格式化且易於閱讀的 Excel 檔案。

## 功能特色

- **雙模式抓取**: 支援匿名與會員登入模式。
- **動態日期設定**: 自動將查詢日期設定為執行日的一週後，方便每日追蹤。
- **錯誤處理與重試**: 當抓取失敗時，會自動重啟瀏覽器並重試，提高成功率。
- **客製化飯店清單**: 可在程式碼中輕鬆修改 `hotel_urls` 列表來指定想查詢的飯店。
- **自動化 Excel 報告**:
    - 將同一飯店、同一日期的所有房型整理在同一區塊。
    - 房型與價格分兩列顯示，清晰明瞭。
    - 自動合併儲存格、設定樣式與調整欄寬，提升報告可讀性。
    - 輸出檔案以時間戳命名，方便管理。

## 專案結構

```
.
├── booking.py              # 標準模式爬蟲主程式
├── booking_login.py        # Cookie 登入模式爬蟲主程式
├── save_cookies.py         # 用於手動登入並儲存 Cookie 的工具
├── booking_cookies.pkl     # (執行 save_cookies.py 後產生) 儲存登入資訊的檔案
└── README.md               # 本說明文件
```

## 環境建置

### 1. 前置需求

- **Python 3.8+**
- **Google Chrome 瀏覽器**
- **ChromeDriver**: 版本需與您的 Chrome 瀏覽器版本對應。
    - ChromeDriver 下載頁面
    - 下載後，請將 `chromedriver` 執行檔放置於專案根目錄下，或任何系統 `PATH` 路徑中。

### 2. 安裝 Python 套件

在專案根目錄下，透過 pip 安裝所需的函式庫：

```bash
pip install selenium pandas openpyxl
```

## 使用教學

### 步驟 1: (僅限登入模式) 儲存登入 Cookie

如果您想使用會員價模式，請先執行此步驟。若只使用標準模式，可跳過此步驟。

1.  在終端機中執行 `save_cookies.py`：
    ```bash
    python save_cookies.py
    ```
2.  程式會自動開啟一個 Chrome 瀏覽器並前往 Booking.com。
3.  請在該瀏覽器中**手動完成登入**操作。
4.  登入成功後，回到終端機視窗，按下 `Enter` 鍵。
5.  程式會將您的登入資訊儲存為 `booking_cookies.pkl` 檔案，並關閉瀏覽器。

> **注意**: Cookie 具有時效性，如果登入模式失效，請重新執行此步驟以更新 Cookie。

### 步驟 2: 修改要查詢的飯店列表

打開 `booking.py` 或 `booking_login.py`，找到 `hotel_urls` 變數，並將其內容替換為您想查詢的飯店網址。

```python
# 位於檔案底部 if __name__ == "__main__": 區塊中
hotel_urls = [
    f"https://www.booking.com/hotel/tw/hotel-1.zh-tw.html?checkin={checkin_date}...",
    f"https://www.booking.com/hotel/tw/hotel-2.zh-tw.html?checkin={checkin_date}..."
]
```

> **提示**: 建議直接從 Booking.com 網站上複製包含 `checkin`、`checkout` 等參數的完整網址，程式會自動處理日期。

### 步驟 3: 執行爬蟲

- **執行標準模式**:
  ```bash
  python booking.py
  ```

- **執行登入模式**:
  ```bash
  python booking_login.py
  ```

程式會開始執行，並在終端機中顯示詳細的抓取進度。

### 步驟 4: 查看結果

執行完畢後，專案目錄下會產生一個以執行時間命名的 `.xlsx` 檔案 (例如 `20231027_153000.xlsx` 或 `20231027_153000_login.xlsx`)，其中包含了所有抓取到的飯店價格資訊。

---

*免責聲明：本專案僅供學術研究與技術交流，請遵守網站的使用條款，勿進行惡意或過於頻繁的抓取。*