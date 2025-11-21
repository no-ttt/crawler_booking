import pickle
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

def save_booking_cookies():
    """
    手動登入 Booking.com 並儲存 cookies。
    """
    options = Options()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--window-size=1200,800")
    
    driver = webdriver.Chrome(options=options)
    
    # 前往 Booking.com 首頁，因為需要在此網域下才能儲存 cookie
    driver.get("https://www.booking.com")

    print("\n" + "="*60)
    print("  請在自動開啟的瀏覽器中手動登入 Booking.com。")
    input("  登入成功後，請回到此視窗並按下 Enter 鍵繼續...")
    print("="*60 + "\n")

    # 儲存 cookie 到檔案
    with open("booking_cookies.pkl", "wb") as file:
        pickle.dump(driver.get_cookies(), file)

    print("✅ Cookies 已成功儲存至 booking_cookies.pkl")
    driver.quit()

if __name__ == "__main__":
    save_booking_cookies()