"""
Script đánh thức app Streamlit (chạy bằng GitHub Actions).
Không dùng làm file chính trên Streamlit Cloud — cấu hình Main file là app.py.
"""
import sys

# Nếu đang chạy trong môi trường Streamlit (Cloud), không dùng selenium
if "streamlit" in sys.modules:
    import streamlit as st
    st.error(
        "Đây không phải ứng dụng chính. Trên Streamlit Cloud, hãy đặt **Main file path** thành **app.py** (Manage app → Settings)."
    )
    st.stop()

import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# --- CÁC THAM SỐ CÓ THỂ ĐIỀU CHỈNH ---
APP_URL = "https://vqgfnjwpybwjtqrh6senci.streamlit.app/"
TOTAL_RUN_TIME_MINUTES = 10  # Tổng thời gian kịch bản sẽ chạy (phút)
REFRESH_INTERVAL_MINUTES = 3   # Khoảng thời gian giữa mỗi lần làm mới (phút)
# ------------------------------------

print("--- Setting up headless Chrome browser ---")
options = webdriver.ChromeOptions()
options.add_argument("--headless")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")

service = ChromeService(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

try:
    print(f"--- Navigating to {APP_URL} ---")
    driver.get(APP_URL)

    # Đánh thức ứng dụng lần đầu nếu cần
    try:
        print("--- Checking if app is asleep (initial check)... ---")
        short_wait = WebDriverWait(driver, 5)
        wakeup_button = short_wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Yes, get this app back up!')]"))
        )
        print("--- App is asleep. Clicking wakeup button. ---")
        wakeup_button.click()
        # Chờ ứng dụng tải xong sau khi nhấn nút
        long_wait = WebDriverWait(driver, 120)
        long_wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
    except TimeoutException:
        print("--- Wakeup button not found, assuming app is already awake. ---")

    # Bắt đầu vòng lặp giữ cho ứng dụng thức
    start_time = time.time()
    total_run_seconds = TOTAL_RUN_TIME_MINUTES * 60
    refresh_interval_seconds = REFRESH_INTERVAL_MINUTES * 60
    next_refresh_time = start_time + refresh_interval_seconds

    print(f"--- Starting keep-awake loop for {TOTAL_RUN_TIME_MINUTES} minutes. Refreshing every {REFRESH_INTERVAL_MINUTES} minutes. ---")

    while time.time() - start_time < total_run_seconds:
        if time.time() > next_refresh_time:
            print(f"--- Refreshing page at {time.strftime('%H:%M:%S')} ---")
            driver.refresh()
            next_refresh_time += refresh_interval_seconds
        time.sleep(15) # Ngủ 15 giây để giảm tải CPU

    print(f"--- Loop finished after {TOTAL_RUN_TIME_MINUTES} minutes. ---")

except Exception as e:
    print("--- An unexpected error occurred. ---")
    print(f"Details: {e}")
    driver.save_screenshot("debug_screenshot.png")
    raise e

finally:
    print("--- Closing browser ---")
    driver.quit()
