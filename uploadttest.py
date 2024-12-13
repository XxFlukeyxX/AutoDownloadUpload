import os
import glob
from selenium import webdriver
from selenium.webdriver.common.by import By
import time

def get_latest_file(directory):
    """ค้นหาไฟล์ล่าสุดในโฟลเดอร์ที่กำหนด"""
    files = glob.glob(os.path.join(directory, "*"))  # หาไฟล์ทั้งหมดในโฟลเดอร์
    if not files:
        raise FileNotFoundError("ไม่มีไฟล์ในโฟลเดอร์ที่กำหนด")
    latest_file = max(files, key=os.path.getctime)  # หาไฟล์ที่มีวันที่แก้ไขล่าสุด
    return latest_file

# โฟลเดอร์ดาวน์โหลด (ปรับตามระบบของคุณ)
download_folder = "C:\\Users\\mikot\\Downloads"

try:
    # หาไฟล์ล่าสุด
    latest_file = get_latest_file(download_folder)
    print(f"ไฟล์ล่าสุดที่พบ: {latest_file}")

    # เริ่มต้น WebDriver
    driver = webdriver.Firefox()
    driver.implicitly_wait(3)
    driver.maximize_window()

    # เปิดหน้าเว็บที่ต้องการ
    driver.get("https://smallpdf.com/share-document")

    # รอให้หน้าเว็บโหลด
    time.sleep(5)

    # หา input[type="file"] ที่อยู่ใน DOM ซึ่งจะให้เราอัปโหลดไฟล์
    file_input = driver.find_element(By.XPATH, "//input[@type='file']")

    # พิมพ์ path ของไฟล์ที่ต้องการอัปโหลด
    file_input.send_keys(latest_file)

    # รอให้การอัปโหลดเสร็จ
    time.sleep(10)

finally:
    # ปิด browser
    driver.quit()
