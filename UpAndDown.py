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

def download_file(download_folder, login_url, username, password):
    """เข้าสู่ระบบและดาวน์โหลดไฟล์"""
    driver = webdriver.Firefox()
    driver.implicitly_wait(3)
    
    try:
        driver.get(login_url)
        driver.maximize_window()
        
        # คลิกเข้าสู่หน้าล็อกอิน
        driver.find_element(By.XPATH, "/html/body/table/tbody/tr[1]/td[1]/table/tbody/tr[4]/td/a").click()
        
        # กรอกข้อมูลล็อกอิน
        driver.find_element(By.XPATH, "/html/body/table/tbody/tr[1]/td[3]/form/div/table[2]/tbody/tr[1]/td[4]/input").send_keys(username)
        driver.find_element(By.XPATH, "/html/body/table/tbody/tr[1]/td[3]/form/div/table[2]/tbody/tr[2]/td[2]/input[1]").send_keys(password)
        driver.find_element(By.XPATH, "/html/body/table/tbody/tr[1]/td[3]/form/div/table[2]/tbody/tr[3]/td[2]/font/input").click()
        
        # คลิกเพื่อดาวน์โหลดไฟล์
        driver.find_element(By.XPATH, "/html/body/table/tbody/tr[1]/td[1]/table/tbody/tr[6]/td/a").click()
        driver.find_element(By.XPATH, "/html/body/table/tbody/tr[1]/td[3]/div/table[2]/tbody/tr/td/table/tbody/tr[2]/td[1]/a").click()
        
        time.sleep(5)  # รอการดาวน์โหลดไฟล์
    finally:
        driver.quit()

def upload_file(latest_file, upload_url):
    """อัปโหลดไฟล์ไปยังเว็บไซต์"""
    driver = webdriver.Firefox()
    driver.implicitly_wait(3)
    
    try:
        driver.get(upload_url)
        driver.maximize_window()
        
        # หา input[type="file"] และอัปโหลดไฟล์
        file_input = driver.find_element(By.XPATH, "//input[@type='file']")
        file_input.send_keys(latest_file)
        
        # รอให้การอัปโหลดเสร็จ
        time.sleep(10)
    finally:
        driver.quit()

def main():
    download_folder = "C:\\Users\\mikot\\Downloads"
    login_url = "http://regis.rmutto.ac.th/registrar/home.asp"
    upload_url = "https://smallpdf.com/share-document"
    username = "026530461019-8"
    password = "1102700828213"
    
    # ขั้นตอนที่ 1: ดาวน์โหลดไฟล์
    download_file(download_folder, login_url, username, password)
    
    # ขั้นตอนที่ 2: หาไฟล์ล่าสุด
    latest_file = get_latest_file(download_folder)
    print(f"ไฟล์ล่าสุดที่พบ: {latest_file}")
    
    # ขั้นตอนที่ 3: อัปโหลดไฟล์
    upload_file(latest_file, upload_url)

if __name__ == "__main__":
    main()
