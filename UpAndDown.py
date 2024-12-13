import os  # ใช้จัดการ path และไฟล์ในระบบปฏิบัติการ
import glob  # ใช้ค้นหาไฟล์ในโฟลเดอร์ตามรูปแบบที่กำหนด
from selenium import webdriver  # ใช้ควบคุมเบราว์เซอร์ด้วย Selenium
from selenium.webdriver.common.by import By  # ใช้ระบุวิธีการค้นหา element ในหน้าเว็บ
import time  # ใช้หน่วงเวลาการทำงานของโปรแกรม

def get_latest_file(directory):
    """ค้นหาไฟล์ล่าสุดในโฟลเดอร์ที่กำหนด"""
    files = glob.glob(os.path.join(directory, "*"))  # หาไฟล์ทั้งหมดในโฟลเดอร์ที่ระบุ
    if not files:  # ถ้าไม่มีไฟล์ในโฟลเดอร์
        raise FileNotFoundError("ไม่มีไฟล์ในโฟลเดอร์ที่กำหนด")  # ส่งข้อผิดพลาดว่าไม่พบไฟล์
    latest_file = max(files, key=os.path.getctime)  # ค้นหาไฟล์ที่มีวันที่แก้ไขล่าสุด
    return latest_file  # ส่งคืน path ของไฟล์ล่าสุด

def download_file(download_folder, login_url, username, password):
    """เข้าสู่ระบบและดาวน์โหลดไฟล์"""
    driver = webdriver.Firefox()  # เริ่มต้นเบราว์เซอร์ Firefox
    driver.implicitly_wait(3)  # กำหนดเวลาให้รอการโหลด element สูงสุด 3 วินาที
    
    try:
        driver.get(login_url)  # เปิดหน้าเว็บที่ระบุใน login_url
        driver.maximize_window()  # ขยายหน้าต่างเบราว์เซอร์ให้เต็มจอ
        
        # คลิกเข้าสู่หน้าล็อกอิน
        driver.find_element(By.XPATH, "/html/body/table/tbody/tr[1]/td[1]/table/tbody/tr[4]/td/a").click()  # คลิกลิงก์ไปยังหน้าล็อกอิน
        
        # กรอกข้อมูลล็อกอิน
        driver.find_element(By.XPATH, "/html/body/table/tbody/tr[1]/td[3]/form/div/table[2]/tbody/tr[1]/td[4]/input").send_keys(username)  # กรอกชื่อผู้ใช้
        driver.find_element(By.XPATH, "/html/body/table/tbody/tr[1]/td[3]/form/div/table[2]/tbody/tr[2]/td[2]/input[1]").send_keys(password)  # กรอกรหัสผ่าน
        driver.find_element(By.XPATH, "/html/body/table/tbody/tr[1]/td[3]/form/div/table[2]/tbody/tr[3]/td[2]/font/input").click()  # คลิกปุ่มเข้าสู่ระบบ
        
        # คลิกเพื่อดาวน์โหลดไฟล์
        driver.find_element(By.XPATH, "/html/body/table/tbody/tr[1]/td[1]/table/tbody/tr[6]/td/a").click()  # คลิกที่ลิงก์ไปยังหน้าดาวน์โหลด
        driver.find_element(By.XPATH, "/html/body/table/tbody/tr[1]/td[3]/div/table[2]/tbody/tr/td/table/tbody/tr[2]/td[1]/a").click()  # คลิกเพื่อดาวน์โหลดไฟล์
        
        time.sleep(5)  # รอให้การดาวน์โหลดเสร็จ
    finally:
        driver.quit()  # ปิดเบราว์เซอร์เพื่อคืนทรัพยากร

def upload_file(latest_file, upload_url):
    """อัปโหลดไฟล์ไปยังเว็บไซต์"""
    driver = webdriver.Firefox()  # เริ่มต้นเบราว์เซอร์ Firefox
    driver.implicitly_wait(3)  # กำหนดเวลาให้รอการโหลด element สูงสุด 3 วินาที
    
    try:
        driver.get(upload_url)  # เปิดหน้าเว็บที่ระบุใน upload_url
        driver.maximize_window()  # ขยายหน้าต่างเบราว์เซอร์ให้เต็มจอ
        
        # หา input[type="file"] และอัปโหลดไฟล์
        file_input = driver.find_element(By.XPATH, "//input[@type='file']")  # ค้นหา input[type="file"] สำหรับการอัปโหลดไฟล์
        file_input.send_keys(latest_file)  # ส่ง path ของไฟล์เข้าไปใน input[type="file"]
        
        # รอให้การอัปโหลดเสร็จ
        time.sleep(10)  # หน่วงเวลา 10 วินาทีเพื่อรอให้การอัปโหลดไฟล์เสร็จสิ้น
    finally:
        driver.quit()  # ปิดเบราว์เซอร์เพื่อคืนทรัพยากร

def main():
    """ฟังก์ชันหลักเพื่อควบคุมการทำงาน"""
    download_folder = "C:\\Users\\mikot\\Downloads"  # โฟลเดอร์ที่ไฟล์จะถูกดาวน์โหลดลงไป
    login_url = "http://regis.rmutto.ac.th/registrar/home.asp"  # URL ของหน้าล็อกอิน
    upload_url = "https://smallpdf.com/share-document"  # URL ของหน้าสำหรับอัปโหลดไฟล์
    username = "026530461019-8"  # ชื่อผู้ใช้ (ปรับให้เหมาะสม)
    password = "1102700828213"  # รหัสผ่าน (ปรับให้เหมาะสม)
    
    # ขั้นตอนที่ 1: ดาวน์โหลดไฟล์
    download_file(download_folder, login_url, username, password)  # ดาวน์โหลดไฟล์จากเว็บ
    
    # ขั้นตอนที่ 2: หาไฟล์ล่าสุด
    latest_file = get_latest_file(download_folder)  # ค้นหาไฟล์ล่าสุดในโฟลเดอร์ดาวน์โหลด
    print(f"ไฟล์ล่าสุดที่พบ: {latest_file}")  # แสดง path ของไฟล์ล่าสุดที่พบ
    
    # ขั้นตอนที่ 3: อัปโหลดไฟล์
    upload_file(latest_file, upload_url)  # อัปโหลดไฟล์ล่าสุดไปยังเว็บ

if __name__ == "__main__":
    main()  # เรียกใช้งานฟังก์ชันหลักเพื่อเริ่มต้นกระบวนการ
