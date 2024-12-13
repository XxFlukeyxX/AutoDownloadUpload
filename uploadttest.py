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

# โฟลเดอร์ดาวน์โหลด (ปรับตามระบบของคุณ)
download_folder = "C:\\Users\\mikot\\Downloads"  # กำหนด path ของโฟลเดอร์ดาวน์โหลด

try:
    # หาไฟล์ล่าสุดในโฟลเดอร์ดาวน์โหลด
    latest_file = get_latest_file(download_folder)  # เรียกฟังก์ชันเพื่อหาไฟล์ล่าสุด
    print(f"ไฟล์ล่าสุดที่พบ: {latest_file}")  # แสดง path ของไฟล์ล่าสุดที่พบ

    # เริ่มต้น WebDriver เพื่อเปิดเบราว์เซอร์
    driver = webdriver.Firefox()  # เปิด Firefox Browser ด้วย Selenium
    driver.implicitly_wait(3)  # กำหนดเวลาให้รอการโหลด element สูงสุด 3 วินาที
    driver.maximize_window()  # ขยายหน้าต่างเบราว์เซอร์ให้เต็มจอ

    # เปิดหน้าเว็บที่ต้องการ
    driver.get("https://smallpdf.com/share-document")  # เปิด URL สำหรับอัปโหลดไฟล์

    # รอให้หน้าเว็บโหลดเสร็จ
    time.sleep(5)  # หน่วงเวลาการทำงาน 5 วินาที เพื่อรอให้หน้าเว็บโหลดเสร็จ

    # หา input ที่ใช้สำหรับอัปโหลดไฟล์ (โดยใช้ XPATH)
    file_input = driver.find_element(By.XPATH, "//input[@type='file']")  # ค้นหา input[type="file"]

    # อัปโหลดไฟล์ล่าสุดโดยส่ง path ของไฟล์เข้าไปใน input
    file_input.send_keys(latest_file)  # ส่ง path ของไฟล์เข้าไปใน input[type="file"]

    # รอให้การอัปโหลดเสร็จ
    time.sleep(10)  # หน่วงเวลา 10 วินาที เพื่อให้การอัปโหลดไฟล์เสร็จสิ้น

finally:
    # ปิดเบราว์เซอร์เพื่อคืนทรัพยากร
    driver.quit()  # ปิดเบราว์เซอร์ไม่ว่าจะเกิดข้อผิดพลาดหรือไม่
