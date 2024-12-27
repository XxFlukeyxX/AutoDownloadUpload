from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.firefox import GeckoDriverManager
import requests
import time

# กำหนดข้อมูลผู้ใช้
email = "pichai_jo"
password = "CSautomation"
line_token = "IbJu3fDcHwWbFXjLMCxaRUbSMtwtWeJClf39I5Yf2Je"
reply_persons = ["นางสาวต้องใจ แย้มผกา", "นายปฐมพงษ์ ริณพัฒน์","นางสาวชนิสรา อ่ำสอาด"]

# ตั้งค่า Firefox Options
firefox_options = Options()
firefox_options.set_preference("browser.download.folderList", 1)  # ใช้โฟลเดอร์ Downloads
firefox_options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf")  # ไม่ถามก่อนดาวน์โหลดไฟล์ PDF
firefox_options.set_preference("pdfjs.disabled", True)  # ปิดการแสดง PDF ในเบราว์เซอร์

# สร้าง WebDriver
service = Service(GeckoDriverManager().install())
driver = webdriver.Firefox(service=service, options=firefox_options)

def send_line_notification(token, message):
    url = "https://notify-api.line.me/api/notify"
    headers = {
        "Authorization": f"Bearer {token}",
    }
    data = {
        "message": message
    }

    try:
        response = requests.post(url, headers=headers, data=data)
        response.raise_for_status()  # Raise an exception for HTTP errors
        print("Notification sent successfully!")
    except requests.exceptions.RequestException as e:
        print(f"Error sending notification: {e}")

# กำหนดข้อมูลผู้ตอบกลับเอกสาร

def check_for_user():
    try:
        # เปิดเว็บไซต์
        driver.get("https://e-doc.rmutto.ac.th/home.aspx")
        print("เปิดเว็บไซต์สำเร็จ!")
        time.sleep(5)
        driver.minimize_window()

        # คลิกเมนู "กล่องส่งออกเอกสาร"
        menu_button = driver.find_element(By.XPATH, '/html/body/form/nav/div/div[2]/ul/li[2]/a')
        menu_button.click()
        print("คลิกเมนูสำเร็จ!")
        time.sleep(5)

        # ล็อกอิน
        driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_txtEmail"]').send_keys(email)
        driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_txtPassword"]').send_keys(password)
        driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_bttLogin"]').click()
        print("เข้าสู่ระบบสำเร็จ!")
        time.sleep(5)

        # คลิกปุ่ม "ถัดไป" ในหน้าเอกสาร
        next_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="mainContentPlaceHolder_notificationSourceGridViewDocument_postBackUrlLinkButton_0"]'))
        )
        next_button.click()
        print("คลิกปุ่มถัดไปสำเร็จ!")

        # สลับไปยังหน้าต่างใหม่ (ถ้ามี)
        WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 1)  # รอให้หน้าต่างใหม่เปิดขึ้น
        windows = driver.window_handles
        for w in windows:
            driver.switch_to.window(w)
            if "documentInbox.aspx" in driver.current_url:
                print("สลับไปยังหน้าต่างใหม่สำเร็จ!")
                break
        time.sleep(5)

        # คลิกแท็บ "เอกสารที่ยังไม่ได้เปิดอ่าน"
        tab_unread = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[3]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div/div[1]/ul/li[2]/a/span'))
        )
        tab_unread.click()
        print("เปลี่ยนไปยังแท็บ 'เอกสารที่ยังไม่ได้เปิดอ่าน'")
        time.sleep(5)

        # ค้นหาคำในองค์ประกอบที่กำหนด
        target_xpath = '//*[@id="mainContentPlaceHolder_eDocumentContentByDocumentDirectoryID1_UpdatePanelData"]'
        root_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, target_xpath))
        )
        print("พบองค์ประกอบที่ต้องการค้นหา!")
                # ค้นหาชื่อใน reply_persons
        for person in reply_persons:
            try:
                # ใช้ XPath ค้นหาชื่อ
                person_element = root_element.find_element(By.XPATH, f".//*[contains(text(), '{person}')]")
                if person_element:
                    print(f"พบชื่อ: {person}")

                    # คลิกเข้าไปยังองค์ประกอบที่พบ
                    person_element.click()
                    print(f"คลิกเข้าไปยังชื่อ {person} สำเร็จ!")
                    break  # หยุดเมื่อเจอชื่อที่ตรง
            except Exception as e:
                print(f"ไม่พบชื่อ {person} ในรายการนี้: {e}")

        else:
            print("ไม่พบชื่อใด ๆ ใน reply_persons")

        # ดำเนินการเพิ่มเติมหลังจากคลิก (ถ้าจำเป็น)
        time.sleep(5)

# เรียกใช้งานฟังก์ชัน
check_for_user()

# ปิด WebDriver
driver.quit()