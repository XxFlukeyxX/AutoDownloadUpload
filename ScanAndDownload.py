from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from webdriver_manager.firefox import GeckoDriverManager
import os
import requests
import time

# 🛠️ กำหนดข้อมูลผู้ใช้
email = "pichai_jo"
password = "CSautomation"
person = "นางสาวต้องใจ แย้มผกา"
line_token = "IbJu3fDcHwWbFXjLMCxaRUbSMtwtWeJClf39I5Yf2Je"

# 🛠️ ตั้งค่า Firefox Options สำหรับดาวน์โหลดในโฟลเดอร์ Downloads
firefox_options = Options()
firefox_options.set_preference("browser.download.folderList", 1)  # 1: ใช้โฟลเดอร์ Downloads ของเครื่อง
firefox_options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf")  # ไม่ถามก่อนดาวน์โหลด PDF
firefox_options.set_preference("pdfjs.disabled", True)  # ปิดตัวแสดง PDF ในเบราว์เซอร์

# 🛠️ สร้าง WebDriver
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

try:
    # 1️⃣ เปิดเว็บไซต์
    driver.get("https://e-doc.rmutto.ac.th/home.aspx")
    print("เปิดเว็บไซต์สำเร็จ!")
    time.sleep(5)

        # 3️⃣ คลิกเมนู "กล่องส่งออกเอกสาร"
    menu_button = driver.find_element("xpath", '/html/body/form/nav/div/div[2]/ul/li[2]/a')
    menu_button.click()
    print("คลิกเมนูสำเร็จ!")
    time.sleep(3)

    # 2️⃣ ล็อกอิน
    driver.find_element("xpath", '//*[@id="mainContentPlaceHolder_txtEmail"]').send_keys(email)
    driver.find_element("xpath", '//*[@id="mainContentPlaceHolder_txtPassword"]').send_keys(password)
    driver.find_element("xpath", '//*[@id="mainContentPlaceHolder_bttLogin"]').click()
    print("เข้าสู่ระบบสำเร็จ!")
    time.sleep(5)

    # 6️⃣ คลิกปุ่มถัดไปในหน้าเอกสาร
    next_button = driver.find_element("xpath", '//*[@id="mainContentPlaceHolder_notificationSourceGridViewDocument_postBackUrlLinkButton_0"]')
    next_button.click()
    print("คลิกปุ่มถัดไปสำเร็จ!")
    time.sleep(10)

    # 7️⃣ สลับไปยังหน้าต่างใหม่ (ถ้ามี)
    windows = driver.window_handles
    for w in windows:
        driver.switch_to.window(w)
        if "documentInbox.aspx" in driver.current_url:
            print("สลับไปยังหน้าต่างใหม่สำเร็จ!")
            break
    time.sleep(5)

    new_page_button = driver.find_element("xpath", '/html/body/form/div[3]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div/div[1]/ul/li[2]/a/span')
    new_page_button.click()
    time.sleep(5)

    rows = driver.find_elements("xpath", '//tr')
    for row in rows:
        if "รายงานขอจัดซื้อจัดจ?าง" in row.text.strip():
            button = row.find_element("xpath", './/a')
            button.click()
            break

    time.sleep(10)

        # 7️⃣ สลับไปยังหน้าต่างใหม่ (ถ้ามี)
    windows = driver.window_handles
    for w in windows:
        driver.switch_to.window(w)
        if "documentDetail.aspx" in driver.current_url:
            print("สลับไปยังหน้าต่างใหม่สำเร็จ!")
            break
    time.sleep(5)

    downloaded_pdfs = set()

    # 4️⃣ ค้นหาแถวที่มีชื่อ "นางสาวต้องใจ แย้มผกา"
    rows = driver.find_elements("xpath", '//tr')  # ดึงทุกแถวในตาราง
    for row in rows:
        if person in row.text:  # ตรวจสอบข้อความในแถว
            print(f"พบแถวที่มีชื่อ: {row.text}")

        # ค้นหาลิงก์ PDF ในแถวที่ตรงกับชื่อ
            pdf_links = row.find_elements("xpath", './/a[contains(@href, "DocumentGenerateFile.ashx")]')

        # ตรวจสอบจำนวนลิงก์ PDF
            if len(pdf_links) > 0:
                print(f"พบลิงก์ PDF ในแถวนี้: {len(pdf_links)}")

            # ดาวน์โหลดไฟล์ PDF ทีละลิงก์
            for index, pdf_link in enumerate(pdf_links):
                pdf_url = pdf_link.get_attribute('href')  # ดึง URL ของไฟล์ PDF
                if pdf_url not in downloaded_pdfs:  # ตรวจสอบว่าไฟล์ถูกดาวน์โหลดหรือยัง
                    try:
                        # เลื่อนหน้าจอไปยังลิงก์ PDF ก่อนคลิก
                        driver.execute_script("arguments[0].scrollIntoView(true);", pdf_link)
                        print(f"กำลังกดลิงก์ PDF ที่ {index + 1}: {pdf_url}")
                        pdf_link.click()  # คลิกลิงก์ PDF
                        downloaded_pdfs.add(pdf_url)  # บันทึกสถานะการดาวน์โหลด
                        time.sleep(5)  # รอให้ดาวน์โหลดเสร็จ
                    except Exception as e:
                        print(f"ไม่สามารถดาวน์โหลดไฟล์ PDF ที่ {index + 1}: {e}")
                else:
                    print(f"ลิงก์ PDF นี้ถูกดาวน์โหลดไปแล้ว: {pdf_url}")
                    message = f"พบการตอบกลับเอกสารจาก '{person}' ทำการดาวน์โหลดไฟล์ที่แนบมา\nจำนวนทั้งหมด {len(downloaded_pdfs)} ไฟล์"
                    send_line_notification(line_token, message)
        else:
            # ไม่แสดงข้อความใดๆ หากไม่มีลิงก์ PDF
            pass
            
finally:
    driver.quit()
    print("ปิด WebDriver เรียบร้อยแล้ว!")

