from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.firefox import GeckoDriverManager
from plyer import notification
import time

# กำหนดข้อมูลผู้ใช้
email = "pichai_jo"
password = "CSautomation"

# ตั้งค่า Firefox Options
firefox_options = Options()
firefox_options.set_preference("browser.download.folderList", 1)  # ใช้โฟลเดอร์ Downloads
firefox_options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf")  # ไม่ถามก่อนดาวน์โหลดไฟล์ PDF
firefox_options.set_preference("pdfjs.disabled", True)  # ปิดการแสดง PDF ในเบราว์เซอร์

# สร้าง WebDriver
service = Service(GeckoDriverManager().install())
driver = webdriver.Firefox(service=service, options=firefox_options)

def check_for_user():
    try:
        # เปิดเว็บไซต์
        driver.get("https://e-doc.rmutto.ac.th/home.aspx")
        print("เปิดเว็บไซต์สำเร็จ!")
        time.sleep(5)

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
            EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[3]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div/div[1]/ul/li[1]/a'))
        )
        tab_unread.click()
        print("เปลี่ยนไปยังแท็บ 'เอกสารที่ยังไม่ได้เปิดอ่าน'")
        time.sleep(5)

        # ค้นหาคำว่า "นางสาวต้องใจ แย้มผกา" เฉพาะในองค์ประกอบที่กำหนด
        target_xpath = '/html/body/form/div[3]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div/div[2]/div[1]/div/div/div/div'
        root_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, target_xpath))
        )
        print("พบองค์ประกอบที่ต้องการค้นหา!")

        # ค้นหาข้อความในองค์ประกอบ
        elements = root_element.find_elements(By.XPATH, ".//*[contains(text(), 'นางสาวชนิสรา อ่ำสอาด')]")

        # ตรวจสอบผลลัพธ์
        if elements:
            print("พบ 'นางสาวชนิสรา อ่ำสอาด' ในองค์ประกอบที่ระบุ!")
            for element in elements:
                print("ข้อความที่พบ:", element.text)
            notification.notify(
                title='คำเตือนจากระบบ',
                message="⚠️ พบคำว่า 'นางสาวชนิสรา อ่ำสอาด' ในองค์ประกอบที่ระบุ!",
                timeout=10
            )
        else:
            print("ไม่พบ 'นางสาวชนิสรา อ่ำสอาด' ในองค์ประกอบที่ระบุ")
            notification.notify(
                title='ผลการตรวจสอบ',
                message="❌ ไม่พบคำว่า 'นางสาวชนิสรา อ่ำสอาด' ในองค์ประกอบที่ระบุ",
                timeout=10
            )
    except Exception as e:
        print(f"เกิดข้อผิดพลาด: {e}")
        notification.notify(
            title='ข้อผิดพลาด',
            message=f"⚠️ เกิดข้อผิดพลาดระหว่างการทำงาน: {e}",
            timeout=10
        )
    finally:
        # ปิด WebDriver
        driver.quit()

# เรียกใช้ฟังก์ชันตรวจสอบ
check_for_user()
