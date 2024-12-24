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

# ตั้งค่า Firefox Options สำหรับดาวน์โหลดไฟล์ในโฟลเดอร์ Downloads
firefox_options = Options()
firefox_options.set_preference("browser.download.folderList", 1)  # ใช้โฟลเดอร์ Downloads ของระบบ
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

        # ตรวจสอบว่าแท็บถูกเลือกสำเร็จ
        try:
            selected_tab = driver.find_element(By.XPATH, '/html/body/form/div[3]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div/div[1]/ul/li[1]')
            if 'active' in selected_tab.get_attribute("class"):
                print("แท็บ 'เอกสารที่ยังไม่ได้เปิดอ่าน' ถูกเลือกสำเร็จ")
            else:
                print("แท็บ 'เอกสารที่ยังไม่ได้เปิดอ่าน' ไม่ถูกเลือก")
        except Exception as e:
            print(f"ไม่สามารถตรวจสอบสถานะของแท็บได้: {e}")

        # ค้นหาคำว่า "นางสาวต้องใจ แย้มผกา" ในโซน "เอกสารที่ยังไม่ได้เปิดอ่าน"
        found = False
        while True:
            print("กำลังค้นหาในหน้านี้...")
            elements = driver.find_elements(By.XPATH, "//div[@id='mainContentPlaceHolder_eDocumentContentByDocumentDirectoryID1_gvData']//span[contains(text(), 'นางสาวต้องใจ แย้มผกา')]")
            if elements:
                print("พบ 'นางสาวต้องใจ แย้มผกา' ในหน้านี้!")
                found = True
                for element in elements:
                    print("องค์ประกอบที่พบ:", element.get_attribute("outerHTML"))
                break

            # คลิกปุ่ม "ถัดไป" ถ้ามี
            try:
                next_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="mainContentPlaceHolder_eDocumentContentByDocumentDirectoryID1_gvData_postBackPager_ctl02"]'))
                )
                next_button.click()
                print("คลิกปุ่มถัดไปสำเร็จ รอให้หน้าโหลดใหม่...")
                WebDriverWait(driver, 10).until(EC.staleness_of(next_button))  # รอให้ปุ่มถัดไปหายไปจาก DOM (หน้าโหลดใหม่)
                time.sleep(2)
            except Exception:
                print("ไม่มีปุ่มถัดไปแล้ว!")
                break

        # แจ้งเตือนผู้ใช้ (พบหรือไม่พบคำที่ต้องการ)
        if found:
            notification.notify(
                title='คำเตือนจากระบบ',
                message="⚠️ พบคำว่า 'นางสาวต้องใจ แย้มผกา' ในเอกสารที่ยังไม่ได้เปิดอ่าน!",
                timeout=10
            )
        else:
            notification.notify(
                title='ผลการตรวจสอบ',
                message="❌ ไม่พบคำว่า 'นางสาวต้องใจ แย้มผกา' ในเอกสารที่ยังไม่ได้เปิดอ่าน",
                timeout=10
            )
        return found

    except Exception as e:
        print(f"เกิดข้อผิดพลาด: {e}")
        notification.notify(
            title='ข้อผิดพลาด',
            message=f"⚠️ เกิดข้อผิดพลาดระหว่างการทำงาน: {e}",
            timeout=10
        )
        return False
    finally:
        driver.quit()

# เรียกใช้ฟังก์ชันตรวจสอบเพียงครั้งเดียว
check_for_user()
