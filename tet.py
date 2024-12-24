from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from webdriver_manager.firefox import GeckoDriverManager
import time
from plyer import notification  # สำหรับการแจ้งเตือนในเครื่อง

# 🛠️ กำหนดข้อมูลผู้ใช้
email = "pichai_jo"
password = "CSautomation"

# 🛠️ ตั้งค่า Firefox Options
firefox_options = Options()
firefox_options.set_preference("browser.download.folderList", 1)  # ใช้โฟลเดอร์ดาวน์โหลด
firefox_options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf")  # ไม่ถามก่อนดาวน์โหลด PDF
firefox_options.set_preference("pdfjs.disabled", True)  # ปิดตัวแสดง PDF ในเบราว์เซอร์

# 🛠️ สร้าง WebDriver
service = Service(GeckoDriverManager().install())
driver = webdriver.Firefox(service=service, options=firefox_options)

# 🛠️ เก็บสถานะของคำที่เคยแจ้งเตือน
notified_set = set()  # ชุดสำหรับเก็บข้อความที่เคยแจ้งเตือนไปแล้ว

def check_for_user():
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

        # 4️⃣ ค้นหาข้อมูลในแต่ละแถว
        rows = driver.find_elements("xpath", "//div[contains(@class, 'row-class')]")  # ปรับ XPath ตามโครงสร้างหน้าเว็บ
        for row in rows:
            try:
                # ค้นหาชื่อและเวลาจากแต่ละแถว
                name_element = row.find_element("xpath", ".//span[contains(text(), 'นางสาวต้องใจ แย้มผกา')]")
                time_element = row.find_element("xpath", ".//span[contains(@class, 'time-class')]")  # ปรับ class ตามหน้าเว็บจริง
                if name_element and time_element:
                    name_text = name_element.text.strip()
                    time_text = time_element.text.strip()

                    # ตรวจสอบว่าชื่อนี้เคยแจ้งเตือนหรือยัง
                    if name_text not in notified_set:
                        # แจ้งเตือนเฉพาะข้อความใหม่
                        notification.notify(
                            title="แจ้งเตือนข้อมูลใหม่",
                            message=f"พบ '{name_text}' ส่งเมื่อ {time_text}",
                            timeout=10  # ระยะเวลาที่แจ้งเตือนจะแสดง
                        )
                        notified_set.add(name_text)  # บันทึกชื่อใน notified_set เพื่อหลีกเลี่ยงการแจ้งเตือนซ้ำ
                        print(f"แจ้งเตือน: {name_text} | {time_text}")
                    else:
                        print(f"ข้อมูลซ้ำ: {name_text} | {time_text}")
            except Exception as e:
                print(f"ข้ามข้อมูลที่ไม่สมบูรณ์: {e}")
        return True

    except Exception as e:
        print(f"เกิดข้อผิดพลาด: {e}")
        return False
    finally:
        driver.quit()  # ปิด WebDriver หลังจากเสร็จสิ้น