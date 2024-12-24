import os
import re
import pdfplumber
import fitz  # PyMuPDF
from pdf2image import convert_from_path
import pytesseract
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.firefox import GeckoDriverManager
import time
import json
from plyer import notification


# 📧 เข้าสู่ระบบ
email = "pichai_jo"
password = "CSautomation"

uploaded_files_record = "uploaded_files.json"

# โหลดรายการไฟล์ที่เคยอัปโหลด
def load_uploaded_files():
    if os.path.exists(uploaded_files_record):
        with open(uploaded_files_record, "r", encoding="utf-8") as f:
            return set(json.load(f))  # โหลดเป็นเซ็ต
    return set()

# บันทึกรายการไฟล์ที่อัปโหลด
def save_uploaded_file(file_path):
    uploaded_files = load_uploaded_files()  # โหลดไฟล์เดิม
    uploaded_files.add(file_path)  # เพิ่มไฟล์ใหม่
    with open(uploaded_files_record, "w", encoding="utf-8") as f:
        json.dump(list(uploaded_files), f, ensure_ascii=False, indent=4)  # บันทึกเป็น JSON
    print(f"บันทึกไฟล์ JSON เสร็จสิ้นที่: {os.path.abspath(uploaded_files_record)}")

def show_notification(title, message):
    notification.notify(
        title=title,
        message=message,
        app_name="PDF Uploader",
        timeout=10  # ระยะเวลาแสดงแจ้งเตือน (วินาที)
    )

# 📁 1️⃣ **ค้นหาไฟล์ทั้งหมดในโฟลเดอร์และโฟลเดอร์ย่อย**
def get_files_in_subfolders(download_folder):
    files = []
    for root, dirs, filenames in os.walk(download_folder):
        for filename in filenames:
            files.append(os.path.join(root, filename))
    return files

# 📘 2️⃣ **อ่านข้อมูลจากไฟล์ PDF**
def extract_pdf_data_with_plumber(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ''.join(page.extract_text(layout=True) or '' for page in pdf.pages)
        return text
    except Exception as e:
        print(f"Error with pdfplumber: {e}")
        return ""

def extract_pdf_data_with_fitz(pdf_path):
    try:
        text = ""
        with fitz.open(pdf_path) as pdf:
            for page in pdf:
                text += page.get_text("text")
        return text
    except Exception as e:
        print(f"Error with PyMuPDF: {e}")
        return ""

def extract_text_with_ocr(pdf_path):
    try:
        pages = convert_from_path(pdf_path)
        text = ""
        for page in pages:
            text += pytesseract.image_to_string(page, lang='tha')
        return text
    except Exception as e:
        print(f"Error with OCR: {e}")
        return ""

def extract_pdf_data(pdf_path):
    # ลองใช้ pdfplumber ก่อน
    text = extract_pdf_data_with_plumber(pdf_path)
    if not text.strip():
        print("ลองใช้ PyMuPDF")
        text = extract_pdf_data_with_fitz(pdf_path)
    if not text.strip():
        print("ลองใช้ OCR")
        text = extract_text_with_ocr(pdf_path)
    return text

def thai_to_arabic_number(text):
    thai_numbers = '๐๑๒๓๔๕๖๗๘๙'
    arabic_numbers = '0123456789'
    translation_table = str.maketrans(thai_numbers, arabic_numbers)
    return text.translate(translation_table)

# ฟังก์ชันสำหรับดึงวันที่ภาษาไทย
def extract_thai_date(text):
    match = re.search(r'\b(\d{1,2})\s*(มกราคม|กุมภาพันธ์|มีนาคม|เมษายน|พฤษภาคม|มิถุนายน|กรกฎาคม|สิงหาคม|กันยายน|ตุลาคม|พฤศจิกายน|ธันวาคม)\s*(\d{4})\b', text)
    if match:
        day, month, year = match.groups()
        thai_months = {
            "มกราคม": "01", "กุมภาพันธ์": "02", "มีนาคม": "03", "เมษายน": "04",
            "พฤษภาคม": "05", "มิถุนายน": "06", "กรกฎาคม": "07", "สิงหาคม": "08",
            "กันยายน": "09", "ตุลาคม": "10", "พฤศจิกายน": "11", "ธันวาคม": "12"
        }
        month_num = thai_months[month]
        return thai_to_arabic_number(f"{day.zfill(2)}/{month_num}/{int(year) - 543}")  # แปลงปีไทยเป็นคริสต์ศักราช
    return "ไม่พบวันที่"

# 📅 4️⃣ **ฟังก์ชันเลือกวันที่ใน Date Picker โดยอิงจากวันที่ในเอกสาร**
def select_date_in_datepicker(driver, day, month, year):
    # แปลงปีเป็นปี พ.ศ.
    thai_year = str(int(year) + 543)  # แปลงปี ค.ศ. เป็นปี พ.ศ.
    thai_months = {
        "01": "มกราคม", "02": "กุมภาพันธ์", "03": "มีนาคม", "04": "เมษายน",
        "05": "พฤษภาคม", "06": "มิถุนายน", "07": "กรกฎาคม", "08": "สิงหาคม",
        "09": "กันยายน", "10": "ตุลาคม", "11": "พฤศจิกายน", "12": "ธันวาคม"
    }
    thai_month = thai_months[month]

    # เปิด Date Picker
    date_input = driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_eDocumentContentCreate1_txtDocumentDate"]')
    date_input.click()
    time.sleep(1)

    # ตรวจสอบเดือนและปีปัจจุบันในปฏิทิน
    while True:
        current_title = driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_eDocumentContentCreate1_txtDocumentDatetCalendarExtender_title"]').text
        if thai_month in current_title and thai_year in current_title:
            break

        # เลื่อนปฏิทินไปยังเดือนและปีที่ต้องการ
        if int(year) < int(current_title[-4:]):  # หากปีต้องการน้อยกว่า ให้ย้อนกลับ
            prev_button = driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_eDocumentContentCreate1_txtDocumentDatetCalendarExtender_prevArrow"]')
            prev_button.click()
        else:  # หากปีมากกว่า หรือเดือนยังไม่ตรง ให้กดถัดไป
            next_button = driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_eDocumentContentCreate1_txtDocumentDatetCalendarExtender_nextArrow"]')
            next_button.click()
        time.sleep(1)

    # เลือกวันที่
    for row in range(6):
        for col in range(7):
            day_xpath = f'//*[@id="mainContentPlaceHolder_eDocumentContentCreate1_txtDocumentDatetCalendarExtender_day_{row}_{col}"]'
            try:
                day_element = driver.find_element(By.XPATH, day_xpath)
                if day_element.text == str(int(day)):
                    day_element.click()
                    return
            except Exception:
                pass
    print(f"ไม่พบวันที่: {day}/{month}/{year} ใน Date Picker")

# 📂 **ตั้งค่าโฟลเดอร์ที่เก็บไฟล์**
download_folder = r"C:\Users\mikot\Desktop\test"
files_in_folder = get_files_in_subfolders(download_folder)

# 📤 5️⃣ **จัดการไฟล์ PDF**
uploaded_files = load_uploaded_files()  # โหลดรายการไฟล์ที่เคยอัปโหลด

# กรองเฉพาะไฟล์ที่ยังไม่ได้อัปโหลด
pdf_files = [file for file in files_in_folder if file.endswith(".pdf") and file not in uploaded_files]

if pdf_files:
    latest_file = pdf_files[0]  # ดึงไฟล์ใหม่ล่าสุดที่ยังไม่ได้อัปโหลด
    pdf_data = extract_pdf_data(latest_file)

    # ดึงข้อมูลสำคัญ
    document_name = re.search(r'เรื่อง[:\s]*(.*)', pdf_data, re.MULTILINE).group(1).strip() if re.search(r'เรื่อง[:\s]*(.*)', pdf_data, re.MULTILINE) else "ไม่พบชื่อเรื่อง"
    document_code = re.search(r'(พส\.อ\s*\d*\s*/\s*\d+)', pdf_data)
    document_code = document_code.group(1).strip() if document_code else "ไม่พบเลขที่เอกสาร"
    impact_name = re.search(r'เรียน[:\s]*(.*)', pdf_data, re.MULTILINE).group(1).strip() if re.search(r'เรียน[:\s]*(.*)', pdf_data, re.MULTILINE) else "ไม่พบชื่อผู้รับ"
    document_date = extract_thai_date(pdf_data)

    print(f"ชื่อเอกสาร: {document_name}")
    print(f"เลขที่เอกสาร: {document_code}")
    print(f"เรียน: {impact_name}")
    print(f"วันที่เอกสาร: {document_date}")

# 📡 5️⃣ **เริ่มต้นกระบวนการ Selenium**
service = Service(GeckoDriverManager().install())
driver = webdriver.Firefox(service=service)

try:
    driver.get("https://e-doc.rmutto.ac.th/home.aspx")
    time.sleep(5)

    # คลิกที่ลิงก์แรก
    xpath_link = "/html/body/form/nav/div/div[2]/ul/li[2]/a"
    link = driver.find_element(By.XPATH, xpath_link)
    link.click()

    # รอให้หน้าใหม่โหลด
    time.sleep(2)

    driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_txtEmail"]').send_keys(email)
    driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_txtPassword"]').send_keys(password)
    driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_bttLogin"]').click()
    time.sleep(5)

    xpath_notification = '//*[@id="mainContentPlaceHolder_notificationSourceGridViewDocument_postBackUrlLinkButton_0"]'
    notification_link = driver.find_element(By.XPATH, xpath_notification)
    notification_link.click()
    time.sleep(15)

    windows = driver.window_handles
    for w in windows:
        driver.switch_to.window(w)
        if driver.current_url == "https://e-doc.rmutto.ac.th/documentInbox.aspx":
            break

    sent_new_document_button = driver.find_element(By.ID, "mainContentPlaceHolder_eDocumentDirectoryByPersonID_DropDrawList1_bttSentNewDocument")
    sent_new_document_button.click()
    time.sleep(5)

    document_name_input = driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_eDocumentContentCreate1_txtName"]')
    document_name_input.send_keys(document_name)

    document_code_input = driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_eDocumentContentCreate1_txtCodeRef"]')
    document_code_input.send_keys(document_code)

    if document_date != "ไม่พบวันที่":
        day, month, year = document_date.split("/")
        select_date_in_datepicker(driver, day, month, year)
    else:
        print("ไม่พบวันที่ในเอกสาร")

    driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_eDocumentContentCreate1_txtImpactName"]').send_keys(impact_name)

    save_button = driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_eDocumentContentCreate1_Button2"]')
    save_button.click()
    time.sleep(2)

    comment_link = driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_eDocumentContentCreate1_gvDataFrom_gvDataFrom_linkComment_0"]')
    comment_link.click()
    time.sleep(2)

    sent_person_button = driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_eDocumentContentCreate1_bttPersonSent"]')
    sent_person_button.click()
    time.sleep(2)

    sent_directory_add_button = driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_eDocumentDirectorySentTo1_bttDirectoryAdd"]')
    sent_directory_add_button.click()
    time.sleep(2)

   
    keyword_input = driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_eDocumentNotebook1_txtKeyword"]')
    keyword_input.send_keys("นางสาวต้องใจ แย้มผกา")

    additional_button = driver.find_element(By.XPATH, '/html/body/form/div[3]/div[2]/div[4]/div/div/div/div[2]/div/div/div/div[1]/div/div[2]/div/div/div/span/a/span')
    additional_button.click()
    time.sleep(2)

    checkbox_xpath = "//tr[.//a[text()='นางสาวต้องใจ แย้มผกา']]//input[@type='checkbox']"
    checkbox_element = driver.find_element(By.XPATH, checkbox_xpath)
    checkbox_element.click()
    time.sleep(5)

    confirm_button = driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_eDocumentNotebook1_bttConfirm"]')
    confirm_button.click()
    time.sleep(2)

    directory_confirm_button = driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_eDocumentDirectorySentTo1_bttConfirm"]')
    directory_confirm_button.click()
    time.sleep(5)

    uploaded_files = load_uploaded_files()  # โหลดรายการไฟล์ที่เคยอัปโหลด

    uploaded_count = 0


    for file_path in pdf_files:  # วนลูปไฟล์ทั้งหมดใน pdf_files
        # ข้ามไฟล์ที่เคยอัปโหลด
        if file_path in uploaded_files:
            print(f"ข้ามไฟล์ที่เคยอัปโหลด: {file_path}")
            continue

        try:
            # รอ input file และเคลียร์สถานะก่อนใส่ไฟล์ใหม่
            file_input = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//*[@id='mainContentPlaceHolder_eDocumentContentCreate1_AjaxFileUpload1_Html5InputFile']")
                )
            )
            driver.execute_script("arguments[0].value = '';", file_input)  # รีเซ็ต input file

            print(f"พาธไฟล์ที่จะอัปโหลด: {file_path}")
            file_input.send_keys(file_path)  # ใส่ไฟล์ลงใน input
            time.sleep(5)  # รอให้ระบบประมวลผลไฟล์ก่อนคลิกปุ่มอัปโหลด

            # ใช้ JavaScript เพื่อคลิกปุ่มอัปโหลดแทน
            upload_button = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//*[@id='mainContentPlaceHolder_eDocumentContentCreate1_AjaxFileUpload1_UploadOrCancelButton']")
                )
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", upload_button)
            driver.execute_script("arguments[0].click();", upload_button)  # คลิกผ่าน JavaScript
            print(f"อัปโหลดไฟล์สำเร็จ: {file_path}")

            # เพิ่มไฟล์ในรายการที่อัปโหลดแล้ว
            save_uploaded_file(file_path)
            uploaded_count += 1  # เพิ่มตัวนับไฟล์ที่ถูกอัปโหลดในรอบนี้
        except Exception as e:
            print(f"เกิดข้อผิดพลาดในการอัปโหลดไฟล์ {file_path}: {e}")

finally:
    time.sleep(5)
    driver.quit()

    if uploaded_count > 0:
        print(f"อัปโหลดไฟล์ใหม่จำนวน {uploaded_count} ไฟล์สำเร็จ!")
    else:
        print("ไม่มีไฟล์ใหม่สำหรับอัปโหลด")

    # แสดงการแจ้งเตือนเมื่ออัปโหลดเสร็จสิ้น เฉพาะไฟล์ที่อัปโหลดในครั้งนี้
    if uploaded_count > 0:
        show_notification("การอัปโหลดเสร็จสิ้น", f"อัปโหลดไฟล์จำนวน {uploaded_count} ไฟล์เรียบร้อยแล้ว")
    else:
        show_notification("การอัปโหลดเสร็จสิ้น", "ไม่มีไฟล์ใหม่ถูกอัปโหลด")
    print("การอัปโหลดเสร็จสิ้น")
    