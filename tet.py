import os
import re
import pdfplumber
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.firefox import GeckoDriverManager
import time
from plyer import notification

# เข้าสู่ระบบ
email = "pichai_jo"
password = "CSautomation"

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
            if filename.endswith(".pdf"):
                files.append(os.path.join(root, filename))
    return files

# 📘 2️⃣ **อ่านข้อมูลจากไฟล์ PDF**
def extract_pdf_data(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = ''.join(page.extract_text() or '' for page in pdf.pages)
    return text

# 📅 3️⃣ **ดึงวันที่จาก PDF และแปลงเลขไทยเป็น ค.ศ.**
def extract_thai_date(pdf_data):
    thai_months = {
        "มกราคม": "01", "กุมภาพันธ์": "02", "มีนาคม": "03", "เมษายน": "04",
        "พฤษภาคม": "05", "มิถุนายน": "06", "กรกฎาคม": "07", "สิงหาคม": "08",
        "กันยายน": "09", "ตุลาคม": "10", "พฤศจิกายน": "11", "ธันวาคม": "12"
    }
    try:
        date_match = re.search(r'(\d{1,2})\s(มกราคม|กุมภาพันธ์|มีนาคม|เมษายน|พฤษภาคม|มิถุนายน|กรกฎาคม|สิงหาคม|กันยายน|ตุลาคม|พฤศจิกายน|ธันวาคม)\s(๒๕\d{2})', pdf_data)
        if date_match:
            day, month_thai, year_thai = date_match.groups()
            day = day.translate(str.maketrans("๐๑๒๓๔๕๖๗๘๙", "0123456789"))
            year = str(int(year_thai.translate(str.maketrans("๐๑๒๓๔๕๖๗๘๙", "0123456789"))) - 543)
            month = thai_months[month_thai]
            return f"{day}/{month}/{year}"
        return "ไม่พบวันที่"
    except Exception as e:
        print(f"Error: {e}")
        return "ไม่พบวันที่"

# 📅 4️⃣ **ฟังก์ชันเลือกวันที่ใน Date Picker โดยอิงจากวันที่ในเอกสาร**
def select_date_in_datepicker(driver, day, month, year):
    thai_year = str(int(year) + 543)  # แปลงปี ค.ศ. เป็นปี พ.ศ.
    thai_months = {
        "01": "มกราคม", "02": "กุมภาพันธ์", "03": "มีนาคม", "04": "เมษายน",
        "05": "พฤษภาคม", "06": "มิถุนายน", "07": "กรกฎาคม", "08": "สิงหาคม",
        "09": "กันยายน", "10": "ตุลาคม", "11": "พฤศจิกายน", "12": "ธันวาคม"
    }
    thai_month = thai_months[month]

    date_input = driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_eDocumentContentCreate1_txtDocumentDate"]')
    date_input.click()
    time.sleep(1)

    while True:
        current_title = driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_eDocumentContentCreate1_txtDocumentDatetCalendarExtender_title"]').text
        if thai_month in current_title and thai_year in current_title:
            break
        prev_button = driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_eDocumentContentCreate1_txtDocumentDatetCalendarExtender_prevArrow"]')
        prev_button.click()
        time.sleep(1)

    for row in range(6):
        for col in range(7):
            try:
                day_element = driver.find_element(By.XPATH, f'//*[@id="mainContentPlaceHolder_eDocumentContentCreate1_txtDocumentDatetCalendarExtender_day_{row}_{col}"]')
                if day_element.text == str(int(day)):
                    day_element.click()
                    return
            except Exception:
                pass

# 📂 **ตั้งค่าโฟลเดอร์ที่เก็บไฟล์**
download_folder = r"C:\\Users\\mikot\\Desktop\\test"
files_in_folder = get_files_in_subfolders(download_folder)

def save_uploaded_files(uploaded_files):
    with open('uploaded_files.json', 'w', encoding='utf-8') as f:
        json.dump(uploaded_files, f)

def load_uploaded_files():
    if os.path.exists('uploaded_files.json'):
        with open('uploaded_files.json', 'r', encoding='utf-8') as f:
            return json.load(f)
    return []

uploaded_files = load_uploaded_files()
pdf_files = [file for file in files_in_folder if file.endswith(".pdf") and file not in uploaded_files]

if pdf_files:
    latest_file = pdf_files[0]
    pdf_data = extract_pdf_data(latest_file)
    document_name_match = re.search(r'เรื่อง[:\s]*(.*)', pdf_data, re.MULTILINE)
    document_name = document_name_match.group(1).strip() if document_name_match else "ไม่พบชื่อเรื่อง"
    document_code_match = re.search(r'ที\s*[^\S\r\n]*[ที่:]*\s*(.*?/\d+)', pdf_data, re.MULTILINE)
    document_code = document_code_match.group(1).strip() if document_code_match else "ไม่พบเลขที่เอกสาร"
    impact_name_match = re.search(r'เรียน[:\s]*(.*)', pdf_data, re.MULTILINE)
    impact_name = impact_name_match.group(1).strip() if impact_name_match else "ไม่พบชื่อผู้รับ"
    document_date = extract_thai_date(pdf_data)

service = Service(GeckoDriverManager().install())
driver = webdriver.Firefox(service=service)

try:
    driver.get("https://e-doc.rmutto.ac.th/home.aspx")
    time.sleep(5)
    driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_txtEmail"]').send_keys(email)
    driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_txtPassword"]').send_keys(password)
    driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_bttLogin"]').click()
    time.sleep(5)

    for file_path in pdf_files:
        file_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='mainContentPlaceHolder_eDocumentContentCreate1_AjaxFileUpload1_Html5InputFile']"))
        )
        file_input.send_keys(file_path)
        uploaded_files.append(file_path)
        save_uploaded_files(uploaded_files)

except Exception as e:
    print(f"Error: {e}")

finally:
    driver.quit()
    show_notification("การอัปโหลดเสร็จสิ้น", f"อัปโหลดไฟล์จำนวน {len(uploaded_files)} ไฟล์เรียบร้อยแล้ว")
