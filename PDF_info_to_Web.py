import os
import re
import pdfplumber
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.firefox import GeckoDriverManager
import time

# 🛠️ ข้อมูลสำหรับล็อกอิน
email = "pichai_jo"
password = "CSautomation"

# 📁 1️⃣ **ค้นหาไฟล์ล่าสุดในโฟลเดอร์**
def get_latest_file(download_folder):
    files = [os.path.join(download_folder, f) for f in os.listdir(download_folder) if os.path.isfile(os.path.join(download_folder, f))]
    if not files:
        raise Exception("ไม่พบไฟล์ใด ๆ ในโฟลเดอร์ดาวน์โหลด")
    latest_file = max(files, key=os.path.getctime)
    latest_file = latest_file.replace("\\", "/")
    return latest_file

# 📘 2️⃣ **อ่านข้อมูลจากไฟล์ PDF**
def extract_pdf_data(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = ''
        for page in pdf.pages:
            text += page.extract_text()
    text = text.replace("", "้").replace("אָ", "า")
    return text

# 📂 **ตั้งค่าโฟลเดอร์ที่เก็บไฟล์**
download_folder = os.path.expanduser("~/Downloads")
latest_file = get_latest_file(download_folder)
print(f"พาธไฟล์ล่าสุด: {latest_file}")

if latest_file.endswith(".pdf"):
    pdf_data = extract_pdf_data(latest_file)
    try:
        document_name = re.search(r'เรื่อง[:\s]*(.*?)(\n|$)', pdf_data).group(1).strip()
    except AttributeError:
        document_name = "ไม่พบชื่อเรื่อง"

    try:
        document_code = re.search(r'ที่[:\s]*(\S*?/\d+)', pdf_data).group(1).strip()
    except AttributeError:
        document_code = "ไม่พบเลขที่เอกสาร"

    try:
        impact_name = re.search(r'เรียน[:\s]*(.*?)(\n|$)', pdf_data).group(1).strip()
    except AttributeError:
        impact_name = "ไม่พบชื่อผู้รับ"
else:
    document_name = "ไม่พบชื่อเรื่อง"
    document_code = "ไม่พบเลขที่เอกสาร"
    impact_name = "ไม่พบชื่อผู้รับ"

print(f"ชื่อเอกสาร: {document_name}")
print(f"เลขที่เอกสาร: {document_code}")
print(f"เรียน: {impact_name}")

# 📡 3️⃣ **เริ่มต้นกระบวนการ Selenium**
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

    email_input = driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_txtEmail"]')
    email_input.send_keys(email)
    password_input = driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_txtPassword"]')
    password_input.send_keys(password)
    login_button = driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_bttLogin"]')
    login_button.click()
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
    time.sleep(2)

    document_name_input = driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_eDocumentContentCreate1_txtName"]')
    document_name_input.send_keys(document_name)

    document_code_input = driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_eDocumentContentCreate1_txtCodeRef"]')
    document_code_input.send_keys(document_code)

    impact_name_input = driver.find_element(By.XPATH, '//*[@id="mainContentPlaceHolder_eDocumentContentCreate1_txtImpactName"]')
    impact_name_input.send_keys(impact_name)

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
    
    # 🔥 **อัปโหลดไฟล์**
    try:
        file_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='mainContentPlaceHolder_eDocumentContentCreate1_AjaxFileUpload1_Html5InputFile']"))
        )
        file_path = os.path.abspath(latest_file)
        print(f"พาธไฟล์ที่จะอัปโหลด: {file_path}")
        
        driver.execute_script("arguments[0].scrollIntoView(true);", file_input)  
        
        file_input.send_keys(file_path)  
        print(f"อัปโหลดไฟล์สำเร็จ: {latest_file}")
        
        upload_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='mainContentPlaceHolder_eDocumentContentCreate1_AjaxFileUpload1_UploadOrCancelButton']"))
        )
        upload_button.click()
        print("คลิกปุ่มอัปโหลดสำเร็จ")
        
        save_and_send_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='mainContentPlaceHolder_eDocumentContentCreate1_bttSaveAndSent']"))
        )
        save_and_send_button.click()
        print("คลิกปุ่ม Save and Sent สำเร็จ")
        
    except Exception as e:
        print(f"เกิดข้อผิดพลาดในการอัปโหลดไฟล์หรือคลิกปุ่ม Save and Sent: {e}")

except Exception as e:
    print(f"เกิดข้อผิดพลาด: {e}")

finally:
    time.sleep(10)
    driver.quit()
