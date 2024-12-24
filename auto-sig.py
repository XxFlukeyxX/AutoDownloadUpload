import os
import fitz  # PyMuPDF
from pdf2image import convert_from_path
from pytesseract import image_to_string

# ฟังก์ชัน Normalize ข้อความ
def normalize_text(text):
    return text.strip().lower().replace(" ", "").replace("", "").replace("\u200b", "")

# ฟังก์ชันแปะลายเซ็นที่ตำแหน่งข้อความ
def add_signature(page, x0, y0, signature_image, fixed_width, fixed_height):
    new_x1 = x0 + fixed_width  # ความกว้างลายเซ็น
    new_y1 = y0 + fixed_height  # ความสูงลายเซ็น
    rect = fitz.Rect(x0, y0, new_x1, new_y1)  # สร้างพื้นที่ลายเซ็น
    print(f"แปะลายเซ็นที่ตำแหน่ง: {rect}")
    page.insert_image(rect, filename=signature_image, overlay=True)

# ฟังก์ชันดึงข้อความจาก PDF
def extract_text(page, pdf_path, page_num):
    text = page.get_text("text")
    if not text.strip():  # หากไม่มีข้อความใน text layer ใช้ OCR
        print(f"หน้า {page_num + 1} ไม่มีข้อความ จะใช้ OCR แทน")
        images = convert_from_path(pdf_path, first_page=page_num + 1, last_page=page_num + 1)
        ocr_text = ""
        for image in images:
            ocr_text += image_to_string(image, lang="tha") + "\n"
        return ocr_text
    return text

# ตั้งค่าพารามิเตอร์
base_dir = "C:/Users/mikot/Desktop/testlogin/doc"
signature_image = "C:/Users/mikot/Desktop/testlogin/sign/signature2.png"
fixed_width = 200  # ความกว้างลายเซ็น
fixed_height = 50  # ความสูงลายเซ็น
target_keywords = ["ขอแสดงความนับถือ", "ลงชื่อ", "ลายเซ็นต์", "เจ้าหน้าที่บริหารงานทั่วไป"]

# ค้นหาไฟล์ PDF ในโฟลเดอร์
pdf_files = [f for f in os.listdir(base_dir) if f.endswith('.pdf')]
if not pdf_files:
    raise FileNotFoundError(f"ไม่พบไฟล์ PDF ในโฟลเดอร์: {base_dir}")

# สร้างโฟลเดอร์ 'Out' ถ้ายังไม่มี
output_dir = os.path.join(base_dir, "Out")
os.makedirs(output_dir, exist_ok=True)

# เลือกไฟล์ PDF แรกในโฟลเดอร์
pdf_path = os.path.join(base_dir, pdf_files[0])
print(f"กำลังเปิดไฟล์ PDF: {pdf_files[0]}")

# เปิดไฟล์ PDF
doc = fitz.open(pdf_path)

# อ่านข้อมูลในแต่ละหน้า
for page_num in range(len(doc)):
    page = doc.load_page(page_num)
    text = extract_text(page, pdf_path, page_num)
    print(f"ข้อความในหน้า {page_num + 1}: {text.strip()}")

    # ค้นหาคำที่เกี่ยวข้องกับลายเซ็น
    for keyword in target_keywords:
        keyword_normalized = normalize_text(keyword)
        blocks = page.get_text("dict")["blocks"]
        for block_num, block in enumerate(blocks):
            if block['type'] == 0:  # ตรวจสอบเฉพาะข้อความ
                for line_num, line in enumerate(block["lines"]):
                    for span in line["spans"]:
                        span_text = normalize_text(span["text"])
                        words = span["text"].split()  # แยกข้อความใน span เป็นคำ ๆ
                        x0, y0, x1, y1 = span["bbox"]  # พิกัดเริ่มต้นของ span
                        current_x = x0  # ตำแหน่ง x เริ่มต้น

                        for word in words:
                            normalized_word = normalize_text(word)
                            if keyword_normalized in normalized_word:
                                print(f"เจอคำ '{keyword}' ใน Block {block_num}, Line {line_num} | ตำแหน่ง: ({current_x}, {y0})")
                                add_signature(page, current_x, y0, signature_image, fixed_width, fixed_height)
                            # คำนวณตำแหน่งถัดไป
                            current_x += (x1 - x0) / len(words)  # หาพิกัดถัดไปตามความยาวของ span

# บันทึก PDF ใหม่ในโฟลเดอร์ Out
output_path = os.path.join(output_dir, f"{pdf_files[0].replace('.pdf', '-signed.pdf')}")
doc.save(output_path)
print(f"ลายเซ็นถูกแปะทับและบันทึกในไฟล์ใหม่เรียบร้อยแล้ว: {output_path}")
