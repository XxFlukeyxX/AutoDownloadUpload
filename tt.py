import os
import fitz  # PyMuPDF
from pdf2image import convert_from_path
from pytesseract import image_to_string

# ฟังก์ชัน Normalize ข้อความ
def normalize_text(text):
    return (
        text.strip()
        .lower()
        .replace(" ", "")
        .replace("\ufeff", "")
        .replace("\u200b", "")
        .replace("\u0e4e", "")  # ลบตัวสะกดพิเศษ เช่น "\u0e4e"
    )

# ฟังก์ชันแปะลายเซ็นที่ตำแหน่งข้อความ
def add_signature(page, x0, y0, signature_image, fixed_width, fixed_height, applied_positions, threshold=10):
    for (ax0, ay0, ax1, ay1) in applied_positions:
        if abs(x0 - ax0) < threshold and abs(y0 - ay0) < threshold:
            print(f"ตำแหน่ง ({x0}, {y0}) ใกล้กับลายเซ็นที่มีอยู่แล้ว ข้ามการแปะลายเซ็น")
            return  # ข้ามการแปะลายเซ็นถ้าตำแหน่งใกล้กับที่มีอยู่แล้ว
    new_x1 = x0 + fixed_width  # ความกว้างลายเซ็น
    new_y1 = y0 + fixed_height  # ความสูงลายเซ็น
    rect = fitz.Rect(x0, y0, new_x1, new_y1)  # สร้างพื้นที่ลายเซ็น
    print(f"แปะลายเซ็นที่ตำแหน่ง: {rect}")
    page.insert_image(rect, filename=signature_image, overlay=True)
    applied_positions.append((x0, y0, new_x1, new_y1))  # บันทึกตำแหน่งที่แปะแล้ว

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
base_dir = "C:/Users/ACER/Desktop/FORBAT"

# ค้นหาไฟล์ PDF ในโฟลเดอร์
pdf_files = [f for f in os.listdir(base_dir) if f.endswith('.pdf')]
if not pdf_files:
    raise FileNotFoundError(f"ไม่พบไฟล์ PDF ในโฟลเดอร์: {base_dir}")

pdf_path = os.path.join(base_dir, pdf_files[0])  # เลือกไฟล์ PDF แรกในโฟลเดอร์
print(f"กำลังเปิดไฟล์ PDF: {pdf_files[0]}")

signature_image = "C:/Users/ACER/Desktop/Auto-Signature/signature/signature.png"
fixed_width = 150  # ความกว้างลายเซ็น
fixed_height = 50  # ความสูงลายเซ็น
target_keywords = ["นายธรรมจักร แสงเพ็ชร์", "นางสาวปวีณา จันทสุข"]

# เปิดไฟล์ PDF
doc = fitz.open(pdf_path)

# อ่านข้อมูลในแต่ละหน้า
for page_num in range(len(doc)):
    page = doc.load_page(page_num)
    applied_positions = []  # รายการตำแหน่งที่แปะลายเซ็นแล้ว

    # ค้นหาคำที่เกี่ยวข้องกับลายเซ็น
    for keyword in target_keywords:
        keyword_normalized = normalize_text(keyword)
        spans = page.search_for(keyword)

        for span_rect in spans:
            x0, y0, x1, y1 = span_rect  # พิกัดของคำที่เจอ
            print(f"เจอคำ '{keyword}' ในหน้า {page_num + 1} | ตำแหน่ง: ({x0}, {y0})")
            add_signature(page, x0, y0, signature_image, fixed_width, fixed_height, applied_positions)

# บันทึก PDF ใหม่
output_path = "C:/Users/ACER/Desktop/FORBAT/Out/ตัวอย่างรายงานขอซื้อขอจ้าง-signed.pdf"
doc.save(output_path)
print(f"ลายเซ็นถูกแปะทับและบันทึกในไฟล์ใหม่เรียบร้อยแล้ว: {output_path}")
