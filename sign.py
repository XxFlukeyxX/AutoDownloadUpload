import os
import fitz  # PyMuPDF
from pdf2image import convert_from_path
from pytesseract import image_to_string
from PIL import Image
import io
 
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
 
# ฟังก์ชันเลือกไฟล์ลายเซ็นตามชื่อ
def get_signature_image_path(name, signature_folder):
    """
    ค้นหาไฟล์ลายเซ็นที่ตรงกับชื่อบุคคล
    """
    normalized_name = normalize_text(name)
    for file in os.listdir(signature_folder):
        if file.endswith(".png") or file.endswith(".jpg"):
            file_name = normalize_text(file.rsplit('.', 1)[0])  # เอาชื่อไฟล์ (ไม่รวม .png/.jpg)
            if file_name == normalized_name:
                return os.path.join(signature_folder, file)
    return None
 
# ฟังก์ชันปรับขนาดรูปภาพ
def resize_image(image_path, width, height):
    with Image.open(image_path) as img:
        resized_img = img.resize((width, height))
        buffer = io.BytesIO()
        resized_img.save(buffer, format="PNG")
        buffer.seek(0)
        return buffer
 
# ฟังก์ชันแปะลายเซ็น (ไม่ปรับ offset)
def add_signature_fixed(page, x0, y0, signature_image_path, fixed_width, fixed_height, applied_positions):
    for (ax0, ay0, ax1, ay1) in applied_positions:
        if abs(x0 - ax0) < 10 and abs(y0 - ay0) < 10:
            print(f"ตำแหน่ง ({x0}, {y0}) ใกล้กับลายเซ็นที่มีอยู่แล้ว ข้ามการแปะลายเซ็น")
            return
 
    adjusted_x1 = x0 + fixed_width
    adjusted_y1 = y0 - fixed_height
    rect = fitz.Rect(x0, y0 - fixed_height, adjusted_x1, y0)
 
    resized_image_stream = resize_image(signature_image_path, fixed_width, fixed_height)
    page.insert_image(rect, stream=resized_image_stream, overlay=True)
 
    print(f"แปะลายเซ็นที่ตำแหน่ง: {rect}")
    applied_positions.append((x0, y0, adjusted_x1, adjusted_y1))
 
# ตั้งค่าพารามิเตอร์
base_dir = "C:/Users/mikot/Desktop/testlogin/doc"
output_dir = "C:/Users/mikot/Desktop/testlogin/doc/Out"
os.makedirs(output_dir, exist_ok=True)  # สร้างโฟลเดอร์ Out ถ้ายังไม่มี
signature_folder = "C:/Users/mikot/Desktop/testlogin/signature"  # โฟลเดอร์เก็บไฟล์ลายเซ็น
 
# ค้นหาไฟล์ PDF ในโฟลเดอร์
pdf_files = [f for f in os.listdir(base_dir) if f.endswith('.pdf')]
if not pdf_files:
    raise FileNotFoundError(f"ไม่พบไฟล์ PDF ในโฟลเดอร์: {base_dir}")
 
target_keywords = ["นายธรรมจักร แสงเพ็ชร์", "นางสาวปวีณา จันทสุข", "นางสาวลัดดาวัลย์ เจริญสุข"]
 
# ใน loop ประมวลผลแต่ละไฟล์
for pdf_file in pdf_files:
    pdf_path = os.path.join(base_dir, pdf_file)
    print(f"กำลังเปิดไฟล์ PDF: {pdf_file}")
 
    doc = fitz.open(pdf_path)
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        applied_positions = []
        page_width, page_height = page.rect.width, page.rect.height
 
        for keyword in target_keywords:
            spans = page.search_for(keyword)
 
            for span_rect in spans:
                x0, y0, x1, y1 = span_rect
                print(f"เจอคำ '{keyword}' ในหน้า {page_num + 1} | ตำแหน่ง: ({x0}, {y0})")
 
                # ค้นหาไฟล์ลายเซ็นที่ตรงกับชื่อ
                signature_path = get_signature_image_path(keyword, signature_folder)
                if not signature_path:
                    print(f"ไม่พบไฟล์ลายเซ็นสำหรับ '{keyword}' ข้ามการเซ็น")
                    continue
 
                # แปะลายเซ็นโดยตรงโดยไม่ปรับ offset
                add_signature_fixed(page, x0, y0, signature_path, 150, 30, applied_positions)
 
    # บันทึก PDF ใหม่ในชื่อไฟล์เดิม
    output_path = os.path.join(output_dir, pdf_file)
    doc.save(output_path)
    print(f"ลายเซ็นถูกแปะทับและบันทึกในไฟล์ใหม่เรียบร้อยแล้ว: {output_path}")