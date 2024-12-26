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

def configure_for_file(pdf_path):
    """
    กำหนดค่าการตั้งค่าตามโครงสร้างของไฟล์ PDF
    """
    import fitz  # PyMuPDF

    doc = fitz.open(pdf_path)
    sample_page = doc[0]
    text_blocks = sample_page.get_text("dict")["blocks"]
    page_width = sample_page.rect.width
    page_height = sample_page.rect.height

    # เงื่อนไขสำหรับเอกสารที่มีโครงสร้างซับซ้อน
    if len(text_blocks) > 30 or (page_width / page_height < 0.75):  # ตัวอย่าง: บล็อกข้อความเยอะหรืออัตราส่วนหน้าไม่ปกติ
        return {
            "offset_x": 5,  # ค่าปรับแต่งเฉพาะ
            "offset_y": -4,
            "fixed_width": 180,
            "fixed_height": 65,
            "page_width": page_width,
            "page_height": page_height,
            "text_blocks": len(text_blocks)
        }

    # ค่าเริ่มต้นสำหรับไฟล์ทั่วไป
    if len(text_blocks) > 15:  # ตัวอย่าง: ไฟล์ที่มีโครงสร้างข้อความแน่น
        return {"offset_x": -5, "offset_y": -5, "fixed_width": 170, "fixed_height": 60, "page_width": page_width, "page_height": page_height, "text_blocks": len(text_blocks)}
    elif len(text_blocks) > 10:  # ไฟล์ที่มีโครงสร้างซับซ้อน
        return {"offset_x": 10, "offset_y": -20, "fixed_width": 150, "fixed_height": 50, "page_width": page_width, "page_height": page_height, "text_blocks": len(text_blocks)}
    elif page_width > 600:  # ไฟล์ที่กว้าง
        return {"offset_x": -5, "offset_y": -5, "fixed_width": 160, "fixed_height": 60, "page_width": page_width, "page_height": page_height, "text_blocks": len(text_blocks)}
    else:  # ค่าเริ่มต้นสำหรับไฟล์ทั่วไป
        return {"offset_x": -5, "offset_y": -20, "fixed_width": 150, "fixed_height": 50, "page_width": page_width, "page_height": page_height, "text_blocks": len(text_blocks)}



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

# ฟังก์ชันแปะลายเซ็น
def add_signature(page, x0, y0, signature_image_path, fixed_width, fixed_height, applied_positions, offset_y=-20, offset_x=0, threshold=10):
    for (ax0, ay0, ax1, ay1) in applied_positions:
        if abs(x0 - ax0) < threshold and abs(y0 - ay0) < threshold:
            print(f"ตำแหน่ง ({x0}, {y0}) ใกล้กับลายเซ็นที่มีอยู่แล้ว ข้ามการแปะลายเซ็น")
            return  # ข้ามการแปะลายเซ็นถ้าตำแหน่งใกล้กับที่มีอยู่แล้ว

    # ปรับตำแหน่ง
    adjusted_x0 = x0 + offset_x
    adjusted_y0 = y0 - fixed_height + offset_y  # เลื่อนตำแหน่งขึ้นด้านบนข้อความ
    adjusted_x1 = adjusted_x0 + fixed_width
    adjusted_y1 = adjusted_y0 + fixed_height
    rect = fitz.Rect(adjusted_x0, adjusted_y0, adjusted_x1, adjusted_y1)  # สร้างพื้นที่ลายเซ็น

    # เปิดรูปภาพและปรับขนาดในหน่วยความจำ
    resized_image_stream = resize_image(signature_image_path, 150, 30)
    page.insert_image(rect, stream=resized_image_stream, overlay=True)

    print(f"แปะลายเซ็นที่ตำแหน่ง: {rect}")
    applied_positions.append((adjusted_x0, adjusted_y0, adjusted_x1, adjusted_y1))  # บันทึกตำแหน่งที่แปะแล้ว

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

# ฟังก์ชันปรับ offset อัตโนมัติ
def auto_adjust_offsets(x0, y0, page_width, page_height):
    """
    ปรับค่า offset อัตโนมัติ โดยพิจารณาตำแหน่งข้อความและขนาดหน้า PDF
    """
    margin_x = -25  # ระยะขอบแนวนอน
    margin_y = -90  # ระยะขอบแนวตั้ง

    # ตรวจสอบขอบด้านขวาและล่าง
    if x0 + fixed_width + margin_x > page_width:
        offset_x = -fixed_width - margin_x
    else:
        offset_x = margin_x

    if y0 - fixed_height - margin_y < 0:
        offset_y = fixed_height + margin_y  # ถ้าขึ้นบนไม่ได้ ให้เลื่อนไปด้านล่าง
    else:
        offset_y = -fixed_height - margin_y  # ค่าเริ่มต้นคือเลื่อนขึ้นด้านบน

    return offset_x, offset_y

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

    # รับค่าการตั้งค่าเฉพาะสำหรับไฟล์นี้
    file_config = configure_for_file(pdf_path)
    offset_x_default = file_config["offset_x"]
    offset_y_default = file_config["offset_y"]
    fixed_width = file_config["fixed_width"]
    fixed_height = file_config["fixed_height"]

    doc = fitz.open(pdf_path)
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        applied_positions = []

        page_width, page_height = page.rect.width, page.rect.height

        for keyword in target_keywords:
            keyword_normalized = normalize_text(keyword)
            spans = page.search_for(keyword)

            for span_rect in spans:
                x0, y0, x1, y1 = span_rect
                print(f"เจอคำ '{keyword}' ในหน้า {page_num + 1} | ตำแหน่ง: ({x0}, {y0})")

                # ค้นหาไฟล์ลายเซ็นที่ตรงกับชื่อ
                signature_path = get_signature_image_path(keyword, signature_folder)
                if not signature_path:
                    print(f"ไม่พบไฟล์ลายเซ็นสำหรับ '{keyword}' ข้ามการเซ็น")
                    continue

                # ปรับ offset อัตโนมัติ
                offset_x, offset_y = auto_adjust_offsets(x0, y0, page_width, page_height)
                offset_x += offset_x_default  # เพิ่ม offset เฉพาะไฟล์
                offset_y += offset_y_default

                add_signature(page, x0, y0, signature_path, fixed_width, fixed_height, applied_positions, offset_y=offset_y, offset_x=offset_x)

    # บันทึก PDF ใหม่ในชื่อไฟล์เดิม
    output_path = os.path.join(output_dir, pdf_file)
    doc.save(output_path)
    print(f"ลายเซ็นถูกแปะทับและบันทึกในไฟล์ใหม่เรียบร้อยแล้ว: {output_path}")
