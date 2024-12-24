@echo off
cd /d "C:\Users\mikot\Desktop\testlogin"  # ไปยังพาธที่เก็บไฟล์ Python ของคุณ

# รัน signature_detection.py ก่อน
python .\sign.py

# รอ 10 วินาที
timeout /t 5

# รัน test.py หลังจากรอ 10 วินาที
python .\PDF_info_to_Web.py

pause  # รอให้คุณเห็นผลลัพธ์
