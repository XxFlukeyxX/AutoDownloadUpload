@echo off
chcp 65001 >nul
cd /d "C:\Users\mikot\Desktop\testlogin"

REM ตั้งค่าตัวแปรโฟลเดอร์
set DOC_FOLDER=C:\Users\mikot\Desktop\testlogin\doc
set OUT_FOLDER=C:\Users\mikot\Desktop\testlogin\doc\Out
set PROCESSED_FILES=C:\Users\mikot\Desktop\testlogin\processed_files.txt

REM สร้างไฟล์สำหรับเก็บรายชื่อไฟล์ที่ประมวลผลแล้ว (ถ้ายังไม่มี) และให้เป็น UTF-8
if not exist "%PROCESSED_FILES%" (
    > "%PROCESSED_FILES%" echo UTF-8 Files List
)

:loop
echo กำลังตรวจสอบไฟล์ใหม่...

REM สะสมรายชื่อไฟล์ใหม่ในตัวแปร
set "new_files="
for %%f in ("%DOC_FOLDER%\*.pdf") do (
    find "%%~nxf" "%PROCESSED_FILES%" >nul
    if errorlevel 1 (
        echo เจอไฟล์ใหม่: %%~nxf
        set "new_files=!new_files! %%f"
        REM เพิ่มชื่อไฟล์ลงใน PROCESSED_FILES ทันที
        echo %%~nxf >> "%PROCESSED_FILES%"
    )
)

REM หากพบไฟล์ใหม่ให้รัน Python Script
if defined new_files (
    REM รัน finish-code-signature.py
    python finish-code-signature.py !new_files!

    REM รอ 3 วินาที
    timeout /t 3 /nobreak >nul

    REM รัน PDF_info_to_Web.py
    python PDF_info_to_Web.py !new_files!
    
)

REM รอ 5 วินาทีก่อนตรวจสอบใหม่
timeout /t 5 /nobreak >nul
goto loop
