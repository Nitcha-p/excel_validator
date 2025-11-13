@echo off
REM ==========================================
REM  Multi-Sheet Excel Validator - Launcher
REM ==========================================

REM เปลี่ยน directory ให้มาอยู่ที่โฟลเดอร์ไฟล์ .bat นี้
cd /d "%~dp0"

REM ถ้ามี venv ให้ใช้ python จาก venv
if exist "venv\Scripts\python.exe" (
    set "PYTHON=venv\Scripts\python.exe"
) else (
    REM ถ้าไม่มี venv จะ fallback ไปที่ python ทั่วไป (กรณี dev คนอื่น)
    set "PYTHON=python"
)

echo Starting Excel Validator with Streamlit...
echo.

"%PYTHON%" -m streamlit run final_excel_sheet_validator.py

echo.
echo Application has stopped. Press any key to close this window.
pause >nul
