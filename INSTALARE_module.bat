@echo off
echo ================================================
echo  Instalare module Python (o singura data)
echo ================================================
python -m pip install --upgrade pip
python -m pip install customtkinter python-docx pypdf pillow pywin32 tkinterdnd2 pytesseract
echo.
echo NOTA: pentru OCR pe imagini (nr. dosar din print screen) este nevoie si de
echo       Tesseract OCR (program separat): https://github.com/UB-Mannheim/tesseract/wiki
echo       Fara Tesseract, aplicatia merge normal - doar OCR-ul pe imagini este dezactivat.
echo.
pause
