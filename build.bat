@echo off
chcp 65001 > nul
echo ===================================================
echo   PDF OCR Tool GR - Δημιουργία Εκτελέσιμου (EXE)
echo ===================================================
echo.

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

if not exist ".venv\Scripts\python.exe" (
    echo [1/3] Δημιουργία εικονικού περιβάλλοντος...
    uv venv .venv
)

echo [2/3] Έλεγχος και εγκατάσταση εξαρτήσεων...
uv pip install --python .venv\Scripts\python.exe -r requirements.txt

echo.
echo [3/3] Συσκευασία εφαρμογής με το PyInstaller...
echo.

".venv\Scripts\pyinstaller.exe" ^
    --noconsole ^
    --onefile ^
    --clean ^
    --name "PDF_OCR_Tool_GR" ^
    --collect-all pymupdf ^
    --collect-all docx ^
    --collect-all rapidocr_onnxruntime ^
    --collect-all onnxruntime ^
    --hidden-import=PIL ^
    --hidden-import=pytesseract ^
    --hidden-import=numpy ^
    "pdf_ocr_tool.py"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ===================================================
    echo  ΕΠΙΤΥΧΙΑ! Το εκτελέσιμο αρχείο δημιουργήθηκε:
    echo  %SCRIPT_DIR%dist\PDF_OCR_Tool_GR.exe
    echo ===================================================
    echo.
) else (
    echo.
    echo ===================================================
    echo  ΣΦΑΛΜΑ κατά τη δημιουργία του εκτελέσιμου!
    echo ===================================================
    echo.
)

pause
