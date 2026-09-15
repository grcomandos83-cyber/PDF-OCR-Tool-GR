@echo off
chcp 65001 > nul
echo ===================================================
echo   PDF OCR Tool GR - Επαγγελματικό Εργαλείο OCR
echo ===================================================
echo.

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

if exist ".venv\Scripts\pythonw.exe" (
    echo Εκκίνηση εφαρμογής...
    start "" ".venv\Scripts\pythonw.exe" "pdf_ocr_tool.py"
    exit /b
)

if exist ".venv\Scripts\python.exe" (
    echo Εκκίνηση εφαρμογής...
    start "" ".venv\Scripts\python.exe" "pdf_ocr_tool.py"
    exit /b
)

echo Δημιουργία περιβάλλοντος και εγκατάσταση εξαρτήσεων...
uv venv .venv
uv pip install --python .venv\Scripts\python.exe -r requirements.txt

echo Εκκίνηση εφαρμογής...
start "" ".venv\Scripts\pythonw.exe" "pdf_ocr_tool.py"
