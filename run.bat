@echo off
chcp 65001 >nul
title Muhawwil - PDF & Images to Office
cd /d "%~dp0"
echo ============================================
echo   Muhawwil  |  http://127.0.0.1:8000
echo ============================================
start "" http://127.0.0.1:8000
if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" -m uvicorn main:app --host 127.0.0.1 --port 8000
) else (
  python -m uvicorn main:app --host 127.0.0.1 --port 8000
)
pause
