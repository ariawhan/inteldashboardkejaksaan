@echo off
title Intel Dashboard
cd /d "%~dp0"
echo Memeriksa dependency...
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo Gagal memasang dependency. Periksa Python dan koneksi internet.
  pause
  exit /b 1
)
echo Menjalankan Intel Dashboard...
python app.py
pause
