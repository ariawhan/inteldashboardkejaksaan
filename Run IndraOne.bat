@echo off
setlocal
title IndraOne Server

cd /d "%~dp0"

set "XAMPP_DIR=C:\xampp"
set "PYTHON_EXE=python"
set "APP_HOST=0.0.0.0"
set "APP_PORT=5050"

if exist ".env" (
  for /f "usebackq tokens=1,* delims==" %%A in (".env") do (
    if /i "%%A"=="FLASK_HOST" set "APP_HOST=%%B"
    if /i "%%A"=="FLASK_PORT" set "APP_PORT=%%B"
  )
)

if exist "%XAMPP_DIR%\xampp_start.exe" (
  echo Menjalankan XAMPP...
  start "" /min "%XAMPP_DIR%\xampp_start.exe"
) else if exist "%XAMPP_DIR%\mysql_start.bat" (
  echo Menjalankan MySQL XAMPP...
  start "" /min "%XAMPP_DIR%\mysql_start.bat"
) else (
  echo XAMPP tidak ditemukan di %XAMPP_DIR%.
  echo Jika XAMPP berada di folder lain, ubah XAMPP_DIR pada file ini.
)

echo Menunggu MySQL siap...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ready=$false; for($i=0; $i -lt 45; $i++){ try { $client=New-Object Net.Sockets.TcpClient('127.0.0.1',3306); $client.Close(); $ready=$true; break } catch { Start-Sleep -Seconds 2 } }; if(-not $ready){ exit 1 }"
if errorlevel 1 (
  echo MySQL belum bisa dihubungi di 127.0.0.1:3306.
  echo Pastikan MySQL di XAMPP menyala, lalu jalankan ulang.
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo Membuat virtual environment...
  %PYTHON_EXE% -m venv .venv
  if errorlevel 1 (
    echo Gagal membuat virtual environment. Pastikan Python sudah terinstall dan masuk PATH.
    pause
    exit /b 1
  )
)

set "PYTHON_EXE=%CD%\.venv\Scripts\python.exe"

echo Memasang dependency...
"%PYTHON_EXE%" -m pip install -r requirements.txt
if errorlevel 1 (
  echo Gagal memasang dependency. Periksa koneksi internet atau instalasi Python.
  pause
  exit /b 1
)

echo Menyiapkan Playwright Chromium...
"%PYTHON_EXE%" -m playwright install chromium

echo Menjalankan migration database...
"%PYTHON_EXE%" -m flask --app app migrate
if errorlevel 1 (
  echo Migration gagal. Periksa konfigurasi database pada .env.
  pause
  exit /b 1
)

echo.
echo IndraOne berjalan.
echo Akses dari server : http://127.0.0.1:%APP_PORT%
echo Akses dari jaringan: http://IP-SERVER:%APP_PORT%
echo.
"%PYTHON_EXE%" app.py

pause
