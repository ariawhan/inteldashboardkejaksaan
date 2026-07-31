# Menjalankan IndraOne di Komputer Server Windows

Panduan ini untuk memindahkan aplikasi ke komputer server, menjalankan dari Desktop, membuat autostart, dan menyetel nomor surat terakhir dari proses manual ke digital.

## 1. Persiapan Server

Install di komputer server:

- Python 3.11 atau lebih baru, centang `Add Python to PATH`
- XAMPP, umumnya di `C:\xampp`
- Google Chrome

Copy folder project `Intel-Dashboard` ke Desktop server, misalnya:

```text
C:\Users\Server\Desktop\Intel-Dashboard
```

Pastikan file `.env` sudah sesuai database XAMPP server:

```env
FLASK_HOST=0.0.0.0
FLASK_PORT=5050
DB_HOST=127.0.0.1
DB_PORT=3306
DB_USER=root
DB_PASSWORD=
DB_NAME=intel_dashboard
```

## 2. Jalankan Dengan Double Click

Double-click file:

```text
Run IndraOne.bat
```

Jika ingin membuat shortcut langsung di Desktop server:

```powershell
.\Create Desktop Shortcut.ps1
```

Setelah itu cukup double-click shortcut `IndraOne Server` di Desktop.

File ini akan:

- menyalakan XAMPP/MySQL
- membuat `.venv` jika belum ada
- memasang dependency Python
- memasang Playwright Chromium untuk ekspor PDF
- menjalankan migration database
- membuka aplikasi di jaringan lokal

Akses dari server:

```text
http://127.0.0.1:5050
```

Akses dari komputer lain:

```text
http://IP-SERVER:5050
```

Contoh:

```text
http://192.168.120.177:5050
```

Jika komputer lain belum bisa membuka, izinkan port `5050` di Windows Defender Firewall.

## 3. Membuat Aplikasi Otomatis Jalan Saat Komputer Hidup

Klik kanan PowerShell, pilih `Run as Administrator`, lalu jalankan:

```powershell
cd "$env:USERPROFILE\Desktop\Intel-Dashboard"
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
.\Install IndraOne Autostart.ps1
```

Autostart ini berjalan saat user Windows login. Ini lebih aman untuk aplikasi yang perlu Chrome/Playwright dibanding service murni sebelum login.

Untuk menghapus autostart:

```powershell
.\Uninstall IndraOne Autostart.ps1
```

## 4. Set Nomor Surat Terakhir Saat Migrasi Manual ke Digital

Nomor surat di sistem memakai konsep `nomor berikutnya`. Jadi yang dimasukkan adalah nomor terakhir yang sudah dipakai manual.

Contoh jika LAPINHAR terakhir manual tahun 2026 adalah:

```text
R-LIH-284/N.1.11/Dek.4/07/2026
```

Jalankan:

```powershell
.\.venv\Scripts\python.exe -m flask --app app set-last-number lapinhar 284 --year 2026
```

Maka nomor LAPINHAR berikutnya menjadi `285`.

Contoh jika LAPINSUS terakhir manual tahun 2026 adalah:

```text
R-LIK-202/N.1.11/Dek.4/07/2026
```

Jalankan:

```powershell
.\.venv\Scripts\python.exe -m flask --app app set-last-number lapinsus 202 --year 2026
```

Maka nomor LAPINSUS berikutnya menjadi `203`.

Jika tahun berganti, sistem otomatis mulai dari 1 untuk tahun baru. Jika ingin menyetel awal tahun tertentu:

```powershell
.\.venv\Scripts\python.exe -m flask --app app set-last-number lapinhar 0 --year 2027
.\.venv\Scripts\python.exe -m flask --app app set-last-number lapinsus 0 --year 2027
```

## 5. Backup Saat Pindah Komputer

Yang perlu dipindahkan/backup:

- database MySQL `intel_dashboard`
- folder `uploads`
- file `.env`
- source aplikasi

Jangan hanya copy source tanpa database, karena nomor surat, user, konfigurasi TTD, dan laporan tersimpan di MySQL.

## 6. Jika Login SIPede/Inteliz Kena Proxy

Jika Chrome biasa bisa membuka SIPede, tetapi popup login otomatis menampilkan `net::ERR_NETWORK_ACCESS_DENIED`, pastikan `.env` berisi:

```env
BROWSER_USE_SYSTEM_PROXY=1
```

Jika proxy Windows masih belum terbaca, isi proxy manual:

```env
BROWSER_PROXY_SERVER=http://alamat-proxy:port
BROWSER_PROXY_BYPASS=localhost;127.0.0.1;<-loopback>
```

Khusus SIPede saja:

```env
SIPEDE_BROWSER_PROXY_SERVER=http://alamat-proxy:port
```

Jika muncul `ERR_PROXY_CONNECTION_FAILED`, matikan proxy khusus SIPede agar langsung mencoba koneksi biasa:

```env
SIPEDE_BROWSER_USE_SYSTEM_PROXY=0
```

Untuk melihat Chrome automasi saat debug:

```env
SIPEDE_BROWSER_HEADLESS=0
SIPEDE_BROWSER_SLOW_MO=250
```

Setelah mengubah `.env`, restart `Run IndraOne.bat`.
