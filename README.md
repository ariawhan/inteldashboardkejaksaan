# Intel Dashboard

Aplikasi Flask untuk menyusun dan mengelola LAPINHAR serta LAPINSUS dengan database MySQL/MariaDB. Fitur mencakup editor laporan, pratinjau folio, ekspor PDF, lampiran foto, penomoran surat tahunan, manajemen pengguna, serta konfigurasi integrasi Inteliz, Sipede, dan WhatsApp.

## Persyaratan

- Python 3.11 atau lebih baru
- MySQL/MariaDB (dapat menggunakan XAMPP)
- Google Chrome untuk ekspor PDF berbasis pratinjau
- Git jika ingin mengelola source code

## Instalasi

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
Copy-Item .env.example .env
```

Ubah `.env`, terutama `SECRET_KEY`, `ADMIN_PASSWORD`, dan konfigurasi database. Jangan mengunggah `.env` ke repository.

Aktifkan MySQL pada XAMPP, kemudian jalankan migration dan aplikasi:

```powershell
python -m flask --app app migrate
python app.py
```

Aplikasi tersedia pada `http://127.0.0.1:5050`. Komputer lain pada jaringan lokal dapat menggunakan alamat IP server, misalnya `http://192.168.1.10:5050`.

Pada Windows, aplikasi juga dapat dijalankan melalui `start.bat`.

## Struktur utama

- `app.py` — aplikasi Flask dan route
- `templates/` — template halaman
- `static/` — CSS, JavaScript, gambar, dan editor
- `migrations/` — migration MySQL berurutan
- `uploads/` — lampiran lokal; tidak masuk Git

## Keamanan repository

File berikut sengaja diabaikan oleh Git:

- `.env` dan kredensial lokal
- upload serta lampiran laporan
- log server dan database lokal
- dokumen contoh, PDF, dan DOCX hasil ekspor
- profil browser, cookie, token, dan session automasi

Sebelum menjadikan repository publik, periksa kembali riwayat commit untuk memastikan tidak pernah ada kredensial atau dokumen sensitif yang terunggah.

## Catatan integrasi

Kredensial Inteliz dan Sipede disimpan terenkripsi di database milik instalasi lokal. Database tersebut tidak boleh dimasukkan ke repository. Automasi eksternal harus digunakan sesuai kewenangan dan kebijakan keamanan instansi.
