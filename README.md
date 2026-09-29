# Scribd → Excel V2 (public-download safe)

Aplikasi Node.js untuk:
- memeriksa URL Scribd;
- mengunduh **file yang memang tersedia secara publik** melalui URL file langsung;
- mengonversi PDF/DOCX/TXT/CSV/XLS/XLSX menjadi Excel;
- preview hasil dan download `hasil-konversi.xlsx`.

## Jalankan
```bash
npm install
npm start
```

Buka `http://localhost:3000`.

## Batasan Scribd
Aplikasi **tidak** membypass login, subscription, CAPTCHA, DRM, paywall, token privat, atau proteksi Scribd.
Jika sebuah dokumen hanya dapat diunduh setelah login/berlangganan, gunakan fitur unduh resmi Scribd atau unggah file yang memang Anda miliki/hak untuk mengolahnya.

## Railway
- Root directory: folder proyek ini.
- Build/install: `npm install`
- Start command: `npm start`
- Port: gunakan `process.env.PORT` (sudah didukung server).

## Alur baru
1. Tempel URL Scribd → Periksa URL.
2. Jika server menemukan URL file yang benar-benar publik, klik **Gunakan tautan publik ini**.
3. Jika tidak, gunakan tombol **Unduh resmi** di Scribd lalu unggah file ke bagian konversi.
