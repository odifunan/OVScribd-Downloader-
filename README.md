# Document Downloader + Excel

Versi proyek ini mengikuti pola antarmuka DocDownloader: pilihan Scribd, Issuu, SlideShare, dan Academia; kolom URL; tombol Get Link; hasil tautan file publik; serta konversi file ke Excel.

## Menjalankan
1. Node.js 20+
2. `npm install`
3. `npm start`
4. Buka `http://localhost:3000`

## Railway
Deploy sebagai Node.js app dan gunakan start command `npm start`.

## Batasan akses
Resolver hanya mencari file yang tersedia secara publik pada HTML halaman atau URL file langsung. Ia tidak membypass login, subscription, CAPTCHA, DRM, paywall, atau pembatasan akses. Untuk dokumen yang hanya dapat di-download melalui akun/izin pengguna, gunakan download resmi lalu upload file ke bagian konversi.
