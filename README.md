# Scribd Excel V8 — Layout Asli

Perubahan utama:
- PDF upload menghasilkan Excel `Tampilan Asli`: satu halaman PDF = satu worksheet.
- Halaman PDF dirender sebagai gambar dengan rasio asli, sehingga posisi tabel/tulisan tidak direkonstruksi menjadi sel.
- Untuk Scribd mode Auto, aplikasi mencoba mode halaman gambar terlebih dahulu agar layout visual dipertahankan; teks hanya menjadi fallback.
- Excel data biasa tetap tersedia untuk ekstraksi/editing.
- Download utama untuk PDF diarahkan ke Excel layout asli.
- Tidak melewati CAPTCHA, Browser Check, paywall, DRM, atau login.

Deploy Railway: gunakan Dockerfile yang disertakan.
