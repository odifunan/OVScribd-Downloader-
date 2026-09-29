# Scribd → Excel V3

Versi ini mempertahankan konversi file ke Excel dan menambahkan integrasi backend dengan `Phoenix124/scribd-downloader` untuk dokumen Scribd yang dapat diakses secara sah.

## Fitur
- UI gaya DocDownloader: URL → Get Link → hasil.
- Server-side Scribd engine melalui Python `scribdl`.
- Mode Auto/Text/Images.
- Hasil teks: download Markdown + Excel.
- Hasil halaman gambar: gabungkan menjadi PDF.
- Upload manual PDF/DOCX/TXT/MD/CSV/XLS/XLSX → Excel.
- Batas upload 25 MB.
- Dockerfile siap untuk Railway.

## Menjalankan dengan Docker / Railway
```bash
docker build -t scribd-excel-v3 .
docker run -p 3000:3000 scribd-excel-v3
```

Railway dapat menggunakan Dockerfile yang sudah disertakan. Tidak perlu memasukkan perintah `pip install ...` secara manual karena Dockerfile memasang dependency saat build.

## Menjalankan lokal tanpa Docker
Node.js 20+ dan Python 3 diperlukan.
```bash
npm install
python3 -m venv .venv
# Windows: .venv\\Scripts\\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
npm start
```

## Catatan penggunaan
Engine pihak ketiga diintegrasikan tanpa menerima username/password Scribd/Everand, CAPTCHA, DRM bypass, atau kredensial subscription. Gunakan hanya pada dokumen yang memang dapat Anda akses dan unduh secara sah.

Repository engine menyatakan proyeknya unsupported dan mendokumentasikan bahwa dokumen Scribd berbasis gambar diproses berbeda dari dokumen teks; karena itu mode Auto dapat mencoba mode gambar sebagai fallback. Lihat dokumentasi upstream sebelum deployment produksi.
