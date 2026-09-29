import express from 'express';
import multer from 'multer';
import fs from 'fs';
import path from 'path';
import os from 'os';
import axios from 'axios';
import * as XLSX from 'xlsx';
import * as pdfjs from 'pdf-parse';
import mammoth from 'mammoth';
import * as cheerio from 'cheerio';

const app = express();
const TMP = path.join(os.tmpdir(), 'scribd-excel-v2');
const OUT = path.join(os.tmpdir(), 'scribd-out');
fs.mkdirSync(TMP, { recursive: true });
fs.mkdirSync(OUT, { recursive: true });

const upload = multer({
  dest: TMP,
  limits: { fileSize: 25 * 1024 * 1024 }
});

app.use(express.json({ limit: '1mb' }));
app.use(express.static('public'));

function rowsFromText(text) {
  return text.split(/\r?\n/)
    .map(s => s.trim())
    .filter(Boolean)
    .map(line => {
      const parts = line.includes('\t') ? line.split('\t') : line.split(/\s{2,}/);
      return parts.map(x => x.trim());
    });
}

function pad(rows) {
  const n = Math.max(1, ...rows.map(r => r.length));
  return rows.map(r => Array.from({ length: n }, (_, i) => r[i] ?? ''));
}

function workbook(rows) {
  const ws = XLSX.utils.aoa_to_sheet(pad(rows));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Data');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

function safeExt(name, contentType = '') {
  const ext = path.extname(new URL(name, 'http://localhost').pathname).toLowerCase();
  if (['.pdf','.doc','.docx','.txt','.csv','.xls','.xlsx'].includes(ext)) return ext;
  const ct = contentType.split(';')[0].toLowerCase();
  const map = {
    'application/pdf': '.pdf',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': '.docx',
    'application/msword': '.doc',
    'text/plain': '.txt',
    'text/csv': '.csv',
    'application/vnd.ms-excel': '.xls',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': '.xlsx'
  };
  return map[ct] || '';
}

async function downloadPublicFile(url) {
  const r = await axios.get(url, {
    responseType: 'arraybuffer',
    timeout: 30000,
    maxRedirects: 5,
    maxContentLength: 25 * 1024 * 1024,
    maxBodyLength: 25 * 1024 * 1024,
    headers: { 'User-Agent': 'Mozilla/5.0' }
  });
  const type = String(r.headers['content-type'] || '');
  const ext = safeExt(url, type);
  if (!ext) {
    throw new Error('URL tersebut tidak mengarah ke file PDF/DOC/DOCX/TXT/CSV/XLS/XLSX yang dapat diakses secara publik.');
  }
  const filename = 'dokumen-' + Date.now() + ext;
  const filePath = path.join(TMP, filename);
  fs.writeFileSync(filePath, Buffer.from(r.data));
  return { filePath, filename, contentType: type };
}

async function extractFile(filePath, originalname) {
  const ext = path.extname(originalname).toLowerCase();
  let rows = [];

  if (['.xlsx','.xls','.csv'].includes(ext)) {
    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  } else if (ext === '.txt' || ext === '.csv') {
    rows = rowsFromText(fs.readFileSync(filePath, 'utf8'));
  } else if (ext === '.pdf') {
    const data = await pdfjs(fs.readFileSync(filePath));
    rows = rowsFromText(data.text);
  } else if (ext === '.docx') {
    const data = await mammoth.extractRawText({ path: filePath });
    rows = rowsFromText(data.value);
  } else if (ext === '.doc') {
    throw new Error('Format DOC lama belum didukung. Simpan sebagai DOCX atau PDF terlebih dahulu.');
  } else {
    throw new Error('Format didukung: PDF, DOCX, TXT, CSV, XLSX, XLS.');
  }

  return rows;
}

async function createResult(filePath, originalname) {
  const rows = await extractFile(filePath, originalname);
  const id = Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
  fs.writeFileSync(path.join(OUT, id + '.xlsx'), workbook(rows));
  return { id, rows: pad(rows).slice(0, 200), count: rows.length };
}

app.post('/api/extract', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'File belum dipilih.' });
  try {
    const result = await createResult(req.file.path, req.file.originalname);
    res.json(result);
  } catch (e) {
    res.status(500).json({ error: 'Gagal membaca file: ' + e.message });
  } finally {
    try { fs.unlinkSync(req.file.path); } catch {}
  }
});

/*
 * Unduh URL file yang memang dapat diakses publik.
 * Ini tidak mencoba melewati login, subscription, CAPTCHA, DRM, paywall,
 * atau proteksi Scribd.
 */
app.post('/api/download-public', async (req, res) => {
  const url = String(req.body?.url || '').trim();
  let u;
  try { u = new URL(url); }
  catch { return res.status(400).json({ error: 'URL tidak valid.' }); }

  if (!['http:', 'https:'].includes(u.protocol)) {
    return res.status(400).json({ error: 'URL harus menggunakan HTTP atau HTTPS.' });
  }

  try {
    const file = await downloadPublicFile(url);
    res.download(file.filePath, file.filename, () => {
      try { fs.unlinkSync(file.filePath); } catch {}
    });
  } catch (e) {
    res.status(400).json({
      error: e.response?.status === 401 || e.response?.status === 403
        ? 'File membutuhkan izin/login sehingga tidak dapat diunduh oleh server. Gunakan tombol Unduh resmi Scribd atau unggah file yang Anda miliki.'
        : e.message
    });
  }
});

app.get('/api/download/:id', (req, res) => {
  const p = path.join(OUT, req.params.id + '.xlsx');
  if (!fs.existsSync(p)) return res.status(404).send('File kedaluwarsa. Silakan proses ulang.');
  res.download(p, 'hasil-konversi.xlsx');
});

app.post('/api/check-url', async (req, res) => {
  const url = String(req.body?.url || '').trim();
  let u;
  try { u = new URL(url); }
  catch { return res.status(400).json({ error: 'URL tidak valid.' }); }

  if (!/(^|\.)scribd\.com$/i.test(u.hostname)) {
    return res.status(400).json({ error: 'Masukkan URL Scribd yang valid.' });
  }

  try {
    const r = await axios.get(url, {
      timeout: 12000,
      headers: { 'User-Agent': 'Mozilla/5.0' },
      maxRedirects: 5
    });
    const $ = cheerio.load(r.data);
    const title =
      $('meta[property="og:title"]').attr('content') ||
      $('title').text().trim() ||
      'Dokumen Scribd';

    const links = [];
    $('a[href]').each((_, a) => {
      const h = $(a).attr('href');
      if (!h) return;
      if (/(download|\.pdf(?:\?|$)|\.docx?(?:\?|$)|\.xlsx?(?:\?|$))/i.test(h)) {
        try { links.push(new URL(h, url).href); } catch {}
      }
    });

    const unique = [...new Set(links)].slice(0, 10);
    res.json({
      title,
      url,
      publicDownloadLinks: unique,
      message: unique.length
        ? 'Ditemukan tautan yang tampak seperti unduhan publik. Jika tautan tersebut dapat diakses tanpa login, Anda dapat membukanya atau menyalinnya ke fitur "Unduh URL File Publik".'
        : 'Tidak ditemukan tautan file publik pada halaman Scribd. Jika dokumen dapat Anda unduh, gunakan tombol Unduh resmi Scribd lalu unggah file tersebut ke aplikasi.'
    });
  } catch {
    res.status(502).json({
      error: 'Halaman Scribd tidak dapat diperiksa dari server. Buka URL di Scribd atau gunakan unduhan resmi Scribd.'
    });
  }
});

app.listen(process.env.PORT || 3000, () =>
  console.log('Scribd Excel V2 running on ' + (process.env.PORT || 3000))
);
