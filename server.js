import express from 'express';
import multer from 'multer';
import fs from 'fs';
import path from 'path';
import os from 'os';
import axios from 'axios';
import * as XLSX from 'xlsx';
import pdf from 'pdf-parse';
import mammoth from 'mammoth';
import * as cheerio from 'cheerio';

const app = express();
const TMP = path.join(os.tmpdir(), 'scribd-excel-v2');
const OUT = path.join(os.tmpdir(), 'scribd-out');
fs.mkdirSync(TMP, { recursive: true });
fs.mkdirSync(OUT, { recursive: true });
const upload = multer({ dest: TMP, limits: { fileSize: 25 * 1024 * 1024 } });
app.use(express.json({ limit: '1mb' }));
app.use(express.static('public'));

const SOURCES = {
  scribd: { label: 'Scribd', hosts: ['scribd.com', 'www.scribd.com', 'id.scribd.com'] },
  issuu: { label: 'Issuu', hosts: ['issuu.com'] },
  slideshare: { label: 'SlideShare', hosts: ['slideshare.net', 'www.slideshare.net'] },
  academia: { label: 'Academia', hosts: ['academia.edu', 'www.academia.edu'] }
};

function rowsFromText(text) {
  return text.split(/\r?\n/).map(s => s.trim()).filter(Boolean).map(line => {
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
function sourceFor(hostname) {
  const h = hostname.toLowerCase();
  return Object.entries(SOURCES).find(([, s]) => s.hosts.includes(h))?.[0] || null;
}
function isAllowedHost(url) {
  try { return !!sourceFor(new URL(url).hostname); } catch { return false; }
}
function isDirectFile(url) {
  return /\.(pdf|docx?|xlsx?|csv|txt)(?:[?#].*)?$/i.test(url);
}
function absoluteLinks($, base) {
  const out = [];
  $('a[href], link[href], meta[content]').each((_, el) => {
    const raw = $(el).attr('href') || $(el).attr('content');
    if (!raw) return;
    try {
      const u = new URL(raw, base).href;
      if (/\.(pdf|docx?|xlsx?|csv|txt)(?:[?#].*)?$/i.test(u)) out.push(u);
    } catch {}
  });
  return [...new Set(out)];
}

app.post('/api/resolve', async (req, res) => {
  const url = String(req.body?.url || '').trim();
  let u;
  try { u = new URL(url); } catch { return res.status(400).json({ error: 'URL tidak valid.' }); }
  const source = sourceFor(u.hostname);
  if (!source) return res.status(400).json({ error: 'URL harus berasal dari Scribd, Issuu, SlideShare, atau Academia.' });

  try {
    if (isDirectFile(url)) {
      return res.json({ source, title: path.basename(u.pathname), url, publicDownloadLinks: [url], downloadable: true, message: 'Tautan file langsung terdeteksi.' });
    }
    const r = await axios.get(url, {
      timeout: 15000,
      maxRedirects: 5,
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/153 Safari/537.36', Accept: 'text/html,application/xhtml+xml' },
      validateStatus: s => s >= 200 && s < 400
    });
    const $ = cheerio.load(r.data);
    const title = $('meta[property="og:title"]').attr('content') || $('meta[name="twitter:title"]').attr('content') || $('title').text().trim() || `Dokumen ${SOURCES[source].label}`;
    const links = absoluteLinks($, url);
    const canonical = $('link[rel="canonical"]').attr('href') || url;
    const text = $('body').text().replace(/\s+/g, ' ').trim();
    const restricted = /sign in|log in|subscribe|premium|download.*not available|not available for download/i.test(text);
    res.json({ source, title, url: canonical, publicDownloadLinks: links.slice(0, 10), downloadable: links.length > 0, restrictedHint: restricted, message: links.length ? 'Tautan file publik yang terlihat pada halaman ditemukan.' : `Tidak ada tautan file publik yang terlihat. Jika ${SOURCES[source].label} menyediakan tombol download untuk akun Anda, gunakan download resmi lalu unggah file ke aplikasi.` });
  } catch (e) {
    res.status(502).json({ error: `Tidak dapat memeriksa halaman ${SOURCES[source].label}: ${e.message}` });
  }
});

app.post('/api/extract', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'File belum dipilih.' });
  const p = req.file.path, ext = path.extname(req.file.originalname).toLowerCase();
  try {
    let rows = [];
    if (['.xlsx', '.xls', '.csv'].includes(ext)) {
      const wb = XLSX.readFile(p); const ws = wb.Sheets[wb.SheetNames[0]];
      rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    } else if (ext === '.txt') rows = rowsFromText(fs.readFileSync(p, 'utf8'));
    else if (ext === '.pdf') rows = rowsFromText((await pdf(fs.readFileSync(p))).text);
    else if (ext === '.docx') rows = rowsFromText((await mammoth.extractRawText({ path: p })).value);
    else return res.status(400).json({ error: 'Format didukung: PDF, DOCX, TXT, CSV, XLSX, XLS.' });
    const id = Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
    fs.writeFileSync(path.join(OUT, id + '.xlsx'), workbook(rows));
    res.json({ id, rows: pad(rows).slice(0, 200), count: rows.length });
  } catch (e) { res.status(500).json({ error: 'Gagal membaca file: ' + e.message }); }
  finally { try { fs.unlinkSync(p); } catch {} }
});

app.get('/api/download/:id', (req, res) => {
  const p = path.join(OUT, req.params.id + '.xlsx');
  if (!fs.existsSync(p)) return res.status(404).send('File kedaluwarsa. Silakan proses ulang.');
  res.download(p, 'hasil-konversi.xlsx');
});

app.listen(process.env.PORT || 3000, () => console.log('Document Downloader + Excel running on ' + (process.env.PORT || 3000)));
