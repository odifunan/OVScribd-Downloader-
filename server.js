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
import { spawn } from 'child_process';

const app = express();
const TMP = path.join(os.tmpdir(), 'scribd-excel-v2');
const OUT = path.join(os.tmpdir(), 'scribd-out');
fs.mkdirSync(TMP, { recursive: true });
fs.mkdirSync(OUT, { recursive: true });
const upload = multer({ dest: TMP, limits: { fileSize: 25 * 1024 * 1024 } });
app.use(express.json({ limit: '1mb' }));
app.use(express.static('public'));

function cleanCell(value) {
  return String(value ?? '')
    .replace(/\u00a0/g, ' ')
    .replace(/[ \t]+$/g, '')
    .trim();
}
function rowsFromText(text) {
  const lines = String(text || '').replace(/\r\n?/g, '\n').split('\n');
  const rows = [];
  for (const raw of lines) {
    const line = cleanCell(raw);
    if (!line) continue;
    let parts;
    if (line.includes('\t')) parts = line.split('\t').map(cleanCell);
    else if (/\s{2,}/.test(line)) parts = line.split(/\s{2,}/).map(cleanCell);
    else parts = [line];
    if (parts.some(Boolean)) rows.push(parts);
  }
  return rows;
}
function normalizeRows(rows) {
  return rows
    .map(r => Array.isArray(r) ? r.map(cleanCell) : [cleanCell(r)])
    .filter(r => r.some(v => v !== ''));
}
function pad(rows) {
  const normalized = normalizeRows(rows);
  const n = Math.max(1, ...normalized.map(r => r.length));
  return normalized.map(r => Array.from({ length: n }, (_, i) => r[i] ?? ''));
}
function workbook(rows, fullText = '') {
  let normalized = pad(rows);
  if (normalized.length && normalized.every(r => r.length === 1)) {
    normalized = normalized.map((r, i) => [i + 1, r[0]]);
  }
  const ws = XLSX.utils.aoa_to_sheet(normalized);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Data');

  // Keep the original text in a second sheet so no content is lost during parsing.
  const full = String(fullText || '').replace(/\r\n?/g, '\n').split('\n').filter(x => x.trim());
  if (full.length) {
    const fullWs = XLSX.utils.aoa_to_sheet([['No.', 'Isi Dokumen'], ...full.map((v, i) => [i + 1, cleanCell(v)])]);
    XLSX.utils.book_append_sheet(wb, fullWs, 'Dokumen Lengkap');
  }
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}
function safeId() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 9); }
function createPdfExactExcel(id, pdfPath) {
  const output = path.join(OUT, id + '.visual.xlsx');
  const renderDir = path.join(OUT, id + '-pages');
  const helper = path.join(process.cwd(), 'scripts', 'pdf_exact_excel.py');
  fs.mkdirSync(renderDir, { recursive: true });
  return new Promise((resolve, reject) => {
    const child = spawn(pythonCommand(), [helper, pdfPath, output, renderDir], { stdio: 'inherit' });
    child.on('error', reject);
    child.on('close', code => code === 0 && fs.existsSync(output)
      ? resolve(output)
      : reject(new Error('Gagal membuat Excel dengan layout PDF asli.')));
  });
}
function createVisualExcel(id, imagePaths) {
  const output = path.join(OUT, id + '.visual.xlsx');
  const helper = path.join(process.cwd(), 'scripts', 'visual_excel.py');
  return new Promise((resolve, reject) => {
    const child = spawn(pythonCommand(), [helper, output, ...imagePaths], { stdio: 'inherit' });
    child.on('error', reject);
    child.on('close', code => code === 0 && fs.existsSync(output)
      ? resolve(output)
      : reject(new Error('Gagal membuat Excel tampilan asli.')));
  });
}
async function polishExcel(filePath) {
  const helper = path.join(process.cwd(), 'scripts', 'format_excel.py');
  const tmpPath = filePath.replace(/\.xlsx$/i, '.formatted.xlsx');
  await new Promise((resolve, reject) => {
    const child = spawn(pythonCommand(), [helper, filePath, tmpPath], { stdio: 'inherit' });
    child.on('error', reject);
    child.on('close', code => code === 0 ? resolve() : reject(new Error('Gagal merapikan format Excel.')));
  });
  fs.renameSync(tmpPath, filePath);
}

function isScribdUrl(raw) {
  try {
    const u = new URL(raw);
    return /^([a-z0-9-]+\.)*scribd\.com$/i.test(u.hostname);
  } catch { return false; }
}
function pythonCommand() {
  if (process.env.PYTHON_BIN) return process.env.PYTHON_BIN;
  return process.platform === 'win32' ? 'python' : 'python3';
}
function runScribdl(url, workDir, mode = 'auto') {
  return new Promise((resolve, reject) => {
    const args = ['-m', 'scribdl'];
    if (mode === 'images') args.push('-i');
    args.push(url);
    const child = spawn(pythonCommand(), args, { cwd: workDir, env: process.env });
    let stdout = '', stderr = '';
    child.stdout.on('data', d => stdout += d.toString());
    child.stderr.on('data', d => stderr += d.toString());
    const timer = setTimeout(() => { child.kill('SIGTERM'); reject(new Error('Proses downloader melebihi batas 90 detik.')); }, 90000);
    child.on('error', e => { clearTimeout(timer); reject(new Error('Python/scribdl belum terpasang di server: ' + e.message)); });
    child.on('close', code => {
      clearTimeout(timer);
      if (code !== 0) return reject(new Error((stderr || stdout || `scribdl berhenti dengan kode ${code}`).trim().slice(-3000)));
      resolve({ stdout, stderr });
    });
  });
}
function newestFiles(dir) {
  return fs.readdirSync(dir).map(name => ({ name, path: path.join(dir, name) }))
    .filter(x => fs.statSync(x.path).isFile())
    .sort((a,b) => fs.statSync(b.path).mtimeMs - fs.statSync(a.path).mtimeMs);
}
function buildPdfFromImages(imagePaths, outPath) {
  const imgs = imagePaths.map(p => requireImage(p));
  if (!imgs.length) throw new Error('Tidak ada halaman gambar yang dihasilkan.');
  const rgb = imgs.map(im => im.convert('RGB'));
  rgb[0].save(outPath, { saveAll: true, appendImages: rgb.slice(1), format: 'PDF', resolution: 150 });
}
// Lazy-load Pillow through a small Python helper instead of adding a Node native image dependency.
function requireImage() { throw new Error('PDF gambar dibuat oleh helper Python.'); }

app.post('/api/extract', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'File belum dipilih.' });
  const p = req.file.path, ext = path.extname(req.file.originalname).toLowerCase();
  try {
    let rows = [];
    if (['.xlsx', '.xls', '.csv'].includes(ext)) {
      const wb = XLSX.readFile(p); const ws = wb.Sheets[wb.SheetNames[0]];
      rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    } else if (ext === '.txt' || ext === '.md') rows = rowsFromText(fs.readFileSync(p, 'utf8'));
    else if (ext === '.pdf') { const data = await pdfjs(fs.readFileSync(p)); rows = rowsFromText(data.text); }
    else if (ext === '.docx') { const data = await mammoth.extractRawText({ path: p }); rows = rowsFromText(data.value); }
    else return res.status(400).json({ error: 'Format didukung: PDF, DOCX, TXT, MD, CSV, XLSX, XLS.' });
    const id = safeId();
    const originalText = ['.txt', '.md', '.pdf', '.docx'].includes(ext) ? (ext === '.docx' ? '' : fs.readFileSync(p, 'utf8')) : '';
    const standardPath = path.join(OUT, id + '.xlsx');
    fs.writeFileSync(standardPath, workbook(rows, originalText));
    await polishExcel(standardPath);
    let exact = null;
    if (ext === '.pdf') exact = await createPdfExactExcel(id, p);
    res.json({ id, rows: pad(rows).slice(0, 200), count: rows.length, exactExcel: exact ? `/api/download-exact/${id}` : null, message: exact ? 'PDF dipertahankan sebagai halaman asli di Excel.' : null });
  } catch (e) { res.status(500).json({ error: 'Gagal membaca file: ' + e.message }); }
  finally { try { fs.unlinkSync(p); } catch {} }
});

app.get('/api/download/:id', (req, res) => {
  const p = path.join(OUT, req.params.id + '.xlsx');
  if (!fs.existsSync(p)) return res.status(404).send('File kedaluwarsa. Silakan proses ulang.');
  res.download(p, 'hasil-konversi.xlsx');
});

app.get('/api/download-exact/:id', (req, res) => {
  const p = path.join(OUT, req.params.id + '.visual.xlsx');
  if (!fs.existsSync(p)) return res.status(404).send('Excel layout asli tidak ditemukan. Silakan proses ulang.');
  res.download(p, 'hasil-konversi-tampilan-asli.xlsx');
});

app.post('/api/check-url', async (req, res) => {
  const url = String(req.body?.url || '').trim();
  if (!isScribdUrl(url)) return res.status(400).json({ error: 'Masukkan URL Scribd yang valid.' });
  try {
    const r = await axios.get(url, { timeout: 12000, headers: { 'User-Agent': 'Mozilla/5.0' }, maxRedirects: 5 });
    const $ = cheerio.load(r.data);
    const title = $('meta[property="og:title"]').attr('content') || $('title').text().trim() || 'Dokumen Scribd';
    res.json({ title, url, engineAvailable: true, message: 'URL valid. Gunakan tombol Download melalui engine server untuk dokumen yang dapat diakses.' });
  } catch (e) {
    res.status(502).json({ error: 'Halaman Scribd tidak dapat diperiksa dari server. Coba lagi atau gunakan unduhan resmi lalu unggah file.' });
  }
});

app.post('/api/scribd/download', async (req, res) => {
  const url = String(req.body?.url || '').trim();
  const mode = ['auto', 'text', 'images'].includes(req.body?.mode) ? req.body.mode : 'auto';
  if (!isScribdUrl(url)) return res.status(400).json({ error: 'URL Scribd tidak valid.' });

  const jobId = safeId();
  const dir = path.join(TMP, 'jobs', jobId);
  fs.mkdirSync(dir, { recursive: true });
  try {
    // The third-party tool is used only for publicly accessible Scribd documents; no account, paywall,
    // CAPTCHA, DRM or subscription credentials are accepted by this application.
    await runScribdl(url, dir, mode === 'text' ? 'text' : 'images');
    let files = newestFiles(dir);
    let md = files.find(x => /\.md$/i.test(x.name));
    let images = files.filter(x => /\.(jpe?g|png|webp)$/i.test(x.name));

    // Auto mode prefers page images so the original visual layout is preserved.
    if (mode === 'auto' && images.length === 0 && !md) {
      await runScribdl(url, dir, 'text');
      files = newestFiles(dir); md = files.find(x => /\.md$/i.test(x.name)); images = files.filter(x => /\.(jpe?g|png|webp)$/i.test(x.name));
    }

    if (md) {
      const id = safeId();
      const text = fs.readFileSync(md.path, 'utf8');
      const base = path.basename(md.name, '.md');
      const mdPath = path.join(OUT, id + '.md');
      const xlsxPath = path.join(OUT, id + '.xlsx');
      fs.writeFileSync(mdPath, text);
      fs.writeFileSync(xlsxPath, workbook(rowsFromText(text), text));
      await polishExcel(xlsxPath);
      return res.json({ id, type: 'text', title: base, textPreview: text.slice(0, 8000), download: `/api/scribd/file/${id}`, excel: `/api/scribd/excel/${id}` });
    }

    if (images.length) {
      const id = safeId();
      const manifest = images.map(x => x.path);
      fs.writeFileSync(path.join(OUT, id + '.json'), JSON.stringify({ images: manifest }));
      await createVisualExcel(id, manifest);
      return res.json({ id, type: 'images', pages: images.length, title: path.basename(images[0].name).replace(/[_-]?(\d+)\.(jpe?g|png|webp)$/i, ''), download: `/api/scribd/file/${id}`, excel: `/api/scribd/exact-excel/${id}`, pdf: `/api/scribd/pdf/${id}`, note: 'Excel ini mempertahankan tampilan halaman sebagai gambar agar posisi tabel dan tulisan sama seperti dokumen asli.' });
    }
    return res.status(422).json({ error: 'Engine tidak menghasilkan dokumen. Pastikan URL menunjuk ke dokumen yang dapat diakses dan Anda memiliki hak untuk mengunduhnya.' });
  } catch (e) {
    const msg = String(e?.message || e);
    const blocked = /Client Challenge|browser check|blocked the automated request|ScribdFetchError/i.test(msg);
    if (blocked) {
      return res.status(502).json({
        code: 'SCRIBD_BROWSER_CHECK',
        error: 'Scribd memblokir permintaan otomatis dari server (Browser Check). Karena itu dokumen belum dapat diambil oleh engine.',
        detail: 'Buka dokumen di Scribd menggunakan browser, gunakan fitur download resmi jika tersedia, lalu unggah file tersebut ke bagian konversi Excel.',
        openUrl: url
      });
    }
    return res.status(502).json({ error: 'Gagal mengambil dokumen: ' + msg.slice(-1200), hint: 'Pastikan Python 3 dan scribd-downloader sudah terpasang. Di Railway, gunakan Dockerfile yang disertakan.' });
  }
});

app.get('/api/scribd/file/:id', (req, res) => {
  const md = path.join(OUT, req.params.id + '.md');
  const manifest = path.join(OUT, req.params.id + '.json');
  if (fs.existsSync(md)) return res.download(md, 'dokumen-scribd.md');
  if (fs.existsSync(manifest)) return res.status(409).json({ error: 'Dokumen berupa halaman gambar. Gunakan endpoint PDF.' });
  res.status(404).send('Hasil tidak ditemukan.');
});
app.get('/api/scribd/excel/:id', (req, res) => {
  const p = path.join(OUT, req.params.id + '.xlsx');
  if (!fs.existsSync(p)) return res.status(404).send('Excel tidak ditemukan.');
  res.download(p, 'hasil-scribd.xlsx');
});
app.get('/api/scribd/exact-excel/:id', (req, res) => {
  const p = path.join(OUT, req.params.id + '.visual.xlsx');
  if (!fs.existsSync(p)) return res.status(404).send('Excel tampilan asli tidak ditemukan.');
  res.download(p, 'dokumen-scribd-tampilan-asli.xlsx');
});
app.get('/api/scribd/pdf/:id', (req, res) => {
  const manifestPath = path.join(OUT, req.params.id + '.json');
  if (!fs.existsSync(manifestPath)) return res.status(404).send('PDF hanya tersedia untuk hasil halaman gambar.');
  const data = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
  const helper = path.join(process.cwd(), 'scripts', 'images_to_pdf.py');
  const output = path.join(OUT, req.params.id + '.pdf');
  const child = spawn(pythonCommand(), [helper, output, ...data.images], { stdio: 'inherit' });
  child.on('error', e => res.status(500).send('Gagal membuat PDF: ' + e.message));
  child.on('close', code => {
    if (code !== 0 || !fs.existsSync(output)) return res.status(500).send('Gagal membuat PDF.');
    res.download(output, 'dokumen-scribd.pdf');
  });
});

app.get('/api/health', (req, res) => res.json({ ok: true, python: pythonCommand() }));

app.listen(process.env.PORT || 3000, () => console.log('Scribd Excel V2 running on ' + (process.env.PORT || 3000)));
