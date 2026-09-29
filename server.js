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

const app=express();
const upload=multer({dest:path.join(os.tmpdir(),'scribd-excel-v2'),limits:{fileSize:25*1024*1024}});
app.use(express.json({limit:'1mb'}));
app.use(express.static('public'));

function rowsFromText(text){
  return text.split(/\r?\n/).map(s=>s.trim()).filter(Boolean).map(line=>{
    const parts=line.includes('\t')?line.split('\t'):line.split(/\s{2,}/);
    return parts.map(x=>x.trim());
  });
}
function pad(rows){
  const n=Math.max(1,...rows.map(r=>r.length));
  return rows.map(r=>Array.from({length:n},(_,i)=>r[i]??''));
}
function workbook(rows){
  const ws=XLSX.utils.aoa_to_sheet(pad(rows));
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'Data');
  return XLSX.write(wb,{type:'buffer',bookType:'xlsx'});
}

app.post('/api/extract',upload.single('file'),async(req,res)=>{
  if(!req.file) return res.status(400).json({error:'File belum dipilih.'});
  const p=req.file.path, ext=path.extname(req.file.originalname).toLowerCase();
  try{
    let rows=[];
    if(['.xlsx','.xls','.csv'].includes(ext)){
      const wb=XLSX.readFile(p); const ws=wb.Sheets[wb.SheetNames[0]]; rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
    } else if(ext==='.txt') rows=rowsFromText(fs.readFileSync(p,'utf8'));
    else if(ext==='.pdf'){
      const data=await pdfjs(fs.readFileSync(p)); rows=rowsFromText(data.text);
    } else if(ext==='.docx'){
      const data=await mammoth.extractRawText({path:p}); rows=rowsFromText(data.value);
    } else return res.status(400).json({error:'Format didukung: PDF, DOCX, TXT, CSV, XLSX, XLS.'});
    const id=Date.now().toString(36)+Math.random().toString(36).slice(2,7);
    fs.mkdirSync(path.join(os.tmpdir(),'scribd-out'),{recursive:true});
    fs.writeFileSync(path.join(os.tmpdir(),'scribd-out',id+'.xlsx'),workbook(rows));
    res.json({id,rows:pad(rows).slice(0,200),count:rows.length});
  }catch(e){res.status(500).json({error:'Gagal membaca file: '+e.message});}
  finally{try{fs.unlinkSync(p)}catch{}}
});
app.get('/api/download/:id',(req,res)=>{
  const p=path.join(os.tmpdir(),'scribd-out',req.params.id+'.xlsx');
  if(!fs.existsSync(p)) return res.status(404).send('File kedaluwarsa. Silakan proses ulang.');
  res.download(p,'hasil-konversi.xlsx');
});

app.post('/api/check-url',async(req,res)=>{
  const url=String(req.body?.url||'').trim();
  let u; try{u=new URL(url)}catch{return res.status(400).json({error:'URL tidak valid.'});}
  if(!/(^|\.)scribd\.com$/i.test(u.hostname)) return res.status(400).json({error:'Masukkan URL Scribd yang valid.'});
  try{
    const r=await axios.get(url,{timeout:12000,headers:{'User-Agent':'Mozilla/5.0'},maxRedirects:5});
    const $=cheerio.load(r.data); const title=$('meta[property="og:title"]').attr('content')||$('title').text().trim()||'Dokumen Scribd';
    const links=[]; $('a[href]').each((_,a)=>{const h=$(a).attr('href'); if(h&&/(download|\.pdf(?:\?|$)|\.docx?(?:\?|$)|\.xlsx?(?:\?|$))/i.test(h)) links.push(new URL(h,url).href)});
    res.json({title,url,publicDownloadLinks:[...new Set(links)].slice(0,10),message:links.length?'Tautan unduhan yang terlihat pada halaman ditemukan.':'Scribd tidak menampilkan tautan unduhan publik pada halaman ini. Jika ada tombol Unduh di akun Anda, gunakan unduhan resmi Scribd lalu unggah file tersebut ke sini.'});
  }catch(e){res.status(502).json({error:'Halaman Scribd tidak dapat diperiksa dari server. Buka URL di Scribd atau unduh file secara resmi, lalu unggah file ke bagian konversi.'});}
});

app.listen(process.env.PORT||3000,()=>console.log('Scribd Excel V2 running on '+(process.env.PORT||3000)));
