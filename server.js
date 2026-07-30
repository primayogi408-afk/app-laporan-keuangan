
const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const multer = require('multer');

const app = express();
const IS_VERCEL = !!process.env.VERCEL;
const DATA_DIR = IS_VERCEL ? path.join('/tmp','data') : path.join(__dirname,'data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

const PERIODS_PATH = path.join(DATA_DIR, 'periods.json');

function configPath(period){ return path.join(DATA_DIR, `config_${(period||'master').replace(/\s+/g,'-')}.json`); }
function editsPath(period){ return path.join(DATA_DIR, `edits_${(period||'master').replace(/\s+/g,'-')}.json`); }
function excelPath(period){
  const key = (period||'master').replace(/\s+/g,'-');
  return path.join(DATA_DIR, `${key}.xlsx`);
}

// INIT copy bundled
(function initBundledData(){
  try {
    const sources = [
      path.join(__dirname,'data','master.xlsx'),
      path.join(__dirname,'master.xlsx'),
      path.join(__dirname,'data','data.xlsx'),
    ];
    for (const src of sources){
      if (fs.existsSync(src)){
        const dest = path.join(DATA_DIR,'master.xlsx');
        if (!fs.existsSync(dest) || fs.statSync(src).size !== fs.statSync(dest).size){
          fs.copyFileSync(src,dest);
          console.log('Copied',src,'->',dest);
        }
        break;
      }
    }
    const bundledDir = path.join(__dirname,'data');
    if (fs.existsSync(bundledDir)){
      for (const f of fs.readdirSync(bundledDir)){
        if (!f.toLowerCase().endsWith('.xlsx')) continue;
        const src = path.join(bundledDir,f);
        const dest = path.join(DATA_DIR,f);
        if (fs.statSync(src).isFile() && !fs.existsSync(dest)) fs.copyFileSync(src,dest);
      }
    }
    if (!fs.existsSync(PERIODS_PATH)){
      const hasMaster = fs.existsSync(path.join(DATA_DIR,'master.xlsx'));
      if (hasMaster) fs.writeFileSync(PERIODS_PATH, JSON.stringify([{key:'master',label:'Utama'}],null,2));
    }
  }catch(e){ console.log('init error',e.message); }
})();

const storage = multer.memoryStorage();
const upload = multer({ storage });

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname,'public')));
app.use(express.static(__dirname));

app.get('/', (req,res)=>{
  const candidates = [path.join(__dirname,'public','index.html'), path.join(__dirname,'index.html')];
  for (const p of candidates) if (fs.existsSync(p)) return res.sendFile(p);
  res.send('<h1>index.html not found</h1><p>Root:'+fs.readdirSync(__dirname).join(', ')+'</p>');
});

app.get('/api', (req,res)=> res.json({status:'OK'}));

function loadPeriods(){
  try{
    if (!fs.existsSync(DATA_DIR)) return [];
    const files = fs.readdirSync(DATA_DIR).filter(f=>f.toLowerCase().endsWith('.xlsx'));
    const periods=[];
    for (const f of files){
      const key = path.basename(f,'.xlsx');
      if (key.startsWith('~')) continue;
      let label=key;
      try{
        if (fs.existsSync(PERIODS_PATH)){
          const saved=JSON.parse(fs.readFileSync(PERIODS_PATH,'utf8'));
          const found=saved.find(p=>p.key===key || p.key.replace(/\s+/g,'-')===key);
          if (found) label=found.label;
        }
      }catch{}
      periods.push({key,label});
    }
    periods.sort((a,b)=> a.key.toLowerCase()==='master'?-1 : b.key.toLowerCase()==='master'?1 : 0);
    return periods;
  }catch{ return []; }
}
function savePeriods(list){ try{ fs.writeFileSync(PERIODS_PATH, JSON.stringify(list,null,2)); }catch{} }

function excelExists(period){
  if (!period) return fs.existsSync(path.join(DATA_DIR,'master.xlsx'));
  const cand=[period, period.replace(/\s+/g,'-'), period.replace(/-/g,' '), period.toUpperCase(), period.replace(/\s+/g,'-').toUpperCase()];
  for (const c of cand){ if (fs.existsSync(path.join(DATA_DIR, `${c}.xlsx`))) return true; if (fs.existsSync(path.join(DATA_DIR, `${c.replace(/\s+/g,'-')}.xlsx`))) return true; }
  return fs.existsSync(path.join(DATA_DIR,'master.xlsx'));
}
function loadEdits(period){
  try{ const ep=editsPath(period); if (fs.existsSync(ep)) return JSON.parse(fs.readFileSync(ep,'utf8')); }catch{}
  return {};
}
function saveEdits(edits,period){ try{ fs.writeFileSync(editsPath(period), JSON.stringify(edits,null,2)); }catch{} }

function loadConfig(kategoriNames, period){
  const defaults={
    neraca:{'AKTIVA LANCAR':[], 'AKTIVA LAIN-LAIN':[], 'AKTIVA TETAP':[], 'HUTANG LANCAR':[], 'MODAL':[]},
    labaRugi:{'PENJUALAN':[], 'HARGA POKOK PENJUALAN':[], 'BIAYA PENJUALAN':[], 'BIAYA ADM UMUM':[], 'PENDAPATAN NON OPERASIONAL':[], 'BIAYA NON OPERASIONAL':[]}
  };
  try{
    const cp=configPath(period);
    if (fs.existsSync(cp)) return JSON.parse(fs.readFileSync(cp,'utf8'));
  }catch{}
  return defaults;
}
function saveConfig(cfg,period){ try{ fs.writeFileSync(configPath(period), JSON.stringify(cfg,null,2)); }catch(e){ console.log(e); } }

function loadExcelData(period){
  try{
    let filePath=null;
    const tryNames=[
      period,
      period ? period.replace(/\s+/g,'-') : null,
      period ? period.replace(/-/g,' ') : null,
      period ? period.toUpperCase() : null,
      period ? period.replace(/\s+/g,'-').toUpperCase() : null,
      'master'
    ].filter(Boolean);
    for (const name of tryNames){
      const p = path.join(DATA_DIR, `${name}.xlsx`);
      if (fs.existsSync(p)){ filePath=p; break; }
    }
    if (!filePath) return null;
    const wb = XLSX.readFile(filePath);
    const edits = loadEdits(period);
    const semuakategori=[];
    const kategoriMap={};
    const kategoriList=[];
    const kategoriNamesSet=new Set();
    // cari sheet Semua Data
    for (const sheetName of wb.SheetNames){
      if (sheetName.toLowerCase().includes('semua')){
        const raw = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], {header:1});
        for (let i=1;i<raw.length;i++){
          const row=raw[i];
          if (!row || !row[0]) continue;
          const kode=String(row[0]??'').trim();
          if (!kode || kode==='GRAND TOTAL' || kode==='TOTAL') continue;
          const nama=String(row[1]??'').trim();
          if (!nama) continue;
          const key=kode+':'+nama;
          const e=edits[key]||{};
          const saldoAwal=e.saldoAwal!==undefined?e.saldoAwal:(parseFloat(row[2])||0);
          const debet=e.debet!==undefined?e.debet:(parseFloat(row[3])||0);
          const kredit=e.kredit!==undefined?e.kredit:(parseFloat(row[4])||0);
          const saldoAkhirVal=parseFloat(row[5])|| (saldoAwal+debet-kredit);
          const kategori=String(row[7]??row[6]??'').trim()||'Lain-lain';
          const obj={kode,nama,saldoAwal,debet,kredit,saldoAkhir:saldoAkhirVal,nettChange:debet-kredit,kategori};
          semuakategori.push(obj);
          kategoriNamesSet.add(kategori);
          if (!kategoriMap[kategori]) kategoriMap[kategori]=[];
          kategoriMap[kategori].push(obj);
        }
        break;
      }
    }
    if (semuakategori.length===0){
      for (const sheetName of wb.SheetNames){
        if (sheetName.toLowerCase().includes('semua')) continue;
        const raw = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], {header:1});
        // cari header KODE
        let headerIdx=-1;
        for (let i=0;i<Math.min(10,raw.length);i++){
          if (raw[i] && raw[i][0] && String(raw[i][0]).toLowerCase().includes('kode')){ headerIdx=i; break; }
        }
        if (headerIdx===-1) continue;
        for (let i=headerIdx+1;i<raw.length;i++){
          const row=raw[i];
          if (!row || !row[0]) continue;
          const kode=String(row[0]??'').trim();
          if (!kode) continue;
          const nama=String(row[1]??'').trim();
          const saldoAwal=parseFloat(row[2])||0;
          const debet=parseFloat(row[3])||0;
          const kredit=parseFloat(row[4])||0;
          const saldoAkhir=parseFloat(row[5])|| (saldoAwal+debet-kredit);
          const obj={kode,nama,saldoAwal,debet,kredit,saldoAkhir,nettChange:debet-kredit,kategori:sheetName};
          semuakategori.push(obj);
          if (!kategoriMap[sheetName]) kategoriMap[sheetName]=[];
          kategoriMap[sheetName].push(obj);
          kategoriNamesSet.add(sheetName);
        }
      }
    }
    for (const name of kategoriNamesSet){
      const list=kategoriMap[name]||[];
      const totalAwal=list.reduce((a,b)=>a+b.saldoAwal,0);
      const totalDebet=list.reduce((a,b)=>a+b.debet,0);
      const totalKredit=list.reduce((a,b)=>a+b.kredit,0);
      const totalAkhir=list.reduce((a,b)=>a+b.saldoAkhir,0);
      const totalNett=list.reduce((a,b)=>a+b.nettChange,0);
      kategoriList.push({nama:name, totalAwal, totalDebet, totalKredit, totalAkhir, totalNett, count:list.length});
    }
    return {semuakategori, kategoriList, kategoriMap};
  }catch(e){ console.log('loadExcel error',e.message, e.stack); return null; }
}

// API
app.get('/api/periods', (req,res)=>{ res.json(loadPeriods()); });

app.get('/api/data', (req,res)=>{
  let data=loadExcelData(req.query.period);
  if (!data) data=loadExcelData('master');
  if (!data) return res.status(404).json({error:'Belum upload Excel'});
  res.json(data.semuakategori);
});
app.get('/api/kategori', (req,res)=>{
  let data=loadExcelData(req.query.period);
  if (!data) data=loadExcelData('master');
  if (!data) return res.status(404).json({error:'Belum upload Excel'});
  res.json(data.kategoriList);
});
app.get('/api/ringkasan', (req,res)=>{
  let data=loadExcelData(req.query.period);
  if (!data) data=loadExcelData('master');
  if (!data) return res.status(404).json({error:'Belum upload Excel'});
  const grand={ totalAwal: data.semuakategori.reduce((a,b)=>a+b.saldoAwal,0), totalDebet: data.semuakategori.reduce((a,b)=>a+b.debet,0), totalKredit: data.semuakategori.reduce((a,b)=>a+b.kredit,0), totalAkhir: data.semuakategori.reduce((a,b)=>a+b.saldoAkhir,0) };
  res.json({ kategoriList:data.kategoriList, grandTotal:grand, totalAkun:data.semuakategori.length });
});
app.get('/api/neraca', (req,res)=>{
  let data=loadExcelData(req.query.period);
  if (!data) data=loadExcelData('master');
  if (!data) return res.status(404).json({error:'Belum upload Excel'});
  const cfg=loadConfig(data.kategoriList.map(k=>k.nama), req.query.period);
  const hitung=(section)=>{
    const names=cfg.neraca[section]||[];
    const items=[];
    for (const n of names){
      const kat=data.kategoriList.find(k=>k.nama===n);
      const list=data.kategoriMap[n]||[];
      for (const it of list){
        const saldo=req.query.mode==='netchange'? it.nettChange : it.saldoAkhir;
        if (Math.abs(saldo)>0.001) items.push({kode:it.kode,nama:it.nama,saldo});
      }
    }
    return items;
  };
  const total=(items)=>items.reduce((a,b)=>a+b.saldo,0);
  const lancar=hitung('AKTIVA LANCAR'); const lain=hitung('AKTIVA LAIN-LAIN'); const tetap=hitung('AKTIVA TETAP'); const hutang=hitung('HUTANG LANCAR'); const modal=hitung('MODAL');
  res.json({ aktivaLancar:lancar, totalAktivaLancar:total(lancar), aktivaLain:lain, totalAktivaLain:total(lain), aktivaTetap:tetap, totalAktivaTetap:total(tetap), totalAktiva:total(lancar)+total(lain)+total(tetap), hutangLancar:hutang, totalHutangLancar:total(hutang), modal, totalModal:total(modal), totalPasiva:Math.abs(total(hutang))+Math.abs(total(modal)), config:cfg });
});
app.get('/api/laba-rugi', (req,res)=>{
  let data=loadExcelData(req.query.period);
  if (!data) data=loadExcelData('master');
  if (!data) return res.status(404).json({error:'Belum upload Excel'});
  const cfg=loadConfig(data.kategoriList.map(k=>k.nama), req.query.period);
  const mode=req.query.mode||'netchange'; const isNett=mode==='netchange';
  const get=(n)=>data.kategoriList.find(k=>k.nama===n);
  const getSection=(s)=>(cfg.labaRugi[s]||[]).map(n=>get(n)).filter(Boolean);
  const katVal=(k)=>Math.abs(isNett?(k.totalNett||0):(k.totalAkhir||0));
  const itemVal=(i)=>Math.abs(isNett?(i.nettChange||0):(i.saldoAkhir||0));
  const penjualanList=getSection('PENJUALAN'); const hppList=getSection('HARGA POKOK PENJUALAN'); const bpList=getSection('BIAYA PENJUALAN'); const bauList=getSection('BIAYA ADM UMUM'); const pnoList=getSection('PENDAPATAN NON OPERASIONAL'); const bnoList=getSection('BIAYA NON OPERASIONAL');
  const penjualanNet=penjualanList.reduce((a,k)=>a+katVal(k),0); const hppVal=hppList.reduce((a,k)=>a+katVal(k),0); const labaKotor=penjualanNet-hppVal;
  const bpVal=bpList.reduce((a,k)=>a+katVal(k),0); const bauVal=bauList.reduce((a,k)=>a+katVal(k),0); const totalBiayaOps=bpVal+bauVal; const labaOperasional=labaKotor-totalBiayaOps;
  const pnoVal=pnoList.reduce((a,k)=>a+katVal(k),0); const bnoVal=bnoList.reduce((a,k)=>a+katVal(k),0); const labaKomersil=labaOperasional+(pnoVal-bnoVal);
  const pct=(v)=>penjualanNet?(v/penjualanNet*100).toFixed(2).replace('.',','):'0,00';
  res.json({ mode, period:req.query.period||'master', items:[
    {label:'PENJUALAN',section:true}, ...penjualanList.map(k=>({label:k.nama,value:katVal(k),pct:pct(katVal(k)),indent:true})),
    {label:'PENJUALAN BERSIH',value:penjualanNet,pct:pct(penjualanNet),bold:true}, {spacer:true},
    {label:'HARGA POKOK PENJUALAN',section:true}, ...hppList.map(k=>({label:k.nama,value:katVal(k),pct:pct(katVal(k)),indent:true})),
    {label:'LABA KOTOR',value:labaKotor,pct:pct(labaKotor),bold:true}, {spacer:true},
    {label:'LABA OPERASIONAL',value:labaOperasional,pct:pct(labaOperasional),bold:true}, {spacer:true},
    {label:'LABA KOMERSIL',value:labaKomersil,pct:pct(labaKomersil),bold:true}, {spacer:true},
    {label:'LABA SEBELUM PAJAK',value:labaKomersil,pct:pct(labaKomersil),bold:true,'class':'grand-total'}
  ]});
});
app.get('/api/config', (req,res)=>{
  let data=loadExcelData(req.query.period); if (!data) data=loadExcelData('master');
  const katNames=data?data.kategoriList.map(k=>k.nama):[];
  res.json({config:loadConfig(katNames, req.query.period), kategori:katNames, period:req.query.period||'master'});
});
app.post('/api/config', (req,res)=>{ saveConfig(req.body.config, req.query.period); res.json({message:'OK'}); });
app.get('/api/status', (req,res)=>{ res.json({adaData:excelExists(req.query.period), period:req.query.period||'master'}); });
app.put('/api/data/batch', (req,res)=>{
  let data=loadExcelData(req.query.period); if (!data) return res.status(404).json({error:'Belum upload Excel'});
  const edits=loadEdits(req.query.period);
  (req.body.items||[]).forEach(({kode,saldoAwal,debet,kredit})=>{
    const item=data.semuakategori.find(a=>a.kode===kode); if(!item) return;
    edits[item.kode+':'+item.nama]={saldoAwal,debet,kredit};
  });
  saveEdits(edits, req.query.period); res.json({message:'OK'});
});
app.put('/api/data/:kode', (req,res)=>{
  let data=loadExcelData(req.query.period); if (!data) return res.status(404).json({error:'Belum upload Excel'});
  const item=data.semuakategori.find(a=>a.kode===req.params.kode); if (!item) return res.status(404).json({error:'Akun tidak ditemukan'});
  const edits=loadEdits(req.query.period); edits[item.kode+':'+item.nama]={saldoAwal:req.body.saldoAwal,debet:req.body.debet,kredit:req.body.kredit};
  saveEdits(edits, req.query.period); res.json({message:'OK'});
});
app.post('/api/upload', upload.single('file'), (req,res)=>{
  try{
    if (!req.file) return res.status(400).json({error:'No file'});
    let rawPeriod=(req.body.period||'master').replace(/[^a-zA-Z0-9\-_ ]/g,'').trim()||'master';
    const period=rawPeriod.replace(/\s+/g,'-').toUpperCase();
    const label=req.body.label||rawPeriod;
    const destPath=excelPath(period);
    fs.writeFileSync(destPath, req.file.buffer);
    try{ const ep=editsPath(period); if (fs.existsSync(ep)) fs.unlinkSync(ep); }catch{}
    let periods=loadPeriods();
    if (!periods.find(p=>p.key===period)) periods.push({key:period,label});
    savePeriods(periods);
    const data=loadExcelData(period);
    if (!data || data.semuakategori.length===0) return res.status(400).json({error:'Excel tidak valid / sheet Semua Data kosong'});
    res.json({message:'OK', period, label, count:data.semuakategori.length});
  }catch(e){ console.log(e); res.status(500).json({error:e.message}); }
});
app.delete('/api/periods/:key', (req,res)=>{
  try{
    const rawKey=req.params.key;
    if (rawKey.toLowerCase()==='master') return res.status(400).json({error:'Master tidak boleh dihapus'});
    const norm=rawKey.replace(/\s+/g,'-').toUpperCase();
    const keysToTry=[rawKey, norm, rawKey.replace(/-/g,' ')];
    let deleted=[];
    for (const k of keysToTry){
      const pths=[path.join(DATA_DIR, `${k}.xlsx`), path.join(DATA_DIR, `${k.replace(/\s+/g,'-')}.xlsx`), editsPath(k), configPath(k)];
      for (const p of pths){ if (fs.existsSync(p)){ try{fs.unlinkSync(p); deleted.push(path.basename(p));}catch{} } }
    }
    let periods=loadPeriods();
    res.json({message:'Deleted', deletedFiles:deleted, periods});
  }catch(e){ res.status(500).json({error:e.message}); }
});

if (!IS_VERCEL){ const PORT=process.env.PORT||3001; app.listen(PORT,'0.0.0.0',()=>console.log('Server http://localhost:'+PORT)); }

module.exports=app;
