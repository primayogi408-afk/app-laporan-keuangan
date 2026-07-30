
const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const multer = require('multer');
const ExcelJS = require('exceljs');

const app = express();

// VERCEL FIX: pakai /tmp kalau di Vercel (read-only filesystem)
const IS_VERCEL = !!process.env.VERCEL;
const DATA_DIR = IS_VERCEL ? path.join('/tmp', 'data') : path.join(__dirname, 'data');
const CONFIG_PATH = path.join(DATA_DIR, 'config.json');
const PERIODS_PATH = path.join(DATA_DIR, 'periods.json');

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

const storage = multer.memoryStorage();
const upload = multer({ storage });

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.get('/api', (req, res) => {
  res.json({ status: 'OK', message: 'App Laporan Keuangan jalan!' });
});


function loadPeriods() {
  if (fs.existsSync(PERIODS_PATH)) {
    try { return JSON.parse(fs.readFileSync(PERIODS_PATH, 'utf8')); } catch { }
  }
  if (fs.existsSync(path.join(DATA_DIR, 'master.xlsx'))) {
    const def = [{ key: 'master', label: 'Utama' }];
    savePeriods(def);
    return def;
  }
  return [];
}
function savePeriods(list) {
  fs.writeFileSync(PERIODS_PATH, JSON.stringify(list, null, 2));
}
function excelPath(period) { return path.join(DATA_DIR, (period || 'master') + '.xlsx'); }
function editsPath(period) { return path.join(DATA_DIR, 'edits_' + (period || 'master') + '.json'); }
function configPath(period) { return path.join(DATA_DIR, 'config_' + (period || 'master') + '.json'); }

function defaultNeracaConfig(kategoriList) {
  const map = {
    'AKTIVA LANCAR': ['KAS','BANK','PIUTANG DAGANG','PIUTANG LAIN-LAIN','PERSEDIAAN BARANG','PPN LEBIH BAYAR','PPN YMH DITERIMA','PPN MASUKAN','PPH 22 DIBAYAR DI MUKA','PPH 23 DIBAYAR DI MUKA','PPH 25 DIBAYAR DI MUKA','BIAYA DIBAYAR DI MUKA','PENDAPATAN YMH DITERIMA','UANG MUKA PEMBELIAN','ASURANSI DIBAYAR DI MUKA','DEPOSITO'],
    'AKTIVA LAIN-LAIN': ['MODAL PENYERTAAN'],
    'AKTIVA TETAP': ['NILAI PEROLEHAN','AKUMULASI PENYUSUTAN'],
    'HUTANG LANCAR': ['HUTANG DAGANG','UANG MUKA PENJUALAN','HUTANG LAIN-LAIN','PPN YMH DIBAYAR','PPN KELUARAN','PPH 21 YMH DIBAYAR','PPH 23 YMH DIBAYAR','PPH 25 YMH DIBAYAR','PPH PS 4(2) YMH DIBAYAR','PPH 29 TERHUTANG','PPH 26 YMH DIBAYAR','BIAYA YMH DIBAYAR'],
    'MODAL': ['MODAL DISETOR','LABA DITAHAN','LABA ( RUGI ) S/D BULAN LALU','LABA ( RUGI ) BULAN INI','AGIO SAHAM']
  };
  if (kategoriList) {
    const allKategori = new Set(kategoriList);
    for (const section of Object.keys(map)) {
      map[section] = map[section].map(k => {
        const match = kategoriList.find(kl => kl.replace(/\s+/g, ' ').trim().toUpperCase() === k.replace(/\s+/g, ' ').trim().toUpperCase());
        return match || k;
      }).filter(k => allKategori.has(k));
    }
  }
  return map;
}
function defaultLabaRugiConfig(kategoriList) {
  const map = {
    'PENJUALAN': ['PENJUALAN'],
    'HARGA POKOK PENJUALAN': ['HARGA POKOK PENJUALAN'],
    'BIAYA PENJUALAN': ['BIAYA PENJUALAN'],
    'BIAYA ADM UMUM': ['BIAYA ADM UMUM'],
    'PENDAPATAN NON OPERASIONAL': ['PENDAPATAN NON OPERASIONAL'],
    'BIAYA NON OPERASIONAL': ['BIAYA NON OPERASIONAL']
  };
  if (kategoriList) {
    const allKategori = new Set(kategoriList);
    for (const section of Object.keys(map)) {
      map[section] = map[section].map(k => {
        const match = kategoriList.find(kl => kl.replace(/\s+/g, ' ').trim().toUpperCase() === k.replace(/\s+/g, ' ').trim().toUpperCase());
        return match || k;
      }).filter(k => allKategori.has(k));
    }
  }
  return map;
}
function isEmptyConfig(cfg) {
  for (const group of ['neraca', 'labaRugi']) {
    if (cfg[group]) {
      for (const section of Object.keys(cfg[group])) {
        if (cfg[group][section].length > 0) return false;
      }
    }
  }
  return true;
}
function loadConfig(kategoriList, period) {
  const p = period || 'master';
  const pcp = configPath(period);
  if (fs.existsSync(pcp)) {
    try {
      const cfg = JSON.parse(fs.readFileSync(pcp, 'utf8'));
      if (!isEmptyConfig(cfg)) return cfg;
    } catch {}
  }
  const legacy = path.join(__dirname, 'config.json');
  if (fs.existsSync(legacy) && !fs.existsSync(CONFIG_PATH)) {
    try {
      const cfg = JSON.parse(fs.readFileSync(legacy, 'utf8'));
      if (cfg.neraca && cfg.labaRugi) return cfg;
    } catch {}
  }
  if (fs.existsSync(CONFIG_PATH)) {
    try {
      const cfg = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));
      if (cfg.neraca && cfg.labaRugi && !isEmptyConfig(cfg)) {
        if (p !== 'master') saveConfig(cfg, period);
        return cfg;
      }
      const migrated = { neraca: cfg, labaRugi: defaultLabaRugiConfig(kategoriList) };
      saveConfig(migrated, period);
      return migrated;
    } catch {}
  }
  const def = { neraca: defaultNeracaConfig(kategoriList), labaRugi: defaultLabaRugiConfig(kategoriList) };
  saveConfig(def, period);
  return def;
}
function saveConfig(cfg, period) {
  const p = period || 'master';
  const dest = p === 'master' ? CONFIG_PATH : configPath(period);
  fs.writeFileSync(dest, JSON.stringify(cfg, null, 2));
}
function excelExists(period) { return fs.existsSync(excelPath(period)); }
function loadEdits(period) {
  const ep = editsPath(period);
  if (fs.existsSync(ep)) {
    try { return JSON.parse(fs.readFileSync(ep, 'utf8')); } catch {}
  }
  return {};
}
function saveEdits(edits, period) {
  fs.writeFileSync(editsPath(period), JSON.stringify(edits, null, 2));
}
function loadExcelData(period) {
  if (!excelExists(period)) return null;
  const wb = XLSX.readFile(excelPath(period));
  const edits = loadEdits(period);
  const semuakategori = [];
  const kategoriNames = [];
  for (const sheetName of wb.SheetNames) {
    if (sheetName === 'Semua Data') {
      const raw = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1 });
      for (let i = 1; i < raw.length; i++) {
        const row = raw[i];
        const kode = String(row[0] ?? '').trim();
        if (!kode || kode === 'GRAND TOTAL' || kode === 'TOTAL') continue;
        const nama = String(row[1] ?? '').trim();
        const key = kode + ':' + nama;
        const e = edits[key] || {};
        const saldoAwal = e.saldoAwal !== undefined ? e.saldoAwal : (parseFloat(row[2]) || 0);
        const debet = e.debet !== undefined ? e.debet : (parseFloat(row[3]) || 0);
        const kredit = e.kredit !== undefined ? e.kredit : (parseFloat(row[4]) || 0);
        semuakategori.push({ kode, nama, saldoAwal, debet, kredit, nettChange: debet - kredit, saldoAkhir: saldoAwal + debet - kredit, kategori: String(row[7] ?? '').trim() });
      }
      continue;
    }
    kategoriNames.push(sheetName);
  }
  if (semuakategori.length === 0) {
    for (const sheetName of wb.SheetNames) {
      if (sheetName === 'Semua Data') continue;
      const ws = wb.Sheets[sheetName];
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });
      let headerRow = -1;
      for (let i = 0; i < raw.length; i++) { if (raw[i][0] === 'KODE') { headerRow = i; break; } }
      if (headerRow === -1) continue;
      for (let i = headerRow + 1; i < raw.length; i++) {
        const row = raw[i];
        const kode = String(row[0] ?? '').trim();
        if (!kode || kode === 'GRAND TOTAL' || kode === 'TOTAL') continue;
        const nama = String(row[1] ?? '').trim();
        const key = kode + ':' + nama;
        const e = edits[key] || {};
        const saldoAwal = e.saldoAwal !== undefined ? e.saldoAwal : (parseFloat(row[2]) || 0);
        const debet = e.debet !== undefined ? e.debet : (parseFloat(row[3]) || 0);
        const kredit = e.kredit !== undefined ? e.kredit : (parseFloat(row[4]) || 0);
        semuakategori.push({ kode, nama, saldoAwal, debet, kredit, nettChange: debet - kredit, saldoAkhir: saldoAwal + debet - kredit, kategori: sheetName });
      }
    }
  }
  const kategoriMap = {};
  for (const item of semuakategori) { const kat = item.kategori || 'LAIN-LAIN'; if (!kategoriMap[kat]) kategoriMap[kat] = []; kategoriMap[kat].push(item); }
  const kategoriList = Object.entries(kategoriMap).map(([nama, items]) => ({ nama, totalAwal: items.reduce((a,b)=>a+b.saldoAwal,0), totalAkhir: items.reduce((a,b)=>a+b.saldoAkhir,0), totalDebet: items.reduce((a,b)=>a+b.debet,0), totalKredit: items.reduce((a,b)=>a+b.kredit,0), totalNett: items.reduce((a,b)=>a+b.nettChange,0), jumlahItem: items.length }));
  return { semuakategori, kategoriList, kategoriMap, kategoriNames };
}

app.post('/api/upload', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Tidak ada file' });
  try {
    const period = req.body.period ? req.body.period.replace(/[^a-zA-Z0-9\-_ ]/g, '').trim() : 'master';
    const label = req.body.label || period;
    const destPath = excelPath(period);
    fs.writeFileSync(destPath, req.file.buffer);
    const data = loadExcelData(period);
    if (!data || data.semuakategori.length === 0) return res.status(400).json({ error: 'File Excel tidak memiliki data yang sesuai' });
    const periods = loadPeriods();
    const existing = periods.find(p => p.key === period);
    if (existing) existing.label = label; else periods.push({ key: period, label });
    savePeriods(periods);
    const katNames = data.kategoriList.map(k => k.nama);
    const cfg = loadConfig(katNames, period);
    for (const group of ['neraca', 'labaRugi']) { if (cfg[group]) { for (const section of Object.keys(cfg[group])) { cfg[group][section] = cfg[group][section].filter(k => katNames.includes(k)); } } }
    saveConfig(cfg, period);
    res.json({ message: 'OK', total: data.semuakategori.length, period, label });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/periods', (req, res) => { res.json(loadPeriods()); });
app.get('/api/data', (req, res) => { const data = loadExcelData(req.query.period); if (!data) return res.status(404).json({ error: 'Belum upload Excel' }); res.json(data.semuakategori); });
app.get('/api/kategori', (req, res) => { const data = loadExcelData(req.query.period); if (!data) return res.status(404).json({ error: 'Belum upload Excel' }); res.json(data.kategoriList); });
app.get('/api/kategori/:nama', (req, res) => { const data = loadExcelData(req.query.period); if (!data) return res.status(404).json({ error: 'Belum upload Excel' }); const items = data.kategoriMap[req.params.nama]; if (!items) return res.status(404).json({ error: 'Kategori tidak ditemukan' }); res.json({ kategori: data.kategoriList.find(k=>k.nama===req.params.nama), items }); });
app.get('/api/ringkasan', (req, res) => { const data = loadExcelData(req.query.period); if (!data) return res.status(404).json({ error: 'Belum upload Excel' }); const cfg = loadConfig(data.kategoriList.map(k=>k.nama), req.query.period); const hitung=(s)=>{ const names=cfg.neraca[s]||[]; return data.kategoriList.filter(k=>names.includes(k.nama)); }; const sum=(items)=>items.reduce((a,b)=>a+b.totalAkhir,0); res.json({ kategori: data.kategoriList, totalAsetLancar: sum(hitung('AKTIVA LANCAR')), totalAsetLain: sum(hitung('AKTIVA LAIN-LAIN')), totalAsetTetap: sum(hitung('AKTIVA TETAP')), totalHutangLancar: sum(hitung('HUTANG LANCAR')), totalModal: sum(hitung('MODAL')), totalAkun: data.semuakategori.length, semuaData: data.semuakategori, kategoriMap: data.kategoriMap, period: req.query.period||'master' }); });
app.get('/api/neraca', (req, res) => {
  const data = loadExcelData(req.query.period); if (!data) return res.status(404).json({ error: 'Belum upload Excel' });
  const cfg = loadConfig(data.kategoriList.map(k=>k.nama), req.query.period);
  const hitung=(s)=>{ const names=cfg.neraca[s]||[]; return data.kategoriList.filter(k=>names.includes(k.nama)).map(k=>({ nama:k.nama, saldo:k.totalAkhir })); };
  const total=(items)=>items.reduce((a,b)=>a+b.saldo,0);
  const lancar=hitung('AKTIVA LANCAR'); const lain=hitung('AKTIVA LAIN-LAIN'); const tetap=hitung('AKTIVA TETAP'); const hutang=hitung('HUTANG LANCAR'); const modal=hitung('MODAL');
  res.json({ aktivaLancar:lancar, totalAktivaLancar:total(lancar), aktivaLain:lain, totalAktivaLain:total(lain), aktivaTetap:tetap, totalAktivaTetap:total(tetap), totalAktiva:total(lancar)+total(lain)+total(tetap), hutangLancar:hutang, totalHutangLancar:total(hutang), modal, totalModal:total(modal), totalPasiva:Math.abs(total(hutang))+Math.abs(total(modal)), config:cfg });
});
app.get('/api/laba-rugi', (req, res) => {
  const data = loadExcelData(req.query.period); if (!data) return res.status(404).json({ error: 'Belum upload Excel' });
  const cfg = loadConfig(data.kategoriList.map(k=>k.nama), req.query.period);
  const mode = req.query.mode || 'netchange'; const isNett = mode==='netchange';
  const get=(n)=>data.kategoriList.find(k=>k.nama===n); const getSection=(s)=>(cfg.labaRugi[s]||[]).map(n=>get(n)).filter(Boolean);
  const katVal=(k)=>Math.abs(isNett ? (k.totalNett||0) : (k.totalAkhir||0)); const itemVal=(i)=>Math.abs(isNett ? (i.nettChange||0) : (i.saldoAkhir||0));
  const penjualanList=getSection('PENJUALAN'); const hppList=getSection('HARGA POKOK PENJUALAN'); const bpList=getSection('BIAYA PENJUALAN'); const bauList=getSection('BIAYA ADM UMUM'); const pnoList=getSection('PENDAPATAN NON OPERASIONAL'); const bnoList=getSection('BIAYA NON OPERASIONAL');
  const penjualanNet=penjualanList.reduce((a,k)=>a+katVal(k),0); const hppVal=hppList.reduce((a,k)=>a+katVal(k),0); const labaKotor=penjualanNet-hppVal;
  const bpVal=bpList.reduce((a,k)=>a+katVal(k),0); const bauVal=bauList.reduce((a,k)=>a+katVal(k),0); const totalBiayaOps=bpVal+bauVal; const labaOperasional=labaKotor-totalBiayaOps;
  const pnoVal=pnoList.reduce((a,k)=>a+katVal(k),0); const bnoVal=bnoList.reduce((a,k)=>a+katVal(k),0); const totalPendNonOps=pnoVal-bnoVal; const labaKomersil=labaOperasional+totalPendNonOps;
  const allPnoItems=[]; pnoList.forEach(k=>{(data.kategoriMap[k.nama]||[]).forEach(i=>allPnoItems.push(i));}); const allBnoItems=[]; bnoList.forEach(k=>{(data.kategoriMap[k.nama]||[]).forEach(i=>allBnoItems.push(i));});
  let koreksiPositifItems = allBnoItems.filter(i=>[93,96].includes(parseInt(i.kode))); let koreksiNegatifItems = allPnoItems.filter(i=>[71,74,78].includes(parseInt(i.kode)));
  const koreksiPositif=koreksiPositifItems.reduce((a,i)=>a+itemVal(i),0); const koreksiNegatif=koreksiNegatifItems.reduce((a,i)=>a+itemVal(i),0);
  const totalKoreksiFiskal=koreksiPositif-koreksiNegatif; const labaSebelumPajak=labaKomersil+totalKoreksiFiskal;
  const pct=(v)=>penjualanNet ? (v/penjualanNet*100).toFixed(2).replace('.',',') : '0,00';
  res.json({ mode, period: req.query.period||'master', items: [ { label:'PENJUALAN', section:true }, ...penjualanList.map(k=>({ label:k.nama, value:katVal(k), pct:pct(katVal(k)), indent:true })), { label:'PENJUALAN BERSIH', value:penjualanNet, pct:pct(penjualanNet), bold:true }, { spacer:true }, { label:'HARGA POKOK PENJUALAN', section:true }, ...hppList.map(k=>({ label:k.nama, value:katVal(k), pct:pct(katVal(k)), indent:true })), { label:'LABA KOTOR', value:labaKotor, pct:pct(labaKotor), bold:true }, { spacer:true }, { label:'LABA OPERASIONAL', value:labaOperasional, pct:pct(labaOperasional), bold:true }, { spacer:true }, { label:'LABA ( RUGI ) KOMERSIL', value:labaKomersil, pct:pct(labaKomersil), bold:true }, { spacer:true }, { label:'LABA ( RUGI ) SEBELUM PAJAK', value:labaSebelumPajak, pct:pct(labaSebelumPajak), bold:true, 'class':'grand-total' } ] });
});
app.get('/api/config', (req, res) => { const data=loadExcelData(req.query.period); const katNames=data?data.kategoriList.map(k=>k.nama):[]; res.json({ config: loadConfig(katNames, req.query.period), kategori: katNames, period: req.query.period||'master' }); });
app.post('/api/config', (req, res) => { saveConfig(req.body.config, req.query.period); res.json({ message:'OK' }); });
app.get('/api/status', (req, res) => { res.json({ adaData: excelExists(req.query.period), file: (req.query.period||'master')+'.xlsx', period: req.query.period||'master' }); });
app.put('/api/data/batch', (req, res) => { const data=loadExcelData(req.query.period); if(!data) return res.status(404).json({ error:'Belum upload Excel' }); const edits=loadEdits(req.query.period); (req.body.items||[]).forEach(({kode,saldoAwal,debet,kredit})=>{ const item=data.semuakategori.find(a=>a.kode===kode); if(!item) return; edits[item.kode+':'+item.nama]={saldoAwal,debet,kredit}; }); saveEdits(edits, req.query.period); res.json({ message:'OK' }); });
app.put('/api/data/:kode', (req, res) => { const data=loadExcelData(req.query.period); if(!data) return res.status(404).json({ error:'Belum upload Excel' }); const item=data.semuakategori.find(a=>a.kode===req.params.kode); if(!item) return res.status(404).json({ error:'Akun tidak ditemukan' }); const edits=loadEdits(req.query.period); edits[item.kode+':'+item.nama]={ saldoAwal:req.body.saldoAwal, debet:req.body.debet, kredit:req.body.kredit }; saveEdits(edits, req.query.period); res.json({ message:'OK' }); });
app.get('/api/neraca/download', async (req,res)=>{ res.status(501).json({ error:'Download neraca pakai mode lokal, di Vercel gunakan export dari frontend' }); });
app.get('/api/laba-rugi/download', async (req,res)=>{ res.status(501).json({ error:'Download laba-rugi pakai mode lokal' }); });

if (!IS_VERCEL) {
  const PORT = process.env.PORT || 3001;
  app.listen(PORT, '0.0.0.0', () => console.log(`Server: http://localhost:${PORT}`));
}

module.exports = app;
