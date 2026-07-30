const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const multer = require('multer');
const ExcelJS = require('exceljs');

const app = express();
const PORT = process.env.PORT || 3001;
const CONFIG_PATH = path.join(__dirname, 'config.json');
const DATA_DIR = path.join(__dirname, 'data');
const PERIODS_PATH = path.join(DATA_DIR, 'periods.json');

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR);

const storage = multer.diskStorage({
  destination: DATA_DIR,
  filename: (req, file, cb) => {
    const period = req.body.period ? req.body.period.replace(/[^a-zA-Z0-9\-_ ]/g, '').trim() : 'master';
    cb(null, period + '.xlsx');
  }
});
const upload = multer({ storage });

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

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

function excelPath(period) {
  return path.join(DATA_DIR, (period || 'master') + '.xlsx');
}

function editsPath(period) {
  return path.join(DATA_DIR, 'edits_' + (period || 'master') + '.json');
}

function configPath(period) {
  return path.join(DATA_DIR, 'config_' + (period || 'master') + '.json');
}

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
    } catch { }
  }
  if (fs.existsSync(CONFIG_PATH)) {
    try {
      const cfg = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));
      if (cfg.neraca && cfg.labaRugi && !isEmptyConfig(cfg)) {
        if (p !== 'master') saveConfig(cfg, period);
        return cfg;
      }
      if (cfg.neraca && cfg.labaRugi) {
        // Shared config is empty, merge with defaults
        const defN = defaultNeracaConfig(kategoriList);
        const defL = defaultLabaRugiConfig(kategoriList);
        for (const section of Object.keys(defN)) {
          if (!cfg.neraca[section] || cfg.neraca[section].length === 0) cfg.neraca[section] = defN[section];
        }
        for (const section of Object.keys(defL)) {
          if (!cfg.labaRugi[section] || cfg.labaRugi[section].length === 0) cfg.labaRugi[section] = defL[section];
        }
        saveConfig(cfg, period);
        return cfg;
      }
      const migrated = { neraca: cfg, labaRugi: defaultLabaRugiConfig(kategoriList) };
      saveConfig(migrated, period);
      return migrated;
    } catch { }
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

function excelExists(period) {
  return fs.existsSync(excelPath(period));
}

function loadEdits(period) {
  const ep = editsPath(period);
  if (fs.existsSync(ep)) {
    try { return JSON.parse(fs.readFileSync(ep, 'utf8')); } catch { }
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
        semuakategori.push({
          kode, nama,
          saldoAwal, debet, kredit,
          nettChange: debet - kredit,
          saldoAkhir: saldoAwal + debet - kredit,
          kategori: String(row[7] ?? '').trim()
        });
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
      let foundHeader = false;
      let headerRow = -1;
      for (let i = 0; i < raw.length; i++) {
        const r = raw[i];
        if (r[0] === 'KODE' || r[0] === 'KODE') {
          foundHeader = true;
          headerRow = i;
          break;
        }
      }
      if (!foundHeader) continue;
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
        semuakategori.push({
          kode, nama,
          saldoAwal, debet, kredit,
          nettChange: debet - kredit,
          saldoAkhir: saldoAwal + debet - kredit,
          kategori: sheetName
        });
      }
    }
  }

  const kategoriMap = {};
  for (const item of semuakategori) {
    const kat = item.kategori || 'LAIN-LAIN';
    if (!kategoriMap[kat]) kategoriMap[kat] = [];
    kategoriMap[kat].push(item);
  }

  const kategoriList = Object.entries(kategoriMap).map(([nama, items]) => ({
    nama,
    totalAwal: items.reduce((a, b) => a + b.saldoAwal, 0),
    totalAkhir: items.reduce((a, b) => a + b.saldoAkhir, 0),
    totalDebet: items.reduce((a, b) => a + b.debet, 0),
    totalKredit: items.reduce((a, b) => a + b.kredit, 0),
    totalNett: items.reduce((a, b) => a + b.nettChange, 0),
    jumlahItem: items.length
  }));

  return { semuakategori, kategoriList, kategoriMap, kategoriNames };
}

app.post('/api/upload', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Tidak ada file' });
  try {
    const period = req.body.period ? req.body.period.replace(/[^a-zA-Z0-9\-_ ]/g, '').trim() : 'master';
    const label = req.body.label || period;
    const data = loadExcelData(period);
    if (!data || data.semuakategori.length === 0) {
      return res.status(400).json({ error: 'File Excel tidak memiliki data yang sesuai' });
    }
    const periods = loadPeriods();
    const existing = periods.find(p => p.key === period);
    if (existing) {
      existing.label = label;
    } else {
      periods.push({ key: period, label });
    }
    savePeriods(periods);
    const katNames = data.kategoriList.map(k => k.nama);
    const cfg = loadConfig(katNames, period);
    for (const group of ['neraca', 'labaRugi']) {
      if (cfg[group]) {
        for (const section of Object.keys(cfg[group])) {
          cfg[group][section] = cfg[group][section].filter(k => katNames.includes(k));
        }
      }
    }
    saveConfig(cfg, period);
    res.json({ message: 'OK', total: data.semuakategori.length, period, label });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/periods', (req, res) => {
  res.json(loadPeriods());
});

app.get('/api/data', (req, res) => {
  const period = req.query.period;
  const data = loadExcelData(period);
  if (!data) return res.status(404).json({ error: 'Belum upload Excel' });
  res.json(data.semuakategori);
});

app.get('/api/kategori', (req, res) => {
  const period = req.query.period;
  const data = loadExcelData(period);
  if (!data) return res.status(404).json({ error: 'Belum upload Excel' });
  res.json(data.kategoriList);
});

app.get('/api/kategori/:nama', (req, res) => {
  const period = req.query.period;
  const data = loadExcelData(period);
  if (!data) return res.status(404).json({ error: 'Belum upload Excel' });
  const nama = req.params.nama;
  const items = data.kategoriMap[nama];
  if (!items) return res.status(404).json({ error: 'Kategori tidak ditemukan' });
  const kategori = data.kategoriList.find(k => k.nama === nama);
  res.json({ kategori, items });
});

app.get('/api/ringkasan', (req, res) => {
  const period = req.query.period;
  const data = loadExcelData(period);
  if (!data) return res.status(404).json({ error: 'Belum upload Excel' });
  const cfg = loadConfig(data.kategoriList.map(k => k.nama), period);

  const hitung = (section) => {
    const names = cfg.neraca[section] || [];
    return data.kategoriList.filter(k => names.includes(k.nama));
  };

  const sum = (items) => items.reduce((a, b) => a + b.totalAkhir, 0);

  res.json({
    kategori: data.kategoriList,
    totalAsetLancar: sum(hitung('AKTIVA LANCAR')),
    totalAsetLain: sum(hitung('AKTIVA LAIN-LAIN')),
    totalAsetTetap: sum(hitung('AKTIVA TETAP')),
    totalHutangLancar: sum(hitung('HUTANG LANCAR')),
    totalModal: sum(hitung('MODAL')),
    totalAkun: data.semuakategori.length,
    semuaData: data.semuakategori,
    kategoriMap: data.kategoriMap,
    period: period || 'master'
  });
});

app.get('/api/neraca', (req, res) => {
  const period = req.query.period;
  const data = loadExcelData(period);
  if (!data) return res.status(404).json({ error: 'Belum upload Excel' });
  const cfg = loadConfig(data.kategoriList.map(k => k.nama), period);

  const hitung = (section) => {
    const names = cfg.neraca[section] || [];
    return data.kategoriList.filter(k => names.includes(k.nama)).map(k => ({
      nama: k.nama,
      saldo: k.totalAkhir
    }));
  };

  const total = (items) => items.reduce((a, b) => a + b.saldo, 0);

  const lancar = hitung('AKTIVA LANCAR');
  const lain = hitung('AKTIVA LAIN-LAIN');
  const tetap = hitung('AKTIVA TETAP');
  const hutang = hitung('HUTANG LANCAR');
  const modal = hitung('MODAL');

  const totalLancar = total(lancar);
  const totalLain = total(lain);
  const totalTetap = total(tetap);
  const totalHutang = total(hutang);
  const totalModal = total(modal);

  const totalAktiva = totalLancar + totalLain + totalTetap;
  const totalPasiva = Math.abs(totalHutang) + Math.abs(totalModal);

  res.json({
    aktivaLancar: lancar, totalAktivaLancar: totalLancar,
    aktivaLain: lain, totalAktivaLain: totalLain,
    aktivaTetap: tetap, totalAktivaTetap: totalTetap,
    totalAktiva,
    hutangLancar: hutang, totalHutangLancar: totalHutang,
    modal, totalModal: totalModal,
    totalPasiva,
    config: cfg
  });
});

app.get('/api/laba-rugi', (req, res) => {
  const period = req.query.period;
  const data = loadExcelData(period);
  if (!data) return res.status(404).json({ error: 'Belum upload Excel' });
  const cfg = loadConfig(data.kategoriList.map(k => k.nama), period);
  const mode = req.query.mode || 'netchange';

  const isNett = mode === 'netchange';
  const kat = data.kategoriList;
  const get = (name) => kat.find(k => k.nama === name);
  const getSection = (section) => (cfg.labaRugi[section] || []).map(n => get(n)).filter(Boolean);
  const katVal = (k) => Math.abs(isNett ? (k.totalNett || 0) : (k.totalAkhir || 0));
  const itemVal = (i) => Math.abs(isNett ? (i.nettChange || 0) : (i.saldoAkhir || 0));

  const penjualanList = getSection('PENJUALAN');
  const hppList = getSection('HARGA POKOK PENJUALAN');
  const bpList = getSection('BIAYA PENJUALAN');
  const bauList = getSection('BIAYA ADM UMUM');
  const pnoList = getSection('PENDAPATAN NON OPERASIONAL');
  const bnoList = getSection('BIAYA NON OPERASIONAL');

  const penjualanNet = penjualanList.reduce((a, k) => a + katVal(k), 0);
  const hppVal = hppList.reduce((a, k) => a + katVal(k), 0);
  const labaKotor = penjualanNet - hppVal;
  const bpVal = bpList.reduce((a, k) => a + katVal(k), 0);
  const bauVal = bauList.reduce((a, k) => a + katVal(k), 0);
  const totalBiayaOps = bpVal + bauVal;
  const labaOperasional = labaKotor - totalBiayaOps;

  const pnoVal = pnoList.reduce((a, k) => a + katVal(k), 0);
  const bnoVal = bnoList.reduce((a, k) => a + katVal(k), 0);
  const totalPendNonOps = pnoVal - bnoVal;
  const labaKomersil = labaOperasional + totalPendNonOps;

  const allPnoItems = [];
  pnoList.forEach(k => { (data.kategoriMap[k.nama] || []).forEach(i => allPnoItems.push(i)); });
  const allBnoItems = [];
  bnoList.forEach(k => { (data.kategoriMap[k.nama] || []).forEach(i => allBnoItems.push(i)); });

  const koreksiPositifCats = (cfg.labaRugi['KOREKSI POSITIF'] || []);
  const koreksiNegatifCats = (cfg.labaRugi['KOREKSI NEGATIF'] || []);

  let koreksiPositifItems = [];
  let koreksiNegatifItems = [];

  if (koreksiPositifCats.length) {
    koreksiPositifCats.forEach(catName => {
      (data.kategoriMap[catName] || []).forEach(i => koreksiPositifItems.push(i));
    });
  } else {
    koreksiPositifItems = allBnoItems.filter(i => [93,96].includes(parseInt(i.kode)));
  }

  if (koreksiNegatifCats.length) {
    koreksiNegatifCats.forEach(catName => {
      (data.kategoriMap[catName] || []).forEach(i => koreksiNegatifItems.push(i));
    });
  } else {
    koreksiNegatifItems = allPnoItems.filter(i => [71,74,78].includes(parseInt(i.kode)));
  }

  const koreksiPositif = koreksiPositifItems.reduce((a, i) => a + itemVal(i), 0);
  const koreksiNegatif = koreksiNegatifItems.reduce((a, i) => a + itemVal(i), 0);
  const totalKoreksiFiskal = koreksiPositif - koreksiNegatif;
  const labaSebelumPajak = labaKomersil + totalKoreksiFiskal;

  const pct = (v) => penjualanNet ? (v / penjualanNet * 100).toFixed(2).replace('.', ',') : '0,00';

  const penjualanItems = penjualanList.length ? penjualanList.map(k => ({ label: k.nama, value: katVal(k), pct: pct(katVal(k)), indent: true })) : [];
  const hppItems = hppList.length ? hppList.map(k => ({ label: k.nama, value: katVal(k), pct: pct(katVal(k)), indent: true })) : [];

  res.json({
    mode, period: period || 'master',
    items: [
      { label: 'PENJUALAN', section: true },
      ...penjualanItems,
      { label: 'RETUR PENJUALAN', value: 0, pct: '-', bold: false },
      { label: 'PENJUALAN BERSIH', value: penjualanNet, pct: pct(penjualanNet), bold: true },
      { spacer: true },
      { label: 'HARGA POKOK PENJUALAN', section: true },
      ...hppItems,
      { label: 'LABA KOTOR', value: labaKotor, pct: pct(labaKotor), bold: true },
      { spacer: true },
      { label: 'BIAYA OPERASIONAL', section: true },
      ...bpList.map(k => ({ label: k.nama, value: katVal(k), pct: pct(katVal(k)), indent: true })),
      ...bauList.map(k => ({ label: k.nama, value: katVal(k), pct: pct(katVal(k)), indent: true })),
      { label: 'TOTAL BIAYA OPERASIONAL', value: totalBiayaOps, pct: pct(totalBiayaOps), bold: true, indent: true },
      { spacer: true },
      { label: 'LABA OPERASIONAL', value: labaOperasional, pct: pct(labaOperasional), bold: true },
      { spacer: true },
      { label: 'PENDAPATAN DAN (BIAYA NON OPERASIONAL)', section: true },
      ...pnoList.map(k => ({ label: k.nama, value: katVal(k), pct: pct(katVal(k)), indent: true })),
      ...bnoList.map(k => ({ label: '( ' + k.nama + ' )', value: -katVal(k), pct: '(' + pct(katVal(k)) + ')', indent: true })),
      { label: 'TOTAL PENDAPATAN DAN (BIAYA NON OPERASIONAL)', value: totalPendNonOps, pct: pct(totalPendNonOps), bold: true, indent: true },
      { spacer: true },
      { label: 'LABA ( RUGI ) KOMERSIL', value: labaKomersil, pct: pct(labaKomersil), bold: true },
      { spacer: true },
      { label: 'KOREKSI POSITIF', section: true },
      ...koreksiPositifItems.map(i => ({ label: '**   ' + i.nama, value: itemVal(i), pct: pct(itemVal(i)), indent: true })),
      { label: 'TOTAL KOREKSI POSITIF', value: koreksiPositif, pct: pct(koreksiPositif), bold: true, indent: true },
      { spacer: true },
      { label: 'KOREKSI NEGATIF', section: true },
      ...koreksiNegatifItems.map(i => ({ label: '**   ' + i.nama, value: -itemVal(i), pct: '(' + pct(itemVal(i)) + ')', indent: true })),
      { label: 'TOTAL KOREKSI NEGATIF', value: -koreksiNegatif, pct: '(' + pct(koreksiNegatif) + ')', bold: true, indent: true },
      { spacer: true },
      { label: 'TOTAL KOREKSI FISKAL (KOREKSI POSITIF + KOREKSI NEGATIF)', value: totalKoreksiFiskal, pct: (totalKoreksiFiskal >= 0 ? '' : '(') + pct(Math.abs(totalKoreksiFiskal)) + (totalKoreksiFiskal >= 0 ? '' : ')'), bold: true },
      { spacer: true },
      { label: 'LABA ( RUGI ) SEBELUM PAJAK', value: labaSebelumPajak, pct: pct(labaSebelumPajak), bold: true, 'class': 'grand-total' }
    ]
  });
});

app.get('/api/config', (req, res) => {
  const period = req.query.period;
  const data = loadExcelData(period);
  const katNames = data ? data.kategoriList.map(k => k.nama) : [];
  const cfg = loadConfig(katNames, period);
  res.json({ config: cfg, kategori: katNames, period: period || 'master' });
});

app.post('/api/config', (req, res) => {
  const cfg = req.body.config;
  const period = req.query.period;
  saveConfig(cfg, period);
  res.json({ message: 'OK' });
});

app.get('/api/status', (req, res) => {
  const period = req.query.period;
  res.json({ adaData: excelExists(period), file: (period || 'master') + '.xlsx', period: period || 'master' });
});

app.put('/api/data/batch', (req, res) => {
  const period = req.query.period;
  const data = loadExcelData(period);
  if (!data) return res.status(404).json({ error: 'Belum upload Excel' });
  const edits = loadEdits(period);
  const { items } = req.body;
  if (!Array.isArray(items)) return res.status(400).json({ error: 'items harus array' });
  items.forEach(({ kode, saldoAwal, debet, kredit }) => {
    const item = data.semuakategori.find(a => a.kode === kode);
    if (!item) return;
    const key = item.kode + ':' + item.nama;
    edits[key] = { saldoAwal, debet, kredit };
  });
  saveEdits(edits, period);
  res.json({ message: 'OK', total: items.length });
});

app.put('/api/data/:kode', (req, res) => {
  const period = req.query.period;
  const data = loadExcelData(period);
  if (!data) return res.status(404).json({ error: 'Belum upload Excel' });
  const { kode } = req.params;
  const { saldoAwal, debet, kredit } = req.body;
  const item = data.semuakategori.find(a => a.kode === kode);
  if (!item) return res.status(404).json({ error: 'Akun tidak ditemukan' });
  const edits = loadEdits(period);
  const key = item.kode + ':' + item.nama;
  edits[key] = {
    saldoAwal: saldoAwal !== undefined ? saldoAwal : item.saldoAwal,
    debet: debet !== undefined ? debet : item.debet,
    kredit: kredit !== undefined ? kredit : item.kredit
  };
  saveEdits(edits, period);
  res.json({ message: 'OK' });
});

function rp(n) {
  if (n === undefined || n === null) return '';
  const abs = Math.abs(n);
  const f = new Intl.NumberFormat('id-ID').format(Math.round(abs));
  return n < 0 ? '(' + f + ')' : f;
}

async function downloadNeraca(req, res) {
  const period = req.query.period;
  const data = loadExcelData(period);
  if (!data) return res.status(404).json({ error: 'Belum upload Excel' });
  const cfg = loadConfig(data.kategoriList.map(k => k.nama), period);

  const hitung = (section) => {
    const names = cfg.neraca[section] || [];
    return data.kategoriList.filter(k => names.includes(k.nama)).map(k => ({ nama: k.nama, saldo: k.totalAkhir }));
  };
  const total = (items) => items.reduce((a, b) => a + b.saldo, 0);

  const lancar = hitung('AKTIVA LANCAR');
  const lain = hitung('AKTIVA LAIN-LAIN');
  const tetap = hitung('AKTIVA TETAP');
  const hutang = hitung('HUTANG LANCAR');
  const modal = hitung('MODAL');

  const totalLancar = total(lancar);
  const totalLain = total(lain);
  const totalTetap = total(tetap);
  const totalHutang = total(hutang);
  const totalModal = total(modal);
  const totalAktiva = totalLancar + totalLain + totalTetap;
  const totalPasiva = Math.abs(totalHutang) + Math.abs(totalModal);
  const pctA = (v) => totalAktiva ? (v / totalAktiva * 100).toFixed(2).replace('.', ',') : '0,00';
  const pctP = (v) => totalPasiva ? (Math.abs(v) / totalPasiva * 100).toFixed(2).replace('.', ',') : '0,00';

  const rows = [];
  const add = (left, lval, lpct, right, rval, rpct, cls) => rows.push({ left, lval: lval || '', lpct: lpct || '', right: right || '', rval: rval || '', rpct: rpct || '', cls: cls || '' });

  const sectionMap = [
    { left: 'AKTIVA LANCAR', right: 'HUTANG LANCAR', leftItems: lancar, rightItems: hutang, leftTotal: totalLancar, rightTotal: totalHutang, pctLeft: pctA, pctRight: pctP },
    { left: 'AKTIVA LAIN-LAIN', right: 'MODAL', leftItems: lain, rightItems: modal, leftTotal: totalLain, rightTotal: totalModal, pctLeft: pctA, pctRight: pctP },
    { left: 'AKTIVA TETAP', right: '', leftItems: tetap, rightItems: [], leftTotal: totalTetap, rightTotal: 0, pctLeft: pctA, pctRight: null }
  ];

  sectionMap.forEach((sec, idx) => {
    if (idx > 0) add('', '', '', '', '', '', 'spacer');
    add(sec.left, '', '', sec.right || '', '', '', 'section');
    const m = Math.max(sec.leftItems.length, sec.rightItems.length || 1);
    for (let i = 0; i < m; i++) {
      const l = sec.leftItems[i];
      const r = sec.rightItems[i];
      add(
        l ? l.nama : '', l ? rp(l.saldo) : '', l ? sec.pctLeft(l.saldo) : '',
        r ? r.nama : '', r ? rp(Math.abs(r.saldo)) : '', r ? sec.pctRight(r.saldo) : '',
        ''
      );
    }
    add('TOTAL ' + sec.left, rp(sec.leftTotal), sec.pctLeft(sec.leftTotal),
      sec.right ? 'TOTAL ' + sec.right : '', sec.right ? rp(Math.abs(sec.rightTotal)) : '', sec.right ? sec.pctRight(sec.rightTotal) : '',
      'total');
  });

  add('', '', '', '', '', '', 'spacer');
  add('TOTAL AKTIVA', rp(totalAktiva), '100,00', 'TOTAL HUTANG + MODAL', rp(totalPasiva), '100,00', 'grand');

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Neraca');
  ws.columns = [
    { width: 28 }, { width: 18 }, { width: 10 },
    { width: 28 }, { width: 18 }, { width: 10 }
  ];

  const titleFont = { name: 'Calibri', size: 14, bold: true };
  const headerFont = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A1A2E' } };
  const sectionFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE3F2FD' } };
  const totalFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8EAF6' } };
  const grandFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC5CAE9' } };
  const thinBorder = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
  const topBorder = { top: { style: 'medium' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
  const bottomBorder = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'medium' }, right: { style: 'thin' } };
  const grandBorder = { top: { style: 'double' }, left: { style: 'thin' }, bottom: { style: 'double' }, right: { style: 'thin' } };

  ws.mergeCells(1, 1, 1, 6);
  const titleRow = ws.getCell('A1');
  titleRow.value = 'PT. CAHAYA BUANA KEMALA';
  titleRow.font = titleFont;
  titleRow.alignment = { horizontal: 'center', vertical: 'middle' };

  ws.mergeCells(2, 1, 2, 6);
  const subRow = ws.getCell('A2');
  subRow.value = 'NERACA';
  subRow.font = { name: 'Calibri', size: 12, bold: true };
  subRow.alignment = { horizontal: 'center', vertical: 'middle' };

  ws.mergeCells(3, 1, 3, 6);
  const perRow = ws.getCell('A3');
  perRow.value = 'PER : 31 JANUARI 2026';
  perRow.font = { name: 'Calibri', size: 10 };
  perRow.alignment = { horizontal: 'center', vertical: 'middle' };

  let r = 4;
  const hRow = ws.getRow(r);
  ['KETERANGAN', 'TOTAL', '%', 'KETERANGAN', 'TOTAL', '%'].forEach((h, i) => {
    const cell = hRow.getCell(i + 1);
    cell.value = h;
    cell.font = headerFont;
    cell.fill = headerFill;
    cell.alignment = { horizontal: i % 3 === 0 ? 'left' : 'right', vertical: 'middle', wrapText: true };
    cell.border = thinBorder;
  });
  r++;

  rows.forEach(row => {
    if (row.cls === 'spacer') { r++; return; }
    const isSection = row.cls === 'section';
    const isTotal = row.cls === 'total';
    const isGrand = row.cls === 'grand';

    const rr = ws.getRow(r);
    const vals = [row.left, row.lval, row.lpct, row.right, row.rval, row.rpct];
    vals.forEach((v, i) => {
      const cell = rr.getCell(i + 1);
      cell.value = v || null;
      if (isSection) {
        cell.font = { name: 'Calibri', size: 10, bold: true };
        cell.fill = sectionFill;
      } else if (isGrand) {
        cell.font = { name: 'Calibri', size: 11, bold: true };
        cell.fill = grandFill;
      } else if (isTotal) {
        cell.font = { name: 'Calibri', size: 10, bold: true };
        cell.fill = totalFill;
      } else {
        cell.font = { name: 'Calibri', size: 10 };
      }
      cell.alignment = { horizontal: i % 3 === 0 ? 'left' : 'right', vertical: 'middle' };
      if (isGrand) cell.border = grandBorder;
      else if (isTotal && i >= 3) cell.border = bottomBorder;
      else if (i % 3 === 0) cell.border = thinBorder;
      else cell.border = thinBorder;
    });
    if (isTotal) {
      [1, 4].forEach(ci => {
        const cell = rr.getCell(ci);
        cell.border = ci >= 4 ? bottomBorder : topBorder;
      });
    }
    r++;
  });

  ws.addRow([]);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=neraca.xlsx');
  await wb.xlsx.write(res);
  res.end();
}

async function downloadLabaRugi(req, res) {
  const period = req.query.period;
  const data = loadExcelData(period);
  if (!data) return res.status(404).json({ error: 'Belum upload Excel' });
  const cfg = loadConfig(data.kategoriList.map(k => k.nama), period);
  const mode = req.query.mode || 'netchange';

  const isNett = mode === 'netchange';
  const kat = data.kategoriList;
  const get = (name) => kat.find(k => k.nama === name);
  const getSection = (section) => (cfg.labaRugi[section] || []).map(n => get(n)).filter(Boolean);
  const katVal = (k) => Math.abs(isNett ? (k.totalNett || 0) : (k.totalAkhir || 0));
  const itemVal = (i) => Math.abs(isNett ? (i.nettChange || 0) : (i.saldoAkhir || 0));

  const penjualanList = getSection('PENJUALAN');
  const hppList = getSection('HARGA POKOK PENJUALAN');
  const bpList = getSection('BIAYA PENJUALAN');
  const bauList = getSection('BIAYA ADM UMUM');
  const pnoList = getSection('PENDAPATAN NON OPERASIONAL');
  const bnoList = getSection('BIAYA NON OPERASIONAL');

  const penjualanNet = penjualanList.reduce((a, k) => a + katVal(k), 0);
  const hppVal = hppList.reduce((a, k) => a + katVal(k), 0);
  const labaKotor = penjualanNet - hppVal;
  const bpVal = bpList.reduce((a, k) => a + katVal(k), 0);
  const bauVal = bauList.reduce((a, k) => a + katVal(k), 0);
  const totalBiayaOps = bpVal + bauVal;
  const labaOperasional = labaKotor - totalBiayaOps;
  const pnoVal = pnoList.reduce((a, k) => a + katVal(k), 0);
  const bnoVal = bnoList.reduce((a, k) => a + katVal(k), 0);
  const totalPendNonOps = pnoVal - bnoVal;
  const labaKomersil = labaOperasional + totalPendNonOps;

  const allPnoItems = [];
  pnoList.forEach(k => { (data.kategoriMap[k.nama] || []).forEach(i => allPnoItems.push(i)); });
  const allBnoItems = [];
  bnoList.forEach(k => { (data.kategoriMap[k.nama] || []).forEach(i => allBnoItems.push(i)); });

  const koreksiPositifCats = (cfg.labaRugi['KOREKSI POSITIF'] || []);
  const koreksiNegatifCats = (cfg.labaRugi['KOREKSI NEGATIF'] || []);
  let koreksiPositifItems = [];
  let koreksiNegatifItems = [];

  if (koreksiPositifCats.length) {
    koreksiPositifCats.forEach(catName => { (data.kategoriMap[catName] || []).forEach(i => koreksiPositifItems.push(i)); });
  } else {
    koreksiPositifItems = allBnoItems.filter(i => [93,96].includes(parseInt(i.kode)));
  }
  if (koreksiNegatifCats.length) {
    koreksiNegatifCats.forEach(catName => { (data.kategoriMap[catName] || []).forEach(i => koreksiNegatifItems.push(i)); });
  } else {
    koreksiNegatifItems = allPnoItems.filter(i => [71,74,78].includes(parseInt(i.kode)));
  }

  const koreksiPositif = koreksiPositifItems.reduce((a, i) => a + itemVal(i), 0);
  const koreksiNegatif = koreksiNegatifItems.reduce((a, i) => a + itemVal(i), 0);
  const totalKoreksiFiskal = koreksiPositif - koreksiNegatif;
  const labaSebelumPajak = labaKomersil + totalKoreksiFiskal;
  const pct = (v) => penjualanNet ? (v / penjualanNet * 100).toFixed(2).replace('.', ',') : '0,00';

  const items = [
    { label: 'PENJUALAN', section: true },
    ...penjualanList.map(k => ({ label: k.nama, value: katVal(k), pct: pct(katVal(k)), indent: true })),
    { label: 'RETUR PENJUALAN', value: 0, pct: '-', bold: false },
    { label: 'PENJUALAN BERSIH', value: penjualanNet, pct: pct(penjualanNet), bold: true },
    { spacer: true },
    { label: 'HARGA POKOK PENJUALAN', section: true },
    ...hppList.map(k => ({ label: k.nama, value: katVal(k), pct: pct(katVal(k)), indent: true })),
    { label: 'LABA KOTOR', value: labaKotor, pct: pct(labaKotor), bold: true },
    { spacer: true },
    { label: 'BIAYA OPERASIONAL', section: true },
    ...bpList.map(k => ({ label: k.nama, value: katVal(k), pct: pct(katVal(k)), indent: true })),
    ...bauList.map(k => ({ label: k.nama, value: katVal(k), pct: pct(katVal(k)), indent: true })),
    { label: 'TOTAL BIAYA OPERASIONAL', value: totalBiayaOps, pct: pct(totalBiayaOps), bold: true, indent: true },
    { spacer: true },
    { label: 'LABA OPERASIONAL', value: labaOperasional, pct: pct(labaOperasional), bold: true },
    { spacer: true },
    { label: 'PENDAPATAN DAN (BIAYA NON OPERASIONAL)', section: true },
    ...pnoList.map(k => ({ label: k.nama, value: katVal(k), pct: pct(katVal(k)), indent: true })),
    ...bnoList.map(k => ({ label: '( ' + k.nama + ' )', value: -katVal(k), pct: '(' + pct(katVal(k)) + ')', indent: true })),
    { label: 'TOTAL PENDAPATAN DAN (BIAYA NON OPERASIONAL)', value: totalPendNonOps, pct: pct(totalPendNonOps), bold: true, indent: true },
    { spacer: true },
    { label: 'LABA ( RUGI ) KOMERSIL', value: labaKomersil, pct: pct(labaKomersil), bold: true },
    { spacer: true },
    { label: 'KOREKSI POSITIF', section: true },
    ...koreksiPositifItems.map(i => ({ label: '**   ' + i.nama, value: itemVal(i), pct: pct(itemVal(i)), indent: true })),
    { label: 'TOTAL KOREKSI POSITIF', value: koreksiPositif, pct: pct(koreksiPositif), bold: true, indent: true },
    { spacer: true },
    { label: 'KOREKSI NEGATIF', section: true },
    ...koreksiNegatifItems.map(i => ({ label: '**   ' + i.nama, value: -itemVal(i), pct: '(' + pct(itemVal(i)) + ')', indent: true })),
    { label: 'TOTAL KOREKSI NEGATIF', value: -koreksiNegatif, pct: '(' + pct(koreksiNegatif) + ')', bold: true, indent: true },
    { spacer: true },
    { label: 'TOTAL KOREKSI FISKAL (KOREKSI POSITIF + KOREKSI NEGATIF)', value: totalKoreksiFiskal, pct: (totalKoreksiFiskal >= 0 ? '' : '(') + pct(Math.abs(totalKoreksiFiskal)) + (totalKoreksiFiskal >= 0 ? '' : ')'), bold: true },
    { spacer: true },
    { label: 'LABA ( RUGI ) SEBELUM PAJAK', value: labaSebelumPajak, pct: pct(labaSebelumPajak), bold: true, 'class': 'grand-total' }
  ];

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Laba Rugi');
  ws.columns = [{ width: 48 }, { width: 22 }, { width: 12 }];

  const titleFont = { name: 'Calibri', size: 14, bold: true };
  const headerFont = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A1A2E' } };
  const sectionFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE3F2FD' } };
  const totalFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8EAF6' } };
  const grandFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC5CAE9' } };
  const thinBorder = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
  const grandBorder = { top: { style: 'double' }, left: { style: 'thin' }, bottom: { style: 'double' }, right: { style: 'thin' } };

  ws.mergeCells(1, 1, 1, 3);
  ws.getCell('A1').value = 'LAPORAN LABA RUGI';
  ws.getCell('A1').font = titleFont;
  ws.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };

  let r = 2;
  const hRow = ws.getRow(r);
  ['KETERANGAN', 'TOTAL', '%'].forEach((h, i) => {
    const cell = hRow.getCell(i + 1);
    cell.value = h;
    cell.font = headerFont;
    cell.fill = headerFill;
    cell.alignment = { horizontal: i === 0 ? 'left' : 'right', vertical: 'middle' };
    cell.border = thinBorder;
  });

  items.forEach(item => {
    if (item.spacer) { r++; return; }
    r++;
    const rr = ws.getRow(r);

    if (item.section) {
      const cell = rr.getCell(1);
      cell.value = item.label;
      cell.font = { name: 'Calibri', size: 10, bold: true };
      cell.fill = sectionFill;
      cell.alignment = { horizontal: 'left', vertical: 'middle' };
      cell.border = thinBorder;
      [2, 3].forEach(ci => { rr.getCell(ci).border = thinBorder; });
      return;
    }

    const cls = item['class'] || (item.bold ? 'total' : '');
    const isGrand = cls === 'grand-total';
    const isTotal = cls === 'total';

    [item.label, item.value !== undefined ? rp(item.value) : '-', item.pct || '-'].forEach((v, i) => {
      const cell = rr.getCell(i + 1);
      cell.value = v;
      if (isGrand) {
        cell.font = { name: 'Calibri', size: 11, bold: true };
        cell.fill = grandFill;
        cell.border = grandBorder;
      } else if (isTotal) {
        cell.font = { name: 'Calibri', size: 10, bold: true };
        cell.fill = totalFill;
        cell.border = thinBorder;
      } else {
        cell.font = { name: 'Calibri', size: 10 };
        cell.border = thinBorder;
      }
      cell.alignment = { horizontal: i === 0 ? 'left' : 'right', vertical: 'middle' };
    });
  });

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=laba-rugi.xlsx');
  await wb.xlsx.write(res);
  res.end();
}

app.get('/api/neraca/download', async (req, res) => {
  try { await downloadNeraca(req, res); } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/laba-rugi/download', async (req, res) => {
  try { await downloadLabaRugi(req, res); } catch (e) { res.status(500).json({ error: e.message }); }
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server: http://localhost:${PORT}`);
});
