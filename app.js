const API = '/api';
let semuaData = [];
let ringkasanData = null;
let neracaData = null;
let configData = null;
let labaRugiData = null;
let configType = 'neraca';
let labaRugiMode = 'netchange';
let selectedPeriod = localStorage.getItem('selectedPeriod') || 'master';

function rp(n) {
  if (n === undefined || n === null) return '-';
  const abs = Math.abs(n);
  const f = new Intl.NumberFormat('id-ID').format(Math.round(abs));
  return n < 0 ? `(${f})` : f;
}

function pct(n, total) {
  if (!total) return '0,00';
  return (Math.abs(n) / Math.abs(total) * 100).toFixed(2).replace('.', ',');
}

function toggleSidebar() {
  document.getElementById('sidebar').classList.toggle('open');
}

async function init() {
  try {
    const periods = await fetch(`${API}/periods`).then(r => r.json());
    if (periods.length === 0) {
      document.getElementById('uploadOverlay').style.display = 'flex';
      return;
    }
    if (!periods.find(p => p.key === selectedPeriod)) {
      selectedPeriod = periods[0].key;
      localStorage.setItem('selectedPeriod', selectedPeriod);
    }
    document.getElementById('uploadOverlay').style.display = 'none';
    document.getElementById('appContainer').classList.remove('hidden');
    loadPeriodsUI();
    try {
      await loadAll();
    } catch {
      document.getElementById('uploadOverlay').style.display = 'flex';
      document.getElementById('appContainer').classList.add('hidden');
    }
  } catch (e) {
    document.querySelector('.upload-box').innerHTML = `<p style="color:#c62828;">Server tidak terhubung. Pastikan server berjalan.</p>`;
  }
}

async function loadPeriodsUI() {
  try {
    const periods = await fetch(`${API}/periods`).then(r => r.json());
    const container = document.getElementById('periodSelector');
    container.innerHTML = periods.map(p =>
      `<button class="period-btn ${p.key === selectedPeriod ? 'active' : ''}" data-key="${p.key}" onclick="gantiPeriod('${p.key}')">${p.label}</button>`
    ).join('');
  } catch {}
}

async function gantiPeriod(key) {
  selectedPeriod = key;
  localStorage.setItem('selectedPeriod', key);
  document.querySelectorAll('.period-btn').forEach(b => b.classList.toggle('active', b.dataset.key === key));
  await loadAll(true);
}

document.getElementById('fileInput').onchange = function () {
  const file = this.files[0];
  if (file) {
    this.classList.add('selected');
    document.getElementById('fileDisplay').textContent = 'File: ' + file.name + ' (' + (file.size / 1024).toFixed(0) + ' KB)';
    document.getElementById('btnUpload').disabled = false;
  }
};

document.getElementById('uploadForm').onsubmit = async (e) => {
  e.preventDefault();
  const file = document.getElementById('fileInput').files[0];
  if (!file) return;
  const periodName = document.getElementById('periodName').value.trim();
  if (!periodName) { document.getElementById('uploadStatus').innerHTML = '<span style="color:#c62828;">Nama periode harus diisi</span>'; return; }
  const formData = new FormData();
  formData.append('period', periodName);
  formData.append('label', periodName);
  formData.append('file', file);
  const status = document.getElementById('uploadStatus');
  status.innerHTML = 'Memproses...';
  try {
    const res = await fetch(`${API}/upload`, { method: 'POST', body: formData });
    const data = await res.json();
    if (data.error) { status.innerHTML = `<span style="color:#c62828;">${data.error}</span>`; return; }
    selectedPeriod = data.period;
    localStorage.setItem('selectedPeriod', selectedPeriod);
    status.innerHTML = `<span style="color:#2e7d32;">Berhasil! ${data.total} akun dimuat.</span>`;
    setTimeout(async () => {
      document.getElementById('uploadOverlay').style.display = 'none';
      document.getElementById('appContainer').classList.remove('hidden');
      loadPeriodsUI();
      await loadAll();
    }, 800);
  } catch (e) {
    status.innerHTML = `<span style="color:#c62828;">Gagal upload: ${e.message}</span>`;
  }
};

function uploadUlang() {
  document.getElementById('appContainer').classList.add('hidden');
  document.getElementById('uploadOverlay').style.display = 'flex';
  document.getElementById('fileInput').value = '';
  document.getElementById('fileInput').classList.remove('selected');
  document.getElementById('btnUpload').disabled = true;
  document.getElementById('uploadStatus').innerHTML = '';
  const p = document.getElementById('periodName');
  p.value = selectedPeriod === 'master' ? '' : selectedPeriod;
  p.focus();
}

function qp(extra) {
  const p = selectedPeriod || 'master';
  const params = new URLSearchParams({ period: p });
  if (extra) Object.entries(extra).forEach(([k, v]) => params.set(k, v));
  const s = params.toString();
  return s ? '?' + s : '';
}

let currentPage = 'dashboard';

async function loadAll(keepPage) {
  showLoading(true);
  try {
    const mode = document.getElementById('labaRugiMode')?.value || 'netchange';
    labaRugiMode = mode;
    const [semua, ringkasan, neraca, config, labaRugi] = await Promise.all([
      fetch(`${API}/data${qp()}`).then(r => r.json()),
      fetch(`${API}/ringkasan${qp()}`).then(r => r.json()),
      fetch(`${API}/neraca${qp()}`).then(r => r.json()),
      fetch(`${API}/config${qp()}`).then(r => r.json()),
      fetch(`${API}/laba-rugi${qp({ mode })}`).then(r => r.json())
    ]);
    semuaData = semua;
    ringkasanData = ringkasan;
    neracaData = neraca;
    configData = config;
    labaRugiData = labaRugi;
    renderDashboard();
    renderNeraca();
    renderLabaRugi();
    renderKonfigurasi();
    isiFilterKategori();
    showPage(keepPage ? currentPage : 'dashboard');
    showLoading(false);
  } catch (e) {
    console.error('loadAll failed', e);
    // auto-recover if period not found -> reset to master
    if (selectedPeriod !== 'master') {
      console.log('Resetting invalid period', selectedPeriod, 'to master');
      selectedPeriod = 'master';
      localStorage.setItem('selectedPeriod', 'master');
      // retry once with master
      try {
        const periods = await fetch(`${API}/periods`).then(r=>r.json());
        if (periods.length>0 && !periods.find(p=>p.key===selectedPeriod)) {
          selectedPeriod = periods[0].key;
          localStorage.setItem('selectedPeriod', selectedPeriod);
        }
        return await loadAll(keepPage);
      } catch (e2) {
        document.getElementById('loading').innerHTML = 'Gagal memuat data. Coba clear cache: localStorage.clear()';
      }
    } else {
      document.getElementById('loading').innerHTML = 'Gagal memuat data.';
    }
  }
}

function showLoading(show) {
  const el = document.getElementById('loading');
  if (show) el.classList.remove('hidden'); else el.classList.add('hidden');
}

function showPage(page) {
  currentPage = page;
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.getElementById(`page-${page}`).classList.add('active');
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  const t = { dashboard: 'Dashboard', 'semua-data': 'Semua Data', neraca: 'Neraca', 'laba-rugi': 'Laba Rugi', konfigurasi: 'Konfigurasi' };
  document.getElementById('pageTitle').textContent = t[page] || page;
  const el = document.querySelector(`[onclick="showPage('${page}')"]`);
  if (el) el.classList.add('active');
  document.getElementById('sidebar').classList.remove('open');
  if (page === 'semua-data') renderSemuaData(semuaData);
  if (page === 'neraca') renderNeraca();
  if (page === 'laba-rugi') renderLabaRugi();
  if (page === 'konfigurasi') { renderKonfigurasi(); showConfigTab(configType); }
}

function renderDashboard() {
  const r = ringkasanData;
  if (!r) return;
  document.getElementById('summaryCards').innerHTML = `
    <div class="summary-card"><div class="label">Aktiva Lancar</div><div class="value positive">${rp(r.totalAsetLancar)}</div></div>
    <div class="summary-card"><div class="label">Aktiva Lain-lain</div><div class="value">${rp(r.totalAsetLain)}</div></div>
    <div class="summary-card"><div class="label">Aktiva Tetap</div><div class="value">${rp(r.totalAsetTetap)}</div></div>
    <div class="summary-card"><div class="label">Hutang Lancar</div><div class="value negative">${rp(Math.abs(r.totalHutangLancar))}</div></div>
    <div class="summary-card"><div class="label">Modal</div><div class="value positive">${rp(Math.abs(r.totalModal))}</div></div>
    <div class="summary-card"><div class="label">Total Akun</div><div class="value">${r.totalAkun}</div></div>
  `;
  renderChartKategori(r.kategori);
  renderChartPie(r);
}

let chartKategori, chartPie;
function renderChartKategori(kategori) {
  if (chartKategori) chartKategori.destroy();
  const ctx = document.getElementById('chartKategori').getContext('2d');
  const labels = kategori.filter(k => Math.abs(k.totalAkhir) > 0).slice(0, 12).map(k => k.nama);
  const values = kategori.filter(k => Math.abs(k.totalAkhir) > 0).slice(0, 12).map(k => Math.abs(k.totalAkhir));
  chartKategori = new Chart(ctx, {
    type: 'bar', data: { labels, datasets: [{ label: 'Saldo Akhir', data: values, backgroundColor: '#42a5f5', borderRadius: 4 }] },
    options: { responsive: true, plugins: { legend: { display: false }, title: { display: true, text: 'Top 12 Kategori' } }, scales: { y: { beginAtZero: true, ticks: { callback: v => rp(v) } } } }
  });
}

function renderChartPie(r) {
  if (chartPie) chartPie.destroy();
  const ctx = document.getElementById('chartPie').getContext('2d');
  chartPie = new Chart(ctx, {
    type: 'doughnut', data: {
      labels: ['Aktiva Lancar', 'Aktiva Lain', 'Aktiva Tetap', 'Hutang Lancar', 'Modal'],
      datasets: [{ data: [r.totalAsetLancar, r.totalAsetLain, Math.abs(r.totalAsetTetap), Math.abs(r.totalHutangLancar), Math.abs(r.totalModal)], backgroundColor: ['#42a5f5','#ab47bc','#66bb6a','#ef5350','#ffa726'] }]
    },
    options: { responsive: true, plugins: { title: { display: true, text: 'Komposisi Neraca' } } }
  });
}

let selectedKeys = new Set();

function renderSemuaData(data) {
  const filtered = data || semuaData;
  document.getElementById('totalData').textContent = `${filtered.length} akun`;
  const sums = { saldoAwal: 0, debet: 0, kredit: 0, nettChange: 0, saldoAkhir: 0 };
  filtered.forEach(a => { sums.saldoAwal += a.saldoAwal; sums.debet += a.debet; sums.kredit += a.kredit; sums.nettChange += a.nettChange; sums.saldoAkhir += a.saldoAkhir; });
  const tbody = document.getElementById('tbodySemuaData');
  tbody.innerHTML = filtered.map(a => {
    const key = a.kode + ':' + a.nama;
    const editing = editingKey === key;
    const checked = selectedKeys.has(key);
    if (editing) {
      const es = editState;
      return `<tr class="editing">
        <td><input type="checkbox" disabled></td>
        <td><span class="edit-actions"><button class="btn-save-edit" onclick="saveEdit()">Simpan</button><button class="btn-cancel-edit" onclick="cancelEdit()">Batal</button></span></td>
        <td>${a.kode}</td><td><strong>${a.nama}</strong></td>
        <td class="num"><input class="edit-input" type="text" id="editSaldoAwal" value="${es.saldoAwal}" oninput="calcEdit()" style="width:110px"></td>
        <td class="num"><input class="edit-input" type="text" id="editDebet" value="${es.debet}" oninput="calcEdit()" style="width:110px"></td>
        <td class="num"><input class="edit-input" type="text" id="editKredit" value="${es.kredit}" oninput="calcEdit()" style="width:110px"></td>
        <td class="num"><span id="editNettChange">${es.nettChange}</span></td>
        <td class="num"><span id="editSaldoAkhir">${es.saldoAkhir}</span></td>
        <td><span class="badge">${a.kategori}</span></td>
      </tr>`;
    }
    return `<tr class="${checked ? 'selected-row' : ''}">
      <td><input type="checkbox" ${checked ? 'checked' : ''} onchange="toggleSelect('${key}')"></td>
      <td><button class="btn-aksi" onclick="startEdit('${a.kode}','${a.nama.replace(/'/g, "\\'")}',${a.saldoAwal},${a.debet},${a.kredit})">Edit</button></td>
      <td>${a.kode}</td><td><strong>${a.nama}</strong></td>
      <td class="num">${rp(a.saldoAwal)}</td><td class="num">${rp(a.debet)}</td><td class="num">${rp(a.kredit)}</td>
      <td class="num">${rp(a.nettChange)}</td><td class="num">${rp(a.saldoAkhir)}</td>
      <td><span class="badge">${a.kategori}</span></td>
    </tr>`;
  }).join('') + `<tr class="total-row">
    <td></td><td></td><td></td><td></td>
    <td class="num">${rp(sums.saldoAwal)}</td>
    <td class="num">${rp(sums.debet)}</td>
    <td class="num">${rp(sums.kredit)}</td>
    <td class="num">${rp(sums.nettChange)}</td>
    <td class="num">${rp(sums.saldoAkhir)}</td>
    <td></td>
  </tr>`;
  updateBatchBar();
}

let editingKey = null;
let editState = {};

function startEdit(kode, nama, saldoAwal, debet, kredit) {
  editingKey = kode + ':' + nama;
  editState = { kode, nama, saldoAwal, debet, kredit, nettChange: debet - kredit, saldoAkhir: saldoAwal + debet - kredit };
  filterData();
}

function calcEdit() {
  const sa = parseFloat(document.getElementById('editSaldoAwal').value.replace(/\./g, '').replace(',', '.')) || 0;
  const d = parseFloat(document.getElementById('editDebet').value.replace(/\./g, '').replace(',', '.')) || 0;
  const k = parseFloat(document.getElementById('editKredit').value.replace(/\./g, '').replace(',', '.')) || 0;
  const nc = d - k;
  const sa2 = sa + d - k;
  document.getElementById('editNettChange').textContent = rp(nc);
  document.getElementById('editSaldoAkhir').textContent = rp(sa2);
  editState = { ...editState, saldoAwal: sa, debet: d, kredit: k, nettChange: nc, saldoAkhir: sa2 };
}

function cancelEdit() {
  editingKey = null;
  editState = {};
  filterData();
}

async function saveEdit() {
  if (!editState.kode) return;
  await fetch(`${API}/data/${editState.kode}${qp()}`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ saldoAwal: editState.saldoAwal, debet: editState.debet, kredit: editState.kredit })
  });
  editingKey = null;
  editState = {};
  await loadAll(true);
}

function toggleSelect(key) {
  if (selectedKeys.has(key)) selectedKeys.delete(key); else selectedKeys.add(key);
  document.getElementById('checkAll').checked = selectedKeys.size === (document.getElementById('tbodySemuaData').childElementCount);
  updateBatchBar();
}

function toggleSelectAll() {
  const checked = document.getElementById('checkAll').checked;
  selectedKeys.clear();
  if (checked) {
    const data = semuaDataFiltred || semuaData;
    data.forEach(a => selectedKeys.add(a.kode + ':' + a.nama));
  }
  filterData();
  updateBatchBar();
}

function updateBatchBar() {
  const bar = document.getElementById('batchBar');
  if (selectedKeys.size === 0) { bar.classList.add('hidden'); return; }
  bar.classList.remove('hidden');
  document.getElementById('selectedCount').textContent = `${selectedKeys.size} item dipilih`;
}

function clearSelection() {
  selectedKeys.clear();
  document.getElementById('checkAll').checked = false;
  filterData();
  updateBatchBar();
}

async function batchEditSaldoAwal() {
  const input = prompt('Masukkan nilai Saldo Awal baru untuk ' + selectedKeys.size + ' item yang dipilih:', '0');
  if (input === null) return;
  const val = parseFloat(input.replace(/\./g, '').replace(',', '.')) || 0;
  const items = [];
  selectedKeys.forEach(key => {
    const [kode] = key.split(':');
    const a = semuaData.find(d => d.kode + ':' + d.nama === key);
    if (a) items.push({ kode, saldoAwal: val, debet: a.debet, kredit: a.kredit });
  });
  if (items.length === 0) return;
  await fetch(`${API}/data/batch${qp()}`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ items })
  });
  selectedKeys.clear();
  document.getElementById('checkAll').checked = false;
  await loadAll(true);
  filterData();
}

let semuaDataFiltred = null;

function isiFilterKategori() {
  const select = document.getElementById('filterKategori');
  const kat = [...new Set(semuaData.map(a => a.kategori))].sort();
  select.innerHTML = '<option value="">Semua Kategori</option>' + kat.map(k => `<option value="${k}">${k}</option>`).join('');
}

function filterData() {
  const kat = document.getElementById('filterKategori').value;
  const q = document.getElementById('searchInput').value.toLowerCase().trim();
  let data = semuaData;
  if (kat) data = data.filter(a => a.kategori === kat);
  if (q) data = data.filter(a => a.nama.toLowerCase().includes(q) || a.kode.toLowerCase().includes(q) || a.kategori.toLowerCase().includes(q));
  semuaDataFiltred = data;
  renderSemuaData(data);
}

function searchData(e) {
  if (e.key === 'Enter') filterData();
}

function renderNeraca() {
  const n = neracaData;
  if (!n) return;

  function row(left, lval, lpct, right, rval, rpctVal, cls, indent) {
    return { left, lval, lpct, right, rval, rpct: rpctVal, cls, indent };
  }
  function sectionRow(left, right) {
    return { section: true, left, right };
  }
  function spacer() {
    return { spacer: true };
  }

  const rs = [];
  const push = (r) => rs.push(r);

  push(sectionRow('AKTIVA LANCAR', 'HUTANG LANCAR'));

  const l1 = n.aktivaLancar.map(k => ({ n: k.nama, v: rp(k.saldo), p: pct(k.saldo, n.totalAktiva) }));
  const r1 = n.hutangLancar.map(k => ({ n: k.nama, v: rp(Math.abs(k.saldo)), p: pct(k.saldo, n.totalPasiva) }));
  const m1 = Math.max(l1.length, r1.length);
  for (let i = 0; i < m1; i++) {
    push(row(
      l1[i]?.n || '', l1[i]?.v || '', l1[i]?.p || '',
      r1[i]?.n || '', r1[i]?.v || '', r1[i]?.p || '',
      '', true
    ));
  }
  push(row('TOTAL AKTIVA LANCAR', rp(n.totalAktivaLancar), pct(n.totalAktivaLancar, n.totalAktiva),
    'TOTAL HUTANG LANCAR', rp(Math.abs(n.totalHutangLancar)), pct(n.totalHutangLancar, n.totalPasiva), 'total-row'));

  push(spacer());

  push(sectionRow('AKTIVA LAIN-LAIN', 'MODAL'));

  const l2 = n.aktivaLain.map(k => ({ n: k.nama, v: rp(k.saldo), p: pct(k.saldo, n.totalAktiva) }));
  const r2 = n.modal.map(k => ({ n: k.nama, v: rp(Math.abs(k.saldo)), p: pct(k.saldo, n.totalPasiva) }));
  const m2 = Math.max(l2.length || 1, r2.length || 1);
  for (let i = 0; i < m2; i++) {
    push(row(
      l2[i]?.n || '', l2[i]?.v || '', l2[i]?.p || '',
      r2[i]?.n || '', r2[i]?.v || '', r2[i]?.p || '',
      '', true
    ));
  }
  push(row('TOTAL AKTIVA LAIN-LAIN', rp(n.totalAktivaLain), pct(n.totalAktivaLain, n.totalAktiva),
    'TOTAL MODAL', rp(Math.abs(n.totalModal)), pct(n.totalModal, n.totalPasiva), 'total-row'));

  push(spacer());

  push(sectionRow('AKTIVA TETAP', ''));

  for (const k of n.aktivaTetap) {
    push(row(k.nama, rp(k.saldo), pct(k.saldo, n.totalAktiva), '', '', '', '', true));
  }
  push(row('TOTAL AKTIVA TETAP', rp(n.totalAktivaTetap), pct(n.totalAktivaTetap, n.totalAktiva), '', '', '', 'total-row'));

  push(spacer());

  push(row('TOTAL AKTIVA', rp(n.totalAktiva), '100,00', 'TOTAL HUTANG + MODAL', rp(n.totalPasiva), '100,00', 'grand-total'));

  const html = rs.map(r => {
    if (r.spacer) return '<tr class="spacer"><td colspan="6"></td></tr>';
    if (r.section) {
      return `<tr class="section-row"><td class="section-title" colspan="3">${r.left}</td><td class="section-title" colspan="3">${r.right}</td></tr>`;
    }
    const cls = r.cls || '';
    const indentStyle = r.indent ? ' style="padding-left:24px"' : '';
    return `<tr class="${cls}">
      <td${indentStyle}>${r.left}</td><td class="num">${r.lval}</td><td class="num pct">${r.lpct}</td>
      <td${indentStyle}>${r.right || ''}</td><td class="num">${r.rval || ''}</td><td class="num pct">${r.rpct || ''}</td>
    </tr>`;
  }).join('');

  document.getElementById('neracaBody').innerHTML = html;
}

function renderLabaRugi() {
  const d = labaRugiData;
  if (!d || !d.items) return;
  const basePct = d.items.find(i => i.label === 'PENJUALAN BERSIH')?.value || 1;

  let html = `<table class="laba-rugi-table">
    <thead><tr><th style="width:55%">KETERANGAN</th><th class="num" style="width:25%">TOTAL</th><th class="num" style="width:20%">%</th></tr></thead>
    <tbody>`;

  d.items.forEach(item => {
    if (item.spacer) {
      html += '<tr class="spacer"><td colspan="3"></td></tr>';
    } else if (item.section) {
      html += `<tr class="section-row"><td class="section-title" colspan="3">${item.label}</td></tr>`;
    } else {
      const cls = item['class'] || (item.bold ? 'total-row' : '');
      const indent = item.indent ? ' style="padding-left:24px"' : '';
      const val = item.value !== undefined && item.value !== null ? rp(item.value) : '-';
      const pct = item.pct || '-';
      html += `<tr class="${cls}"><td${indent}>${item.label}</td><td class="num">${val}</td><td class="num pct">${pct}</td></tr>`;
    }
  });

  html += '</tbody></table>';
  document.getElementById('labaRugiContent').innerHTML = html;
}

function showConfigTab(type) {
  configType = type;
  document.querySelectorAll('.config-tab').forEach(t => t.classList.toggle('active', t.dataset.type === type));
  document.getElementById('configDesc').textContent = type === 'neraca'
    ? 'Atur kategori mana yang masuk ke bagian Neraca. Drag & drop kategori antar bagian.'
    : 'Atur kategori mana yang masuk ke bagian Laba Rugi. Drag & drop kategori antar bagian.';
  renderKonfigurasi();
}

function renderKonfigurasi() {
  const cfg = configData;
  if (!cfg) return;
  const isNeraca = configType === 'neraca';
  const sections = isNeraca
    ? ['AKTIVA LANCAR', 'AKTIVA LAIN-LAIN', 'AKTIVA TETAP', 'HUTANG LANCAR', 'MODAL']
    : ['PENJUALAN', 'HARGA POKOK PENJUALAN', 'BIAYA PENJUALAN', 'BIAYA ADM UMUM', 'PENDAPATAN NON OPERASIONAL', 'BIAYA NON OPERASIONAL', 'KOREKSI POSITIF', 'KOREKSI NEGATIF'];
  const group = isNeraca ? 'neraca' : 'labaRugi';
  const allMapped = new Set();
  const container = document.getElementById('configContainer');
  const saldoMap = {};
  if (ringkasanData) {
    ringkasanData.kategori.forEach(k => { saldoMap[k.nama] = k.totalAkhir; });
  }

  const optAll = sections.map(s => `<option value="${s}">${s}</option>`).join('');
  const selTpl = (cur) => `<select class="move-select" onchange="pindahKategori(this,'${group}')">${sections.map(s => `<option value="${s}" ${s === cur ? 'selected' : ''}>${s}</option>`).join('')}<option value="" ${!cur ? 'selected' : ''}>-- Tidak Digunakan --</option></select>`;

  container.innerHTML = sections.map(s => {
    const items = ((cfg.config[group] || {})[s] || []).filter(k => cfg.kategori.includes(k));
    items.forEach(k => allMapped.add(k));
    const id = 'drop-' + group + '-' + s.replace(/\s/g, '');
    return `<div class="config-section">
      <h4>${s}</h4>
      <div class="drop-area" id="${id}" data-section="${s}" data-group="${group}">
        ${items.length ? items.map((k, idx) => `
          <div class="config-item active" draggable="true" data-name="${k}" data-section="${s}" data-group="${group}" data-index="${idx}">
            <span class="drag-handle">&#9776;</span>
            <span class="item-name">${k}</span>
            <span class="item-value">${rp(saldoMap[k] ?? 0)}</span>
            ${selTpl(s)}
          </div>
        `).join('') : '<div class="empty-hint">Kosong - seret kategori ke sini</div>'}
      </div>
    </div>`;
  }).join('');

  const mappedInThisGroup = new Set();
  sections.forEach(s => {
    ((cfg.config[group] || {})[s] || []).forEach(k => mappedInThisGroup.add(k));
  });
  const unused = cfg.kategori.filter(k => !mappedInThisGroup.has(k));
  const unusedEl = document.getElementById('unusedList');
  unusedEl.innerHTML = unused.length
    ? unused.map((k, idx) => `
        <div class="config-item inactive" draggable="true" data-name="${k}" data-section="unused" data-group="${group}" data-index="${idx}">
          <span class="drag-handle">&#9776;</span>
          <span class="item-name">${k}</span>
          <span class="item-value">${rp(saldoMap[k] ?? 0)}</span>
          ${selTpl('')}
        </div>
      `).join('')
    : '<div class="empty-hint">Semua kategori sudah terpakai</div>';

  document.querySelectorAll('.config-item').forEach(el => {
    el.addEventListener('dragstart', onDragStart);
    el.addEventListener('dragend', onDragEnd);
  });
  document.querySelectorAll('.drop-area').forEach(el => {
    el.addEventListener('dragover', onDragOver);
    el.addEventListener('drop', onDrop);
  });
}

function pindahKategori(sel, group) {
  const item = sel.closest('.config-item');
  const name = item.dataset.name;
  const fromSection = item.dataset.section;
  const toSection = sel.value;

  fetch(`${API}/config${qp()}`).then(r => r.json()).then(cfg => {
    const config = JSON.parse(JSON.stringify(cfg.config));
    if (fromSection !== 'unused' && config[group]) {
      config[group][fromSection] = (config[group][fromSection] || []).filter(k => k !== name);
    }
    if (!config[group]) config[group] = {};
    if (toSection) {
      const list = config[group][toSection] || [];
      list.push(name);
      config[group][toSection] = list;
    }
    fetch(`${API}/config${qp()}`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ config }) })
      .then(() => Promise.all([
        fetch(`${API}/config${qp()}`).then(r => r.json()),
        fetch(`${API}/neraca${qp()}`).then(r => r.json()),
        fetch(`${API}/ringkasan${qp()}`).then(r => r.json()),
        fetch(`${API}/laba-rugi${qp({ mode: labaRugiMode })}`).then(r => r.json())
      ]))
      .then(([cfg2, neraca, ringkasan, labaRugi]) => {
        configData = cfg2;
        neracaData = neraca;
        ringkasanData = ringkasan;
        labaRugiData = labaRugi;
        renderKonfigurasi();
        renderNeraca();
        renderDashboard();
        renderLabaRugi();
      });
  });
}

let dragged = null;

function onDragStart(e) {
  dragged = {
    name: e.target.dataset.name,
    section: e.target.dataset.section,
    group: e.target.dataset.group || configType,
    el: e.target
  };
  e.target.classList.add('dragging');
  e.dataTransfer.effectAllowed = 'move';
}

function onDragEnd(e) {
  e.target.classList.remove('dragging');
  document.querySelectorAll('.config-item.drag-over').forEach(el => el.classList.remove('drag-over'));
}

function onDragOver(e) {
  e.preventDefault();
  e.dataTransfer.dropEffect = 'move';
  const target = e.target.closest('.config-item');
  document.querySelectorAll('.config-item.drag-over').forEach(el => el.classList.remove('drag-over'));
  if (target && target !== dragged?.el) {
    const rect = target.getBoundingClientRect();
    const mid = rect.top + rect.height / 2;
    if (e.clientY > mid) target.classList.add('drag-over');
    else target.classList.add('drag-over');
  }
}

function onDrop(e) {
  e.preventDefault();
  if (!dragged) return;
  const dropArea = e.target.closest('.drop-area');
  if (!dropArea) return;
  const toSection = dropArea.dataset.section;
  const toGroup = dropArea.dataset.group || configType;
  const fromGroup = dragged.group;

  fetch(`${API}/config${qp()}`).then(r => r.json()).then(cfg => {
    const config = JSON.parse(JSON.stringify(cfg.config));

    if (dragged.section !== 'unused' && config[fromGroup]) {
      config[fromGroup][dragged.section] = (config[fromGroup][dragged.section] || []).filter(k => k !== dragged.name);
    }

    if (!config[toGroup]) config[toGroup] = {};
    const dropTarget = e.target.closest('.config-item');
    const list = config[toGroup][toSection] || [];
    if (dropTarget && dropTarget.dataset.section === toSection && dropTarget.dataset.group === toGroup) {
      const insertIdx = parseInt(dropTarget.dataset.index);
      const after = e.clientY > dropTarget.getBoundingClientRect().top + dropTarget.offsetHeight / 2;
      list.splice(after ? insertIdx + 1 : insertIdx, 0, dragged.name);
    } else {
      list.push(dragged.name);
    }
    config[toGroup][toSection] = list;

    fetch(`${API}/config${qp()}`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ config }) })
      .then(() => {
        dragged = null;
        return Promise.all([
          fetch(`${API}/config${qp()}`).then(r => r.json()),
          fetch(`${API}/neraca${qp()}`).then(r => r.json()),
          fetch(`${API}/ringkasan${qp()}`).then(r => r.json()),
          fetch(`${API}/laba-rugi${qp({ mode: labaRugiMode })}`).then(r => r.json())
        ]);
      })
      .then(([cfg, neraca, ringkasan, labaRugi]) => {
        configData = cfg;
        neracaData = neraca;
        ringkasanData = ringkasan;
        labaRugiData = labaRugi;
        renderKonfigurasi();
        renderNeraca();
        renderDashboard();
        renderLabaRugi();
      });
  });
}

function simpanConfig() {
  const groups = ['neraca', 'labaRugi'];
  const sectionMap = {
    neraca: ['AKTIVA LANCAR', 'AKTIVA LAIN-LAIN', 'AKTIVA TETAP', 'HUTANG LANCAR', 'MODAL'],
    labaRugi: ['PENJUALAN', 'HARGA POKOK PENJUALAN', 'BIAYA PENJUALAN', 'BIAYA ADM UMUM', 'PENDAPATAN NON OPERASIONAL', 'BIAYA NON OPERASIONAL', 'KOREKSI POSITIF', 'KOREKSI NEGATIF']
  };
  const config = {};
  groups.forEach(g => {
    config[g] = {};
    sectionMap[g].forEach(s => {
      const area = document.getElementById('drop-' + g + '-' + s.replace(/\s/g, ''));
      config[g][s] = area ? [...area.querySelectorAll('.config-item')].map(el => el.dataset.name) : [];
    });
  });
  fetch(`${API}/config${qp()}`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ config }) })
    .then(() => {
      alert('Konfigurasi tersimpan!');
      return Promise.all([
        fetch(`${API}/neraca${qp()}`).then(r => r.json()),
        fetch(`${API}/ringkasan${qp()}`).then(r => r.json()),
        fetch(`${API}/laba-rugi${qp({ mode: labaRugiMode })}`).then(r => r.json())
      ]);
    })
    .then(([neraca, ringkasan, labaRugi]) => {
      neracaData = neraca;
      ringkasanData = ringkasan;
      labaRugiData = labaRugi;
      renderNeraca();
      renderDashboard();
      renderLabaRugi();
    });
}

document.addEventListener('keydown', function(e) {
  if (e.key === 'Enter' && editingKey) saveEdit();
  if (e.key === 'Escape' && editingKey) cancelEdit();
});

function downloadNeraca() {
  window.open(`${API}/neraca/download${qp()}`, '_blank');
}

async function ubahModeLabaRugi() {
  const mode = document.getElementById('labaRugiMode').value;
  labaRugiMode = mode;
  const res = await fetch(`${API}/laba-rugi${qp({ mode })}`);
  labaRugiData = await res.json();
  renderLabaRugi();
}

function downloadLabaRugi() {
  const mode = labaRugiMode || 'netchange';
  window.open(`${API}/laba-rugi/download${qp({ mode })}`, '_blank');
}

document.addEventListener('DOMContentLoaded', init);
