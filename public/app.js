// ===================== State =====================
let db = null;
let defaults = { wali: {}, kepsek: 'Tri Wijayanti, M.Pd' };
let currentKelas = '';
let currentSiswaList = [];
let visibleList = [];
let currentNi = null;

const CFG_KEY = 'eraporConfig';
const IDENT_KEY = 'identitasSekolah';
const DEF_TGL_RAPOR = '2026-06-02';
const DEF_TGL_KELULUSAN = '2026-06-03';

// ===================== Elemen =====================
const $ = (id) => document.getElementById(id);
const uploadForm = $('upload-form');
const fileInput = $('file-input');
const uploadStatus = $('upload-status');
const setupSection = $('setup-section');
const workspace = $('workspace');
const kelasSelect = $('kelas-select');
const searchInput = $('search-input');
const listHint = $('list-hint');
const siswaTbody = $('siswa-tbody');
const previewFrame = $('preview-frame');
const previewScaler = $('preview-scaler');
const previewBody = $('preview-body');
const previewEmpty = $('preview-empty');
const previewTitle = $('preview-title');
const btnPrint = $('btn-print');
const btnPrintKelas = $('btn-print-kelas');
const btnNewtab = $('btn-newtab');
const btnPrev = $('btn-prev');
const btnNext = $('btn-next');
const navPos = $('nav-pos');
const loading = $('loading');
const loadingText = $('loading-text');
const btnLoadRapor = $('btn-load-rapor');
const btnSettings = $('btn-settings');

// ===================== Util =====================
function showLoading(t) { loadingText.textContent = t || 'Memproses...'; loading.style.display = 'flex'; }
function hideLoading() { loading.style.display = 'none'; }
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
function num(v, d) { const n = parseInt(v, 10); return isNaN(n) ? d : n; }
function esc(t) { return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

function getCfg() { try { return JSON.parse(localStorage.getItem(CFG_KEY) || '{}'); } catch (e) { return {}; } }
function setCfg(cfg) { localStorage.setItem(CFG_KEY, JSON.stringify(cfg)); }
function ensureCfg() {
  const cfg = getCfg();
  if (!cfg.kepsek) cfg.kepsek = { name: '', sig: '', w: 110, x: 0, y: 0 };
  if (!cfg.wali) cfg.wali = {};
  return cfg;
}
function targetSlice(which) {
  const cfg = ensureCfg();
  if (which === 'ks') return { cfg, obj: cfg.kepsek };
  const k = $('wk-kelas').value;
  if (!cfg.wali[k]) cfg.wali[k] = { name: '', sig: '', w: 110, x: 0, y: 0 };
  return { cfg, obj: cfg.wali[k] };
}

function getIdent() { try { return JSON.parse(localStorage.getItem(IDENT_KEY) || '{}'); } catch (e) { return {}; } }
function setIdent(obj) { localStorage.setItem(IDENT_KEY, JSON.stringify(obj)); }

function updateHeaderSchool() {
  const ident = getIdent();
  const el = $('header-school');
  if (ident.nama) {
    el.textContent = ident.nama;
  } else {
    el.textContent = 'V2.5 — Upload Excel';
  }
}

// ===================== Upload Leger =====================
uploadForm.onsubmit = async (e) => {
  e.preventDefault();
  const file = fileInput.files[0];
  if (!file) { uploadStatus.textContent = 'Pilih file terlebih dahulu'; uploadStatus.className = 'error'; return; }
  const fd = new FormData(); fd.append('file', file);
  showLoading('Membaca data leger...');
  uploadStatus.textContent = ''; uploadStatus.className = '';
  try {
    const res = await fetch('/api/upload', { method: 'POST', body: fd });
    const data = await res.json();
    if (!data.success) throw new Error(data.error);
    hideLoading();
    uploadStatus.innerHTML = 'Berhasil: <strong>' + (data.kelas || []).length + '</strong> kelas terbaca.';
    uploadStatus.className = 'success';
    btnLoadRapor.style.display = '';
  } catch (err) {
    hideLoading();
    uploadStatus.textContent = 'Error: ' + err.message; uploadStatus.className = 'error';
  }
};

async function loadData() {
  showLoading('Memuat data...');
  try {
    const res = await fetch('/api/data', { cache: 'no-store' });
    const data = await res.json();
    if (data.error) throw new Error(data.error);
    db = data;
    defaults = data.defaults || defaults;
    hideLoading();
    setupSection.style.display = 'none';
    workspace.style.display = 'grid';
    document.body.classList.add('workspace-active');
    btnSettings.style.display = '';
    $('btn-upload-lagi').style.display = '';
    populateKelas(data.daftarKelas);
  } catch (err) {
    hideLoading();
    uploadStatus.textContent = 'Gagal muat data: ' + err.message;
    uploadStatus.className = 'error';
  }
}

// ===================== Buka Rapor =====================
btnLoadRapor.onclick = loadData;

// ===================== Daftar Siswa =====================
function populateKelas(daftar) {
  kelasSelect.innerHTML = '<option value="">-- Pilih Kelas --</option>';
  daftar.forEach(k => {
    const o = document.createElement('option');
    o.value = k; o.textContent = 'Kelas ' + k;
    kelasSelect.appendChild(o);
  });
  kelasSelect.onchange = onKelasChange;
  searchInput.oninput = renderList;
}

function onKelasChange() {
  currentKelas = kelasSelect.value;
  searchInput.value = '';
  if (!currentKelas) {
    currentSiswaList = []; siswaTbody.innerHTML = '';
    listHint.style.display = 'block'; searchInput.disabled = true;
    btnPrintKelas.disabled = true;
    return;
  }
  currentSiswaList = (db.kelas[currentKelas] || []).slice();
  searchInput.disabled = false; listHint.style.display = 'none';
  btnPrintKelas.disabled = false;
  renderList();
}

function renderList() {
  const q = (searchInput.value || '').toLowerCase().trim();
  siswaTbody.innerHTML = '';
  const rows = currentSiswaList.filter(s => !q || s.nama.toLowerCase().includes(q));
  visibleList = rows;
  if (!rows.length) {
    siswaTbody.innerHTML = '<tr><td colspan="4" style="text-align:center;color:#8a97a3;padding:18px">Tidak ada siswa.</td></tr>';
    return;
  }
  rows.forEach(s => {
    const tr = document.createElement('tr');
    const ni = s.nisn || String(s.noUrut);
    tr.dataset.ni = ni;
    tr.innerHTML =
      '<td class="c-no">' + s.noUrut + '</td>' +
      '<td>' + s.nama + '</td>' +
      '<td class="c-nisn">' + (s.nisn || '-') + '</td>' +
      '<td class="c-act"><button class="btn-mini">Lihat</button></td>';
    tr.onclick = () => openRapor(ni, tr);
    siswaTbody.appendChild(tr);
  });
  if (currentNi) {
    const act = siswaTbody.querySelector('tr[data-ni="' + CSS.escape(currentNi) + '"]');
    if (act) act.classList.add('active');
  }
  updateNav();
}

function openRapor(ni, tr) {
  currentNi = ni;
  siswaTbody.querySelectorAll('tr.active').forEach(t => t.classList.remove('active'));
  if (tr) tr.classList.add('active');
  previewEmpty.style.display = 'none';
  previewScaler.style.display = 'block';
  previewFrame.onload = () => { fitPreview(); setTimeout(fitPreview, 300); setTimeout(fitPreview, 900); };
  const ident = getIdent();
  var qs = '?semester=' + encodeURIComponent(ident.semester || '2') + '&tahunAjaran=' + encodeURIComponent(ident.tahunAjaran || '');
  previewFrame.src = '/rapor/' + encodeURIComponent(currentKelas) + '/' + encodeURIComponent(ni) + qs;
  const s = currentSiswaList.find(x => (x.nisn || String(x.noUrut)) === ni);
  previewTitle.textContent = s ? s.nama + ' — ' + currentKelas : 'Pratinjau Rapor';
  btnPrint.disabled = false; btnNewtab.disabled = false;
  updateNav();
}

// ===================== Navigasi antar siswa =====================
function niOf(s) { return s.nisn || String(s.noUrut); }
function updateNav() {
  const idx = visibleList.findIndex(s => niOf(s) === currentNi);
  if (idx < 0) {
    navPos.textContent = '–'; btnPrev.disabled = true; btnNext.disabled = true; return;
  }
  navPos.textContent = (idx + 1) + ' / ' + visibleList.length;
  btnPrev.disabled = idx <= 0;
  btnNext.disabled = idx >= visibleList.length - 1;
}
function gotoOffset(d) {
  const idx = visibleList.findIndex(s => niOf(s) === currentNi);
  if (idx < 0) return;
  const ni = niOf(visibleList[Math.max(0, Math.min(visibleList.length - 1, idx + d))]);
  const tr = siswaTbody.querySelector('tr[data-ni="' + CSS.escape(ni) + '"]');
  openRapor(ni, tr);
  if (tr) tr.scrollIntoView({ block: 'nearest' });
}
btnPrev.onclick = () => gotoOffset(-1);
btnNext.onclick = () => gotoOffset(1);
document.addEventListener('keydown', (e) => {
  if (modal.style.display === 'flex') return;
  const tag = (e.target.tagName || '').toLowerCase();
  if (tag === 'input' || tag === 'select' || tag === 'textarea') return;
  if (!currentNi) return;
  if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') { e.preventDefault(); gotoOffset(-1); }
  else if (e.key === 'ArrowRight' || e.key === 'ArrowDown') { e.preventDefault(); gotoOffset(1); }
});

// ===================== Fit-to-width preview =====================
function fitPreview() {
  const doc = previewFrame.contentDocument;
  if (!doc || !doc.body) return;
  const sheet = doc.querySelector('.sheet');
  const LW = sheet ? sheet.offsetWidth : 794;
  const H = doc.documentElement.scrollHeight;
  const avail = previewBody.clientWidth - 24;
  let scale = avail / LW;
  if (scale > 1) scale = 1;
  previewFrame.style.width = LW + 'px';
  previewFrame.style.height = H + 'px';
  previewFrame.style.transform = 'scale(' + scale + ')';
  previewScaler.style.width = (LW * scale) + 'px';
  previewScaler.style.height = (H * scale) + 'px';
  $('zoom-label').textContent = Math.round(scale * 100) + '%';
}
let fitTimer;
window.addEventListener('resize', () => { clearTimeout(fitTimer); fitTimer = setTimeout(fitPreview, 120); });

btnPrint.onclick = () => { if (previewFrame.src) { previewFrame.contentWindow.focus(); previewFrame.contentWindow.print(); } };
let printFrame = null;
btnPrintKelas.onclick = () => {
  if (!currentKelas) return;
  showLoading('Menyiapkan cetak ' + (currentSiswaList.length || '') + ' rapor...');
  if (printFrame) printFrame.remove();
  printFrame = document.createElement('iframe');
  printFrame.style.cssText = 'position:fixed; left:-10000px; top:0; width:210mm; height:297mm; border:0;';
  document.body.appendChild(printFrame);
  printFrame.onload = () => {
    const win = printFrame.contentWindow;
    let last = -1, stable = 0, ticks = 0;
    const timer = setInterval(() => {
      ticks++;
      const root = win.document && win.document.getElementById('rapor-root');
      const n = root ? root.children.length : 0;
      if (n > 0 && n === last) stable++; else { stable = 0; last = n; }
      if ((stable >= 3) || ticks > 60) {
        clearInterval(timer);
        hideLoading();
        try { win.focus(); win.print(); } catch (e) {}
      }
    }, 250);
  };
  const ident = getIdent();
  var qs = '?semester=' + encodeURIComponent(ident.semester || '2') + '&tahunAjaran=' + encodeURIComponent(ident.tahunAjaran || '');
  printFrame.src = '/rapor-batch/' + encodeURIComponent(currentKelas) + qs;
};
btnNewtab.onclick = () => { if (previewFrame.src) window.open(previewFrame.src, '_blank'); };

$('btn-upload-lagi').onclick = () => {
  workspace.style.display = 'none';
  setupSection.style.display = 'block';
  document.body.classList.remove('workspace-active');
  btnSettings.style.display = 'none';
  $('btn-upload-lagi').style.display = 'none';
  btnLoadRapor.style.display = 'none';
  currentKelas = '';
  currentNi = null;
  db = null;
  uploadStatus.textContent = '';
  uploadStatus.className = '';
};

// reload preview
let reloadTimer;
function reloadPreview(immediate) {
  if (!currentNi || previewScaler.style.display === 'none') return;
  clearTimeout(reloadTimer);
  const go = () => { previewFrame.onload = () => { fitPreview(); setTimeout(fitPreview, 300); }; previewFrame.src = previewFrame.src; };
  if (immediate) go(); else reloadTimer = setTimeout(go, 350);
}

// ===================== Modal Pengaturan =====================
const modal = $('settings-modal');
const wkKelas = $('wk-kelas');

// ===== Tab switching =====
document.querySelectorAll('.tab').forEach(tab => {
  tab.onclick = () => {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    tab.classList.add('active');
    document.querySelectorAll('.tab-pane').forEach(p => p.style.display = 'none');
    const pane = $('pane-' + tab.dataset.tab);
    if (pane) pane.style.display = 'block';
  };
});

btnSettings.onclick = openSettings;
$('btn-close-settings').onclick = closeSettings;
$('btn-save-settings').onclick = closeSettings;
modal.onclick = (e) => { if (e.target === modal) closeSettings(); };
function closeSettings() { modal.style.display = 'none'; updateHeaderSchool(); reloadPreview(true); }

function openSettings() {
  // --- Load identitas sekolah ---
  const ident = getIdent();
  if (!ident.semester) ident.semester = '2';
  if (!ident.tahunAjaran) {
    var d = new Date(), y = d.getFullYear();
    ident.tahunAjaran = d.getMonth() >= 7 ? y + '/' + (y + 1) : (y - 1) + '/' + y;
  }
  setIdent(ident);
  $('id-nama').value = ident.nama || '';
  $('id-alamat').value = ident.alamat || '';
  $('id-kota').value = ident.kota || '';
  $('id-kepsek').value = ident.kepsek || '';
  $('id-nip-kepsek').value = ident.nipKepsek || '';
  $('id-semester').value = ident.semester;
  $('id-tahun').value = ident.tahunAjaran;
  if (ident.logo) {
    $('logo-img').src = ident.logo;
    $('logo-preview').style.display = 'flex';
  } else {
    $('logo-preview').style.display = 'none';
  }

  // --- Load TTD Kepsek & tanggal ---
  const cfg = ensureCfg();
  const ks = cfg.kepsek || { name: '', sig: '', w: 110, x: 0, y: 0 };
  $('ks-w').value = ks.w || 110; $('ks-w-val').textContent = ks.w || 110;
  renderStage('ks', ks);
  $('tgl-rapor').value = cfg.tanggalRaporISO || DEF_TGL_RAPOR;

  // --- Load wali kelas ---
  wkKelas.innerHTML = '';
  (db ? db.daftarKelas : []).forEach(k => {
    const o = document.createElement('option');
    o.value = k; o.textContent = 'Kelas ' + k;
    wkKelas.appendChild(o);
  });
  if (currentKelas) wkKelas.value = currentKelas;
  loadWkFields();

  modal.style.display = 'flex';
}

// ===== Identitas Sekolah — auto-save =====
['id-nama','id-alamat','id-kota','id-kepsek','id-nip-kepsek','id-tahun'].forEach(id => {
  $(id).oninput = function() {
    const ident = getIdent();
    ident[id.replace('id-','')] = this.value;
    setIdent(ident);
  };
});
$('id-semester').onchange = function() {
  const ident = getIdent();
  ident.semester = this.value;
  setIdent(ident);
};
$('id-logo').onchange = function(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(ev) {
    const dataUrl = ev.target.result;
    const ident = getIdent();
    ident.logo = dataUrl;
    setIdent(ident);
    $('logo-img').src = dataUrl;
    $('logo-preview').style.display = 'flex';
  };
  reader.readAsDataURL(file);
};
$('btn-logo-hapus').onclick = function() {
  const ident = getIdent();
  delete ident.logo;
  setIdent(ident);
  $('logo-preview').style.display = 'none';
  $('id-logo').value = '';
};

// ===== TTD Kepsek =====
$('ks-sig-file').onchange = (e) => readSig(e, (url) => {
  const { cfg, obj } = targetSlice('ks');
  obj.sig = url; setCfg(cfg); renderStage('ks', obj); reloadPreview();
});
$('ks-w').oninput = () => {
  const { cfg, obj } = targetSlice('ks');
  obj.w = num($('ks-w').value, 110); $('ks-w-val').textContent = obj.w;
  setCfg(cfg); renderStage('ks', obj); reloadPreview();
};
$('ks-center').onclick = () => {
  const { cfg, obj } = targetSlice('ks');
  obj.x = 0; obj.y = 0; setCfg(cfg); renderStage('ks', obj); reloadPreview();
};

// ===== Tanggal =====
$('tgl-rapor').oninput = () => {
  const cfg = ensureCfg();
  cfg.tanggalRaporISO = $('tgl-rapor').value;
  cfg.tanggalKelulusanISO = cfg.tanggalRaporISO;
  setCfg(cfg); reloadPreview();
};

// ===== Wali Kelas =====
function loadWkFields() {
  const cfg = ensureCfg();
  const k = wkKelas.value;
  const w = cfg.wali[k] || { name: '', sig: '', w: 110, x: 0, y: 0 };
  $('wk-name').value = w.name || '';
  $('wk-name').placeholder = (defaults.wali && defaults.wali[k]) || 'Nama wali kelas...';
  $('wk-w').value = w.w || 110; $('wk-w-val').textContent = w.w || 110;
  renderStage('wk', w);
}
wkKelas.onchange = loadWkFields;
$('wk-name').oninput = () => { const { cfg, obj } = targetSlice('wk'); obj.name = $('wk-name').value.trim(); setCfg(cfg); renderStage('wk', obj); reloadPreview(); };
$('wk-w').oninput = () => { const { cfg, obj } = targetSlice('wk'); obj.w = num($('wk-w').value, 110); $('wk-w-val').textContent = obj.w; setCfg(cfg); renderStage('wk', obj); reloadPreview(); };
$('wk-center').onclick = () => { const { cfg, obj } = targetSlice('wk'); obj.x = 0; obj.y = 0; setCfg(cfg); renderStage('wk', obj); reloadPreview(); };
$('wk-sig-file').onchange = (e) => readSig(e, (url) => { const { cfg, obj } = targetSlice('wk'); obj.sig = url; setCfg(cfg); renderStage('wk', obj); reloadPreview(); });

// ===== Clear TTD (ks & wk) =====
document.querySelectorAll('.btn-clear-sig').forEach(btn => {
  btn.onclick = () => {
    const which = btn.dataset.target;
    const { cfg, obj } = targetSlice(which);
    obj.sig = ''; setCfg(cfg); renderStage(which, obj);
    $(which + '-sig-file').value = ''; reloadPreview();
  };
});

function readSig(e, cb) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => cb(reader.result);
  reader.readAsDataURL(file);
}

function tanggalRapor() {
  const t = new Date(), m = t.getMonth() + 1;
  return m >= 7 ? '20 Desember ' + t.getFullYear() : '20 Juni ' + t.getFullYear();
}

function renderStage(which, c) {
  const stage = $(which + '-stage');
  stage.innerHTML = '';
  stage.classList.toggle('has-img', !!(c && c.sig));

  const top = document.createElement('div');
  top.className = 'stage-top';
  const nameEl = document.createElement('div');
  nameEl.className = 'stage-name';

  const identKota = getIdent().kota || 'Surakarta';
  if (which === 'wk') {
    top.innerHTML = identKota + ', ' + tanggalRapor() + '<br>Wali Kelas,';
    const nm = $('wk-name').value.trim() || (defaults.wali && defaults.wali[$('wk-kelas').value]) || '(nama wali kelas)';
    nameEl.innerHTML = '<span class="sig-id">' + esc(nm) + '<br><small>NIP. -</small></span>';
  } else {
    const nipKepsek = getIdent().nipKepsek || '-';
    top.innerHTML = 'Kepala Sekolah';
    const nm = $('id-kepsek').value.trim() || defaults.kepsek || '(nama kepala sekolah)';
    nameEl.innerHTML = '<span class="sig-id">' + esc(nm) + '<br><small>NIP. ' + esc(nipKepsek) + '</small></span>';
  }

  const slot = document.createElement('div');
  slot.className = 'stage-slot';
  if (c && c.sig) {
    const img = document.createElement('img');
    img.src = c.sig;
    applyImgStyle(img, c);
    attachDrag(img, which);
    slot.appendChild(img);
  } else {
    const hint = document.createElement('span');
    hint.className = 'stage-hint';
    hint.textContent = 'Belum ada tanda tangan';
    slot.appendChild(hint);
  }

  stage.appendChild(top);
  stage.appendChild(slot);
  stage.appendChild(nameEl);
}
function applyImgStyle(img, c) {
  img.style.width = (c.w || 110) + 'px';
  img.style.transform = 'translate(calc(-50% + ' + (c.x || 0) + 'px), calc(-50% + ' + (c.y || 0) + 'px))';
}
function attachDrag(img, which) {
  img.onpointerdown = (e) => {
    e.preventDefault();
    const { cfg, obj } = targetSlice(which);
    const sx = e.clientX, sy = e.clientY, bx = obj.x || 0, by = obj.y || 0;
    img.classList.add('dragging');
    try { img.setPointerCapture(e.pointerId); } catch (_) {}
    img.onpointermove = (ev) => {
      obj.x = Math.round(bx + (ev.clientX - sx));
      obj.y = Math.round(by + (ev.clientY - sy));
      applyImgStyle(img, obj);
    };
    img.onpointerup = () => {
      img.classList.remove('dragging');
      img.onpointermove = null; img.onpointerup = null;
      setCfg(cfg);
      reloadPreview();
    };
  };
}

// ===================== Checklist dokumen =====================
function loadDocChecks() {
  const cfg = ensureCfg();
  const d = cfg.docs || { rapor: true, mutasi: false };
  $('doc-rapor').checked = d.rapor !== false;
  $('doc-mutasi').checked = !!d.mutasi;
}
function saveDocChecks() {
  const cfg = ensureCfg();
  cfg.docs = {
    rapor: $('doc-rapor').checked,
    mutasi: $('doc-mutasi').checked
  };
  setCfg(cfg);
  reloadPreview(true);
}
$('doc-rapor').onchange = saveDocChecks;
$('doc-mutasi').onchange = saveDocChecks;
loadDocChecks();

// ===================== Init =====================
updateHeaderSchool();
