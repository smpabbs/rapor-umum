const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const cors = require('cors');
const app = express();
const PORT = process.env.PORT || 3005;
// V2.5: semua lokal — upload Excel adalah satu-satunya sumber data.
const WRITABLE_DIR = process.env.VERCEL ? require('os').tmpdir() : __dirname;
const UPLOAD_DIR = path.join(WRITABLE_DIR, 'uploads');
try { fs.mkdirSync(UPLOAD_DIR, { recursive: true }); } catch (e) {}

app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));

const upload = multer({
  dest: UPLOAD_DIR,
  limits: { fileSize: 50 * 1024 * 1024 }
});

let db = null;

const SHEET_MAP = {};
for (let k = 7; k <= 9; k++) {
  for (let i = 0; i < 6; i++) {
    const letter = String.fromCharCode(65 + i);
    SHEET_MAP[`Total NIlai ${k}${letter}`] = `${k}${letter}`;
    SHEET_MAP[`Total Nilai ${k}${letter}`] = `${k}${letter}`;
    SHEET_MAP[`Total NIlai ${k}${letter}$`] = `${k}${letter}`;
    SHEET_MAP[`Total Nilai ${k}${letter}$`] = `${k}${letter}`;
  }
}

// Nama wali kelas per kelas — sesuaikan jika ada perubahan
const WALI_KELAS_MAP = {
  '7A': 'Yoki Wirawan, S.Pd.',
  '7B': 'Yoki Wirawan, S.Pd.',
  '7C': 'Yoki Wirawan, S.Pd.',
  '7D': 'Yoki Wirawan, S.Pd.',
  '7E': 'Yoki Wirawan, S.Pd.',
  '7F': 'Yoki Wirawan, S.Pd.',
  '8A': 'Yoki Wirawan, S.Pd.',
  '8B': 'Yoki Wirawan, S.Pd.',
  '8C': 'Yoki Wirawan, S.Pd.',
  '8D': 'Yoki Wirawan, S.Pd.',
  '8E': 'Yoki Wirawan, S.Pd.',
  '8F': 'Yoki Wirawan, S.Pd.',
  '9A': 'Yoki Wirawan, S.Pd.',
  '9B': 'Yoki Wirawan, S.Pd.',
  '9C': 'Yoki Wirawan, S.Pd.',
  '9D': 'Yoki Wirawan, S.Pd.',
  '9E': 'Yoki Wirawan, S.Pd.',
  '9F': 'Fatma Roudhotul Rafida Kolis, S.Pd.',
};

function normalizeKey(k) {
  return k.replace(/[\s\-_\/]/g, '').toLowerCase();
}

// Cocokkan nama sheet ke kode kelas (7A..9F) secara toleran.
// Dukung variasi penamaan: "Total Nilai 7A", "Total NIlai 7A", "Nilai 7 A", "7A", dll.
function resolveClassName(sheetName) {
  if (!sheetName) return '';
  if (SHEET_MAP[sheetName]) return SHEET_MAP[sheetName];
  const m = String(sheetName).toUpperCase().match(/(?:^|[^0-9A-Z])([7-9])\s*([A-F])(?![0-9A-Z])/);
  return m ? m[1] + m[2] : '';
}

function parseSheetData(ws, sheetName) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (rows.length < 3) return [];

  const className = resolveClassName(sheetName);
  if (!className) return [];

  const students = [];
  let headerRow = null;
  let headerRowIdx = -1;

  for (let i = 0; i < Math.min(5, rows.length); i++) {
    const row = rows[i];
    if (!row) continue;
    for (let j = 0; j < row.length; j++) {
      const v = String(row[j] || '').trim();
      if (normalizeKey(v) === 'nama') {
        headerRow = row;
        headerRowIdx = i;
        break;
      }
    }
    if (headerRow) break;
  }

  if (!headerRow) {
    headerRow = rows[1] || rows[0];
    headerRowIdx = 1;
  }

  const colMap = {};
  const colMapAll = {}; // semua indeks per key — untuk header yang bentrok (mis. 'predikat' muncul 2×)
  headerRow.forEach((h, i) => {
    if (h && String(h).trim()) {
      const key = normalizeKey(String(h));
      colMap[key] = i; // kolom terakhir menang (kompatibel dgn perilaku lama)
      (colMapAll[key] = colMapAll[key] || []).push(i);
    }
  });

  // Pilih kolom 'predikat' yang paling dekat SETELAH kolom nama blok terkait,
  // agar predikat SD (kolom 33) tidak tertukar dgn predikat Pramuka (kolom 37).
  function predikatAfter(afterIdx) {
    const arr = colMapAll[normalizeKey('predikat')] || [];
    let best;
    for (const idx of arr) {
      if (idx > afterIdx && (best === undefined || idx < best)) best = idx;
    }
    return best;
  }

  // === Kolom blok Pramuka/PMR (dideteksi sekali dari header, toleran variasi antar kelas) ===
  // Pola umum: [opsional NAMA], NILAI, Predikat, Deskripsi.
  // Sebagian kelas tak punya kolom NAMA; kolom Deskripsi paling andal terdeteksi,
  // jadi Predikat Pramuka = kolom 'predikat' tepat SEBELUM kolom Deskripsi Pramuka.
  let pramNameIdx, pramNilaiIdx, pramDescIdx;
  headerRow.forEach((h, i) => {
    const t = String(h || '').toLowerCase();
    if (!/pramuka|pmr|kepramukaan/.test(t)) return;
    if (t.includes('deskripsi') || t.includes('des_') || t.includes('des ')) { if (pramDescIdx === undefined) pramDescIdx = i; }
    else if (t.includes('nilai')) { if (pramNilaiIdx === undefined) pramNilaiIdx = i; }
    else { if (pramNameIdx === undefined) pramNameIdx = i; } // kolom nama murni: "PRAMUKA /PMR", "Kepramukaan/PMR"
  });
  let pramPredIdx;
  if (pramDescIdx !== undefined && pramDescIdx - 1 >= 0
      && normalizeKey(String(headerRow[pramDescIdx - 1] || '')) === 'predikat') {
    pramPredIdx = pramDescIdx - 1;
  } else {
    const anchor = pramNilaiIdx !== undefined ? pramNilaiIdx : pramNameIdx;
    if (anchor !== undefined) pramPredIdx = predikatAfter(anchor);
  }
  const hasPramuka = pramNameIdx !== undefined || pramNilaiIdx !== undefined || pramDescIdx !== undefined;

  const dataStart = headerRowIdx + 1;
  for (let r = dataStart; r < rows.length; r++) {
    const row = rows[r];
    if (!row || !row[0]) continue;

    const firstVal = String(row[0]).trim();
    if (!firstVal || firstVal === 'No') continue;

    const noUrut = parseInt(firstVal);
    if (isNaN(noUrut)) {
      if (['rata', 'total', 'jumlah', 'rerata', 'aver', 'sum'].some(s => firstVal.toLowerCase().includes(s))) break;
      continue;
    }

    const nisIdx  = colMap[normalizeKey('NIS')];
    const nisnIdx = colMap[normalizeKey('NISN')];
    const namaIdx = colMap[normalizeKey('Nama')];

    if (namaIdx === undefined) continue;
    const nama = String(row[namaIdx] || '').trim();
    if (!nama) continue;

    const nis  = nisIdx  !== undefined ? String(row[nisIdx]  || '').trim() : '';
    const nisn = nisnIdx !== undefined ? String(row[nisnIdx] || '').trim() : '';

    const student = {
      noUrut, nis, nisn, nama, kelas: className,
      mapel: {}, ekstrakurikuler: [], kokurikuler: '',
      ketidakhadiran: { sakit: 0, izin: 0, tanpaKeterangan: 0 },
      catatanWalas: '', rataRata: 0
    };

    const mapelDefs = [
      { id: 'PAI', colNilai: 'PAI', colDesc: 'des_PAI' },
      { id: 'Pkn', colNilai: 'Pkn', colDesc: 'des_Pkn' },
      { id: 'Indo', colNilai: 'Indo', colDesc: 'des_indo' },
      { id: 'Math', colNilai: 'Math', colDesc: 'des_Math' },
      { id: 'IPA', colNilai: 'IPA', colDesc: 'des_IPA' },
      { id: 'IPS', colNilai: 'IPS', colDesc: 'des_IPS' },
      { id: 'English', colNilai: 'English', colDesc: 'des_english' },
      { id: 'SBK', colNilai: 'NILAI_SBK', colDesc: 'des_SBK' },
      { id: 'Sport', colNilai: 'Sport', colDesc: 'des_Sport' },
      { id: 'ICT', colNilai: 'ICT', colDesc: 'des_ICT' },
      { id: 'Jawa', colNilai: 'jawa', colDesc: 'des_jawa' },
    ];

    for (const def of mapelDefs) {
      const nilaiCol = colMap[normalizeKey(def.colNilai)];
      const descCol = colMap[normalizeKey(def.colDesc)];
      let nilai = 0, deskripsi = '';
      if (nilaiCol !== undefined && row[nilaiCol]) {
        nilai = parseFloat(String(row[nilaiCol]).trim()) || 0;
      }
      if (descCol !== undefined && row[descCol]) {
        deskripsi = String(row[descCol]).trim();
      }
      student.mapel[def.id] = { nilai, deskripsi };
    }

    // Catatan: SBK (NAMA_SBK/des_SBK) TIDAK dimasukkan ke ekstrakurikuler —
    // SBK hanya muncul sebagai mapel intrakurikuler "Seni dan Budaya".

    const sdNameIdx = colMap[normalizeKey('NAMA_SD')];
    if (sdNameIdx !== undefined && row[sdNameIdx]) {
      const name = String(row[sdNameIdx]).trim();
      let desc = '', predikat = '', nilai = 0;
      const predIdx = predikatAfter(sdNameIdx);
      if (predIdx !== undefined && row[predIdx]) {
        const raw = String(row[predIdx]).trim();
        if (isNaN(parseFloat(raw))) predikat = raw;
        else nilai = parseFloat(raw);
      }
      const sdDescIdx = colMap[normalizeKey('des_SD')];
      if (sdDescIdx !== undefined && row[sdDescIdx]) {
        const raw = String(row[sdDescIdx]).trim();
        const m = raw.match(/^(\d+(?:\.\d+)?)\s+(.*)/);
        if (m) { if (!nilai) nilai = parseFloat(m[1]); desc = m[2].trim(); }
        else if (!predikat && !isNaN(parseFloat(raw))) nilai = parseFloat(raw);
        else desc = raw;
      }
      student.ekstrakurikuler.push({ nama: name, nilai, deskripsi: desc, predikat });
    }

    if (hasPramuka) {
      // Nama kegiatan dari kolom NAMA Pramuka (mis. "Kepramukaan"). Bila kolom NAMA kosong,
      // nama DIBIARKAN KOSONG mengikuti leger — TIDAK dipaksa jadi label "Pramuka / PMR".
      let nama = (pramNameIdx !== undefined && row[pramNameIdx]) ? String(row[pramNameIdx]).trim() : '';
      let predikat = '', nilai = 0;
      if (pramPredIdx !== undefined && row[pramPredIdx]) {
        const raw = String(row[pramPredIdx]).trim();
        if (isNaN(parseFloat(raw))) predikat = raw; else nilai = parseFloat(raw);
      }
      if (pramNilaiIdx !== undefined && row[pramNilaiIdx]) {
        const n = parseFloat(String(row[pramNilaiIdx]).trim());
        if (!isNaN(n)) nilai = n;
      }
      const desc = (pramDescIdx !== undefined && row[pramDescIdx]) ? String(row[pramDescIdx]).trim() : '';
      // Tampilkan baris bila ADA isi di salah satu kolom (nama/predikat/deskripsi). Kolom yang
      // kosong dibiarkan kosong; baris tidak dihilangkan. (Kelas 7/8: kolom NAMA kosong → sel Kegiatan
      // kosong, tapi Predikat & Keterangan tetap tampil. Kelas 9: nama "Kepramukaan" tetap tampil.)
      if (nama || predikat || desc) {
        student.ekstrakurikuler.push({ nama, nilai, deskripsi: desc, predikat });
      }
    }

    const kokuIdx = colMap[normalizeKey('DeskripsiKokurikuler')] ?? colMap[normalizeKey('deskripsikokurikuler')];
    if (kokuIdx !== undefined && row[kokuIdx]) {
      student.kokurikuler = String(row[kokuIdx]).trim();
    }

    const sakitIdx = colMap[normalizeKey('SAKIT')] ?? colMap[normalizeKey('sakit')] ?? colMap[normalizeKey('SAKIT ')] ?? colMap[normalizeKey('sakit ')];
    if (sakitIdx !== undefined && row[sakitIdx]) {
      const v = String(row[sakitIdx]).trim();
      if (v !== '-' && v !== '') student.ketidakhadiran.sakit = parseInt(v) || 0;
    }

    const izinIdx = colMap[normalizeKey('IJIN')] ?? colMap[normalizeKey('Izin')] ?? colMap[normalizeKey('IJIN ')] ?? colMap[normalizeKey('izin')];
    if (izinIdx !== undefined && row[izinIdx]) {
      const v = String(row[izinIdx]).trim();
      if (v !== '-' && v !== '') student.ketidakhadiran.izin = parseInt(v) || 0;
    }

    const tkKeys = ['TANPAKETERANGAN', 'TANPA KETERANGAN', 'tanpaketerangan', 'tanpa keterangan', 'TANPA KETERANGAN '];
    for (const k of tkKeys) {
      const idx = colMap[normalizeKey(k)];
      if (idx !== undefined && row[idx]) {
        const v = String(row[idx]).trim();
        if (v !== '-' && v !== '') student.ketidakhadiran.tanpaKeterangan = parseInt(v) || 0;
        break;
      }
    }

    const walasKeys = ['CATATANWALAS', 'CATATAN WALAS', 'catatan walas', 'catatanwalas'];
    for (const k of walasKeys) {
      const idx = colMap[normalizeKey(k)];
      if (idx !== undefined && row[idx]) {
        student.catatanWalas = String(row[idx]).trim();
        break;
      }
    }

    const nilaiMapel = Object.values(student.mapel).map(v => v.nilai).filter(v => v > 0);
    if (nilaiMapel.length > 0) {
      student.rataRata = Math.round((nilaiMapel.reduce((a, b) => a + b, 0) / nilaiMapel.length) * 100) / 100;
    }

    students.push(student);
  }

  return students;
}

// Muat satu workbook (apa pun sumbernya) ke db; kembalikan laporan sheet.
function loadWorkbook(wb) {
  const kelas = {};
  const report = { matched: [], skipped: [] };
  for (const sn of wb.SheetNames) {
    const className = resolveClassName(sn);
    if (!className) { report.skipped.push({ sheet: sn, alasan: 'bukan sheet kelas' }); continue; }
    const students = parseSheetData(wb.Sheets[sn], sn);
    if (students.length > 0) {
      kelas[className] = students;
      report.matched.push({ sheet: sn, kelas: className, siswa: students.length });
    } else {
      report.skipped.push({ sheet: sn, alasan: 'tidak ada data siswa' });
    }
  }
  // Hanya ganti db bila ada data baru — muat-ulang yang gagal tak mengosongkan data lama.
  if (report.matched.length > 0) db = { kelas };
  return report;
}

function loadLeger(filePath) {
  if (!fs.existsSync(filePath)) return false;
  const report = loadWorkbook(XLSX.readFile(filePath));
  return report.matched.length > 0;
}

// ===== V2.5: tidak ada Google Sheets / admin / auth — semua lokal =====

app.get('/api/download-sample', (req, res) => {
  const wb = XLSX.utils.book_new();
  const sample = [
    ['No', 'NISN', 'Nama', 'PAI', 'des_PAI', 'Pkn', 'des_Pkn', 'Indo', 'des_indo', 'Math', 'des_Math', 'IPA', 'des_IPA', 'IPS', 'des_IPS', 'English', 'des_english', 'NILAI_SBK', 'des_SBK', 'Sport', 'des_Sport', 'ICT', 'des_ICT', 'jawa', 'des_jawa', 'NAMA_SD', 'predikat', 'des_SD', 'PRAMUKA /PMR', 'Nilai Pramuka', 'predikat', 'Deskripsi Pramuka', 'DeskripsiKokurikuler', 'SAKIT', 'IZIN', 'TANPA_KETERANGAN', 'Catatan_Walas', 'RATA_RATA'],
    [1, '1234567890', 'Ahmad Fauzi', 90, 'Menguasai materi PAI dg baik', 88, 'Aktif dalam diskusi', 85, 'Mampu menulis & membaca lancar', 92, 'Unggul pemecahan masalah', 87, 'Memahami konsep IPA dg baik', 90, 'Analisis tajam fenomena sosial', 88, 'Membaca & menulis sangat baik', 86, 'Kreatif dalam berkarya', 89, 'Aktif dalam olahraga', 85, 'Menguasai materi dg baik', 90, 'Aktif pembelajar bahasa Jawa', 'Seni Tari', 'A', 'Sangat kreatif dalam seni tari tradisional', 'Kepramukaan', 88, 'A', 'Aktif dan disiplin dalam kegiatan Pramuka', 'Mengikuti program pengayaan bahasa Inggris', 0, 1, 0, 'Ahmad anak yang rajin dan disiplin.', 87.5],
    [2, '1234567891', 'Siti Nurhaliza', 87, 'Menguasai materi PAI', 90, 'Partisipasi aktif', 88, 'Baik dlm menulis & membaca', 85, 'Baik dlm pemecahan masalah', 90, 'Pemahaman IPA baik', 85, 'Analisis sosial baik', 92, 'Kemampuan Inggris unggul', 88, 'Kreatif', 85, 'Aktif berolahraga', 88, 'Baik', 87, 'Aktif', 'Seni Musik', 'A', 'Berbakat dalam seni musik vokal', 'Kepramukaan', 85, 'A', 'Aktif dan bertanggung jawab dalam Pramuka', 'Mengikuti lomba cerdas cermat', 1, 0, 0, 'Siti anak yang pintar dan sopan.', 88.0],
    [3, '1234567892', 'Budi Santoso', 85, 'Cukup menguasai PAI', 82, 'Cukup aktif', 80, 'Cukup lancar membaca', 78, 'Cukup dlm pemecahan masalah', 83, 'Cukup memahami IPA', 85, 'Cukup dlm analisis', 80, 'Cukup', 82, 'Cukup kreatif', 84, 'Cukup aktif', 80, 'Cukup', 82, 'Cukup aktif', 'Seni Lukis', 'B', 'Cukap kreatif dalam seni lukis', 'Kepramukaan', 80, 'B', 'Cukup aktif dalam kegiatan Pramuka', 'Mengikuti remedial matematika', 2, 1, 0, 'Budi perlu meningkatkan kedisiplinan.', 82.0],
  ];
  const ws = XLSX.utils.aoa_to_sheet(sample);
  ws['!cols'] = sample[0].map(() => ({ wch: 14 }));
  XLSX.utils.book_append_sheet(wb, ws, '7A');
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Disposition', 'attachment; filename="sample-leger-rapor.xlsx"');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});

app.post('/api/upload', (req, res) => {
  const uploadSingle = upload.single('file');
  uploadSingle(req, res, (err) => {
    if (err) return res.status(400).json({ error: err.message });
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
    try {
      const success = loadLeger(req.file.path);
      try { fs.unlinkSync(req.file.path); } catch (e) {}
      if (success) res.json({ success: true, kelas: Object.keys(db.kelas).sort() });
      else res.status(400).json({ error: 'No valid data found in file' });
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: err.message });
    }
  });
});

// ===== V2.5: semua rute Sheets/admin dihapus — hanya upload + rapor =====

app.get('/api/data', (req, res) => {
  if (!db) return res.status(503).json({ error: 'Belum ada data. Upload leger terlebih dahulu.', loading: true });
  const kelasOnly = {};
  for (const [kelas, students] of Object.entries(db.kelas)) {
    kelasOnly[kelas] = students.map(s => ({
      noUrut: s.noUrut, nis: s.nis, nisn: s.nisn, nama: s.nama, rataRata: s.rataRata,
      mapel: s.mapel, ekstrakurikuler: s.ekstrakurikuler,
      kokurikuler: s.kokurikuler, ketidakhadiran: s.ketidakhadiran,
      catatanWalas: s.catatanWalas
    }));
  }
  res.json({
    kelas: kelasOnly,
    daftarKelas: Object.keys(db.kelas).sort(),
    defaults: { wali: WALI_KELAS_MAP, kepsek: 'Tri Wijayanti, M.Pd' }
  });
});

app.get('/api/rapor/:kelas/:ni', (req, res) => {
  if (!db) return res.status(400).json({ error: 'No data loaded' });
  const students = db.kelas[req.params.kelas];
  if (!students) return res.status(404).json({ error: 'Class not found' });
  const student = students.find(s => s.nisn === req.params.ni || String(s.noUrut) === req.params.ni);
  if (!student) return res.status(404).json({ error: 'Student not found' });
  const html = generateRaporHTML(student, req.query);
  res.json({ html });
});

app.get('/rapor/:kelas/:ni', (req, res) => {
  if (!db) return res.status(400).send('No data loaded');
  const students = db.kelas[req.params.kelas];
  if (!students) return res.status(404).send('Class not found');
  const student = students.find(s => s.nisn === req.params.ni || String(s.noUrut) === req.params.ni);
  if (!student) return res.status(404).send('Student not found');
  res.set('Content-Type', 'text/html; charset=UTF-8');
  res.send(generateRaporHTML(student, req.query));
});

// Batch: semua siswa satu kelas dalam satu halaman (tiap siswa mulai halaman baru)
app.get('/rapor-batch/:kelas', (req, res) => {
  if (!db) return res.status(400).send('No data loaded');
  const students = db.kelas[req.params.kelas];
  if (!students || !students.length) return res.status(404).send('Class not found or empty');
  res.set('Content-Type', 'text/html; charset=UTF-8');
  res.send(generateRaporBatchHTML(req.params.kelas, students, req.query));
});

// Jalankan listen hanya bila dieksekusi langsung (node server.js). Di Vercel file ini
// di-"require" sebagai serverless function, sehingga listen tidak dipanggil.
if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
    if (db) console.log('Data ready: ' + Object.keys(db.kelas).length + ' classes');
  });
}

// Ekspor app sebagai handler serverless (Vercel). Utilitas pengujian ditempel sbg properti.
module.exports = app;
module.exports.utils = { parseSheetData, loadLeger, resolveClassName, generateRaporHTML, generateRaporBatchHTML, get db() { return db; } };

// Gabungkan beberapa siswa ke satu dokumen dgn memakai ulang shell (style+script)
// dari rapor tunggal, lalu menyisipkan blok .rapor-source siswa lainnya.
function generateRaporBatchHTML(kelas, students, opts) {
  opts = opts || {};
  // Ambil seluruh blok siswa: dari .rapor-source sampai akhir .mutasi-source.
  const SRC_START = '<div class="rapor-source"';
  const SRC_END = '</div><!-- #mutasi-source -->';
  function extractSource(html) {
    const a = html.indexOf(SRC_START);
    const b = html.indexOf(SRC_END);
    if (a < 0 || b < 0) return '';
    return html.slice(a, b + SRC_END.length);
  }
  let doc = generateRaporHTML(students[0], opts);
  const extra = students.slice(1).map(s => extractSource(generateRaporHTML(s, opts))).join('\n');
  if (extra) {
    const idx = doc.indexOf(SRC_END) + SRC_END.length;
    doc = doc.slice(0, idx) + '\n' + extra + '\n' + doc.slice(idx);
  }
  const safeKelas = String(kelas).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  doc = doc.replace(/<title>[^<]*<\/title>/, `<title>Rapor Kelas ${safeKelas} (${students.length} siswa)</title>`);
  return doc;
}

function generateRaporHTML(s, opts) {
  opts = opts || {};
  const mapelList = [
    { id: 'PAI', nama: 'Pendidikan Agama Islam dan Budi Pekerti' },
    { id: 'Pkn', nama: 'Pendidikan Pancasila' },
    { id: 'Indo', nama: 'Bahasa Indonesia' },
    { id: 'Math', nama: 'Matematika (Umum)' },
    { id: 'IPA', nama: 'Ilmu Pengetahuan Alam (IPA)' },
    { id: 'IPS', nama: 'Ilmu Pengetahuan Sosial (IPS)' },
    { id: 'English', nama: 'Bahasa Inggris' },
    { id: 'SBK', nama: 'Seni, Budaya dan Prakarya' },
    { id: 'Sport', nama: 'Pendidikan Jasmani, Olahraga, dan Kesehatan' },
    { id: 'ICT', nama: 'Informatika' },
    { id: 'Jawa', nama: 'Bahasa Jawa' },
  ];
  const today = new Date();
  const months = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
  const month = today.getMonth() + 1;
  let semester = opts.semester || '';
  let tahunAjaran = opts.tahunAjaran || '';
  if (!semester) {
    semester = month >= 7 ? '1' : '2';
    tahunAjaran = month >= 7 ? `${today.getFullYear()}/${today.getFullYear() + 1}` : `${today.getFullYear() - 1}/${today.getFullYear()}`;
  }
  // Tanggal default (dapat diatur di Pengaturan → di-override via applyConfig di klien).
  const tanggalRapor = '2 Juni 2026';      // tampil di TTD Wali Kelas (rapor)
  const tanggalKelulusan = '3 Juni 2026';  // tampil di lembar Mutasi (kelulusan kelas 9)
  const waliKelas = WALI_KELAS_MAP[s.kelas] || 'Wali Kelas';
  const nisDisplay = s.nis && s.nisn ? `${s.nis} / ${s.nisn}` : (s.nisn || s.nis || '-');
  function esc(t) { return String(t).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

  const subRows = mapelList.map((m, i) => {
    const d = s.mapel[m.id] || {};
    return `<tr>
      <td class="td-no">${i + 1}</td>
      <td class="td-mapel">${esc(m.nama)}</td>
      <td class="td-nilai">${d.nilai || ''}</td>
      <td class="td-desc">${esc(d.deskripsi || '')}</td>
    </tr>`;
  }).join('\n');

  let ekstraRows = s.ekstrakurikuler.map((e, i) => {
    const pred = e.predikat || '';
    return `<tr>
      <td class="td-no">${i + 1}</td>
      <td class="td-desc">${esc(e.nama)}</td>
      <td class="td-pred">${esc(pred)}</td>
      <td class="td-desc">${esc(e.deskripsi || '')}</td>
    </tr>`;
  }).join('\n');
  if (!ekstraRows) ekstraRows = '<tr><td class="td-no"></td><td class="td-desc"></td><td class="td-pred"></td><td class="td-desc"></td></tr>';

  const sk = s.ketidakhadiran.sakit || '-';
  const iz = s.ketidakhadiran.izin || '-';
  const tk = s.ketidakhadiran.tanpaKeterangan || '-';

  // ===== Halaman Mutasi (Keterangan Pindah Sekolah / KELUAR) =====
  const gradeNum = parseInt(String(s.kelas)) || 0;
  const romawiMap = { 7: 'VII (Tujuh)', 8: 'VIII (Delapan)', 9: 'IX (Sembilan)' };
  const kelasRomawi = romawiMap[gradeNum] || '';
  const isLulus = gradeNum === 9; // baris pertama terisi LULUS hanya untuk kelas 9 (kelulusan)

  // ===== Kotak Keputusan / Kelulusan (bawah rapor) =====
  // Hanya muncul di SEMESTER 2 (kenaikan akhir tahun). `semester` sudah mengikuti pilihan user.
  // Semester 1 → kosong (tak ada kotak) untuk semua jenjang.
  let keputusanBox = '';
  if (semester === '2') {
    if (gradeNum === 9) {
      // Kelas 9: tetap kotak "Keterangan Kelulusan : Lulus" (rata tengah, tebal)
      keputusanBox = `<table class="data-table" style="margin-top:8pt"><tr><td class="box-cell" style="border:1px solid #000;min-height:22pt;text-align:center;vertical-align:middle;font-weight:bold">Keterangan Kelulusan : Lulus</td></tr></table>`;
    } else if (gradeNum === 7 || gradeNum === 8) {
      // Kelas 7/8: kotak "Keputusan ... NAIK KE KELAS : <tingkat berikutnya>" (kapital)
      const naikKe = (romawiMap[gradeNum + 1] || '').toUpperCase();
      keputusanBox = `<table class="data-table" style="margin-top:8pt"><tr><td class="box-cell" style="border:1px solid #000;text-align:left;vertical-align:middle;padding:6pt">Keputusan :<br>Berdasarkan hasil yang dicapai pada semester 1 dan 2, peserta didik ditetapkan :<br><span style="font-weight:bold">NAIK KE KELAS : ${naikKe}</span></td></tr></table>`;
    }
  }
  const dots = '.............................';
  const barisLulus = `<tr>
      <td style="text-align:center"><span class="sklh-kota"></span>,<br><span class="tgl-kelulusan">${tanggalKelulusan}</span></td>
      <td style="text-align:center">${kelasRomawi}</td>
      <td style="text-align:center;font-weight:bold">LULUS</td>
      <td class="mutasi-ttd">
        <span class="sklh-kota"></span>, <span class="tgl-kelulusan">${tanggalKelulusan}</span><br>Kepala Sekolah,
        <div class="sig-slot ks-sig"></div>
        <span class="sig-id"><span class="ks-name">Tri Wijayanti, M.Pd</span><br><span class="ks-nip">NIP. -</span></span>
        <div class="mutasi-ortu">Orang Tua/Wali,<br><br>${dots}</div>
      </td>
    </tr>`;
  const barisKosong = `<tr>
      <td></td><td></td><td></td>
      <td class="mutasi-ttd">
        ${dots}<br>Kepala Sekolah,<br><br><br>
        <span class="sig-id">${dots}<br><span class="sig-nip">NIP.</span></span>
        <div class="mutasi-ortu">Orang Tua/Wali,<br><br>${dots}</div>
      </td>
    </tr>`;
  const mutasiRows = (isLulus ? barisLulus : barisKosong) + barisKosong + barisKosong;

  return `<!DOCTYPE html>
<html lang="id"><head><meta charset="UTF-8"><title>Rapor - ${esc(s.nama)} - ${s.kelas}</title><style>
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; }

  /* ===== LEMBAR / SHEET A4 (dipakai screen & print sama persis) ===== */
  .sheet {
    width: 210mm; height: 297mm; padding: 1.5cm;
    background: #fff; position: relative; overflow: hidden;
    display: flex; flex-direction: column;
    font-family: Arial, Helvetica, sans-serif;
    font-size: 10pt; line-height: 1.2; color: #000;
  }
  .sheet-header { flex: 0 0 auto; }
  .sheet-content { flex: 1 1 auto; min-height: 0; overflow: hidden; }
  .sheet-footer {
    flex: 0 0 auto; display: flex; justify-content: space-between; align-items: flex-end;
    font-family: Arial, Helvetica, sans-serif; font-size: 8pt; color: #000;
    border-top: 1px solid #000; padding-top: 3pt; margin-top: 4pt;
  }

  @page { size: A4; margin: 0; }
  @media print {
    body { background: #fff; }
    .sheet { box-shadow: none; margin: 0; break-after: page; page-break-after: always; }
    .sheet:last-child { break-after: auto; page-break-after: auto; }
  }
  @media screen {
    body { background: #c8c8c8; padding: 20px 0; }
    .sheet { margin: 0 auto 20px auto; box-shadow: 0 2px 10px rgba(0,0,0,0.30); }
  }
  .rapor-source { display: none; }

  /* ===== Komponen isi ===== */
  /* Info header (kop) - tanpa border, 2 kolom */
  .info-hdr { width: 100%; margin-bottom: 4pt; border-collapse: collapse; table-layout: fixed; }
  .info-hdr td { border: none !important; padding: 1pt 0; font-size: 10pt; vertical-align: top; }
  .info-hdr .lbl { width: 76pt; }
  .info-hdr .sep-col { width: 8pt; }
  .info-hdr .val { width: 230pt; }     /* kolom kiri kini memuat Alamat → diperlebar */
  .info-hdr .lbl2 { width: 88pt; }
  .info-hdr .sep-col2 { width: 8pt; }
  .info-hdr .val2 { width: 100pt; }    /* kolom kanan kini hanya nilai pendek */

  .kop-sep { border: none; border-top: 2px solid #000; margin: 5pt 0 4pt 0; width: 100%; }
  .kop-title { font-size: 13pt; font-weight: bold; text-align: center; margin: 4pt 0 8pt 0; letter-spacing: 1px; }

  /* Section titles - rata kiri, bold */
  .section-title { font-weight: bold; font-size: 10pt; text-align: left; margin: 8pt 0 2pt 0; }

  /* Tabel umum */
  table.data-table { border-collapse: collapse; width: 100%; }
  table.data-table td,
  table.data-table th { border: 1px solid #000; padding: 2pt 4pt; vertical-align: top; font-size: 10pt; }
  .th-bl { background: #dbe5f1; font-weight: bold; text-align: center; }

  /* Kolom tabel intrakurikuler */
  .td-no { text-align: center; width: 24pt; font-weight: bold; }
  .td-mapel { font-weight: bold; text-align: left; width: 150pt; }
  .td-nilai { text-align: center; width: 50pt; }
  .td-desc { text-align: justify; }

  /* Kolom tabel ekstrakurikuler */
  .td-pred { text-align: center; width: 55pt; }

  /* Kokurikuler / Catatan Walas */
  .box-cell { border: 1px solid #000; padding: 4pt 6pt; font-size: 10pt; min-height: 30pt; text-align: justify; }

  /* Ketidakhadiran */
  .td-absen-label { width: 100pt; font-weight: bold; font-size: 10pt; }
  .td-absen-val { font-size: 10pt; }

  /* Tanda tangan */
  .sig-table { margin-top: 14pt; width: 100%; border-collapse: collapse; }
  .sig-table td { border: none !important; text-align: center; vertical-align: top; font-size: 10pt; padding: 1pt 4pt; }
  .sig-space { height: 52pt; }
  .sig-name { font-weight: bold; }
  /* Nama + NIP rata kiri sebagai satu blok yang ditengahkan,
     agar 'NIP. -' sejajar dengan awal nama */
  .sig-id { display: inline-block; text-align: left; }
  .sig-nip { font-weight: normal; font-size: 9pt; }
  /* Slot tanda tangan (gambar di-overlay, tidak menggeser layout) */
  .sig-slot { position: relative; height: 52pt; }
  .sig-slot img { position: absolute; left: 50%; top: 50%; }

  /* ===== Halaman Mutasi (Keterangan Pindah Sekolah) — dokumen terpisah ===== */
  .mutasi-source { display: none; }
  .mutasi-title { font-weight: bold; font-size: 11pt; text-align: center; margin: 0 0 10pt 0; letter-spacing: 0.5px; }
  .mutasi-id { width: 100%; border-collapse: collapse; margin-bottom: 8pt; }
  .mutasi-id td { border: none !important; font-size: 10pt; padding: 1pt 0; vertical-align: top; }
  .mutasi-id .mlbl { width: 150pt; }
  .mutasi-keluar { font-weight: bold; font-size: 10pt; margin: 4pt 0 2pt 0; }
  .mutasi-table td { font-size: 9pt; vertical-align: top; height: 108pt; }
  .mutasi-table .th-bl { font-size: 9pt; }
  .mutasi-ttd { text-align: center; }
  .mutasi-ttd .sig-slot { height: 40pt; }
  .mutasi-ttd .sig-id { margin-top: 2pt; }
  .mutasi-ortu { margin-top: 12pt; }
</style></head><body>

<!-- Hasil paginasi muncul di sini (diisi oleh JS) -->
<div id="rapor-root"></div>
<noscript><p style="font-family:sans-serif;padding:20px">Aktifkan JavaScript untuk menampilkan rapor.</p></noscript>

<!-- Sumber konten (disembunyikan, dibaca oleh paginator) -->
<div class="rapor-source" data-kelas="${esc(s.kelas)}" data-foot-left="${esc(s.kelas)} | ${esc(s.nama)} | ${esc(s.nis || '')}">

<!-- Header berulang tiap halaman: info siswa + sekolah + garis -->
<div class="rep-header">
<img class="sklh-logo" style="display:none;float:left;margin-right:10pt;max-height:48pt" alt="Logo Sekolah">
<table class="info-hdr">
  <tr>
    <td class="lbl">Nama Murid</td>
    <td class="sep-col">:</td>
    <td class="val">${esc(s.nama)}</td>
    <td class="lbl2">Kelas</td>
    <td class="sep-col2">:</td>
    <td class="val2">${s.kelas}</td>
  </tr>
  <tr>
    <td class="lbl">NIS/NISN</td>
    <td class="sep-col">:</td>
    <td class="val">${esc(nisDisplay)}</td>
    <td class="lbl2">Fase</td>
    <td class="sep-col2">:</td>
    <td class="val2">D</td>
  </tr>
  <tr>
    <td class="lbl">Sekolah</td>
    <td class="sep-col">:</td>
    <td class="val"><span class="sklh-nama"></span></td>
    <td class="lbl2">Semester</td>
    <td class="sep-col2">:</td>
    <td class="val2">${semester}</td>
  </tr>
  <tr>
    <td class="lbl">Alamat</td>
    <td class="sep-col">:</td>
    <td class="val" style="white-space:nowrap"><span class="sklh-alamat"></span></td>
    <td class="lbl2">Tahun Ajaran</td>
    <td class="sep-col2">:</td>
    <td class="val2">${tahunAjaran}</td>
  </tr>
</table>

<hr class="kop-sep">
</div><!-- .rep-header -->

<div class="kop-title">LAPORAN HASIL BELAJAR</div>

<!-- A. INTRAKURIKULER -->
<div class="section-title">A. Intrakurikuler</div>
<table class="data-table">
  <thead>
    <tr class="th-bl">
      <th class="td-no">No</th>
      <th class="td-mapel">Mata Pelajaran</th>
      <th class="td-nilai">Nilai Akhir</th>
      <th>Capaian Kompetensi</th>
    </tr>
  </thead>
  <tbody>${subRows}</tbody>
</table>

<!-- B. KOKURIKULER -->
<div class="section-title">B. Kokurikuler</div>
<table class="data-table"><tr><td class="box-cell" style="border:1px solid #000">${esc(s.kokurikuler || '')}</td></tr></table>

<!-- C. EKSTRAKURIKULER -->
<div class="section-title">C. Ekstrakurikuler</div>
<table class="data-table keep-together">
  <thead>
    <tr class="th-bl">
      <th class="td-no">No</th>
      <th>Kegiatan Ekstrakurikuler</th>
      <th class="td-pred">Predikat</th>
      <th>Keterangan</th>
    </tr>
  </thead>
  <tbody>${ekstraRows}</tbody>
</table>

<!-- D. KETIDAKHADIRAN -->
<div class="section-title">D. Ketidakhadiran</div>
<table class="data-table">
  <tr><td class="td-absen-label">Sakit</td><td class="td-absen-val">: ${sk} hari</td></tr>
  <tr><td class="td-absen-label">Izin</td><td class="td-absen-val">: ${iz} hari</td></tr>
  <tr><td class="td-absen-label">Tanpa Keterangan</td><td class="td-absen-val">: ${tk} hari</td></tr>
</table>

<!-- E. CATATAN WALI KELAS -->
<div class="section-title cluster-start">E. Catatan Wali Kelas</div>
<table class="data-table"><tr><td class="box-cell" style="border:1px solid #000;min-height:40pt">${esc(s.catatanWalas || '')}</td></tr></table>

<!-- KETERANGAN KELULUSAN / KEPUTUSAN NAIK KELAS (hanya semester 2; kosong di semester 1) -->
${keputusanBox}

<!-- TANGGAPAN ORANG TUA/WALI MURID (kosong, diisi tangan oleh orang tua) -->
<div class="section-title">Tanggapan Orang Tua/Wali Murid</div>
<table class="data-table"><tr><td class="box-cell" style="border:1px solid #000;height:50px"></td></tr></table>

<!-- TANDA TANGAN -->
<table class="sig-table">
  <tr>
    <td style="width:50%">
      Orang Tua Murid
    </td>
    <td style="width:50%">
      <span class="sklh-kota"></span>, <span class="tgl-rapor">${tanggalRapor}</span><br>
      Wali Kelas,
    </td>
  </tr>
  <tr>
    <td class="sig-space"></td>
    <td class="sig-space"><div class="sig-slot wk-sig"></div></td>
  </tr>
  <tr>
    <td>…………………………….</td>
    <td class="sig-name"><span class="sig-id"><span class="wk-name">${esc(waliKelas)}</span><br><span class="sig-nip">NIP. -</span></span></td>
  </tr>
  <tr>
    <td colspan="2" style="padding-top:8pt">
      Kepala Sekolah
    </td>
  </tr>
  <tr>
    <td colspan="2" class="sig-space"><div class="sig-slot ks-sig" style="margin:0 auto;width:200pt"></div></td>
  </tr>
  <tr>
    <td colspan="2" class="sig-name"><span class="sig-id"><span class="ks-name">Tri Wijayanti, M.Pd</span><br><span class="ks-nip">NIP. -</span></span></td>
  </tr>
</table>

</div><!-- #rapor-source -->

<!-- ===== KETERANGAN MUTASI (KELUAR) — DOKUMEN TERPISAH: lembar sendiri (tanpa kop/footer rapor, nomor halaman tak menyambung). Hanya ditempatkan setelah rapor. ===== -->
<div class="mutasi-source" data-kelas="${esc(s.kelas)}">
<div class="mutasi-title">KETERANGAN PINDAH SEKOLAH</div>
<table class="mutasi-id">
  <tr><td class="mlbl">NAMA PESERTA DIDIK</td><td>: ${esc(s.nama)}</td></tr>
  <tr><td class="mlbl">NIS / NISN</td><td>: ${esc(nisDisplay)}</td></tr>
</table>
<div class="mutasi-keluar">KELUAR</div>
<table class="data-table mutasi-table">
  <thead>
    <tr class="th-bl">
      <th style="width:82pt">Tanggal</th>
      <th style="width:74pt">Kelas yang Ditinggalkan</th>
      <th>Sebab-sebab Keluar atau Atas Permintaan (Tertulis)</th>
      <th style="width:178pt">Tanda Tangan Kepala Sekolah, Stempel Sekolah, dan Tanda Tangan Orang Tua/Wali</th>
    </tr>
  </thead>
  <tbody>${mutasiRows}</tbody>
</table>
</div><!-- #mutasi-source -->

<script>
(function() {
  // ===== Paginator: pecah konten menjadi lembar A4 nyata =====
  // Tiap lembar: header berulang di atas + konten + footer di bawah.
  // Print mencetak tiap lembar = 1 halaman → screen == print.

  // Dokumen mana yang dibuat (checklist). Default: Rapor saja.
  function selectedDocs() {
    var d = { identitas: false, rapor: true, mutasi: false };
    try {
      var c = JSON.parse(localStorage.getItem('eraporConfig') || '{}');
      if (c.docs) d = { identitas: !!c.docs.identitas, rapor: !!c.docs.rapor, mutasi: !!c.docs.mutasi };
    } catch (e) {}
    if (!d.identitas && !d.rapor && !d.mutasi) d.rapor = true; // jangan sampai kosong semua
    return d;
  }

  function build() {
    var root = document.getElementById('rapor-root');
    if (!root) return;
    root.innerHTML = '';
    var docs = selectedDocs();
    // Proses tiap blok sesuai urutan dokumen & checklist (rapor lalu mutasi tiap siswa, dst).
    var sources = document.querySelectorAll('.rapor-source, .mutasi-source');
    Array.prototype.forEach.call(sources, function(source) {
      if (source.classList.contains('mutasi-source')) { if (docs.mutasi) paginateMutasi(source, root); }
      else { if (docs.rapor) paginateSource(source, root); }
    });
  }

  // Mutasi = dokumen terpisah: lembar polos TANPA kop & footer rapor, tanpa nomor halaman rapor.
  function paginateMutasi(source, root) {
    var blocks = Array.prototype.slice.call(source.children);
    var sheet, content;
    function newPlainSheet() {
      sheet = document.createElement('div');
      sheet.className = 'sheet';
      content = document.createElement('div');
      content.className = 'sheet-content';
      sheet.appendChild(content);
      root.appendChild(sheet);
    }
    function overflows() { return content.scrollHeight > content.clientHeight + 1; }
    newPlainSheet();
    for (var i = 0; i < blocks.length; i++) {
      var item = blocks[i];
      var splittable = item.tagName === 'TABLE' && item.querySelector('thead');
      if (!splittable) {
        var node = item.cloneNode(true);
        content.appendChild(node);
        if (overflows() && content.children.length > 1) {
          content.removeChild(node);
          newPlainSheet();
          content.appendChild(node);
        }
      } else {
        var thead = item.querySelector('thead');
        var rows = Array.prototype.slice.call(item.querySelectorAll('tbody > tr'));
        var cls = item.className;
        function openTableM() {
          var t = document.createElement('table');
          t.className = cls;
          t.appendChild(thead.cloneNode(true));
          var tb = document.createElement('tbody');
          t.appendChild(tb);
          content.appendChild(t);
          return tb;
        }
        var tbody = openTableM();
        for (var r = 0; r < rows.length; r++) {
          var tr = rows[r].cloneNode(true);
          tbody.appendChild(tr);
          if (overflows()) {
            tbody.removeChild(tr);
            if (tbody.children.length === 0) { tbody.appendChild(tr); }
            else { newPlainSheet(); tbody = openTableM(); tbody.appendChild(tr); }
          }
        }
      }
    }
  }

  // Paginasi satu blok rapor (satu siswa) → lembar A4, ditambahkan ke root.
  // Nomor halaman direset per siswa (tiap rapor punya penomoran sendiri).
  function paginateSource(source, root) {
    var headerTpl = source.querySelector('.rep-header');
    var footLeft = source.getAttribute('data-foot-left') || '';

    // Kumpulkan blok konten (selain header berulang)
    var placeables = [];
    Array.prototype.forEach.call(source.children, function(ch) {
      if (ch.classList && ch.classList.contains('rep-header')) return;
      placeables.push(ch);
    });

    var pageNum = 0;
    var sheet = null, content = null;

    function newSheet() {
      pageNum++;
      sheet = document.createElement('div');
      sheet.className = 'sheet';

      var h = headerTpl.cloneNode(true);
      h.classList.remove('rep-header');
      h.classList.add('sheet-header');
      sheet.appendChild(h);

      content = document.createElement('div');
      content.className = 'sheet-content';
      sheet.appendChild(content);

      var f = document.createElement('div');
      f.className = 'sheet-footer';
      var sl = document.createElement('span');
      sl.textContent = footLeft;
      var sr = document.createElement('span');
      sr.textContent = 'Halaman : ' + pageNum;
      f.appendChild(sl);
      f.appendChild(sr);
      sheet.appendChild(f);

      root.appendChild(sheet);
    }

    function overflows() {
      return content.scrollHeight > content.clientHeight + 1;
    }

    function lastIsSectionTitle() {
      var lc = content.lastElementChild;
      return lc && lc.classList && lc.classList.contains('section-title');
    }

    newSheet();

    // Potong tabel besar (mis. Intrakurikuler) per baris; thead diulang tiap halaman.
    function splitTable(item) {
      var thead = item.querySelector('thead');
      var rows = Array.prototype.slice.call(item.querySelectorAll('tbody > tr'));
      var cls = item.className;
      function openTable() {
        var t = document.createElement('table');
        t.className = cls;
        t.appendChild(thead.cloneNode(true));
        var tb = document.createElement('tbody');
        t.appendChild(tb);
        content.appendChild(t);
        return tb;
      }
      var tbody = openTable();
      for (var r = 0; r < rows.length; r++) {
        var tr = rows[r].cloneNode(true);
        tbody.appendChild(tr);
        if (overflows()) {
          tbody.removeChild(tr);
          if (tbody.children.length === 0) {
            // Tak satu pun baris muat → pindah tabel + judul section ke halaman baru
            content.removeChild(tbody.parentNode);
            var carry2 = lastIsSectionTitle() ? content.lastElementChild : null;
            if (carry2) content.removeChild(carry2);
            newSheet();
            if (carry2) content.appendChild(carry2);
            tbody = openTable();
            tbody.appendChild(tr);
          } else {
            // Tutup tabel di halaman ini, buka lagi di halaman baru dgn header
            newSheet();
            tbody = openTable();
            tbody.appendChild(tr);
          }
        }
      }
    }

    for (var i = 0; i < placeables.length; i++) {
      var item = placeables[i];
      var isTable = item.tagName === 'TABLE' && !!item.querySelector('thead');
      // Tabel "keep-together" (mis. Ekstrakurikuler) TIDAK dipotong — pindah utuh bila tak muat.
      var keep = item.classList && item.classList.contains('keep-together');

      if (isTable && !keep) {
        splitTable(item);
        continue;
      }

      // ---- Blok atomik (judul, section-title, box, absen, TTD) ATAU tabel keep-together ----
      var node = item.cloneNode(true);
      var isSig = node.classList && node.classList.contains('sig-table');
      content.appendChild(node);
      if (overflows()) {
        content.removeChild(node);
        if (isSig) {
          // OPSI 1 — TTD tak boleh sendirian. Pindah ke halaman baru, lalu TARIK hanya cluster bawah
          // (mulai "E. Catatan Wali Kelas" ke bawah: Catatan -> Keputusan -> Tanggapan -> TTD) turun
          // menemaninya. BERHENTI di blok ber-class cluster-start -> tak menarik D. Ketidakhadiran
          // / tabel mapel di atasnya. Juga selama masih muat & halaman lama tak jadi kosong.
          var prev = content;            // konten halaman sebelumnya (sebelum newSheet)
          newSheet();                    // content -> halaman baru (kosong)
          content.appendChild(node);     // TTD ditaruh dulu
          while (prev.children.length > 1) {
            var cand = prev.lastElementChild;
            var atTop = cand.classList && cand.classList.contains('cluster-start');
            content.insertBefore(cand, content.firstChild);   // pindah ke ATAS TTD (urutan terjaga)
            if (overflows()) { prev.appendChild(cand); break; } // tak muat → kembalikan & berhenti
            if (atTop) break;            // sudah sampai batas atas cluster (Catatan Wali Kelas) → cukup
          }
        } else {
          // Jangan biarkan section-title yatim di dasar halaman → bawa serta ke halaman baru
          var carry = lastIsSectionTitle() ? content.lastElementChild : null;
          if (carry) content.removeChild(carry);
          newSheet();
          if (carry) content.appendChild(carry);
          content.appendChild(node);
          // Tabel keep-together yang ternyata lebih tinggi dari 1 halaman penuh → fallback potong per baris
          if (keep && isTable && overflows()) {
            content.removeChild(node);
            splitTable(item);
          }
        }
      }
    }
  }

  // ===== Terapkan config (nama wali kelas, kepala sekolah, tanda tangan) =====
  // Dibaca dari localStorage (origin sama dgn halaman utama) sebelum paginasi,
  // agar hasil kloning lembar sudah memuat nama & tanda tangan.
  var configApplied = false;
  function applyConfig() {
    if (configApplied) return;
    configApplied = true;
    var cfg = {};
    try { cfg = JSON.parse(localStorage.getItem('eraporConfig') || '{}'); } catch (e) {}
    var ident = {};
    try { ident = JSON.parse(localStorage.getItem('identitasSekolah') || '{}'); } catch (e) {}
    var sources = document.querySelectorAll('.rapor-source, .mutasi-source');
    Array.prototype.forEach.call(sources, function(source) {
      var kelas = source.getAttribute('data-kelas') || '';
      var w = (cfg.wali && cfg.wali[kelas]) || {};
      var k = cfg.kepsek || {};
      if (w.name) setText(source, '.wk-name', w.name);
      if (k.name) setText(source, '.ks-name', k.name);
      // NIP kepsek dari identitas atau pengaturan
      var nipVal = ident.nipKepsek || (k && k.nip) || '';
      if (nipVal) setText(source, '.ks-nip', (String(nipVal).indexOf('NIP') >= 0 ? '' : 'NIP. ') + nipVal);
      if (cfg.tanggalRaporISO) setText(source, '.tgl-rapor', fmtTgl(cfg.tanggalRaporISO));
      if (cfg.tanggalKelulusanISO) setText(source, '.tgl-kelulusan', fmtTgl(cfg.tanggalKelulusanISO));
      placeSig(source, '.wk-sig', w);
      placeSig(source, '.ks-sig', k);
      // V2.5: isi identitas sekolah
      if (ident.nama) setText(source, '.sklh-nama', ident.nama);
      if (ident.alamat) setText(source, '.sklh-alamat', ident.alamat);
      if (ident.kota) setText(source, '.sklh-kota', ident.kota);
      if (ident.logo) {
        var logoEls = source.querySelectorAll('.sklh-logo');
        for (var li = 0; li < logoEls.length; li++) {
          logoEls[li].src = ident.logo;
          logoEls[li].style.display = '';
        }
      }
    });

    function fmtTgl(iso) {
      var mm = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
      var p = String(iso).split('-'); // format YYYY-MM-DD dari <input type=date>
      if (p.length !== 3) return iso;
      var d = parseInt(p[2], 10), mo = parseInt(p[1], 10) - 1;
      if (isNaN(d) || mo < 0 || mo > 11) return iso;
      return d + ' ' + mm[mo] + ' ' + p[0];
    }

    function setText(root, sel, val) {
      var els = root.querySelectorAll(sel);
      for (var i = 0; i < els.length; i++) els[i].textContent = val;
    }
    function placeSig(root, sel, c) {
      if (!c || !c.sig) return;
      var slots = root.querySelectorAll(sel);
      for (var i = 0; i < slots.length; i++) {
        var img = document.createElement('img');
        img.src = c.sig;
        img.style.width = (c.w || 110) + 'px';
        var x = (c.x || 0), y = (c.y || 0);
        img.style.transform = 'translate(calc(-50% + ' + x + 'px), calc(-50% + ' + y + 'px))';
        slots[i].appendChild(img);
      }
    }
  }

  function run() { applyConfig(); requestAnimationFrame(build); }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', run);
  } else {
    run();
  }
  window.addEventListener('load', build);
  var rt;
  window.addEventListener('resize', function() {
    clearTimeout(rt);
    rt = setTimeout(build, 200);
  });
})();
</script>
</body></html>`;
}
