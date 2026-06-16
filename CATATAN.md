# Catatan Pengembangan

Dokumen ini mencatat proses dan keputusan selama pengembangan e-Rapor V2.5  
berdasarkan sesi percakapan dengan AI (opencode).

---

## Ringkasan Sesi

### Langkah 1 — Setup awal
- Diminta membaca `RENCANA_V2.5.md` dan menjelaskan implementasi
- Arahan: "lanjutkan langsung semuanya"
- **Implementasi:**
  - Port diubah dari 3000 → 3005
  - Semua kode Google Sheets, Admin, auth, crypto dihapus dari `server.js`
  - Halaman landing, admin modal, input URL Google Sheets dihapus dari `index.html`
  - `app.js` disederhanakan — tidak ada landing, admin, Sheets
  - Form identitas sekolah ditambahkan di halaman awal, auto-save ke localStorage
  - Hardcoded values sekolah di `generateRaporHTML` diparameterisasi (`span.sklh-nama`, dll)
  - `applyConfig` di server.js membaca `identitasSekolah` dari localStorage
  - Logo upload + preview ditambahkan ke form identitas
  - Semester/tahunAjaran default diisi otomatis saat pertama load

### Langkah 2 — Merge Settings ke Setup Page (salah paham)
- User: "gabungkan pengaturan ke halaman awal"
- **Salah paham:** Semua input (TTD Kepsek, tanggal) dipindah ke halaman awal, modal disederhanakan jadi hanya Wali Kelas
- **Koreksi user:** Semua isian harusnya di **dalam modal Pengaturan setelah Buka Rapor**, bukan di halaman awal
- **Perbaikan:**
  - Halaman awal dikembalikan hanya upload + Buka Rapor
  - Modal Pengaturan diperluas dengan SEMUA isian (identitas, TTD, tanggal, logo, wali kelas)
  - Form identitas di halaman awal dihapus

### Langkah 3 — Rapikan modal
- Modal diatur jadi **3 tab**: Identitas, Kepala Sekolah, Wali Kelas
- Setiap tab fokus satu kategori
- Tanggal Kelulusan dihapus (otomatis = tanggal rapor)
- CSS tidak terpakai dibersihkan (`.landing-*`, `.admin-links`, `.field-row`/`.fld`, dll)

### Langkah 4 — Download Sample Excel
- Route `GET /api/download-sample` dibuat
- File `.xlsx` dengan 38 kolom dan 3 sample siswa (sheet "7A")
- Link download "📥 Download sample Excel" di halaman awal
- **Revisi:** Data sample awalnya salah (NAMA_SD keisi data Pramuka). Diperbaiki alignment kolom.

### Langkah 5 — Publish GitHub
- Nama repo: `rapor-umum`
- Akun GitHub: `smpabbs`
- Public repo dengan `.gitignore` (node_modules, uploads, data)
- File `RENCANA_V2.5.md` tidak di-commit

### Langkah 6 — Dokumentasi
- `README.md` — fitur project, cara pakai, tech stack
- `CATATAN.md` — mencatat proses percakapan ini

---

## Arsitektur

```
server.js          → Express server, rute API, generate HTML rapor
public/
├── index.html     → Halaman awal (upload) + workspace + modal pengaturan
├── app.js         → Client logic: upload, preview, modal, navigasi, drag TTD
└── style.css      → Semua styling
package.json       → Dependencies: express, multer, xlsx
```

### Aliran Data
```
Excel (.xlsx) → POST /api/upload → parsing sheet → db (memory)
                                           ↓
GET /api/data → populate dropdown kelas + daftar siswa
                                           ↓
Klik siswa → GET /rapor/:kelas/:ni → generate HTML rapor (server-side)
                                           ↓
Preview di iframe → localStorage (identitas, TTD, tanggal) → applyConfig()
                                           ↓
Cetak → iframe batch → print()
```

### localStorage Keys
| Key | Isi |
|-----|-----|
| `identitasSekolah` | nama, alamat, kota, kepsek, nipKepsek, semester, tahunAjaran, logo (dataUrl) |
| `eraporConfig` | kepsek (name, sig, w, x, y), wali (per kelas), tanggalRaporISO, tanggalKelulusanISO, docs (rapor/mutasi) |

---

## Format Excel Leger Nilai (38 kolom)

| # | Kolom | # | Kolom | # | Kolom | # | Kolom |
|---|-------|---|-------|---|-------|---|-------|
| 1 | No | 11 | des_Math | 21 | des_Sport | 31 | Deskripsi Pramuka |
| 2 | NISN | 12 | IPA | 22 | ICT | 32 | DeskripsiKokurikuler |
| 3 | Nama | 13 | des_IPA | 23 | des_ICT | 33 | SAKIT |
| 4 | PAI | 14 | IPS | 24 | jawa | 34 | IZIN |
| 5 | des_PAI | 15 | des_IPS | 25 | des_jawa | 35 | TANPA_KETERANGAN |
| 6 | Pkn | 16 | English | 26 | NAMA_SD | 36 | Catatan_Walas |
| 7 | des_Pkn | 17 | des_english | 27 | predikat | 37 | RATA_RATA |
| 8 | Indo | 18 | NILAI_SBK | 28 | des_SD | 38 | (dst) |
| 9 | des_indo | 19 | des_SBK | 29 | PRAMUKA /PMR | | |
| 10 | Math | 20 | Sport | 30 | Nilai Pramuka | | |

Nama sheet harus mengandung kode kelas (7A–9F), mis. "Total Nilai 7A".

---

## Catatan Penting

- **Tidak ada migrasi database** — semua file proyek siap pakai setelah `npm install`
- **Semua data konfigurasi** tersimpan di localStorage — hapus browser cache akan menghilangkan pengaturan
- **File upload** hanya parsing, tidak disimpan permanen
- **Cetak batch** menggunakan iframe tersembunyi dengan polling untuk deteksi render selesai
- **Drag TTD** menggunakan Pointer Events API
