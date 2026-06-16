# e-Rapor V2.5

Aplikasi cetak rapor berbasis web — **100% offline**, tanpa database, tanpa instalasi rumit.  
Cukup upload file Excel leger nilai, atur pengaturan, lalu cetak rapor per siswa atau per kelas.

---

## Fitur

### Upload & Olah Data
- Upload file Excel (.xlsx) leger nilai — sheet otomatis dikenali per kelas (7A–9F)
- Parsing kolom: NISN, Nama, nilai mapel, deskripsi, ekstrakurikuler (SD & Pramuka), kokurikuler, ketidakhadiran, catatan wali kelas, rata-rata
- Data tersimpan di memori server selama sesi

### Pratinjau & Cetak Rapor
- Pilih kelas → tampilkan daftar siswa
- Klik siswa → lihat pratinjau rapor langsung (fit-to-width)
- Cetak rapor individu atau **cetak satu kelas penuh** (batch print)
- Dockumentasi: Cetak Rapor dan/atau Cetak Mutasi
- Zoom otomatis menyesuaikan lebar panel

### Pengaturan (modal 3 tab)
| Tab | Isian |
|-----|-------|
| 🏫 Identitas | Nama sekolah, alamat, kota, semester, tahun ajaran, logo |
| 👤 Kepala Sekolah | Nama, NIP, TTD (upload + drag atur posisi), tanggal rapor |
| 👥 Wali Kelas | Per kelas: nama wali kelas, TTD (upload + drag) |

- Tanggal kelulusan otomatis mengikuti tanggal rapor
- TTD bisa diatur ukuran dan posisinya (seret langsung di stage)
- Semua tersimpan otomatis ke localStorage browser

### Keamanan & Privasi
- **100% offline** — tidak ada data dikirim ke server eksternal
- Tidak ada database — data hanya di memori (upload) dan localStorage (konfigurasi)
- Single-user, tidak perlu login

---

## Cara Pakai

1. **Download sample Excel** dari halaman awal → isi nilai sesuai kolom
2. **Upload file Excel** — tunggu konfirmasi kelas terbaca
3. Klik **Buka Rapor**
4. Pilih kelas → pilih siswa → lihat pratinjau
5. Atur identitas, TTD, wali kelas via tombol **🔧 Pengaturan**
6. Cetak dengan **🖨 Cetak / PDF** atau **🖨 Cetak 1 Kelas**

---

## Tech Stack

- **Backend:** Node.js, Express, Multer, xlsx (SheetJS)
- **Frontend:** Vanilla JS, CSS, localStorage
- **Cetak:** iframe + `window.print()` — format A4 portrait

---

## Instalasi

```bash
npm install
npm start
# → http://localhost:3005
```

Tidak perlu database, tidak perlu konfigurasi.  
Cukup Node.js 18+ dan browser modern.
