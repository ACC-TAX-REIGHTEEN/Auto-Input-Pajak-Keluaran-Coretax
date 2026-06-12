# 🔍 Auto-Input Pajak Keluaran Coretax

> **Rekonsiliasi & analisa otomatis data Pajak Keluaran antara Accurate dan Coretax DJP**

Skrip Python berbasis pipeline yang membaca data ekspor dari **Accurate** (Safi/Fella `.xls`) dan data rekap Pajak Keluaran dari **Coretax DJP** (`data_export` CSV/XLSX), lalu melakukan pencocokan otomatis menggunakan algoritma **6-tahap matching**, dan menghasilkan **Laporan Analisa Pajak** Excel berisi ringkasan selisih, data tidak cocok, dan semua data yang berhasil dicocokkan — lengkap per bulan, siap untuk keperluan audit.

---

## 📋 Daftar Isi

- [Fitur Utama](#-fitur-utama)
- [Prasyarat](#-prasyarat)
- [Struktur Folder](#-struktur-folder)
- [Cara Penggunaan](#-cara-penggunaan)
- [Alur Pipeline (Step-by-Step)](#-alur-pipeline-step-by-step)
- [Algoritma 6-Tahap Matching](#-algoritma-6-tahap-matching)
- [Output & Struktur Laporan](#-output--struktur-laporan)
- [Folder Addon (Utilitas Tambahan)](#-folder-addon-utilitas-tambahan)
- [Troubleshooting](#-troubleshooting)
- [Catatan Penting](#-catatan-penting)

---

## ✨ Fitur Utama

- **Dual input mode** — Mendukung dua format ekspor Accurate (`Safi.xls` atau `Fella.xls`) dan dua format ekspor Coretax (`data_export*.csv` atau `data_export*.xlsx`) secara sekaligus, dengan deteksi otomatis format yang tersedia.
- **Smart CSV parser** — Parser khusus yang menangani CSV bermasalah (nama pelanggan dengan koma di tengah) tanpa merusak struktur data.
- **Algoritma 6-tahap matching** — Pencocokan berjenjang dari yang paling ketat (TIN + jumlah + tanggal exact) hingga yang paling toleran (prefix nama + jumlah), meminimalkan data yang lolos tanpa pasangan.
- **Normalisasi data lintas sistem** — Membersihkan format NPWP/NIK yang berbeda, menstandarkan nama perusahaan (strip prefix PT/CV/UD), dan mengonversi format tanggal Indonesia ke objek datetime.
- **Deteksi bulan otomatis** — Nama file output ditentukan berdasarkan bulan yang paling dominan dalam data, tanpa konfigurasi manual.
- **Laporan Excel multi-sheet** — Output berisi 5 sheet terstruktur: Ringkasan, Rincian (A/B/C), Rincian 2 (gabungan tidak cocok), Rincian 3 (selisih nama+NPWP), Data Asli Accurate, Data Asli Coretax.
- **Auto-cleanup** — Semua file sementara dihapus otomatis setelah proses selesai; file sumber asli tidak disentuh.

---

## 🔧 Prasyarat

### Python
Python **3.8+** disarankan.

### Library yang dibutuhkan
Install semua dependensi dengan:

```bash
pip install pandas openpyxl xlsxwriter xlrd
```

> **Catatan:** `difflib` sudah termasuk dalam Python standard library, tidak perlu di-install terpisah.

| Library | Kegunaan |
|---|---|
| `pandas` | Baca, transformasi, gabung, dan simpan data Excel/CSV |
| `openpyxl` | Baca/tulis `.xlsx`, styling sel (border, fill, font) |
| `xlsxwriter` | Buat `.xlsx` baru dengan formatting lanjutan |
| `xlrd` | Baca file legacy `.xls` dari Accurate (Safi/Fella) |
| `difflib` | Fuzzy matching nama perusahaan (Tahap 4 matching) |

---

## 📁 Struktur Folder

```
📦 Auto-Input-Pajak-Keluaran-Coretax/
│
├── 📄 Keluaran_Pajak.py              ← File utama. Jalankan ini untuk memulai
│
├── 📄 Safi.xls                       ← [INPUT] Ekspor Accurate via Safi (salah satu)
├── 📄 Fella.xls                      ← [INPUT] Ekspor Accurate via Fella (salah satu)
├── 📄 data_export*.csv               ← [INPUT] Ekspor Coretax format CSV (salah satu)
├── 📄 data_export*.xlsx              ← [INPUT] Ekspor Coretax format XLSX (salah satu)
│
├── 📁 Dapur/                         ← Folder pipeline utama (jangan diubah strukturnya)
│   ├── 📄 __init__.py
│   ├── 📄 1_Cleaner&MergerACC.py     ← Bersihkan Safi/Fella + gabungkan data_export
│   ├── 📄 2_HMA_ex2ex_analytics_third.py  ← Analisa & matching (input XLSX)
│   ├── 📄 2_HMA_csv2ex_analytics_third.py ← Analisa & matching (input CSV)
│   ├── 📄 3_Rincian2.py              ← Susun sheet "Rincian 2" (data tidak cocok)
│   └── 📄 4_rincian3.py              ← Susun sheet "Rincian 3" (selisih nama+NPWP)
│
├── 📁 Addon/                         ← Utilitas tambahan (jalankan manual sesuai kebutuhan)
│   ├── 📄 cleaner&merger.py          ← Versi standalone dari step 1 (Cleaner+Merger)
│   ├── 📄 csv2ex.py                  ← Konversi data_export*.csv → HMA_csv2ex.xlsx
│   ├── 📄 ma_ex2ex.py                ← Gabungkan data_export*.xlsx → HMA_ex2ex.xlsx
│   ├── 📄 tang2ddmmmyy_ex2ex_keluaran.py ← Konversi format tanggal ke DD/MMM/YYYY
│   ├── 📄 hitung_ppn_all_file_keluaran.py ← Hitung total PPN dari semua file Excel
│   └── 📄 audit_data_gabungan_keluaran_ex.py ← Audit integritas data gabungan merger
│
└── 📁 File Dibutuhkan/               ← Placeholder & referensi (bisa diabaikan)
    ├── 📄 Fella.xls                  ← (kosong, contoh nama file)
    ├── 📄 Safi.xls                   ← (kosong, contoh nama file)
    ├── 📄 data_export.csv            ← (kosong, contoh nama file)
    └── 📄 data_export.xlsx           ← (kosong, contoh nama file)
```

---

## 🚀 Cara Penggunaan

### Langkah 1 — Siapkan file input

Letakkan semua file input di **folder utama** (sejajar dengan `Keluaran_Pajak.py`):

**Dari Accurate** — letakkan **salah satu** (atau keduanya):
- `Safi.xls` → ekspor laporan pajak keluaran via modul Safi di Accurate
- `Fella.xls` → ekspor laporan pajak keluaran via modul Fella di Accurate

**Dari Coretax DJP** — letakkan **salah satu** (atau keduanya, bisa lebih dari satu file):
- `data_export.csv` / `data_export_januari.csv` / `data_export_feb.csv` — semua file yang diawali `data_export` dan berekstensi `.csv` akan digabungkan otomatis
- `data_export.xlsx` / `data_export_2.xlsx` — sama, untuk format XLSX

> Nama file `data_export` boleh memiliki suffix apapun selama diawali dengan `data_export`. Program akan mendeteksi semua file dengan pola tersebut secara otomatis.

### Langkah 2 — Jalankan

```bash
python Keluaran_Pajak.py
```

atau klik dua kali file tersebut jika Python sudah terasosiasi di sistem.

Program akan melakukan pengecekan awal (validasi keberadaan semua file dan folder), lalu menjalankan pipeline secara otomatis tanpa perlu input dari pengguna.

### Langkah 3 — Ambil hasil

Setelah proses selesai, file laporan akan muncul di folder utama:

```
Laporan_Analisa_Pajak_MARET.xlsx    ← (nama bulan menyesuaikan mayoritas data)
```

---

## 🔄 Alur Pipeline (Step-by-Step)

Pipeline dijalankan berurutan oleh `Keluaran_Pajak.py`:

```
[Mulai]
   │
   ├─── Validasi Awal (Keluaran_Pajak.py)
   │       Cek: Safi.xls atau Fella.xls ada?
   │       Cek: data_export* (csv/xlsx) ada?
   │       Cek: folder Dapur/ dan Addon/ ada?
   │       Cek: semua skrip Dapur/ dan Addon/ lengkap?
   │       Jika ada yang kurang → tampilkan daftar file hilang & berhenti
   │
   ├─── Salin file input → folder Dapur/
   │       Salin Safi.xls atau Fella.xls → Dapur/
   │       Salin semua data_export* → Dapur/
   │
   ├─── [STEP 1] 1_Cleaner&MergerACC.py
   │       │
   │       ├─ Modul CSV: Baca data_export*.csv dengan smart parser
   │       │   (Tangani nama dengan koma, gabungkan multi-file)
   │       │   Output: HMA_csv2ex.xlsx (jika ada file CSV)
   │       │
   │       ├─ Modul XLSX: Baca data_export*.xlsx dari sheet "data"
   │       │   (Gabungkan multi-file, auto-detect kolom angka)
   │       │   Output: HMA_ex2ex.xlsx (jika ada file XLSX)
   │       │
   │       └─ Modul Safi/Fella: Baca .xls dengan header baris ke-4
   │           Ekstrak: Tanggal, Tgl.Pajak, No.Referensi, No.FakturPajak,
   │                    NamaPelanggan, Negara, JumlahPajak, NomorPajak
   │           Output: Hasil_CleanerACC.xlsx
   │           (File .xls asli di Dapur/ dihapus setelah sukses)
   │
   ├─── Deteksi jalur [STEP 2]
   │       Jika HMA_ex2ex.xlsx ada → gunakan 2_HMA_ex2ex_analytics_third.py
   │       Jika HMA_csv2ex.xlsx ada → gunakan 2_HMA_csv2ex_analytics_third.py
   │
   ├─── [STEP 2] 2_HMA_*_analytics_third.py
   │       Baca: Hasil_CleanerACC.xlsx + HMA_ex2ex.xlsx (atau HMA_csv2ex.xlsx)
   │       Standarisasi: nama perusahaan, NPWP/NIK, format tanggal
   │       Deteksi bulan dominan → tentukan nama file output
   │       Jalankan ALGORITMA 6-TAHAP MATCHING (lihat bagian berikutnya)
   │       Susun laporan: Ringkasan, Rincian (blok A/B/C), Data Asli
   │       Output: Laporan_Analisa_Pajak_[BULAN].xlsx
   │       (File sementara dihapus setelah sukses)
   │
   ├─── [STEP 3] 3_Rincian2.py
   │       Baca sheet "Rincian" dari file laporan output
   │       Ekstrak blok A (hanya di Coretax) + blok B (hanya di Accurate)
   │       Gabungkan, urutkan per tanggal & jumlah
   │       Styling: header hijau, baris B di-highlight oranye
   │       Output: tambah sheet "Rincian 2" ke file laporan yang sama
   │
   ├─── [STEP 4] 4_rincian3.py
   │       Baca sheet "Rincian" → bagian C (Data Matched)
   │       Filter hanya baris dengan catatan "Nama Beda; Format NPWP Beda"
   │       Styling: header + border
   │       Output: tambah sheet "Rincian 3" ke file laporan (jika ada data)
   │
   └─── Finalisasi (Keluaran_Pajak.py)
           Salin Laporan_Analisa_Pajak_*.xlsx → folder utama
           Hapus file laporan dari folder Dapur/
           Selesai ✅
```

---

## 🎯 Algoritma 6-Tahap Matching

Inti dari proyek ini adalah algoritma pencocokan berjenjang yang mencocokkan setiap baris data Accurate dengan baris yang sesuai di data Coretax. Pencocokan dilakukan dari yang paling ketat ke paling toleran, dan setiap baris hanya bisa dicocokkan dengan satu pasangan.

**Pra-proses standarisasi data (sebelum matching):**
- **Nama perusahaan** — lowercase, hapus prefix (PT., CV., UD., TB., Toko, Bengkel), hilangkan tanda baca dan karakter non-alphanumeric
- **NPWP/NIK** — strip semua non-digit, hapus 1 digit terakhir (toleransi perbedaan format antar sistem)
- **Tanggal** — parse format Indonesia (DD Mmm YYYY) ke objek `datetime`
- **Jumlah pajak** — konversi ke `float`

| Tahap | Kriteria Pencocokan | Keterangan |
|---|---|---|
| **1** | NPWP + Jumlah (±5) + Tanggal **exact** | Paling ketat. Cocok sempurna. |
| **2** | Nama + Jumlah (±5) + Tanggal **exact** | Untuk transaksi tanpa NPWP valid. |
| **3** | NPWP + Jumlah (±5) + Tanggal **±3 hari** | Toleransi perbedaan tanggal input. |
| **4** | Nama (fuzzy ≥80%) + Jumlah (±5) + Tanggal **±3 hari** | Nama sedikit berbeda ejaan antar sistem. |
| **5** | Jumlah (±5) + Tanggal **exact** | Fallback jika nama dan NPWP bermasalah. |
| **6** | Prefix nama (5 karakter pertama) + Jumlah (±5) | Tahap terakhir, paling toleran. |

Setelah 6 tahap, data yang masih belum cocok dikategorikan sebagai:
- **"Not in Accurate"** — ada di Coretax, tidak ada di Accurate → masuk **Blok A**
- **"Not in Coretax"** — ada di Accurate, tidak ada di Coretax → masuk **Blok B**

---

## 📤 Output & Struktur Laporan

### Nama file output

```
Laporan_Analisa_Pajak_MARET.xlsx
```
Nama bulan (JANUARI s.d. DESEMBER) ditentukan otomatis dari bulan yang paling banyak muncul di kolom tanggal data Accurate.

### Sheet 1 — Ringkasan

Ikhtisar 5 baris:

| Uraian | Nilai (Rp) |
|---|---|
| Total Pajak Accurate | ... |
| Total Pajak Coretax | ... |
| Selisih Total | ... |
| Total Tidak Ada di Accurate | ... |
| Total Tidak Ada di Coretax | ... |

### Sheet 2 — Rincian

Tiga blok data dalam satu sheet dengan baris total ber-highlight kuning di akhir tiap blok:

**Blok A — Data hanya ada di Coretax**

10 kolom: Tanggal Coretax, Nama Coretax, Nomor Pajak Coretax, Jumlah Coretax, Perbedaan/Selisih, Catatan (`Hanya di Coretax`), serta kolom kosong untuk sisi Accurate.

**Blok B — Data hanya ada di Accurate**

10 kolom: Tanggal Accurate, Nama Accurate, Nomor Pajak Accurate, Jumlah Accurate, Perbedaan/Selisih, Catatan (`Hanya di Accurate`), serta kolom kosong untuk sisi Coretax.

**Blok C — Data Matched**

10 kolom: Tanggal Accurate, Tanggal Coretax, Nama Accurate, Nama Coretax, Nomor Pajak Accurate, Nomor Pajak Coretax, Jumlah Accurate, Jumlah Coretax, Perbedaan/Selisih, Catatan.

Kolom **Catatan** berisi kombinasi flag anomali yang terdeteksi saat matching:
- `Nama Beda` — nama pelanggan berbeda antar sistem
- `Format NPWP Beda` — format NPWP/NIK tidak identik
- `Beda Tanggal` — tanggal tidak sama persis (cocok via toleransi hari)
- `Selisih: [angka]` — jumlah PPN tidak sama persis

### Sheet 3 — Rincian 2

Gabungan Blok A dan Blok B dari sheet Rincian, diurutkan berdasarkan tanggal kemudian jumlah. Berguna sebagai daftar kerja untuk perbaikan data.

Styling: baris data **"Tidak Ada Di Coretax"** di-highlight warna hijau terang untuk perhatian khusus.

### Sheet 4 — Rincian 3 *(kondisional)*

Subset dari Blok C — hanya berisi baris yang memiliki catatan `"Nama Beda; Format NPWP Beda"` secara bersamaan. Berguna untuk identifikasi dan normalisasi master data di salah satu sistem.

> Sheet ini **hanya dibuat jika ada data** yang memenuhi kriteria tersebut.

### Sheet 5 & 6 — Data Asli Accurate / Data Asli Coretax

Data mentah dari kedua sumber setelah normalisasi awal (sebelum matching), disimpan untuk keperluan verifikasi manual. Kolom NPWP/NIK dipastikan tersimpan sebagai teks untuk menghindari konversi notasi ilmiah.

---

## 🛠️ Folder Addon (Utilitas Tambahan)

Folder `Addon/` berisi skrip yang **tidak dijalankan otomatis** oleh pipeline, melainkan digunakan secara manual sesuai kebutuhan tertentu.

### `csv2ex.py`
Konversi satu atau banyak file `data_export*.csv` menjadi satu file `HMA_csv2ex.xlsx`. Gunakan untuk menjalankan langkah ini secara mandiri atau untuk debugging step 1.
```bash
# Jalankan dari folder yang berisi file data_export*.csv
python Addon/csv2ex.py
```

### `ma_ex2ex.py`
Menggabungkan satu atau banyak file `data_export*.xlsx` (sheet bernama `data`) menjadi satu file `HMA_ex2ex.xlsx`. Versi standalone dari modul XLSX di step 1.
```bash
python Addon/ma_ex2ex.py
```

### `cleaner&merger.py`
Versi standalone lengkap dari `Dapur/1_Cleaner&MergerACC.py` — berisi ketiga modul (CSV, XLSX, Safi/Fella) dalam satu file. Berguna untuk dijalankan terpisah dari pipeline utama.

### `hitung_ppn_all_file_keluaran.py`
Membaca semua file `.xlsx`/`.xls` di folder yang ditentukan (default: folder saat ini), lalu menjumlahkan nilai di **kolom G** (kolom ke-7) dari setiap file. Berguna untuk quick-check total PPN sebelum rekonsiliasi.
```bash
python Addon/hitung_ppn_all_file_keluaran.py
# Output: subtotal per file + TOTAL KESELURUHAN
```

### `tang2ddmmmyy_ex2ex_keluaran.py`
Konversi kolom tanggal di file `Hasil_merger_all_ex2ex_keluaran.xlsx` ke format `DD/MMM/YYYY` (contoh: `01/Mar/2025`). Berguna jika file gabungan akan diupload ke sistem lain yang mensyaratkan format tanggal tertentu.

### `audit_data_gabungan_keluaran_ex.py`
Tool audit integritas: membandingkan file-file sumber individual dengan file gabungan hasil merger (`Hasil_merger_all_ex2ex_keluaran.xlsx`). Mendeteksi dua jenis masalah:
- **DATA HILANG** — nomor faktur ada di sumber tapi tidak ada di file gabungan
- **SELISIH NILAI** — nomor faktur ada di keduanya tapi nilai PPN berbeda

Output: `Laporan_audit_merger_all_ex2ex_keluaran.xlsx` dengan sheet Ringkasan dan Detail Kesalahan.
```bash
python Addon/audit_data_gabungan_keluaran_ex.py
```

---

## 🛠️ Troubleshooting

### ❌ `Gagal: File Safi.xls atau Fella.xls tidak ditemukan`
Pastikan file ekspor Accurate sudah ada di folder utama (sejajar dengan `Keluaran_Pajak.py`) dengan nama persis `Safi.xls` atau `Fella.xls`. Program membutuhkan minimal salah satu.

### ❌ `Gagal: File data_export (csv/xlsx) tidak ditemukan`
Pastikan file ekspor dari Coretax DJP ada di folder utama dengan nama yang diawali `data_export`. Contoh: `data_export_maret.csv`, `data_export.xlsx`, `data_export_2.csv` — semua terdeteksi.

### ❌ `Gagal: File Hasil HMA_ex2ex.xlsx atau HMA_csv2ex.xlsx tidak ditemukan`
Step 1 tidak berhasil menghasilkan file HMA. Jalankan secara manual dari folder `Dapur/` dan periksa error:
```bash
cd Dapur
python "1_Cleaner&MergerACC.py"
```
Kemungkinan penyebab: file `data_export*.xlsx` tidak memiliki sheet bernama `"data"` (huruf kecil), atau file CSV memiliki encoding yang tidak terbaca.

### ❌ Error `xlrd` saat membaca Safi/Fella
Install versi `xlrd` yang kompatibel dengan format `.xls` lama:
```bash
pip install "xlrd>=1.0.0,<2.0.0"
```
`xlrd` versi 2.x ke atas tidak lagi mendukung format `.xls` (BIFF) yang digunakan oleh ekspor Accurate.

### ❌ Hampir semua data masuk Blok A atau Blok B, sangat sedikit yang Matched
Kemungkinan format tanggal di data Coretax tidak terbaca dengan benar oleh `pd.to_datetime()`. Buka file `HMA_ex2ex.xlsx` atau `HMA_csv2ex.xlsx` di folder `Dapur/` dan periksa kolom `DocumentDate` — pastikan berisi format tanggal yang dapat diparsing (ISO 8601 `YYYY-MM-DD` atau `DD/MM/YYYY`).

### ❌ `File Laporan_Analisa_Pajak tidak ditemukan setelah proses selesai`
Salah satu step gagal tanpa menghentikan program secara eksplisit. Jalankan step 2 secara manual untuk melihat error lengkap:
```bash
cd Dapur
python 2_HMA_ex2ex_analytics_third.py
```

---

## 📌 Catatan Penting

- **Jangan ubah struktur folder** `Dapur/` dan `Addon/` — semua skrip menggunakan path relatif dan saling bergantung.
- **File input asli aman** — `Keluaran_Pajak.py` **menyalin** (bukan memindahkan) file dari folder utama ke `Dapur/`. File `Safi.xls`/`Fella.xls` dan `data_export*` yang ada di folder `Dapur/` (salinannya) yang akan dihapus setelah diproses.
- **File output akan tertimpa** — jika `Laporan_Analisa_Pajak_MARET.xlsx` sudah ada dari run sebelumnya, file baru akan menimpa file lama. Buat backup terlebih dahulu jika diperlukan.
- **Matching tidak memfilter bulan** — jika file input mengandung data dari beberapa bulan, semua akan diproses bersama-sama. Deteksi bulan hanya digunakan untuk penamaan file output.
- **Selalu verifikasi Blok B** — data di Blok B (Hanya di Accurate) adalah yang paling kritis: faktur tersebut sudah ada di pembukuan Accurate tetapi belum tercatat di Coretax, yang berpotensi berdampak pada pelaporan pajak.

---

## 📜 Lisensi

Proyek ini dikembangkan untuk keperluan internal perpajakan. Silakan sesuaikan dengan kebutuhan organisasi Anda.

---

*Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)*
