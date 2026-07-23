# 📋 Ambil AR — Lembar Tagihan & Sinkronisasi Google Sheets

> **Satu klik: ekspor AR Accurate → lembar tagihan per penagih siap cetak + inject ke Google Sheets**

Pipeline Python lima langkah yang membaca ekspor daftar piutang dari Accurate dari LAPORAN PIUTANG dengan nama (`ExportFile.xls`), mengelompokkan tagihan per **penagih/sales** sesuai mapping kode pelanggan di `piutang.conf`, menghasilkan **`Print_AR.xlsm`** — lembar tagihan per penagih berformat template resmi macro-enabled siap cetak — sekaligus menyuntikkan seluruh data ke **Google Sheets** sebagai rekap digital.

---

## 📋 Daftar Isi

- [Gambaran Umum](#-gambaran-umum)
- [Fitur Utama](#-fitur-utama)
- [Prasyarat](#-prasyarat)
- [Struktur Folder & File](#-struktur-folder--file)
- [Cara Penggunaan](#-cara-penggunaan)
- [Alur Kerja Pipeline](#-alur-kerja-pipeline)
- [Detail Tiap Skrip](#-detail-tiap-skrip)
- [Konfigurasi `piutang.conf`](#-konfigurasi-piutangconf)
- [Setup Google Sheets API](#-setup-google-sheets-api)
- [Output](#-output)
- [Troubleshooting](#-troubleshooting)
- [Catatan Penting](#-catatan-penting)

---

## 🗂️ Gambaran Umum

`Ambil AR` mengotomasi dua pekerjaan manual yang biasa dilakukan admin setiap hari: mencetak lembar tagihan per penagih dan merekap data ke spreadsheet tim. Cukup siapkan satu file ekspor dari Accurate, jalankan satu skrip, dan keduanya selesai sekaligus.

Pipeline menghasilkan dua output:

| Output | Format | Digunakan untuk |
|---|---|---|
| `Print_AR.xlsm` | Excel macro-enabled | Dicetak dan dibagikan ke masing-masing penagih |
| Baris baru di Google Sheets | — | Rekap digital tim, monitoring, dan histori |

---

## ✨ Fitur Utama

- **Mapping penagih dari konfigurasi** — Setiap kode pelanggan dipetakan ke nama penagih melalui `piutang.conf`. Tidak perlu mengubah kode; cukup edit konfigurasi.
- **Kalkulasi otomatis umur piutang** — Kolom `Umur JT` dihitung ulang berdasarkan selisih antara `Tgl Faktur` dan tanggal hari ini, sehingga nilai selalu akurat meskipun file ekspor sudah lama.
- **Kolom Terbayar otomatis** — Menghitung `Nilai Faktur − Sisa Piutang` per baris. Kolom ini hanya ditampilkan jika nilainya lebih dari nol.
- **Template `.xlsm` dengan style, formula, dan shape tersalin** — Skrip 3 menyalin blok header (baris 1–4), baris data (baris 5), baris total berisi formula `=SUM()`, dan footer TTD dari `TEMPLATE.xlsm` via xlwings — termasuk gambar/logo yang ada di template.
- **Auto-fit lebar kolom & tinggi baris TTD** — Kolom disesuaikan otomatis; baris yang berisi teks "TTD SALES & COLLECTOR" secara otomatis diberi tinggi 115 poin untuk area tanda tangan.
- **Helper cleaning sebelum inject** — Langkah 4 meratakan merge cell, mengisi nama penagih ke kolom A tiap baris data, dan menghapus baris non-data sebelum dikirim ke Google Sheets.
- **Insert sebelum baris terakhir** — Data disisipkan tepat sebelum baris terakhir sheet Google Sheets, mempertahankan baris footer/total permanen yang ada di bawah.
- **Auto-cleanup** — Semua file sementara (`*temp.xlsx`, `ExportFile.xls`, `Print_AR.xlsm` di Dapur/) dihapus otomatis setelah seluruh proses selesai.
- **Validasi awal menyeluruh** — Sebelum memulai, orkestrator memverifikasi keberadaan folder `Dapur/` dan semua 8 file dependensi.

---

## 🔧 Prasyarat

### Python
Python **3.8+** disarankan.

### Library yang dibutuhkan

```bash
pip install pandas openpyxl xlrd xlsxwriter xlwings gspread google-auth
```

| Library | Digunakan di | Kegunaan |
|---|---|---|
| `pandas` | Skrip 1, 2 | Baca `.xls`, bersihkan, filter, hitung Terbayar |
| `xlsxwriter` | Skrip 1, 2 | Buat `.xlsx` sementara dengan format angka |
| `xlrd` | Skrip 1 | Baca file legacy `.xls` dari Accurate |
| `xlwings` | Skrip 3 | Buka dan tulis ke template `.xlsm` via Excel COM |
| `openpyxl` | Skrip 4 | Unmerge sel, bersihkan, baca `Print_AR.xlsm` |
| `gspread` | Skrip 5 | Klien Google Sheets API |
| `google-auth` | Skrip 5 | Autentikasi via Service Account |
| `numpy`, `re`, `os`, `glob`, `shutil`, `subprocess`, `sys`, `traceback`, `datetime` | Semua | Standard/bundled library |

### Aplikasi wajib
- **Microsoft Excel** — Wajib terinstall. Skrip 3 menggunakan `xlwings` yang memanggil Excel secara background untuk membuka, memodifikasi, dan menyimpan `TEMPLATE.xlsm`. Tanpa Excel, Skrip 3 akan gagal.

> **Catatan `xlrd`:** Gunakan versi yang kompatibel dengan `.xls`:
> ```bash
> pip install "xlrd>=1.0.0,<2.0.0"
> ```

---

## 📁 Struktur Folder & File

```
📦 Ambil-AR/
│
├── 📄 Ambil AR.py               ← Orkestrator utama. Jalankan ini
├── 📄 ExportFile.xls            ← [INPUT] Ekspor piutang dari Accurate (taruh di sini)
│
└── 📁 Dapur/                    ← Folder pipeline (jangan diubah strukturnya)
    ├── 📄 __init__.py
    ├── 📄 1_CleanerAcc.py       ← Bersihkan ExportFile.xls → ExportFile_clean_temp.xlsx
    ├── 📄 2_FilterAR.py         ← Filter per penagih + hitung Terbayar & total
    ├── 📄 3_CalculateAR.py      ← Susun ke TEMPLATE.xlsm → Print_AR.xlsm
    ├── 📄 4_HelperCleaningData.py ← Ratakan merge, isi nama penagih, hapus non-data
    ├── 📄 5_InjectDataToSS.py   ← Sisipkan 14 kolom ke Google Sheets
    ├── 📄 TEMPLATE.xlsm         ← Template resmi lembar tagihan (jangan dihapus)
    ├── 📄 credentials.json      ← Kredensial Google Service Account (rahasia!)
    └── 📄 piutang.conf          ← Konfigurasi mapping penagih, metadata, Google Sheets
```

> Output `Print_AR.xlsm` disalin ke folder utama selama proses berlangsung, lalu **dihapus dari `Dapur/`** di akhir. File di folder utama tetap ada sebagai hasil cetak.

---

## 🚀 Cara Penggunaan

### Langkah 1 — Siapkan file input

1. Export laporan piutang dari **Accurate** ke format `.xls`.
2. Simpan dengan nama **persis** `ExportFile.xls` di folder utama (sejajar dengan `Ambil AR.py`).

File ekspor harus mengandung kolom-kolom berikut (posisi bebas, skrip mendeteksi otomatis):
`No. Faktur`, `Tgl Faktur`, `Kode`, `Nama Pelanggan`, `Nilai Faktur`, `Sisa Piutang`, `Umur JT`

### Langkah 2 — Sesuaikan `piutang.conf`

Buka `Dapur/piutang.conf` dan perbarui:
- Pasangan `[NAMA SALES]` + `[KODE PELANGGAN]` sesuai kondisi terkini
- Metadata: `[PERUSAHAAN]`, `[DIVISI]`, `[TANGGAL]`, `[INPUT]`
- URL dan nama sheet Google Sheets di seksi `[SS]`

Lihat panduan lengkap di [Konfigurasi `piutang.conf`](#-konfigurasi-piutangconf).

### Langkah 3 — Pasang kredensial Google Sheets

Ganti isi `Dapur/credentials.json` dengan file JSON Service Account Anda. Lihat [Setup Google Sheets API](#-setup-google-sheets-api).

### Langkah 4 — Jalankan

```bash
python "Ambil AR.py"
```

### Langkah 5 — Pantau progress di terminal

```
--> Memulai eksekusi pembersihan data
--> Sedang memproses file: ExportFile.xls...
--> SUKSES! File tersimpan rapi di: ExportFile_clean_temp.xlsx
--> Proses selesai!
--> Memulai eksekusi menyalin dan menyusun data pada template
--> Proses selesai, file telah disimpan sebagai Print_AR.xlsm
--> Memulai persiapan data untuk di susun ke Spreadsheets
--> Sedang memproses file: Print_AR.xlsm
--> Proses selesai, file telah disimpan sebagai: Print_AR_temp.xlsx
--> Memulai unggah data ke Spreadsheets
--> Menyiapkan NN baris untuk disisipkan pada baris ke-X di Sheet '...'...
--> Proses selesai, data berhasil disisipkan ke Spreadsheets dengan format bawaan.
--> Semua proses telah selesai dijalankan.
--> Tekan enter untuk keluar.
```

### Langkah 6 — Ambil hasil

- **`Print_AR.xlsm`** tersedia di folder utama → buka Excel, cetak per penagih
- Data sudah tersisip di Google Sheets pada posisi sebelum baris terakhir

---

## 🔄 Alur Kerja Pipeline

```
[Mulai: Ambil AR.py]
   │
   ├─── Validasi folder & file
   │       Cek folder Dapur/ ada
   │       Cek 8 file syarat di Dapur/ ada:
   │         __init__.py, 1_CleanerAcc.py, 2_FilterAR.py,
   │         3_CalculateAR.py, 4_HelperCleaningData.py,
   │         5_InjectDataToSS.py, credentials.json, piutang.conf
   │       Cek ExportFile.xls ada di folder utama
   │       Jika gagal → tampilkan nama file yang hilang & berhenti
   │
   ├─── Bersihkan Dapur/ dari file lama
   │       Hapus *temp.xlsx dan ExportFile.xls sisa run sebelumnya
   │
   ├─── Salin ExportFile.xls → Dapur/ExportFile.xls
   │
   ├─── [1] 1_CleanerAcc.py
   │       Baca ExportFile.xls (header=None, scan hingga baris ke-150)
   │       Deteksi otomatis posisi 7 kolom target berdasarkan nama
   │       Buang baris: No. Faktur kosong, NaN, atau mengandung kata header
   │       Buang baris: Kode pelanggan kosong atau NaN
   │       Parse angka: Nilai Faktur & Sisa Piutang (format Indonesia/Inggris)
   │       Format angka: #,##0.00 | Auto-fit lebar kolom
   │       Output: ExportFile_clean_temp.xlsx (sheet: Data Bersih)
   │
   ├─── [2] 2_FilterAR.py
   │       Baca piutang.conf → bangun map {Kode Pelanggan: Nama Penagih}
   │       Filter: hanya baris yang kodenya ada di map
   │       Normalisasi bulan Indonesia → parse Tgl Faktur ke datetime
   │       Recalkulasi Umur JT = (hari ini - Tgl Faktur).hari
   │       Format Tgl Faktur → 'DD/MM/YYYY'
   │       Hitung Terbayar = Nilai Faktur - Sisa Piutang
   │         (kosongkan jika = 0)
   │       Tambah kolom Penagih → urutkan: Penagih, Nama, Kode, No.Faktur
   │       Sisipkan baris TOTAL [Penagih] + baris kosong antar grup
   │       Format angka: #,##0.00 | Auto-fit lebar kolom
   │       Output: Laporan_Piutang_Penagih_temp.xlsx
   │
   ├─── [3] 3_CalculateAR.py
   │       Baca piutang.conf → ambil PERUSAHAAN, DIVISI, TANGGAL, INPUT
   │       Buka TEMPLATE.xlsm via xlwings (Excel background)
   │       Rename sheet aktif → "TEMP_DESIGN"
   │       Buat sheet baru "Print AR"
   │       Loop per penagih:
   │         ├─ Salin baris 1–4 template (header) → tulis metadata & nama penagih
   │         ├─ Loop per baris faktur:
   │         │     Salin baris 5 template (baris data) → isi 9 kolom
   │         ├─ Salin baris 6 template (baris total) → tulis =SUM() per kolom nilai
   │         └─ Salin baris 7–N template (footer TTD) → tambah setelah total
   │       Auto-fit kolom, lebar spesifik kolom F–P
   │       Atur tinggi baris yang berisi "TTD SALES & COLLECTOR" = 115 poin
   │       Salin semua shapes (logo/gambar) dari template ke sheet output
   │       Hapus sheet TEMP_DESIGN
   │       Simpan sebagai Print_AR.xlsm
   │
   ├─── Salin Print_AR.xlsm dari Dapur/ → folder utama
   │
   ├─── [4] 4_HelperCleaningData.py
   │       Buka Print_AR.xlsm via openpyxl
   │       Unmerge semua sel yang tergabung
   │       Loop baris: jika menemukan baris header "Nama..." → catat nama penagih
   │       Isi kolom A tiap baris data dengan nama penagih yang sedang aktif
   │       Hapus baris: kosong, LAPORAN HASIL, Nama, Di input oleh,
   │                    No., TOTAL TAGIHAN, TTD SALES & COLLECTOR
   │       Ratakan tinggi semua baris → 18.75 poin
   │       Output: Print_AR_temp.xlsx
   │
   ├─── [5] 5_InjectDataToSS.py
   │       Baca piutang.conf → ambil URL dan SHEET_NAME dari seksi [SS]
   │       Baca Print_AR_temp.xlsx → susun 14 kolom per baris:
   │         Perusahaan, Nama Penagih, Divisi, Tanggal, Input,
   │         No., Kode, Nama Pelanggan, Umur JT, No. Faktur,
   │         Tgl Faktur, Nilai Faktur, Terbayar, Sisa Piutang
   │       Autentikasi via credentials.json (Service Account)
   │       Cari sheet target (case-insensitive, normalisasi spasi)
   │       Hitung baris akhir sheet → sisipkan di posisi (total_baris - 1)
   │       Kirim dengan inherit_from_before=True (mewarisi format baris sebelumnya)
   │
   └─── Cleanup akhir
           Hapus *temp.xlsx, ExportFile.xls, Print_AR.xlsm dari Dapur/
           (Print_AR.xlsm di folder utama tetap ada)
           "Semua proses telah selesai dijalankan."
```

---

## 🔍 Detail Tiap Skrip

### Skrip 1 — `1_CleanerAcc.py`

Membaca `ExportFile.xls` tanpa asumsi posisi header. Scan tiap sel dalam 150 baris pertama untuk menemukan posisi kolom berdasarkan nama. Baris yang `No. Faktur` atau `Kode`-nya kosong, NaN, atau berisi label header ulang (seperti "Total", "Halaman", "Page") otomatis dibuang. Parser angka `parse_to_float()` menangani semua kombinasi pemisah titik/koma format lokal maupun internasional.

**7 kolom output:**

| Kolom | Keterangan |
|---|---|
| `No. Faktur` | Nomor faktur |
| `Tgl Faktur` | Tanggal faktur (teks asli dari Accurate) |
| `Kode` | Kode pelanggan |
| `Nama Pelanggan` | Nama pelanggan |
| `Nilai Faktur` | Nilai faktur, format `#,##0.00` |
| `Sisa Piutang` | Sisa piutang, format `#,##0.00` |
| `Umur JT` | Umur jatuh tempo (dari Accurate, akan di-recalkulasi di Skrip 2) |

---

### Skrip 2 — `2_FilterAR.py`

Membaca `piutang.conf` untuk membangun mapping `{kode: penagih}`, lalu:

- Hanya memproses baris yang kode pelanggannya terdaftar
- **Recalkulasi `Umur JT`** dari `(hari ini − Tgl Faktur)` dalam hari — nilai dari Accurate tidak dipakai
- Menghitung `Terbayar = Nilai Faktur − Sisa Piutang` (dikosongkan jika = 0)
- Menyisipkan baris `TOTAL [Nama Penagih]` dan satu baris pemisah kosong antar kelompok penagih

---

### Skrip 3 — `3_CalculateAR.py`

Menyusun lembar cetak dari template `.xlsm` menggunakan xlwings. Struktur template yang diharapkan:

| Baris template | Fungsi |
|---|---|
| Baris 1–4 | Blok header (logo, nama perusahaan, tanggal, nama penagih) |
| Baris 5 | Baris data (disalin satu kali per faktur) |
| Baris 6 | Baris total (dengan formula `=SUM()` otomatis) |
| Baris 7–N | Footer (area TTD Sales & Collector) |

Setiap shape (gambar/logo) di template juga disalin ke sheet output dan diposisikan ulang di koordinat yang sama.

---

### Skrip 4 — `4_HelperCleaningData.py`

Mempersiapkan `Print_AR.xlsm` agar siap diunggah ke Google Sheets. Karena template menggunakan merge cell untuk header, semua merge dibuka terlebih dahulu. Baris non-data (judul, header kolom, total, TTD) dihapus. Nama penagih diisi ke kolom A setiap baris data aktif berdasarkan baris "Nama..." yang terdeteksi saat scanning.

---

### Skrip 5 — `5_InjectDataToSS.py`

Membaca 14 kolom dari `Print_AR_temp.xlsx` dan menyuntikkannya ke Google Sheets. Pencarian sheet target bersifat **case-insensitive** dan toleran terhadap spasi ganda. Data disisipkan pada posisi `total_baris - 1` (sebelum baris terakhir) dengan `inherit_from_before=True` agar format bawaan sheet Google Sheets (warna, font) diwarisi oleh baris baru.

---

## ⚙️ Konfigurasi `piutang.conf`

File ini menggunakan dua sistem format sekaligus: blok bernama (tanpa `=`) untuk mapping penagih, dan format `key = value` untuk metadata serta Google Sheets.

### Mapping penagih & pelanggan

```ini
[NAMA SALES]
C - M. Fahmi Akbar

[KODE PELANGGAN]
SG-4126
SG-4141
PW-1019
TGL-1038

[NAMA SALES]
DSR - Hasan Kholidin

[KODE PELANGGAN]
TGL-2021
PW-1089
PW-1064
```

**Aturan:**
- Setiap blok `[NAMA SALES]` diikuti **langsung** oleh blok `[KODE PELANGGAN]` berikutnya
- Satu nama per blok `[NAMA SALES]`, satu kode per baris di blok `[KODE PELANGGAN]`
- Kode yang tidak terdaftar di sini otomatis **diabaikan** oleh Skrip 2
- Boleh ada berapapun pasangan `[NAMA SALES]` + `[KODE PELANGGAN]`

### Metadata laporan

```ini
[PERUSAHAAN]
PTM

[DIVISI]
PCMO

[TANGGAL]
21/07/2026

[INPUT]
FEBIKA
```

Nilai ini dicetak di header lembar tagihan dan disertakan sebagai kolom tetap di Google Sheets.

### Konfigurasi Google Sheets

```ini
[SS]
url = https://docs.google.com/spreadsheets/d/ID_SPREADSHEET/edit
sheet_name = NamaSheetTarget
```

| Key | Keterangan |
|---|---|
| `url` | URL lengkap Google Sheets target |
| `sheet_name` | Nama sheet/tab di spreadsheet tersebut |

---

## 🔑 Setup Google Sheets API

### 1. Buat Service Account

1. Buka [Google Cloud Console](https://console.cloud.google.com/) → pilih atau buat project.
2. Aktifkan **Google Sheets API** dan **Google Drive API**.
3. Buka **IAM & Admin → Service Accounts** → buat Service Account baru.
4. Di tab **Keys** → buat key baru tipe **JSON** → file terunduh otomatis.

### 2. Pasang kredensial

Ganti isi `Dapur/credentials.json` dengan file JSON yang diunduh:

```json
{
  "type": "service_account",
  "project_id": "...",
  "private_key_id": "...",
  "private_key": "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n",
  "client_email": "nama@project.iam.gserviceaccount.com",
  ...
}
```

### 3. Berikan akses ke spreadsheet

Buka Google Sheets target → **Share/Bagikan** → tambahkan `client_email` dari `credentials.json` sebagai **Editor**.

### 4. Isi konfigurasi

Di `piutang.conf` seksi `[SS]`:

```ini
[SS]
url = https://docs.google.com/spreadsheets/d/1BxiM.../edit
sheet_name = Data Piutang
```

> **Cara mendapatkan URL:** Salin langsung dari address bar browser saat Google Sheets terbuka.

---

## 📤 Output

### 1. `Print_AR.xlsm` — Lembar tagihan siap cetak

Disalin ke folder utama. Berisi satu blok per penagih secara vertikal, masing-masing terdiri dari:
- **Header** — nama perusahaan, divisi, tanggal, nama penagih, nama penginput
- **Tabel faktur** — No., Kode, Nama Pelanggan, Umur JT, No. Faktur, Tgl Faktur, Nilai Faktur, Terbayar, Sisa Piutang
- **Baris TOTAL TAGIHAN** — formula `=SUM()` otomatis untuk Nilai Faktur, Terbayar, dan Sisa Piutang
- **Footer TTD** — area tanda tangan Sales & Collector (tinggi baris 115 poin)

### 2. Baris baru di Google Sheets

Setiap baris faktur disisipkan sebagai satu baris dengan **14 kolom**:

| # | Kolom | Sumber |
|---|---|---|
| 1 | Perusahaan | `[PERUSAHAAN]` di piutang.conf |
| 2 | Nama Penagih | Nama sales dari mapping |
| 3 | Divisi | `[DIVISI]` di piutang.conf |
| 4 | Tanggal | `[TANGGAL]` di piutang.conf |
| 5 | Input | `[INPUT]` di piutang.conf |
| 6 | No. | Nomor urut dalam grup penagih |
| 7 | Kode | Kode pelanggan |
| 8 | Nama Pelanggan | Nama pelanggan |
| 9 | Umur JT | Umur piutang (hari, dihitung ulang dari hari ini) |
| 10 | No. Faktur | Nomor faktur |
| 11 | Tgl Faktur | Tanggal faktur (DD/MM/YYYY) |
| 12 | Nilai Faktur | Nilai faktur asli |
| 13 | Terbayar | Nilai Faktur − Sisa Piutang |
| 14 | Sisa Piutang | Sisa piutang saat ini |

---

## 🛠️ Troubleshooting

### ❌ `File ExportFile.xls tidak ditemukan untuk diproses`
Pastikan file ada di folder utama dengan nama **persis** `ExportFile.xls` (bukan `.xlsx`).

### ❌ `Error: Kolom No. Faktur tidak ditemukan`
Skrip mencari header `"No. Faktur"` dalam 150 baris pertama. Jika tidak ditemukan, kemungkinan nama kolom di ekspor Accurate berbeda atau encoding file bermasalah. Periksa nama kolom di file `.xls` secara manual.

### ❌ Semua data hilang / hasil filter kosong
Skrip 2 hanya memproses kode yang terdaftar di `piutang.conf`. Periksa apakah kode pelanggan di `ExportFile.xls` persis sama (termasuk huruf kapital dan tanda hubung) dengan yang ada di blok `[KODE PELANGGAN]`.

### ❌ `PermissionError` atau Excel tidak bisa membuka TEMPLATE.xlsm
`TEMPLATE.xlsm` mungkin sedang terbuka di Excel, atau Excel belum terinstall. Tutup semua file Excel yang terbuka, lalu coba lagi.

### ❌ `Print_AR.xlsm` tidak terbentuk
Skrip 3 bergantung pada `TEMPLATE.xlsm` di dalam `Dapur/`. Pastikan file template tidak dihapus atau dipindahkan. File ini tidak termasuk dalam daftar validasi awal orkestrator sehingga kesalahannya baru terdeteksi saat Skrip 3 berjalan.

### ❌ `Error: URL Google Spreadsheet tidak ditemukan di piutang.conf`
Seksi `[SS]` di `piutang.conf` belum diisi. Tambahkan nilai `url = ` dan `sheet_name = `.

### ❌ `Error: Spreadsheets tidak ditemukan. Pastikan email Service Account sudah diberi akses Editor`
Buka Google Sheets target → **Share** → tambahkan `client_email` dari `credentials.json` sebagai Editor.

### ❌ `Error: Sheet 'X' tidak ditemukan` + daftar sheet yang tersedia
Nilai `sheet_name` di `piutang.conf` tidak cocok dengan nama sheet di Google Sheets. Pencarian bersifat case-insensitive, tapi nama harus mengandung kata yang sama. Gunakan nama yang ditampilkan dalam daftar error.

### ❌ Data di Google Sheets bertumpuk / posisi salah
Skrip 5 selalu menyisipkan di posisi `total_baris - 1` (sebelum baris terakhir). Jika baris terakhir bukan baris footer/total yang diinginkan, sesuaikan struktur sheet Google Sheets.

### ❌ Error autentikasi Google (`DefaultCredentialsError`)
Periksa `credentials.json` — pastikan berisi JSON yang valid dan `private_key` tersalin lengkap termasuk `-----BEGIN PRIVATE KEY-----` dan `-----END PRIVATE KEY-----`.

---

## 📌 Catatan Penting

- **`ExportFile.xls` disalin, bukan dipindahkan** — File asli di folder utama aman setelah proses.
- **`TEMPLATE.xlsm` wajib ada dan tidak boleh diubah strukturnya** — Baris 1–4 = header, baris 5 = baris data, baris 6 = baris total, baris 7+ = footer. Jika ingin mengubah tampilan cetak, edit langsung di `TEMPLATE.xlsm`.
- **`credentials.json` bersifat rahasia** — Tambahkan ke `.gitignore`. Jangan pernah commit ke repositori publik.
- **`Print_AR.xlsm` di folder utama tidak dihapus** — Hanya salinan di `Dapur/` yang dihapus. File hasil cetak di folder utama tetap tersedia.
- **Nilai Terbayar dikosongkan jika nol** — Faktur yang belum ada pembayaran sama sekali tidak menampilkan nilai di kolom Terbayar, baik di lembar cetak maupun di Google Sheets.
- **Umur JT selalu dihitung ulang** — Nilai kolom `Umur JT` di output mencerminkan kondisi hari ini, bukan nilai yang tersimpan di Accurate saat ekspor dilakukan.
- **Jangan ubah struktur `Dapur/`** — Semua skrip menggunakan path relatif dan bergantung pada nama file sementara yang sudah ditentukan.

---

## 📜 Lisensi

Proyek ini dikembangkan untuk keperluan internal perusahaan. Silakan sesuaikan dengan kebutuhan organisasi Anda.

---

*Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)*
