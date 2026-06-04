# 🧾 Auto-Input XML Pajak Coretax

**Versi:** 1.2.3  
**Bahasa:** Python 3  
**Platform:** Windows (wajib, karena menggunakan `xlwings` dengan Microsoft Excel)

Skrip otomasi Python untuk mengonversi data ekspor dari sistem akuntansi (format `.xls`) menjadi file Excel XML Coretax yang siap diimpor ke portal **DJP Coretax**. Mendukung dua skenario: **Faktur Biasa** dan **Faktur dengan Diskon**.

---

## 📋 Daftar Isi

- [Latar Belakang](#latar-belakang)
- [Fitur Utama](#fitur-utama)
- [Struktur Folder](#struktur-folder)
- [Persyaratan Sistem](#persyaratan-sistem)
- [Instalasi](#instalasi)
- [File Input yang Dibutuhkan](#file-input-yang-dibutuhkan)
- [Cara Penggunaan](#cara-penggunaan)
  - [Mode 1: Laporan XML Biasa](#mode-1-laporan-xml-biasa)
  - [Mode 2: Laporan XML Diskon](#mode-2-laporan-xml-diskon)
- [Alur Kerja Internal (Pipeline)](#alur-kerja-internal-pipeline)
- [Penjelasan File di Folder Dapur](#penjelasan-file-di-folder-dapur)
- [File Helper](#file-helper)
- [File Output](#file-output)
- [Kalkulasi Pajak](#kalkulasi-pajak)
- [Penanganan Error](#penanganan-error)
- [Catatan Penting](#catatan-penting)
- [Lisensi](#lisensi)

---

## 📌 Latar Belakang

Sistem Coretax DJP mengharuskan pelaporan faktur pajak dalam format XML terstruktur. Proses input manual sangat memakan waktu dan rawan kesalahan, terutama saat volume transaksi tinggi. Proyek ini mengotomasi seluruh proses: dari pembersihan data ekspor akuntansi, kalkulasi pajak (DPP, PPN 12%, PPnBM), hingga pengisian template Excel Coretax — hanya dengan satu kali klik.

---

## ✨ Fitur Utama

- ✅ Mendukung **dua jenis laporan**: Faktur Biasa & Faktur dengan Diskon
- ✅ **Auto-detect posisi kolom** pada file sumber (toleran terhadap variasi format ekspor)
- ✅ Normalisasi format tanggal otomatis (`DD/MM/YY` → `DD/MM/YYYY`)
- ✅ Pembersihan data NPWP, ID TKU, dan Nomor Dokumen Pembeli
- ✅ Lookup otomatis **Kode Barang/Jasa** dari file referensi Helper
- ✅ Penghapusan otomatis item sesuai daftar `Helper_Del.txt`
- ✅ Filter otomatis item dengan **harga satuan negatif**
- ✅ Kalkulasi otomatis: **DPP, DPP Nilai Lain, PPN 12%, PPnBM**
- ✅ Kalkulasi diskon per item (khusus mode Diskon)
- ✅ Auto-update file Helper dari **GitHub** saat ada koneksi internet
- ✅ Penamaan file output otomatis berdasarkan **tanggal transaksi**
- ✅ Pembersihan otomatis semua file sementara setelah proses selesai

---

## 🗂️ Struktur Folder

```
Auto-Input-XML-Pajak-Coretax-1.2.3/
│
├── Buat Laporan XML Biasa.py       ← Script utama untuk faktur biasa
├── Buat Laporan XML Diskon.py      ← Script utama untuk faktur diskon
│
├── AccEFaktur.xls                  ← [INPUT] Data detail item (dari e-Faktur/sistem akuntansi)
├── AccCtxFaktur.xls                ← [INPUT] Data header faktur (dari Accurate/Ctx)
│
└── Dapur/                          ← Folder engine pemrosesan (jangan diubah)
    ├── 0_Ftch_github@AccTaxREighteen.py
    ├── 1_AccCtxFaktur_cleaner.py
    ├── 1_AccDiscFaktur_cleaner.py
    ├── 2_AccEFaktur_cleaner.py
    ├── 2_AccDiscEFaktur_cleaner.py
    ├── 3_MkNwFile.py
    ├── 3_DiscMkNwFile.py
    ├── 4_CpInvRefCalcXlook.py
    ├── 4_DiscCpInvRefCalcXlook.py
    ├── 5_CpData2Fktr.py            ← Ubah nomor Penjual di sini
    ├── 5_DiscCpData2Fktr.py        ← Ubah nomor Penjual di sini
    ├── 6_CpData2DtlFktr.py
    ├── 6_DiscCpData2DtlFktr.py
    ├── 7_HelperDate.py
    ├── 8_HelperKTPOth.py
    │
    ├── TEMPLATE_1.3.25.xlsx        ← Template Excel Coretax (jangan diubah), cukup input nomor penjual di bagian atas
    ├── KTP.txt                     ← Daftar NPWP yang perlu diganti ke National ID
    ├── KTP-OTH.txt                 ← Daftar NPWP yang perlu diganti ke Other ID
    │
    ├── Helper_160100.txt           ← Referensi kode barang (auto-update dari GitHub, ganti alamatnya di python 0_Ftch_github@AccTaxREighteen.py jika mau)
    ├── Helper_200104.txt
    ├── Helper_340300.txt
    ├── Helper_400800.txt
    ├── Helper_401100.txt
    ├── Helper_401300.txt
    ├── Helper_830900.txt
    ├── Helper_842100.txt
    ├── Helper_A.txt                ← Daftar kode yang termasuk grup Barang (A)
    ├── Helper_B.txt                ← Daftar kode yang termasuk grup Jasa (B)
    └── Helper_Del.txt              ← Daftar nama barang/jasa yang harus dihapus
```

---

## 💻 Persyaratan Sistem

| Komponen | Keterangan |
|---|---|
| OS | Windows (wajib, karena `xlwings` memerlukan Microsoft Excel terinstall) |
| Python | 3.8 atau lebih baru |
| Microsoft Excel | Wajib terinstall (digunakan `xlwings` untuk manipulasi template) |
| Koneksi Internet | Opsional (untuk auto-update file Helper dari GitHub) |

---

## 📦 Instalasi

**1. Clone atau download repositori ini**

```bash
git clone https://github.com/ACC-TAX-REIGHTEEN/Auto-Input-XML-Pajak-Coretax.git
```

**2. Install semua dependensi Python**

```bash
pip install pandas openpyxl xlwings xlsxwriter requests
```

Atau jika tersedia file `requirements.txt`:

```bash
pip install -r requirements.txt
```

**Ringkasan library yang digunakan:**

| Library | Kegunaan |
|---|---|
| `pandas` | Membaca, membersihkan, dan mentransformasi data Excel |
| `openpyxl` | Membaca dan menulis file `.xlsx` |
| `xlwings` | Manipulasi template Excel via COM (membutuhkan Excel) |
| `xlsxwriter` | Membuat file Excel baru dengan format kolom otomatis |
| `requests` | Mengunduh pembaruan file Helper dari GitHub |

---

## 📥 File Input yang Dibutuhkan

### Mode Biasa (`Buat Laporan XML Biasa.py`)

Letakkan kedua file berikut di **folder utama** (sejajar dengan script `.py`):

| Nama File | Isi | Sumber |
|---|---|---|
| `AccCtxFaktur.xls` | Data header faktur: tanggal, referensi, nama, NPWP, alamat, nilai faktur, nama penjual | Ekspor dari sistem Accurate/Ctx |
| `AccEFaktur.xls` | Data detail item: nomor invoice, barang/jasa, qty, harga, diskon | Ekspor dari modul e-Faktur/eFaktur |

### Mode Diskon (`Buat Laporan XML Diskon.py`)

| Nama File | Isi | Sumber |
|---|---|---|
| `Faktur.xls` | Data header faktur (format sama dengan `AccCtxFaktur.xls`) | Ekspor dari sistem akuntansi |
| `EFaktur.xls` | Data detail item dengan kolom diskon (format sama dengan `AccEFaktur.xls`) | Ekspor dari modul e-Faktur |

> **Catatan:** Skrip menggunakan **auto-detect header** sehingga posisi kolom tidak harus sama persis, selama nama header-nya cocok dengan yang dikenali sistem.

---

## 🚀 Cara Penggunaan

### Mode 1: Laporan XML Biasa

Gunakan ini untuk faktur **tanpa diskon tingkat item** (diskon invoice diabaikan atau nol).

1. Letakkan file `AccEFaktur.xls` dan `AccCtxFaktur.xls` di folder utama.
2. Double-click **`Buat Laporan XML Biasa.py`**, atau jalankan via terminal:

```bash
python "Buat Laporan XML Biasa.py"
```

3. Tunggu proses selesai. Output berupa file seperti:

```
XML-04-JUN-26.xlsx
```

### Mode 2: Laporan XML Diskon

Gunakan ini untuk faktur yang **mengandung diskon** dan perlu dihitung per-item.

1. Letakkan file `Faktur.xls` dan `EFaktur.xls` di folder utama.
2. Double-click **`Buat Laporan XML Diskon.py`**, atau:

```bash
python "Buat Laporan XML Diskon.py"
```

3. Output berupa file seperti:

```
XML-DISC-04-JUN-26.xlsx
```

---

## ⚙️ Alur Kerja Internal (Pipeline)

Kedua script utama menjalankan pipeline yang sama, hanya berbeda pada modul yang dipanggil:

```
[START]
   │
   ├─ Validasi keberadaan file input & folder Dapur
   │
   ├─ Salin file input ke folder Dapur
   │
   ├── [Step 0] 0_Ftch_github.py
   │     └─ Auto-update Helper_*.txt dari GitHub (jika internet tersedia)
   │
   ├── [Step 1] 1_*Faktur_cleaner.py
   │     └─ Bersihkan & normalisasi data header faktur → AccCtxFaktur_temp.xlsx
   │
   ├── [Step 2] 2_*EFaktur_cleaner.py
   │     └─ Bersihkan data detail item, hitung HJTNP (H.Jual / 1.11) → AccEFaktur_temp.xlsx
   │       [Mode Diskon tambahan: hitung DISC.TANPA, TOTAL QTY, DISC.SATUAN, DISC.ITEM]
   │
   ├── [Step 3] 3_*MkNwFile.py
   │     └─ Buat file kerja kosong MkNwFile_temp.xlsx (3 sheet: Faktur, DetailFaktur, Kalkulasi)
   │
   ├── [Step 4] 4_*CpInvRefCalcXlook.py
   │     └─ Mapping No.Inv (detail) → REFERENSI (faktur) via logika XLOOKUP manual
   │        Hasil ditulis ke sheet Kalkulasi
   │
   ├── [Step 5] 5_*CpData2Fktr.py
   │     └─ Isi sheet Faktur di MkNwFile_temp.xlsx dengan data header yang sudah bersih
   │
   ├── [Step 6] 6_*CpData2DtlFktr.py
   │     └─ Isi sheet DetailFaktur dengan data item
   │        Lookup kode barang dari Helper_*.txt
   │        Hitung DPP, DPP Nilai Lain, PPN, PPnBM
   │        Hapus item dari Helper_Del.txt & item dengan harga minus
   │
   ├── [Step 7] 7_HelperDate.py
   │     └─ Perbaiki format sel tanggal di sheet Faktur → format Excel date DD/MM/YYYY
   │
   ├── [Step 8] 8_HelperKTPOth.py
   │     └─ Koreksi Jenis ID Pembeli berdasarkan KTP.txt & KTP-OTH.txt
   │
   ├─ Transfer data MkNwFile_temp.xlsx → TEMPLATE_1.3.25.xlsx (via xlwings)
   │    ├─ Sheet Faktur: masukkan data mulai baris ke-4, tambahkan baris jika perlu
   │    └─ Sheet DetailFaktur: masukkan data mulai baris ke-2, format kolom harga
   │
   ├─ Simpan output sebagai XML-DD-MMM-YY.xlsx (atau XML-DISC-...)
   │
   └─ Bersihkan semua file sementara dari folder Dapur
[SELESAI]
```

---

## 📄 Penjelasan File di Folder Dapur

| File | Fungsi |
|---|---|
| `0_Ftch_github@AccTaxREighteen.py` | Mengunduh versi terbaru file Helper dari repositori GitHub `ACC-TAX-REIGHTEEN/Helper-For-Tax-Automation` menggunakan perbandingan MD5 checksum |
| `1_AccCtxFaktur_cleaner.py` | Membersihkan data ekspor Accurate/Ctx: normalisasi tanggal, konversi angka, generate kolom IDTKU dari NPWP |
| `1_AccDiscFaktur_cleaner.py` | Versi diskon dari cleaner di atas (input: `Faktur.xls`) |
| `2_AccEFaktur_cleaner.py` | Membersihkan data detail item dari e-Faktur, menghitung kolom `HJTNP` (Harga Jual Tanpa PPN = H.Jual / 1.11) |
| `2_AccDiscEFaktur_cleaner.py` | Versi diskon: tambahan kalkulasi `DISC.TANPA`, `TOTAL QTY per Invoice`, `DISC.SATUAN`, `DISC.ITEM` |
| `3_MkNwFile.py` / `3_DiscMkNwFile.py` | Membuat file Excel antara (`MkNwFile_temp.xlsx`) dengan 3 sheet kosong sebagai wadah data terstruktur |
| `4_CpInvRefCalcXlook.py` / `4_DiscCpInvRefCalcXlook.py` | Menghubungkan data detail (berdasarkan No. Invoice) dengan data faktur (berdasarkan Referensi) — pengganti fungsi `XLOOKUP` di Excel |
| `5_CpData2Fktr.py` / `5_DiscCpData2Fktr.py` | Mengisi sheet **Faktur** di `MkNwFile_temp.xlsx` dengan semua kolom yang dibutuhkan Coretax: Baris, Tanggal, Jenis, Kode Transaksi, ID TKU Penjual, data pembeli, dll. |
| `6_CpData2DtlFktr.py` / `6_DiscCpData2DtlFktr.py` | Mengisi sheet **DetailFaktur**: lookup kode barang, grup Barang/Jasa, satuan ukur, kalkulasi DPP, PPN, PPnBM |
| `7_HelperDate.py` | Mengonversi kolom tanggal dari string ke tipe data Excel date native |
| `8_HelperKTPOth.py` | Mengganti data NPWP/NIK berdasarkan lookup `KTP.txt` (National ID) dan `KTP-OTH.txt` (Other ID) |

---

## 📂 File Helper

File-file `.txt` di folder `Dapur` berfungsi sebagai **tabel referensi** yang dapat dikustomisasi:

| File | Fungsi |
|---|---|
| `Helper_160100.txt` s/d `Helper_842100.txt` | Daftar nama barang/jasa beserta kode komoditasnya untuk lookup otomatis |
| `Helper_A.txt` | Daftar kode komoditas yang masuk kategori **Barang** (grup A, satuan `UM.0021`) |
| `Helper_B.txt` | Daftar kode komoditas yang masuk kategori **Jasa** (grup B, satuan `UM.0030`) |
| `Helper_Del.txt` | Daftar nama barang/jasa yang akan **dihapus otomatis** dari output (misal: item retur, item non-faktur) |
| `KTP.txt` | Daftar NPWP pembeli yang harus dikonversi ke `National ID` (NIK KTP) |
| `KTP-OTH.txt` | Daftar NPWP pembeli yang harus dikonversi ke `Other ID` |

> File Helper dengan prefix `Helper_16xxxx` dst. **diperbarui otomatis** dari GitHub setiap kali script dijalankan (jika ada koneksi internet). Perubahan manual pada file-file ini dapat ter-overwrite.
>
> File `KTP.txt`, `KTP-OTH.txt`, dan `Helper_Del.txt` **tidak di-overwrite** oleh GitHub update — aman untuk dikustomisasi secara lokal.

---

## 📤 File Output

| Kondisi | Nama File Output |
|---|---|
| Mode Biasa, tanggal terbaca dari data | `XML-DD-MMM-YY.xlsx` (contoh: `XML-04-JUN-26.xlsx`) |
| Mode Biasa, tanggal tidak terbaca | `XML-DD-MMM-YY.xlsx` (berdasarkan tanggal hari ini) |
| Mode Diskon, tanggal terbaca dari data | `XML-DISC-DD-MMM-YY.xlsx` (contoh: `XML-DISC-04-JUN-26.xlsx`) |
| Mode Diskon, tanggal tidak terbaca | `XML-DISC-DD-MMM-YY.xlsx` (berdasarkan tanggal hari ini) |

File disimpan di **folder utama** (sejajar dengan script). File yang sudah ada dengan nama sama akan diganti otomatis.

---

## 🧮 Kalkulasi Pajak

### Mode Biasa

| Kolom | Formula |
|---|---|
| `HJTNP` (Harga Jual Tanpa PPN) | `H.Jual / 1.11` |
| `Harga Satuan` | = `HJTNP` |
| `Total Diskon` | = `Discount Faktur` dari data sumber |
| `DPP` | `(Harga Satuan × Qty) - Total Diskon` |
| `DPP Nilai Lain` | `DPP × (11/12)` |
| `Tarif PPN` | `12%` |
| `PPN` | `DPP Nilai Lain × 12%` |
| `Tarif PPnBM` | `0%` |
| `PPnBM` | `0` |

### Mode Diskon (tambahan)

| Kolom | Formula |
|---|---|
| `DISC. TANPA` | `Discount Faktur / 1.11` |
| `TOTAL QTY` | Jumlah total qty semua item dalam satu nomor invoice |
| `DISC. SATUAN` | `DISC. TANPA / TOTAL QTY` |
| `DISC. ITEM` | `DISC. SATUAN × Qty item` |
| `Total Diskon` (DetailFaktur) | = `DISC. ITEM` |

---

## ⚠️ Penanganan Error

Script menampilkan pesan `[ERROR]` dan menghentikan proses jika:

- Folder `Dapur` tidak ditemukan
- File input (`.xls`) tidak ditemukan di folder utama
- Salah satu file wajib di dalam folder `Dapur` hilang
- Salah satu sub-script (`0_` s/d `8_`) gagal dieksekusi
- `MkNwFile_temp.xlsx` tidak terbentuk setelah proses pipeline
- Template Excel tidak dapat dibuka atau disalin

Setiap error menampilkan pesan deskriptif dan meminta konfirmasi `Enter` sebelum keluar, sehingga pengguna dapat membaca pesan kesalahan.

---

## 📝 Catatan Penting

1. **Jangan rename atau pindahkan file di folder `Dapur`** — semua file dipanggil berdasarkan nama yang sudah ditentukan.
2. **ID TKU Penjual** saat ini di-hardcode sebagai `0000000000000000000000` di `5_CpData2Fktr.py`. Ubah sesuai dengan NPWP + ID TKU perusahaan Anda jika berbeda.
3. **Kode Transaksi** di-hardcode sebagai `04` (Penyerahan Kepada Selain Pemungut). Sesuaikan jika transaksi Anda berbeda.
4. Script tidak membuat backup file input secara otomatis. Pastikan menyimpan cadangan sebelum menjalankan.
5. **Pesan moral** dari developer: *"Jangan pernah percaya kepada siapapun, bahkan jika itu dirimu sendiri. Selalu lakukan cek ulang, karena mesin juga bisa salah."*
6. Apabila tidak ingin ada pembersihan, hapus saja python pembersihan dan sesuaikan datanya langsung, supaya dapat digunakan untuk umum dan hapus python di Orkestrator utama.

---

## 📜 Lisensi

Proyek ini dikembangkan untuk keperluan internal perpajakan. Silakan sesuaikan dengan kebutuhan organisasi Anda.

---

*Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)*
