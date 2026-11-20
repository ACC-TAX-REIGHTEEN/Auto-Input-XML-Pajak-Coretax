# 🧾 Auto-Input XML Pajak Coretax

![Python](https://img.shields.io/badge/Python-3.8%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-ETL%20Pipeline-150458?style=for-the-badge&logo=pandas&logoColor=white)
![Excel](https://img.shields.io/badge/Microsoft%20Excel-xlwings%20%7C%20openpyxl-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Coretax](https://img.shields.io/badge/DJP%20Coretax-Template%20v1.6.1-0A4D8C?style=for-the-badge)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white)

Sistem otomasi ETL berbasis Python (*Pandas* & *xlwings*) untuk mengonversi data ekspor faktur penjualan dari sistem akuntansi (seperti Accurate / e-Faktur berformat `.xls`) menjadi berkas Excel siap konversi XML untuk portal **DJP Coretax** (`Template_v.1.6.1.xlsx`). Dilengkapi arsitektur pembacaan cepat `calamine` dengan *auto-fallback*, konfigurasi terpusat `config.conf`, sinkronisasi kamus barang otomatis via GitHub, serta modul **Audit Kelayakan Data Coretax & Kepatuhan SOP** sebelum unggah.

---

## Daftar Isi

- [Fitur Utama](#fitur-utama)
- [Arsitektur & Alur Kerja](#arsitektur--alur-kerja)
- [Teknologi yang Digunakan](#teknologi-yang-digunakan)
- [Struktur Direktori](#struktur-direktori)
- [Prasyarat Sistem](#prasyarat-sistem)
- [Panduan Instalasi & Konfigurasi](#panduan-instalasi--konfigurasi)
- [Cara Penggunaan](#cara-penggunaan)
  - [File Input yang Dibutuhkan](#file-input-yang-dibutuhkan)
  - [Mode 1: Laporan XML Biasa](#mode-1-laporan-xml-biasa)
  - [Mode 2: Laporan XML Diskon](#mode-2-laporan-xml-diskon)
  - [Kalkulasi Pajak & Diskon](#kalkulasi-pajak--diskon)
  - [File Output](#file-output)
- [Referensi Konfigurasi](#referensi-konfigurasi)
  - [Parameter `Dapur/config.conf`](#parameter-dapurconfigconf)
  - [Daftar File Helper (`Dapur/*.txt`)](#daftar-file-helper-dapurtxt)
- [Penanganan Masalah Umum (Troubleshooting)](#penanganan-masalah-umum-troubleshooting)
- [Kontribusi & Lisensi](#kontribusi--lisensi)

---

## Fitur Utama

### 1. Orkestrator Ganda (Mode Biasa & Mode Diskon)
- **`Buat Laporan XML Biasa.py`**: Mengolah faktur penjualan standar tanpa pembagian diskon proporsional tingkat item (menggunakan input `AccCtxFaktur.xls` & `AccEFaktur.xls`).
- **`Buat Laporan XML Diskon.py`**: Menghitung distribusi diskon per item berdasarkan proporsi kuantitas (`TOTAL QTY`) dalam satu nomor invoice (menggunakan input `Faktur.xls` & `EFaktur.xls`).

### 2. Pembersihan & Transformasi Berbasis Pandas dengan High-Speed Engine
- **Auto-Detect Header**: Memindai 20 baris pertama file sumber secara dinamis sehingga tahan terhadap perubahan posisi kolom atau baris judul ekspor akuntansi.
- **Engine Adaptif (`calamine` → `xlrd` / `openpyxl`)**: Memprioritaskan `python-calamine` untuk pembacaan `.xls`/`.xlsx` berkecepatan tinggi, lalu otomatis beralih ke `xlrd` (untuk `.xls`), `openpyxl` (untuk `.xlsx`), atau `pd.read_csv` (fallback pada pembersih detail) apabila `calamine` belum terpasang di komputer cabang.
- **Normalisasi Data Pajak**: Menstandarkan format tanggal (`DD/MM/YYYY`), angka desimal lokal, pembentukan otomatis `ID TKU Pembeli`, klasifikasi `Jenis ID Pembeli` (`TIN`, `National ID`, `Other ID`), serta relasi `No.Inv` ke nomor baris faktur dengan logika *XLOOKUP* mandiri.

### 3. Konfigurasi Terpusat (`Dapur/config.conf`)
- Seluruh parameter perpajakan (`Jenis Faktur`, `Kode Transaksi`, `ID TKU Penjual`, `Keterangan Tambahan`, `Dokumen Pendukung`, `Cap Fasilitas`), sakelar pembagi PPN (`1.11` vs `1.0`), kriteria audit, hingga token GitHub diatur dari satu berkas konfigurasi tanpa perlu menyunting kode Python.

### 4. Sinkronisasi Helper Otomatis dari GitHub
- Modul `0_Ftch_github@AccTaxREighteen.py` memverifikasi integritas 15 berkas kamus referensi (`Helper_*.txt`, `KTP.txt`, `KTP-OTH.txt`) menggunakan **MD5 checksum** terhadap repositori `ACC-TAX-REIGHTEEN/Helper-For-Tax-Automation` dan hanya mengunduh berkas yang berubah (mendukung `git_token` opsional untuk akses terautentikasi).

### 5. Modul Audit Kelayakan Coretax & Kepatuhan SOP (`9_EndReCalcRecon.py`)
- Memindai hasil olahan sebelum impor untuk mendeteksi baris berisiko ditolak portal Coretax atau menyalahi SOP nota, seperti:
  - **`DITOLAK CORETAX`**: Nilai `DPP` atau `PPN` minus akibat diskon melebihi total harga, `Harga Satuan` negatif, atau `Jumlah Barang Jasa <= 0`.
  - **`DITERIMA (BERISIKO SOP)`**: Item dengan `Harga Satuan == 0` (misalnya barang bonus/sampel yang semestinya dipisah notanya).
- Menghasilkan workbook laporan terpisah (`AUDIT_DATA_CTX-*.xlsx`) berisi 3 sheet: `Temuan_Detail`, `Faktur_Terdampak`, dan `Ringkasan_Audit`.

### 6. Injeksi Template Coretax v1.6.1 dengan Proteksi Tipe Data
- Menyuntikkan hasil akhir ke `Dapur/Template_v.1.6.1.xlsx` menggunakan `xlwings` sambil menjaga penanda baris `END` tetap utuh.
- Memaksa `ID TKU Penjual` dan `Kode Barang Jasa` (seperti `000000`) disimpan sebagai teks murni berawalan apostrof (`'`) agar angka nol di depan tidak hilang di Excel, serta memformat `Harga Satuan` ke tipe numerik (`###0,00`).
- Menamai file hasil akhir otomatis sesuai tanggal faktur pertama (format `DD-MMM-YY` berbahasa Indonesia) dan membersihkan seluruh file sementara di `Dapur/`.

---

## Arsitektur & Alur Kerja

```text
+-----------------------------------------------------------------------------+
|                         INPUT FILE (FOLDER UTAMA)                           |
|  Mode Biasa  : AccCtxFaktur.xls + AccEFaktur.xls                            |
|  Mode Diskon : Faktur.xls       + EFaktur.xls                               |
+--------------------------------------+--------------------------------------+
                                       |
                                       v
+-----------------------------------------------------------------------------+
|  ORKESTRATOR: Buat Laporan XML Biasa.py / Buat Laporan XML Diskon.py        |
|  - Validasi folder Dapur/, file input, dan 27 berkas wajib di Dapur/        |
|  - Menyalin file input ke Dapur/ lalu mengeksekusi pipeline secara sekuen   |
+--------------------------------------+--------------------------------------+
                                       |
    +----------------------------------+----------------------------------+
    |                          PIPELINE DAPUR/                            |
    |                                                                     |
    |  [0] 0_Ftch_github@AccTaxREighteen.py                               |
    |      └─ Cek internet & MD5 checksum -> auto-update 15 file Helper   |
    |         dari GitHub (opsional git_token dari config.conf)           |
    |                                                                     |
    |  [1] 1_AccCtxFaktur_cleaner.py / 1_AccDiscFaktur_cleaner.py         |
    |      └─ Auto-detect kolom header faktur, normalisasi TGL & angka,   |
    |         generate IDTKU -> AccCtxFaktur_temp.xlsx                    |
    |                                                                     |
    |  [2] 2_AccEFaktur_cleaner.py / 2_AccDiscEFaktur_cleaner.py          |
    |      └─ Auto-detect kolom detail, baca [CALC] calc_11 (1.11 / 1.0), |
    |         hitung HJTNP (+ DISC. TANPA, TOTAL QTY, DISC. SATUAN,       |
    |         DISC. ITEM khusus Mode Diskon) -> AccEFaktur_temp.xlsx      |
    |                                                                     |
    |  [3] 3_MkNwFile.py / 3_DiscMkNwFile.py                              |
    |      └─ Buat kerangka MkNwFile_temp.xlsx                            |
    |         (Sheet: Faktur, DetailFaktur, Kalkulasi)                    |
    |                                                                     |
    |  [4] 4_CpInvRefCalcXlook.py / 4_DiscCpInvRefCalcXlook.py            |
    |      └─ Mapping No.Inv (detail) -> indeks baris REFERENSI (faktur)  |
    |         ke sheet Kalkulasi (logika XLOOKUP)                         |
    |                                                                     |
    |  [5] 5_CpData2Fktr.py / 5_DiscCpData2Fktr.py                        |
    |      └─ Isi 18 kolom sheet Faktur + parameter [DATA] config.conf    |
    |         (TIN / Other ID, Negara IDN, Period Dok Pendukung MMYYYY)   |
    |                                                                     |
    |  [6] 6_CpData2DtlFktr.py / 6_DiscCpData2DtlFktr.py                  |
    |      └─ Isi 14 kolom DetailFaktur, lookup 10 Helper_*.txt,          |
    |         grup A (UM.0021) / B (UM.0030), hitung DPP, DPP Nilai Lain  |
    |         (11/12), PPN 12%, hapus item Helper_Del.txt & harga < 0     |
    |                                                                     |
    |  [7] 7_HelperDate.py                                                |
    |      └─ Konversi kolom Tanggal Faktur -> tipe Excel Date DD/MM/YYYY |
    |                                                                     |
    |  [8] 8_HelperKTPOth.py                                              |
    |      └─ Konversi NPWP/NIK -> National ID (KTP.txt) /                |
    |         Other ID (KTP-OTH.txt)                                      |
    |                                                                     |
    |  [9] 9_EndReCalcRecon.py                                            |
    |      └─ Baca [AUDIT] config.conf, join referensi Faktur, klasifikasi|
    |         DITOLAK CORETAX / BERISIKO SOP -> AUDIT_DATA_CTX.xlsx       |
    +----------------------------------+----------------------------------+
                                       |
                                       v
+-----------------------------------------------------------------------------+
|  INJEKSI TEMPLATE & PEMBERSIHAN (xlwings + Microsoft Excel)                 |
|  - Salin Dapur/Template_v.1.6.1.xlsx -> TEMPLATE_temp.xlsx                  |
|  - Sisipkan baris dinamis sebelum penanda 'END' di Faktur & DetailFaktur    |
|  - Sanitasi ID TKU Penjual ('), Kode Barang Jasa ('), Harga Satuan (float)  |
|  - Pindahkan output & file audit ke folder utama, hapus semua file temp     |
+--------------------+-----------------------------------+--------------------+
                     |                                   |
                     v                                   v
   +-----------------------------------+   +-----------------------------------+
   | OUTPUT UTAMA (SIAP IMPOR CORETAX) |   | OUTPUT AUDIT (OPSIONAL VIA CONF)  |
   | - XML-DD-MMM-YY.xlsx              |   | - AUDIT_DATA_CTX-DD-MMM-YY.xlsx   |
   | - XML-DISC-DD-MMM-YY.xlsx         |   | - AUDIT_DATA_CTX-DISC-DD-MMM-YY   |
   +-----------------------------------+   +-----------------------------------+
```

---

## Teknologi yang Digunakan

| Komponen / Pustaka | Peran dalam Sistem |
|---|---|
| **Python 3.8+** | Runtime utama untuk orkestrator dan seluruh modul ETL di `Dapur/`. |
| **`pandas`** | Transformasi tabular, *auto-detect* header, agregasi diskon per invoice (`groupby`), penggabungan data audit (`merge`), dan kalkulasi pajak. |
| **`python-calamine`** *(Direkomendasikan)* | Engine utama pembacaan berkas `.xls` dan `.xlsx` berkecepatan tinggi berbasis Rust. |
| **`xlrd`** *(Fallback `.xls`)* | Engine cadangan otomatis apabila `python-calamine` tidak tersedia saat membaca file sumber `.xls`. |
| **`openpyxl`** | Penulisan, pemformatan tanggal native (`7_HelperDate.py`), modifikasi KTP/Other ID (`8_HelperKTPOth.py`), auto-fit lebar kolom, dan fallback baca `.xlsx`. |
| **`xlsxwriter`** | Pembuatan kerangka awal `MkNwFile_temp.xlsx` (3 sheet terstruktur) pada tahap 3. |
| **`xlwings`** | Otomasi COM Microsoft Excel untuk menyisipkan baris dan menyalin data ke `Template_v.1.6.1.xlsx` tanpa merusak struktur validasi, *dropdown*, maupun penanda `END` resmi Coretax. |
| **`requests`** | Pengecekan konektivitas dan pengunduhan pembaruan file Helper dari GitHub Raw (`0_Ftch_github@AccTaxREighteen.py`). |
| **`configparser` & `hashlib`** *(Stdlib)* | Pembacaan parameter `Dapur/config.conf` dan validasi *MD5 checksum* file Helper. |

---

## Struktur Direktori

```text
Auto-Input-XML-Pajak-Coretax/
├── Buat Laporan XML Biasa.py            # Orkestrator utama mode Faktur Biasa
├── Buat Laporan XML Diskon.py           # Orkestrator utama mode Faktur Diskon
├── README.md                            # Dokumentasi proyek
│
├── [Input Mode Biasa]
│   ├── AccCtxFaktur.xls                 # Header faktur (ekspor Accurate/Ctx)
│   └── AccEFaktur.xls                   # Detail item faktur (ekspor e-Faktur)
│
├── [Input Mode Diskon]
│   ├── Faktur.xls                       # Header faktur untuk mode diskon
│   └── EFaktur.xls                      # Detail item faktur dengan diskon
│
└── Dapur/                               # Modul pemrosesan internal (jangan rename file)
    ├── __init__.py                      # Penanda paket modul (wajib ada untuk validasi)
    ├── config.conf                      # Konfigurasi terpusat ([CALC], [DATA], [AUDIT], [GIT])
    ├── Template_v.1.6.1.xlsx            # Template resmi Coretax v1.6.1 (8 sheet)
    │
    ├── 0_Ftch_github@AccTaxREighteen.py # Auto-update 15 file Helper dari GitHub via MD5
    ├── 1_AccCtxFaktur_cleaner.py        # Pembersih header faktur (Mode Biasa)
    ├── 1_AccDiscFaktur_cleaner.py       # Pembersih header faktur (Mode Diskon)
    ├── 2_AccEFaktur_cleaner.py          # Pembersih detail item & kalkulasi HJTNP (Mode Biasa)
    ├── 2_AccDiscEFaktur_cleaner.py      # Pembersih detail item & kalkulasi diskon per item (Mode Diskon)
    ├── 3_MkNwFile.py                    # Generator file kerja MkNwFile_temp.xlsx (Mode Biasa)
    ├── 3_DiscMkNwFile.py                # Generator file kerja MkNwFile_temp.xlsx (Mode Diskon)
    ├── 4_CpInvRefCalcXlook.py           # Pemetaan XLOOKUP No.Inv -> REFERENSI (Mode Biasa)
    ├── 4_DiscCpInvRefCalcXlook.py       # Pemetaan XLOOKUP No.Inv -> REFERENSI (Mode Diskon)
    ├── 5_CpData2Fktr.py                 # Pengisi sheet Faktur dari config.conf (Mode Biasa)
    ├── 5_DiscCpData2Fktr.py             # Pengisi sheet Faktur dari config.conf (Mode Diskon)
    ├── 6_CpData2DtlFktr.py              # Pengisi DetailFaktur, lookup Helper & kalkulasi pajak (Mode Biasa)
    ├── 6_DiscCpData2DtlFktr.py          # Pengisi DetailFaktur dengan DISC. ITEM (Mode Diskon)
    ├── 7_HelperDate.py                  # Konversi teks tanggal -> format Excel Date DD/MM/YYYY
    ├── 8_HelperKTPOth.py                # Klasifikasi pembeli National ID (KTP) & Other ID
    ├── 9_EndReCalcRecon.py              # Modul audit kelayakan Coretax & kepatuhan SOP
    │
    ├── Helper_000000.txt                # Daftar nama barang/jasa berkode 000000
    ├── Helper_120101.txt                # Daftar nama barang/jasa berkode 120101
    ├── Helper_160100.txt                # Daftar nama barang/jasa berkode 160100
    ├── Helper_200104.txt                # Daftar nama barang/jasa berkode 200104
    ├── Helper_340300.txt                # Daftar nama barang/jasa berkode 340300
    ├── Helper_400800.txt                # Daftar nama barang/jasa berkode 400800
    ├── Helper_401100.txt                # Daftar nama barang/jasa berkode 401100
    ├── Helper_401300.txt                # Daftar nama barang/jasa berkode 401300
    ├── Helper_830900.txt                # Daftar nama barang/jasa berkode 830900
    ├── Helper_842100.txt                # Daftar nama barang/jasa berkode 842100
    ├── Helper_A.txt                     # Daftar kode komoditas kategori Barang (A -> UM.0021)
    ├── Helper_B.txt                     # Daftar kode komoditas kategori Jasa (B -> UM.0030)
    ├── Helper_Del.txt                   # Daftar nama item yang dihapus otomatis dari detail
    ├── KTP.txt                          # Daftar nomor ID pembeli yang diubah ke National ID
    └── KTP-OTH.txt                      # Daftar nomor ID pembeli yang diubah ke Other ID
```

---

## Prasyarat Sistem

| Komponen | Persyaratan Minimum | Keterangan |
|---|---|---|
| **Sistem Operasi** | Windows 10 / Windows 11 | Wajib Windows karena tahap akhir menggunakan `xlwings` (COM Automation Microsoft Excel). |
| **Microsoft Excel** | Microsoft Excel Desktop terpasang & aktif | Digunakan secara *headless* (`visible=False`) oleh `xlwings`. |
| **Python** | Python 3.8 atau lebih baru | Disarankan Python 3.10+ agar kompatibel optimal dengan `python-calamine`. |
| **Koneksi Internet** | Opsional | Dibutuhkan oleh tahap `0_` untuk auto-update file Helper dari GitHub. Jika offline, script tetap berjalan menggunakan file lokal. |

---

## Panduan Instalasi & Konfigurasi

### 1. Clone Repositori

```bash
git clone [https://github.com/ACC-TAX-REIGHTEEN/Auto-Input-XML-Pajak-Coretax.git](https://github.com/ACC-TAX-REIGHTEEN/Auto-Input-XML-Pajak-Coretax.git)
cd Auto-Input-XML-Pajak-Coretax
```

### 2. Buat & Aktifkan Virtual Environment *(Direkomendasikan)*

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 3. Pasang Dependensi Python

Pasang paket utama beserta `python-calamine` (untuk performa baca tinggi) dan `xlrd` (cadangan pembaca `.xls`):

```bash
pip install pandas openpyxl xlwings xlsxwriter requests python-calamine xlrd
```

> **Catatan untuk PC Cabang Lama:** Apabila instalasi `python-calamine` gagal karena arsitektur atau versi Python lama, cukup pasang tanpa `python-calamine`:
> ```bash
> pip install pandas openpyxl xlwings xlsxwriter requests xlrd
> ```
> Seluruh modul secara otomatis akan beralih menggunakan `xlrd` (untuk `.xls`) dan `openpyxl` (untuk `.xlsx`).

### 4. Sesuaikan Identitas Penjual pada Template & `config.conf`

1. **Ubah `Dapur/config.conf`** sesuai identitas perpajakan perusahaan Anda, terutama `dat_id_tku_penjual` (22 digit NPWP + ID TKU Penjual):
   ```ini
   [CALC]
   calc_11 = Ya

   [DATA]
   dat_jenis_faktur = Normal
   dat_kode_transaksi = 04
   dat_keterangan_tambahan = 
   dat_dokumen_pendukung = 
   dat_cap_fasilitas = 
   dat_id_tku_penjual = 0012345678999000000000

   [AUDIT]
   aud = Ya
   aud_sheet = DetailFaktur
   aud_col = Harga Satuan
   aud_col_val = 0
   aud_col_ref = Ya

   [GIT]
   git_token = 
   ```
2. **Ubah NPWP Penjual di Template Excel**: Buka `Dapur/Template_v.1.6.1.xlsx`, lalu perbarui sel **`C1`** pada sheet **`Faktur`** (`NPWP Penjual`) dengan 16 digit NPWP perusahaan Anda. Jangan mengubah susunan header atau menghapus baris penanda `END`.

---

## Cara Penggunaan

### File Input yang Dibutuhkan

Letakkan file ekspor `.xls` langsung di **folder utama** (sejajar dengan file `Buat Laporan XML *.py`).

| Mode | File Header Faktur | File Detail Item | Header Minimum yang Dikenali Sistem |
|---|---|---|---|
| **Mode Biasa** | `AccCtxFaktur.xls` | `AccEFaktur.xls` | **Header:** `TGL`, `REFERENSI`, `NAMA`, `NPWP`, `No PKP Pelanggan`, `ALAMAT PAJAK 1`, `Negara Pelanggan`, `DISC INV`, `Nilai Faktur`, `Tax 1 Base Currency`, `Nama Penjual`<br>**Detail:** `No.Inv`, `Tgl Faktur`, `Nama Pelanggan`, `Alamat Pajak 1 Pelanggan`, `Kota`, `No. Barang`, `Nama Barang`, `Qty`, `Unit 1 Barang`, `H.Jual`, `Harga Sat`, `Total Harga`, `Nilai Faktur`, `Tax 1`, `Discount Faktur`, `No. Pelanggan`, `Nomor Pajak Pelanggan` |
| **Mode Diskon** | `Faktur.xls` | `EFaktur.xls` | Sama dengan di atas (`Unit 1 Barang` tidak wajib di detail). Khusus `Faktur.xls`, format kolom `TGL` dianjurkan berpola `DD/MM/YY`. |

### Mode 1: Laporan XML Biasa

Gunakan mode ini untuk transaksi penjualan tanpa pembagian diskon faktur ke masing-masing item.

1. Pastikan `AccCtxFaktur.xls` dan `AccEFaktur.xls` berada di folder utama.
2. Jalankan via klik ganda pada **`Buat Laporan XML Biasa.py`** atau melalui PowerShell/CMD:
   ```powershell
   python "Buat Laporan XML Biasa.py"
   ```
3. Pantau log eksekusi tahap `0_` hingga `9_` serta ringkasan audit di jendela terminal.
4. Tekan `Enter` setelah muncul pesan `--> PROSES SELESAI. Folder Dapur bersih.`

### Mode 2: Laporan XML Diskon

Gunakan mode ini untuk transaksi yang memiliki diskon di tingkat faktur dan harus didistribusikan secara proporsional ke tiap baris item berdasarkan kuantitas.

1. Pastikan `Faktur.xls` dan `EFaktur.xls` berada di folder utama.
2. Jalankan via klik ganda pada **`Buat Laporan XML Diskon.py`** atau melalui PowerShell/CMD:
   ```powershell
   python "Buat Laporan XML Diskon.py"
   ```
3. Pantau log proses hingga selesai, lalu tekan `Enter` untuk menutup jendela terminal.

### Kalkulasi Pajak & Diskon

Pembagi dasar (`pembagi`) dikontrol oleh `calc_11` di `Dapur/config.conf`: bernilai **`1.11`** jika `calc_11 = Ya` (harga sumber sudah termasuk PPN 11%), atau **`1.0`** jika `calc_11 = No` (harga sumber sudah berupa DPP murni).

| Field Output (`DetailFaktur`) | Mode Biasa | Mode Diskon |
|---|---|---|
| `HJTNP` (`Harga Satuan`) | `H.Jual / pembagi` | `H.Jual / pembagi` |
| `DISC. TANPA` *(antara)* | — | `Discount Faktur / pembagi` |
| `TOTAL QTY` *(antara)* | — | `SUM(Qty)` per `No.Inv` |
| `DISC. SATUAN` *(antara)* | — | `DISC. TANPA / TOTAL QTY` |
| `Total Diskon` | `Discount Faktur` | `DISC. ITEM = DISC. SATUAN × Qty` |
| `DPP` | `(Harga Satuan × Qty) - Total Diskon` | `(Harga Satuan × Qty) - Total Diskon` |
| `DPP Nilai Lain` | `(11 / 12) × DPP` | `(11 / 12) × DPP` |
| `Tarif PPN` | `12` | `12` |
| `PPN` | `(DPP Nilai Lain × 12) / 100` | `(DPP Nilai Lain × 12) / 100` |
| `Tarif PPnBM` & `PPnBM` | `0` | `0` |

### File Output

Seluruh hasil disimpan di **folder utama** (menimpa file lama apabila memiliki nama tanggal yang sama):

| Jenis Output | Mode Biasa | Mode Diskon | Isi Berkas |
|---|---|---|---|
| **File Coretax Utama** | `XML-DD-MMM-YY.xlsx` *(cth: `XML-20-NOV-26.xlsx`)* | `XML-DISC-DD-MMM-YY.xlsx` *(cth: `XML-DISC-20-NOV-26.xlsx`)* | Berbasis `Template_v.1.6.1.xlsx` yang telah terisi sheet `Faktur` & `DetailFaktur` siap konversi ke XML Coretax. |
| **File Laporan Audit** *(jika `aud = Ya`)* | `AUDIT_DATA_CTX-DD-MMM-YY.xlsx` | `AUDIT_DATA_CTX-DISC-DD-MMM-YY.xlsx` | 3 sheet hasil pemeriksaan: `Temuan_Detail`, `Faktur_Terdampak`, dan `Ringkasan_Audit`. |

> Penamaan tanggal diambil otomatis dari tanggal faktur baris pertama pada data. Jika tanggal tidak dapat dibaca, sistem menggunakan tanggal sistem hari ini dengan singkatan bulan Indonesia (`JAN`, `FEB`, `MAR`, `APR`, `MEI`, `JUN`, `JUL`, `AGU`, `SEP`, `OKT`, `NOV`, `DES`).

---

## Referensi Konfigurasi

### Parameter `Dapur/config.conf`

| Section | Parameter | Contoh Nilai | Default Fallback | Penjelasan |
|---|---|---|---|---|
| `[CALC]` | `calc_11` | `Ya` / `No` | `1.0` (tanpa bagi) | Jika diisi `Ya`, `Yes`, `True`, atau `1`, harga jual (`H.Jual`) dan diskon (`Discount Faktur`) dibagi `1.11` untuk mengeluarkan PPN. Jika `No`, dibagi `1.0`. |
| `[DATA]` | `dat_jenis_faktur` | `Normal` | `Normal` | Nilai tetap kolom `Jenis Faktur` pada sheet `Faktur`. |
| `[DATA]` | `dat_kode_transaksi` | `04` | `04` | Kode transaksi faktur pajak Coretax (cth: `04` untuk DPP Nilai Lain). |
| `[DATA]` | `dat_keterangan_tambahan` | *(kosong)* / `TD.00501` | `""` | Kode keterangan tambahan (lihat sheet `REF-KetTambahan` di template, relevan untuk kode transaksi `07`/`08`). |
| `[DATA]` | `dat_dokumen_pendukung` | *(kosong)* | `""` | Nomor dokumen pendukung jika dipersyaratkan transaksi. |
| `[DATA]` | `dat_cap_fasilitas` | *(kosong)* / `TD.01101` | `""` | Kode cap fasilitas perpajakan (lihat sheet `REF-CapFasilitas` di template). |
| `[DATA]` | `dat_id_tku_penjual` | `0012345678999000000000` | `12345` | ID TKU Penjual 22 digit (otomatis dibersihkan dan disimpan sebagai teks di template). |
| `[AUDIT]` | `aud` | `Ya` / `Tidak` | `Tidak` *(atau `Ya` jika file conf tidak ada)* | Mengaktifkan atau melewati tahap 9 (`9_EndReCalcRecon.py`). |
| `[AUDIT]` | `aud_sheet` | `DetailFaktur` | `DetailFaktur` | Nama sheet pada `MkNwFile_temp.xlsx` yang diaudit. |
| `[AUDIT]` | `aud_col` | `Harga Satuan` | `Harga Satuan` | Nama kolom acuan utama yang diuji oleh modul audit. |
| `[AUDIT]` | `aud_col_val` | `0` / `1` / angka kustom | `0` | `0` = cari anomali (`aud_col == 0` atau `DPP < 0` atau `PPN < 0`); `1` = tampilkan data non-nol (`aud_col != 0`); nilai numerik lain = filter tepat sama dengan nilai tersebut. |
| `[AUDIT]` | `aud_col_ref` | `Ya` | `Ya` | Penanda pengayaan kolom referensi faktur pada hasil audit (penggabungan `Referensi`, `Tanggal Faktur`, `Nama Pembeli`, `NPWP/NIK Pembeli` dilakukan otomatis dari sheet `Faktur`). |
| `[GIT]` | `git_token` | `ghp_xxxxxxxx` / *(kosong)* | `""` | Personal Access Token GitHub opsional untuk menghindari *rate limit* atau mengakses repositori Helper yang bersifat privat. |

### Daftar File Helper (`Dapur/*.txt`)

| File Helper | Fungsi dalam Pipeline | Pemetaan Output |
|---|---|---|
| `Helper_000000.txt`, `Helper_120101.txt`, `Helper_160100.txt`, `Helper_200104.txt`, `Helper_340300.txt`, `Helper_400800.txt`, `Helper_401100.txt`, `Helper_401300.txt`, `Helper_830900.txt`, `Helper_842100.txt` | Kamus pemetaan nama barang/jasa ke 6 digit **Kode Barang Jasa** Coretax (1 nama item per baris). | Mengisi kolom `Kode Barang Jasa` pada sheet `DetailFaktur`. |
| `Helper_A.txt` | Daftar kode barang yang masuk kategori **Barang (`A`)**. | `Barang/Jasa = A`, `Nama Satuan Ukur = UM.0021`. |
| `Helper_B.txt` | Daftar kode jasa yang masuk kategori **Jasa (`B`)**. | `Barang/Jasa = B`, `Nama Satuan Ukur = UM.0030`. |
| `Helper_Del.txt` | Daftar nama item yang dikecualikan (dihapus otomatis sebelum injeksi template). | Baris dengan `Nama Barang/Jasa` cocok akan dibuang dari `DetailFaktur`. |
| `KTP.txt` | Daftar nomor identitas pembeli yang menggunakan NIK/KTP. | `Jenis ID Pembeli = National ID`, nilai asli dipindah ke `Nomor Dokumen Pembeli`, `NPWP/NIK Pembeli = 0000000000000000`, `ID TKU Pembeli = 000000`. |
| `KTP-OTH.txt` | Daftar nomor identitas pembeli yang diperlakukan sebagai identitas lainnya. | `Jenis ID Pembeli = Other ID`, `NPWP/NIK Pembeli = 0000000000000000`, `Nomor Dokumen Pembeli = 0000000000000000`, `ID TKU Pembeli = 000000`. |

> **Perhatian Sinkronisasi GitHub:** Seluruh 15 berkas `.txt` di atas terdaftar di `FILES_TO_UPDATE` pada `0_Ftch_github@AccTaxREighteen.py`. Ketika komputer terhubung ke internet dan *MD5 checksum* file lokal berbeda dengan repositori GitHub `ACC-TAX-REIGHTEEN/Helper-For-Tax-Automation`, berkas lokal akan otomatis diperbarui mengikuti versi GitHub. Jika ingin menggunakan daftar lokal murni tanpa sinkronisasi otomatis, perbarui URL target di `0_Ftch_github@AccTaxREighteen.py` ke repositori Anda sendiri atau nonaktifkan pemanggilannya di `EXECUTION_ORDER`.

---

## Penanganan Masalah Umum (Troubleshooting)

| Gejala / Pesan Error | Penyebab Umum | Solusi |
|---|---|---|
| `File 'AccEFaktur.xls' / 'Faktur.xls' tidak ditemukan di folder utama.` | Nama file input tidak sesuai dengan mode yang dijalankan atau file masih di folder lain. | Pastikan nama persis: `AccCtxFaktur.xls` + `AccEFaktur.xls` (Mode Biasa) atau `Faktur.xls` + `EFaktur.xls` (Mode Diskon) sejajar dengan script utama. |
| `File '<nama_file>' hilang dari folder Dapur.` | Salah satu dari 27 file wajib di `Dapur/` (termasuk `__init__.py`, `Template_v.1.6.1.xlsx`, atau `Helper_000000.txt`) terhapus atau diganti nama. | Kembalikan berkas yang hilang dari repositori asli tanpa mengubah nama berkasnya. |
| `Error membaca file` saat membaca `.xls` pada PC tanpa `calamine` | Paket `python-calamine` dan fallback `xlrd` sama-sama belum terpasang. | Jalankan `pip install python-calamine xlrd`. |
| `Error saat manipulasi Excel` / proses macet di `xlwings` | Microsoft Excel belum terpasang, lisensi/dialog aktivasi Excel muncul di latar belakang, atau file output sebelumnya sedang terbuka di Excel. | Tutup seluruh file output Excel yang sedang terbuka, pastikan aplikasi Excel dapat dibuka normal tanpa *popup*, lalu jalankan ulang script. |
| Kolom `Tanggal Faktur` menjadi kosong pada **Mode Diskon** | `1_AccDiscFaktur_cleaner.py` mem-parsing `TGL` secara ketat menggunakan format `DD/MM/YY` (`%d/%m/%y`). Jika file `Faktur.xls` memakai tahun 4 digit (`DD/MM/YYYY`), tanggal gagal di-parse. | Pastikan format ekspor tanggal pada `Faktur.xls` adalah `DD/MM/YY`, atau samakan fungsi normalisasi tanggal di `1_AccDiscFaktur_cleaner.py` dengan `1_AccCtxFaktur_cleaner.py`. |
| Kolom `Kode Barang Jasa` kosong pada beberapa baris output | Nama barang di file sumber belum terdaftar pada 10 file `Helper_*.txt` (pencocokan bersifat *exact match* & *case-sensitive*). | Tambahkan nama barang persis seperti di `AccEFaktur.xls` ke file `Helper_<kode>.txt` yang sesuai (dan perbarui di repositori Helper GitHub bila menggunakan auto-update). |
| Muncul status `DITOLAK CORETAX` di file `AUDIT_DATA_CTX-*.xlsx` | Terdapat item dengan `DPP < 0`, `PPN < 0` (diskon melebihi harga total), `Harga Satuan < 0`, atau `Qty <= 0`. | Buka sheet `Faktur_Terdampak` dan `Temuan_Detail` pada file `AUDIT_DATA_CTX-*.xlsx`, koreksi transaksi terkait di Accurate/Excel sumber, lalu jalankan ulang pipeline sebelum impor ke Coretax. |
| File Helper tidak ter-update otomatis dari GitHub | Tidak ada koneksi internet, firewall kantor memblokir `raw.githubusercontent.com`, atau repositori Helper berstatus *private*. | Cek koneksi internet atau isi parameter `git_token` di bagian `[GIT]` pada `Dapur/config.conf` dengan Personal Access Token GitHub yang valid. |

---

## Kontribusi & Lisensi

### Alur Kontribusi
1. *Fork* repositori ini ke akun GitHub Anda.
2. Buat *branch* fitur atau perbaikan baru (`git checkout -b fitur/nama-fitur`).
3. Pastikan setiap perubahan pada skrip `Dapur/` mempertahankan kompatibilitas *fallback engine* (`calamine` → `xlrd`/`openpyxl`) dan penanda baris `END` pada `Template_v.1.6.1.xlsx`.
4. *Commit* perubahan Anda dengan pesan yang jelas (`git commit -m "feat: deskripsi perubahan"`).
5. *Push* ke *branch* Anda dan ajukan *Pull Request* ke repositori utama.

### Lisensi & Maintainer
Proyek ini dikembangkan dan dikelola oleh **[ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)** (`me@ihfandica.com`) untuk mendukung otomasi administrasi perpajakan DJP Coretax.
