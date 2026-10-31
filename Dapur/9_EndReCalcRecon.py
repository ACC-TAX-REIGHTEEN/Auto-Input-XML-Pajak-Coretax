import os
import configparser
import datetime
import pandas as pd
from openpyxl.utils import get_column_letter

def clean_config_val(raw_val):
    if raw_val is None:
        return ""
    cleaned = str(raw_val).split('#')[0].strip()
    return cleaned

def load_audit_config(config_path='config.conf'):
    config = configparser.ConfigParser()
    if not os.path.exists(config_path):
        print(f"--> [PERINGATAN] File '{config_path}' tidak ditemukan! Menggunakan pengaturan default.")
        return {
            'aud': 'Ya',
            'aud_sheet': 'DetailFaktur',
            'aud_col': 'Harga Satuan',
            'aud_col_val': '0',
            'aud_col_ref': 'Tidak'
        }
    
    config.read(config_path, encoding='utf-8')
    
    if 'AUDIT' not in config:
        print(f"--> [INFO] Bagian [AUDIT] tidak ditemukan di '{config_path}'. Proses audit dilewati.")
        return {'aud': 'Tidak'}

    return {
        'aud': clean_config_val(config.get('AUDIT', 'aud', fallback='Tidak')),
        'aud_sheet': clean_config_val(config.get('AUDIT', 'aud_sheet', fallback='DetailFaktur')),
        'aud_col': clean_config_val(config.get('AUDIT', 'aud_col', fallback='Harga Satuan')),
        'aud_col_val': clean_config_val(config.get('AUDIT', 'aud_col_val', fallback='0')),
        'aud_col_ref': clean_config_val(config.get('AUDIT', 'aud_col_ref', fallback='Tidak'))
    }

def format_tanggal_indonesia(val):
    if pd.isna(val) or val is None or str(val).strip() == '':
        return ''
    
    dt = None
    if isinstance(val, (datetime.datetime, datetime.date)):
        dt = val
    else:
        val_str = str(val).strip()
        try:
            parts = val_str.split('/')
            if len(parts) == 3:
                dt = datetime.date(int(parts[2]), int(parts[1]), int(parts[0]))
        except Exception:
            pass
        
        if dt is None:
            try:
                dt = pd.to_datetime(val_str, dayfirst=True)
            except Exception:
                return val_str

    if dt:
        bulan_indo = {
            1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "Mei", 6: "Jun",
            7: "Jul", 8: "Agu", 9: "Sep", 10: "Okt", 11: "Nov", 12: "Des"
        }
        return f"{dt.day} {bulan_indo.get(dt.month, '')} {dt.year}"
    
    return str(val)

def klasifikasi_temuan(row, aud_col):
    alasan = []
    status_coretax = "DITERIMA (BERISIKO SOP)"
    
    harga_target = float(row.get(aud_col, 0))
    dpp = float(row.get('DPP', 0))
    ppn = float(row.get('PPN', 0))
    diskon = float(row.get('Total Diskon', 0))
    qty = float(row.get('Jumlah Barang Jasa', 0))

    if harga_target == 0:
        alasan.append(f"{aud_col} bernilai 0 (Nota barang bonus/sample harus dipisah)")
    elif harga_target < 0:
        alasan.append(f"{aud_col} bernilai negatif ({harga_target})")
        status_coretax = "DITOLAK CORETAX"

    if dpp < 0 or ppn < 0:
        alasan.append(f"Nilai DPP ({dpp:,.2f}) atau PPN ({ppn:,.2f}) MINUS akibat Diskon ({diskon:,.2f}) > Total")
        status_coretax = "DITOLAK CORETAX"

    if qty <= 0:
        alasan.append(f"Jumlah Barang Jasa tidak valid ({qty})")
        status_coretax = "DITOLAK CORETAX"

    if not alasan:
        alasan.append("Memenuhi kriteria filter pengecekan data non-nol")
        status_coretax = "NORMAL / SESUAI FILTER"

    return pd.Series([status_coretax, " | ".join(alasan)])

def autofit_excel_columns(writer, sheet_name):
    worksheet = writer.sheets[sheet_name]
    for column in worksheet.columns:
        max_length = 0
        col_idx = column[0].column
        column_letter = get_column_letter(col_idx) if isinstance(col_idx, int) else col_idx
        
        for cell in column:
            try:
                if cell.value is not None:
                    cell_len = len(str(cell.value))
                    if cell_len > max_length:
                        max_length = cell_len
            except Exception:
                pass
        worksheet.column_dimensions[column_letter].width = min(max_length + 3, 60)

def jalankan_audit_coretax():
    file_input = "MkNwFile_temp.xlsx"
    file_output = "AUDIT_DATA_CTX.xlsx"

    print("--> AUDIT KELAYAKAN DATA CORETAX & KEPATUHAN SOP NOTA")

    cfg = load_audit_config('config.conf')
    
    if cfg['aud'].lower() != 'ya':
        print(f"--> Status Audit di config.conf: '{cfg['aud']}'. Proses audit dilewati.")
        return

    sheet_target = cfg['aud_sheet']
    col_target = cfg['aud_col']
    val_mode = cfg['aud_col_val']
    cek_ref_sop = cfg['aud_col_ref'].lower() == 'ya'

    print(f"--> Konfigurasi Aktif : Sheet='{sheet_target}' | Kolom='{col_target}' | Mode Nilai='{val_mode}' | Filter Ref SOP='{cfg['aud_col_ref']}'")

    if not os.path.exists(file_input):
        print(f"--> [ERROR] File '{file_input}' tidak ditemukan! Pastikan proses sebelumnya sudah dijalankan.")
        return

    try:
        df_detail = pd.read_excel(file_input, sheet_name=sheet_target)
    except Exception as e:
        print(f"--> [ERROR] Gagal membaca sheet '{sheet_target}' di '{file_input}': {e}")
        return

    df_detail.columns = [str(c).strip() for c in df_detail.columns]
    df_detail = df_detail[df_detail['Baris'].astype(str).str.strip().str.upper() != 'END'].copy()

    if col_target not in df_detail.columns:
        print(f"--> [ERROR] Kolom '{col_target}' tidak ditemukan di sheet '{sheet_target}'!")
        print(f"--> Kolom yang tersedia: {list(df_detail.columns)}")
        return

    kolom_numerik = [col_target, 'Jumlah Barang Jasa', 'Total Diskon', 'DPP', 'DPP Nilai Lain', 'PPN']
    for kn in kolom_numerik:
        if kn in df_detail.columns:
            if df_detail[kn].dtype == object:
                df_detail[kn] = df_detail[kn].astype(str).str.strip().str.replace(',', '.', regex=False)
            df_detail[kn] = pd.to_numeric(df_detail[kn], errors='coerce').fillna(0.0)

    try:
        df_faktur = pd.read_excel(file_input, sheet_name='Faktur')
        df_faktur.columns = [str(c).strip() for c in df_faktur.columns]
        
        if 'Tanggal Faktur' in df_faktur.columns:
            df_faktur['Tanggal Faktur'] = df_faktur['Tanggal Faktur'].apply(format_tanggal_indonesia)

        df_detail['Baris_Key'] = pd.to_numeric(df_detail['Baris'], errors='coerce').fillna(-1).astype(int)
        df_faktur['Baris_Key'] = pd.to_numeric(df_faktur['Baris'], errors='coerce').fillna(-2).astype(int)

        kolom_info_faktur = ['Baris_Key', 'Referensi', 'Tanggal Faktur', 'Nama Pembeli', 'NPWP/NIK Pembeli']
        kolom_tersedia = [k for k in kolom_info_faktur if k in df_faktur.columns]

        df_detail = pd.merge(
            df_detail,
            df_faktur[kolom_tersedia].drop_duplicates(subset=['Baris_Key']),
            on='Baris_Key',
            how='left'
        )
        df_detail.drop(columns=['Baris_Key'], inplace=True)
    except Exception as e:
        print(f"--> [INFO] Sheet 'Faktur' tidak dapat digabungkan: {e}")

    if val_mode == '0':
        kondisi_filter = (df_detail[col_target] == 0) | (df_detail.get('DPP', 0) < 0) | (df_detail.get('PPN', 0) < 0)
        deskripsi_mode = f"Data Anomali ({col_target} == 0 atau Nilai DPP/PPN Minus)"
    elif val_mode == '1':
        kondisi_filter = (df_detail[col_target] != 0)
        deskripsi_mode = f"Data Bernilai ({col_target} > 0 atau < 0)"
    else:
        try:
            angka_kustom = float(val_mode)
            kondisi_filter = (df_detail[col_target] == angka_kustom)
            deskripsi_mode = f"Data dengan {col_target} == {angka_kustom}"
        except ValueError:
            print(f"--> [ERROR] Nilai aud_col_val='{val_mode}' tidak valid. Gunakan 0 atau 1.")
            return

    if cek_ref_sop and 'Referensi' in df_detail.columns and 'Tanggal Faktur' in df_detail.columns:
        print("--> Mengecek Referensi & Tanggal Faktur untuk deteksi nota yang sudah patuh SOP pisah nota...")
        
        group_ref = df_detail.groupby(['Referensi', 'Tanggal Faktur'])[col_target]
        nota_semua_nol = group_ref.transform(lambda s: (s == 0).all())
        
        jumlah_nota_sop = df_detail[nota_semua_nol]['Referensi'].nunique()
        baris_keluar = nota_semua_nol.sum()
        
        if baris_keluar > 0:
            print(f"--> [SOP VALID] Ditemukan {jumlah_nota_sop} nota ({baris_keluar} baris) yang Harga Satuannya 0 semua.")
            print("--> Data nota tersebut dikeluarkan dari audit karena sudah dipisah nota sesuai SOP.")
        
        df_temuan = df_detail[kondisi_filter & (~nota_semua_nol)].copy()
    else:
        df_temuan = df_detail[kondisi_filter].copy()

    if not df_temuan.empty:
        df_temuan[['STATUS_CORETAX', 'CATATAN_AUDIT']] = df_temuan.apply(
            lambda r: klasifikasi_temuan(r, col_target), axis=1
        )

        kolom_prioritas = ['Baris', 'Referensi', 'Tanggal Faktur', 'Nama Pembeli', 'STATUS_CORETAX', 'CATATAN_AUDIT']
        kolom_awal = [c for c in kolom_prioritas if c in df_temuan.columns]
        kolom_sisa = [c for c in df_temuan.columns if c not in kolom_awal]
        df_temuan = df_temuan[kolom_awal + kolom_sisa]

    total_baris_detail = len(df_detail)
    total_baris_temuan = len(df_temuan)
    faktur_terdampak = df_temuan['Baris'].nunique() if not df_temuan.empty else 0
    jml_ditolak = len(df_temuan[df_temuan['STATUS_CORETAX'] == 'DITOLAK CORETAX']) if not df_temuan.empty else 0
    jml_peringatan = len(df_temuan[df_temuan['STATUS_CORETAX'] == 'DITERIMA (BERISIKO SOP)']) if not df_temuan.empty else 0

    print("--> -----------------------------------------------------")
    print(f"--> Filter Audit               : {deskripsi_mode}")
    print(f"--> Total Baris Diperiksa      : {total_baris_detail} baris")
    print(f"--> Total Temuan Audit         : {total_baris_temuan} baris (dalam {faktur_terdampak} Nomor Faktur/Baris)")
    print(f"-->   - Pasti Ditolak Coretax  : {jml_ditolak} baris (DPP/PPN Minus)")
    print(f"-->   - Campur Nota Bermasalah : {jml_peringatan} baris (Harga 0 digabung dengan barang lain)")
    print("--> -----------------------------------------------------")

    if not df_temuan.empty:
        print("--> SEKILAS DATA TEMUAN (Maks. 10 Baris Pertama):")
        kolom_preview = [c for c in ['Baris', 'Referensi', 'Tanggal Faktur', 'Nama Barang/Jasa', col_target, 'Total Diskon', 'DPP', 'STATUS_CORETAX'] if c in df_temuan.columns]
        print(f"--> {df_temuan[kolom_preview].head(10).to_string(index=False)}")
        print("--> -----------------------------------------------------")
    else:
        print("--> [AMAN] Tidak ditemukan data anomali atau nota campuran yang belum dipisah.")

    df_ringkasan = pd.DataFrame({
        'METRIK AUDIT': [
            'File Sumber Pemeriksaan',
            'Sheet yang Diperiksa',
            'Kolom Acuan (aud_col)',
            'Kode Filter (aud_col_val)',
            'Pengecekan SOP Pisah Nota (aud_col_ref)',
            'Total Baris Detail Faktur',
            'Total Baris Temuan Audit',
            'Jumlah Faktur Terdampak (Belum Dipisah)',
            'Jumlah Baris Pasti Ditolak Coretax (DPP/PPN Minus)',
            'Jumlah Baris Belum Dipisah Nota (Harga Satuan = 0)'
        ],
        'NILAI / HASIL': [
            file_input,
            sheet_target,
            col_target,
            f"{val_mode} ({deskripsi_mode})",
            cfg['aud_col_ref'],
            total_baris_detail,
            total_baris_temuan,
            faktur_terdampak,
            jml_ditolak,
            jml_peringatan
        ]
    })

    if not df_temuan.empty:
        group_cols = [c for c in ['Baris', 'Referensi', 'Tanggal Faktur', 'Nama Pembeli'] if c in df_temuan.columns]
        df_rekap_faktur = df_temuan.groupby(group_cols, dropna=False).agg(
            Jumlah_Item_Bermasalah=('Nama Barang/Jasa', 'count'),
            Total_Diskon_Bermasalah=('Total Diskon', 'sum'),
            Total_DPP_Bermasalah=('DPP', 'sum'),
            Status_Terparah=('STATUS_CORETAX', lambda x: 'DITOLAK CORETAX' if 'DITOLAK CORETAX' in x.values else 'DITERIMA (BERISIKO SOP)')
        ).reset_index()
    else:
        df_rekap_faktur = pd.DataFrame(columns=['Baris', 'Referensi', 'Tanggal Faktur', 'Status_Terparah'])

    try:
        with pd.ExcelWriter(file_output, engine='openpyxl', mode='w') as writer:
            df_temuan.to_excel(writer, sheet_name='Temuan_Detail', index=False)
            autofit_excel_columns(writer, 'Temuan_Detail')

            df_rekap_faktur.to_excel(writer, sheet_name='Faktur_Terdampak', index=False)
            autofit_excel_columns(writer, 'Faktur_Terdampak')

            df_ringkasan.to_excel(writer, sheet_name='Ringkasan_Audit', index=False)
            autofit_excel_columns(writer, 'Ringkasan_Audit')

        print(f"--> [BERHASIL] Hasil audit lengkap telah disimpan ke: '{file_output}'")
    except Exception as e:
        print(f"--> [ERROR] Gagal menyimpan file '{file_output}': {e}")

if __name__ == "__main__":
    jalankan_audit_coretax()