import pandas as pd
import os
from openpyxl.utils import get_column_letter

def olah_data_faktur_autofit():
    file_sumber = 'AccCtxFaktur_temp.xlsx'
    file_tujuan = 'MkNwFile_temp.xlsx'

    print("--> Sedang membaca data sumber...")
    
    try:
        df_sumber = pd.read_excel(file_sumber, sheet_name=0, dtype={'NPWP': str, 'IDTKU': str})
    except FileNotFoundError:
        print(f"--> Error: File {file_sumber} tidak ditemukan.")
        return

    df_sumber['TGL'] = pd.to_datetime(df_sumber['TGL'], dayfirst=True)
    df_target = pd.DataFrame()

    print("--> Sedang memproses data...")

    df_target['Baris'] = range(1, len(df_sumber) + 1)
    df_target['Tanggal Faktur'] = df_sumber['TGL'].dt.strftime('%d/%m/%Y')
    df_target['Jenis Faktur'] = 'Normal'
    df_target['Kode Transaksi'] = '04'
    df_target['Keterangan Tambahan'] = ''
    df_target['Dokumen Pendukung'] = ''
    df_target['Period Dok Pendukung'] = df_sumber['TGL'].dt.strftime('%m%Y')
    df_target['Referensi'] = df_sumber['REFERENSI']
    df_target['Cap Fasilitas'] = ''
    df_target['ID TKU Penjual'] = '#GANTI INI DENGAN KODE PENJUAL'
    df_target['NPWP/NIK Pembeli'] = df_sumber['NPWP'].fillna('')

    def cek_jenis_id(npwp):
        npwp_str = str(npwp).strip()
        if npwp_str == '0000000000000000':
            return 'Other ID'
        elif len(npwp_str) >= 10:
            return 'TIN'
        else:
            return 'Other ID'

    df_target['Jenis ID Pembeli'] = df_target['NPWP/NIK Pembeli'].apply(cek_jenis_id)
    df_target['Negara Pembeli'] = 'IDN'

    def isi_nomor_dokumen(jenis_id):
        if jenis_id == 'Other ID':
            return '0000000000000000'
        elif jenis_id == 'TIN':
            return '-'
        else:
            return ''

    df_target['Nomor Dokumen Pembeli'] = df_target['Jenis ID Pembeli'].apply(isi_nomor_dokumen)
    df_target['Nama Pembeli'] = df_sumber['NAMA']
    df_target['Alamat Pembeli'] = df_sumber['ALAMAT PAJAK 1']
    df_target['Email Pembeli'] = ''
    df_target['ID TKU Pembeli'] = df_sumber['IDTKU'].fillna('')

    print(f"--> Sedang menyimpan dan merapikan kolom (Auto-fit) di {file_tujuan}...")

    mode = 'a' if os.path.exists(file_tujuan) else 'w'
    if_sheet_exists = 'replace' if mode == 'a' else None

    with pd.ExcelWriter(file_tujuan, engine='openpyxl', mode=mode, if_sheet_exists=if_sheet_exists) as writer:
        sheet_name = 'Faktur'
        df_target.to_excel(writer, sheet_name=sheet_name, index=False)
        
        worksheet = writer.sheets[sheet_name]

        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column) 
            
            for cell in column:
                try:
                    if cell.value:
                        cell_length = len(str(cell.value))
                        if cell_length > max_length:
                            max_length = cell_length
                except:
                    pass
            
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column_letter].width = adjusted_width

    print("--> Selesai! Data berhasil disalin dan kolom sudah di-autofit.")

if __name__ == "__main__":
    olah_data_faktur_autofit()
