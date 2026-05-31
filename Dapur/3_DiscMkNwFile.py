import pandas as pd

def buat_file_midprocess():
    nama_file = "MkNwFile_temp.xlsx"

    kolom_faktur = [
        "Baris",
        "Tanggal Faktur",
        "Jenis Faktur",
        "Kode Transaksi", 
        "Keterangan Tambahan",
        "Dokumen Pendukung",
        "Period Dok Pendukung", 
        "Referensi",
        "Cap Fasilitas",
        "ID TKU Penjual",
        "NPWP/NIK Pembeli", 
        "Jenis ID Pembeli",
        "Negara Pembeli",
        "Nomor Dokumen Pembeli", 
        "Nama Pembeli",
        "Alamat Pembeli",
        "Email Pembeli",
        "ID TKU Pembeli"
    ]

    kolom_detail = [
        "Baris",
        "Barang/Jasa",
        "Kode Barang Jasa",
        "Nama Barang/Jasa", 
        "Nama Satuan Ukur",
        "Harga Satuan",
        "Jumlah Barang Jasa", 
        "Total Diskon",
        "DPP",
        "DPP Nilai Lain",
        "Tarif PPN", 
        "PPN",
        "Tarif PPnBM",
        "PPnBM"
    ]

    kolom_kalkulasi = [
        "DETAIL", 
        "FAKTUR", 
        "URUTAN", 
        "HASIL"
    ]

    df_faktur = pd.DataFrame(columns=kolom_faktur)
    df_detail = pd.DataFrame(columns=kolom_detail)
    df_kalkulasi = pd.DataFrame(columns=kolom_kalkulasi)

    print(f"--> Membuat file {nama_file}...")

    try:
        with pd.ExcelWriter(nama_file, engine='xlsxwriter') as writer:
            
            sheet_name1 = 'Faktur'
            df_faktur.to_excel(writer, sheet_name=sheet_name1, index=False)
            worksheet1 = writer.sheets[sheet_name1]
            
            for i, col in enumerate(df_faktur.columns):
                max_len = len(str(col)) + 2
                worksheet1.set_column(i, i, max_len)

            sheet_name2 = 'DetailFaktur'
            df_detail.to_excel(writer, sheet_name=sheet_name2, index=False)
            worksheet2 = writer.sheets[sheet_name2]

            for i, col in enumerate(df_detail.columns):
                max_len = len(str(col)) + 2
                worksheet2.set_column(i, i, max_len)

            sheet_name3 = 'Kalkulasi'
            df_kalkulasi.to_excel(writer, sheet_name=sheet_name3, index=False)
            worksheet3 = writer.sheets[sheet_name3]

            for i, col in enumerate(df_kalkulasi.columns):
                max_len = len(str(col)) + 2
                worksheet3.set_column(i, i, max_len)

        print(f"--> Berhasil! File '{nama_file}' telah dibuat dengan 3 sheet.")

    except Exception as e:
        print(f"--> Terjadi kesalahan: {e}")

if __name__ == "__main__":
    buat_file_midprocess()
