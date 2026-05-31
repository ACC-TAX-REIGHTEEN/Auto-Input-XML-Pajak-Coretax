import os
import sys
import shutil
import subprocess
import datetime
import xlwings as xw

FOLDER_DAPUR = "Dapur"
INPUT_FILES = ["Faktur.xls", "EFaktur.xls"]

REQUIRED_DAPUR_FILES = [
	"0_Ftch_github@AccTaxREighteen.py",
    "1_AccDiscFaktur_cleaner.py",
    "2_AccDiscEFaktur_cleaner.py",
    "3_DiscMkNwFile.py",
    "4_DiscCpInvRefCalcXlook.py",
    "5_DiscCpData2Fktr.py",
    "6_DiscCpData2DtlFktr.py",
    "7_HelperDate.py",
    "8_HelperKTPOth.py",
    "__init__.py",
    "Helper_160100.txt",
    "Helper_200104.txt",
    "Helper_340300.txt",
    "Helper_400800.txt",
    "Helper_401100.txt",
    "Helper_401300.txt",
    "Helper_830900.txt",
    "Helper_842100.txt",
    "Helper_A.txt",
    "Helper_B.txt",
    "Helper_Del.txt",
    "KTP.txt",
    "KTP-OTH.txt",
    "TEMPLATE_1.3.25.xlsx"
]

EXECUTION_ORDER = [
	"0_Ftch_github@AccTaxREighteen.py",
    "1_AccDiscFaktur_cleaner.py",
    "2_AccDiscEFaktur_cleaner.py",
    "3_DiscMkNwFile.py",
    "4_DiscCpInvRefCalcXlook.py",
    "5_DiscCpData2Fktr.py",
    "6_DiscCpData2DtlFktr.py",
    "7_HelperDate.py",
    "8_HelperKTPOth.py",
]

TEMP_FILES_TO_REMOVE = [
    "MkNwFile_temp.xlsx",
    "AccCtxFaktur_temp.xlsx",
    "AccEFaktur_temp.xlsx",
    "Faktur.xls",
    "EFaktur.xls"
]

def log(message):
    print(f"--> {message}")

def error_exit(message):
    print(f"\n[ERROR] {message}")
    print("Proses Dibatalkan.")
    input("\nTekan Enter untuk keluar...")
    sys.exit(1)

def get_default_filename():
    now = datetime.datetime.now()
    months = {
        1: "JAN", 2: "FEB", 3: "MAR", 4: "APR", 5: "MEI", 6: "JUN",
        7: "JUL", 8: "AGU", 9: "SEP", 10: "OKT", 11: "NOV", 12: "DES"
    }
    year_short = now.year % 100
    return f"XML-DISC-{now.day:02d}-{months[now.month]}-{year_short}"

def generate_filename_from_data(date_val):
    dt = None
    
    try:
        if isinstance(date_val, datetime.datetime) or isinstance(date_val, datetime.date):
            dt = date_val
        elif isinstance(date_val, str):
            parts = date_val.split('/')
            if len(parts) == 3:
                d = int(parts[0])
                m = int(parts[1])
                y = int(parts[2])
                dt = datetime.date(y, m, d)
    except Exception:
        pass

    if dt:
        months = {
            1: "JAN", 2: "FEB", 3: "MAR", 4: "APR", 5: "MEI", 6: "JUN",
            7: "JUL", 8: "AGU", 9: "SEP", 10: "OKT", 11: "NOV", 12: "DES"
        }
        
        year_short = dt.year % 100
        return f"XML-DISC-{dt.day:02d}-{months[dt.month]}-{year_short}"
    else:
        return None

def clean_id_tku_penjual(raw_data):
    cleaned_data = []
    for row in raw_data:
        row_list = list(row)
        id_tku = row_list[9]
        if id_tku is not None:
            str_val = str(id_tku)
            if str_val.endswith('.0'):
                str_val = str_val[:-2]
            str_val = str_val.replace('.', '')
            row_list[9] = "'" + str_val
        cleaned_data.append(row_list)
    return cleaned_data

def force_harga_satuan_to_number(raw_data):
    cleaned_data = []
    for row in raw_data:
        row_list = list(row)
        val = row_list[5] 
        try:
            if val is not None:
                if isinstance(val, str):
                    val = val.replace(',', '.')
                row_list[5] = float(val)
        except ValueError:
            pass
        cleaned_data.append(row_list)
    return cleaned_data

def main():
    final_filename_base = get_default_filename()

    try:
        base_dir = os.getcwd()
        dapur_path = os.path.join(base_dir, FOLDER_DAPUR)

        log("Memeriksa kelengkapan file utama dan folder...")
        if not os.path.exists(dapur_path):
            error_exit(f"Folder '{FOLDER_DAPUR}' tidak ditemukan.")

        for f in INPUT_FILES:
            if not os.path.exists(os.path.join(base_dir, f)):
                error_exit(f"File '{f}' tidak ditemukan di folder utama.")

        log("Memeriksa kelengkapan file di dalam folder Dapur...")
        for f in REQUIRED_DAPUR_FILES:
            if not os.path.exists(os.path.join(dapur_path, f)):
                error_exit(f"File '{f}' hilang dari folder {FOLDER_DAPUR}.")

        log("Menyalin file input ke folder Dapur...")
        try:
            for f in INPUT_FILES:
                src = os.path.join(base_dir, f)
                dst = os.path.join(dapur_path, f)
                shutil.copy2(src, dst)
        except Exception as e:
            error_exit(f"Gagal menyalin file: {e}")

        log("Mulai menjalankan script pengolahan data...")
        python_exe = sys.executable 

        for script in EXECUTION_ORDER:
            script_path = os.path.join(dapur_path, script)
            log(f"Menjalankan: {script} ...")
            try:
                subprocess.run([python_exe, script], cwd=dapur_path, check=True)
            except subprocess.CalledProcessError:
                error_exit(f"Gagal saat menjalankan {script}. Proses berhenti.")

        source_file = os.path.join(dapur_path, "MkNwFile_temp.xlsx")
        if not os.path.exists(source_file):
            error_exit("File 'MkNwFile_temp.xlsx' tidak terbentuk setelah proses.")

        template_ori = os.path.join(dapur_path, "TEMPLATE_1.3.25.xlsx")
        template_temp = os.path.join(dapur_path, "TEMPLATE_temp.xlsx")
        
        try:
            shutil.copy2(template_ori, template_temp)
        except Exception as e:
            error_exit(f"Gagal menyalin template: {e}")

        log("Mulai proses transfer data ke Excel (xlwings)...")
        app = xw.App(visible=False)
        app.screen_updating = False
        app.display_alerts = False
        wb_source = None
        wb_target = None

        try:
            wb_source = app.books.open(source_file)
            wb_target = app.books.open(template_temp)

            log("Mengolah Sheet: Faktur")
            sh_src_faktur = wb_source.sheets['Faktur']
            sh_tgt_faktur = wb_target.sheets['Faktur']

            last_row_src_faktur = sh_src_faktur.range('A' + str(sh_src_faktur.cells.last_cell.row)).end('up').row
            
            if last_row_src_faktur >= 2:
                data_count = last_row_src_faktur - 1 
                raw_data_faktur = sh_src_faktur.range(f'A2:R{last_row_src_faktur}').options(ndim=2).value

                try:
                    first_date = raw_data_faktur[0][1]
                    generated_name = generate_filename_from_data(first_date)
                    if generated_name:
                        final_filename_base = generated_name
                        log(f"Nama file akan diubah menjadi: {final_filename_base}.xlsx (berdasarkan isi data)")
                except Exception as e:
                    log(f"Gagal membaca tanggal dari data, menggunakan tanggal hari ini. Error: {e}")

                final_data_faktur = clean_id_tku_penjual(raw_data_faktur)

                rows_to_insert = data_count - 1
                anchor_row = 5 
                
                if rows_to_insert > 0:
                    sh_tgt_faktur.range(f'{anchor_row}:{anchor_row + rows_to_insert - 1}').insert('down')
                
                sh_tgt_faktur.range('A4').value = final_data_faktur

            log("Mengolah Sheet: DetailFaktur")
            sh_src_detail = wb_source.sheets['DetailFaktur']
            sh_tgt_detail = wb_target.sheets['DetailFaktur']

            last_row_src_detail = sh_src_detail.range('A' + str(sh_src_detail.cells.last_cell.row)).end('up').row

            if last_row_src_detail >= 2:
                data_count_detail = last_row_src_detail - 1
                raw_data_detail = sh_src_detail.range(f'A2:N{last_row_src_detail}').options(ndim=2).value

                final_data_detail = force_harga_satuan_to_number(raw_data_detail)

                rows_to_insert_dtl = data_count_detail - 1
                anchor_row_dtl = 3

                if rows_to_insert_dtl > 0:
                    sh_tgt_detail.range(f'{anchor_row_dtl}:{anchor_row_dtl + rows_to_insert_dtl - 1}').insert('down')

                sh_tgt_detail.range('A2').value = final_data_detail

                end_row_dtl_target = 2 + data_count_detail - 1
                sh_tgt_detail.range(f'F2:F{end_row_dtl_target}').number_format = '###0,00'

            wb_target.save()
            log("Data berhasil dipindahkan dan diformat.")

        except Exception as e:
            error_exit(f"Error saat manipulasi Excel: {e}")
        finally:
            if wb_source: wb_source.close()
            if wb_target: wb_target.close()
            app.quit()

        log("Melakukan pembersihan file...")
        
        final_name_xlsx = f"{final_filename_base}.xlsx"
        final_path = os.path.join(base_dir, final_name_xlsx)

        if os.path.exists(template_temp):
            if os.path.exists(final_path):
                os.remove(final_path)
            shutil.move(template_temp, final_path)
            log(f"File hasil akhir disimpan sebagai: {final_name_xlsx}")
        else:
            error_exit("Gagal menemukan file hasil akhir TEMPLATE_temp.xlsx")

        for temp_file in TEMP_FILES_TO_REMOVE:
            t_path = os.path.join(dapur_path, temp_file)
            if os.path.exists(t_path):
                try:
                    os.remove(t_path)
                except:
                    pass

        for f in INPUT_FILES:
            f_path = os.path.join(dapur_path, f)
            if os.path.exists(f_path):
                os.remove(f_path)

        log("PROSES SELESAI. Folder Dapur bersih.")

    except Exception as general_error:
        error_exit(f"Terjadi kesalahan fatal yang tidak terduga: {general_error}")

    input("\nTekan Enter untuk keluar. Pesan moral: Jangan pernah percaya kepada siapapun, bahkan jika itu dirimu sendiri. Selalu lakukan cek ulang, karena mesin juga bisa salah.")

if __name__ == "__main__":
    main()
