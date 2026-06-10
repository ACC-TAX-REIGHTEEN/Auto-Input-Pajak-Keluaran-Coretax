import pandas as pd
import glob
import io
import os
import datetime

output_excel = 'HMA_csv2ex.xlsx'

def print_debug(message):
    waktu = datetime.datetime.now().strftime("%H:%M:%S")
    print(f"--> [{waktu}] {message}")

def clean_and_parse_csv_smart(file_path):
    cleaned_rows = []
    filename = os.path.basename(file_path)
    print_debug(f"--> Memproses file: {filename}")
    
    with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
        lines = f.readlines()
        if len(lines) == 0: return pd.DataFrame()
        header_raw = lines[0].strip()
        header_cols = header_raw.split(',')
        target_col_count = len(header_cols)
        cleaned_rows.append("|".join(header_cols))
        
        for line in lines[1:]:
            line = line.strip()
            if not line: continue 
            parts = line.split(',')
            
            if len(parts) == target_col_count:
                cleaned_rows.append("|".join(parts))
            
            elif len(parts) > target_col_count:
                cols_right_count = target_col_count - 2
                col_index = parts[0]
                cols_right = parts[-cols_right_count:]
                name_parts = parts[1 : -cols_right_count]
                col_name_fixed = " ".join(name_parts).replace('"', '') 
                new_row = [col_index, col_name_fixed] + cols_right
                cleaned_rows.append("|".join(new_row))
            else:
                cleaned_rows.append("|".join(parts))

    full_str = "\n".join(cleaned_rows)
    return pd.read_csv(io.StringIO(full_str), sep='|', dtype=str)

try:
    print_debug("=== MEMULAI PROSES ===")
    pattern = "data_export*.csv"
    all_files = glob.glob(pattern)
    
    if not all_files: 
        raise ValueError(f"--> Tidak ditemukan file CSV dengan pola '{pattern}'")

    print_debug(f"--> Ditemukan {len(all_files)} file untuk diproses.")

    df_list = []
    processed_files = []

    for filename in all_files:
        try:
            df_parsed = clean_and_parse_csv_smart(filename)
            df_list.append(df_parsed)
            processed_files.append(filename)
        except Exception as e:
            print_debug(f"--> Gagal memproses {filename}: {e}")

    if not df_list:
        raise ValueError("--> Tidak ada data yang berhasil diproses dari file CSV.")

    print_debug("--> Menggabungkan data...")
    df_combined = pd.concat(df_list, ignore_index=True)
    
    if 'Unnamed: 0' in df_combined.columns:
        df_combined = df_combined.drop(columns=['Unnamed: 0'])
    elif len(df_combined.columns) > 0:
        if df_combined.columns[0] == '' or 'Unnamed' in df_combined.columns[0]:
             df_combined = df_combined.drop(columns=[df_combined.columns[0]])

    col_names = df_combined.columns
    start_numeric_index = 4
    if 'TaxBase' in col_names:
        start_numeric_index = df_combined.columns.get_loc('TaxBase')
        print_debug(f"--> Deteksi Kolom Angka (TaxBase) di kolom index: {start_numeric_index}")
    else:
        print_debug(f"--> Menggunakan default kolom angka mulai index: {start_numeric_index}")

    for i in range(start_numeric_index, len(col_names)):
        col = col_names[i]
        df_combined[col] = pd.to_numeric(df_combined[col], errors='coerce')

    print_debug(f"--> Menyimpan ke {output_excel}...")
    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
        df_combined.to_excel(writer, index=False, sheet_name='Data')
        workbook = writer.book
        worksheet = writer.sheets['Data']
        
        fmt_text = workbook.add_format({'num_format': '@'})
        fmt_num = workbook.add_format({'num_format': '#,##0'})
        
        for i, col in enumerate(df_combined.columns):
            try: max_len = max(df_combined[col].astype(str).map(len).max(), len(str(col))) + 2
            except: max_len = 15
            
            if i < start_numeric_index: 
                worksheet.set_column(i, i, max_len, fmt_text)
            else: 
                worksheet.set_column(i, i, max_len, fmt_num)

    print_debug(f"--> SUKSES! File tersimpan: {output_excel}")

    print_debug("--- MEMBERSIHKAN FILE CSV ASLI ---")
    deleted_count = 0
    for csv_file in processed_files:
        try:
            if os.path.exists(csv_file):
                os.remove(csv_file)
                print_debug(f"--> Hapus {csv_file}")
                deleted_count += 1
        except Exception as e:
            print_debug(f"--> Gagal hapus {csv_file}: {e}")

    print_debug(f"--> Selesai. {deleted_count} file CSV telah dihapus.")

except Exception as e:
    print(f"\n--> FATAL ERROR: {e}")
    print("--> File CSV TIDAK dihapus karena terjadi error.")