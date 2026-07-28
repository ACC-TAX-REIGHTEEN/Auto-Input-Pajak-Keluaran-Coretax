import pandas as pd

input_file = 'Hasil_merger_all_ex2ex_keluaran.xlsx'
output_file = 'Hasil_gabungan_keluaran_tgl.xlsx'

try:
    df = pd.read_excel(input_file, dtype=str)
    
    col_date_name = df.columns[3]
    df[col_date_name] = pd.to_datetime(df[col_date_name], errors='coerce')
    df[col_date_name] = df[col_date_name].dt.strftime('%d/%b/%Y')
    
    numeric_indices = [4, 5, 6]
    
    for idx in numeric_indices:
        if idx < len(df.columns):
            col_name = df.columns[idx]
            df[col_name] = pd.to_numeric(df[col_name], errors='coerce')
            
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        fmt_text = workbook.add_format({'num_format': '@'})
        fmt_number = workbook.add_format({'num_format': '#,##0'})

        for i, col in enumerate(df.columns):
            if i in numeric_indices:
                max_len = max(
                    df[col].astype(str).map(len).max(),
                    len(str(col))
                ) + 2
                worksheet.set_column(i, i, max_len, fmt_number)
            else:
                max_len = max(
                    df[col].astype(str).map(len).max(),
                    len(str(col))
                ) + 2
                worksheet.set_column(i, i, max_len, fmt_text)

    print(f"--> Selesai! File tersimpan: {output_file}")

except Exception as e:
    print(f"--> Terjadi kesalahan: {e}")