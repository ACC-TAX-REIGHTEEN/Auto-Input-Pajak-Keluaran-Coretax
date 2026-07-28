import pandas as pd
import numpy as np
import re
import os
from datetime import datetime, timedelta
import difflib

print("--> Memulai proses analisa data (Update: 6 Tahap Matching)...")

print("--> Sedang membaca file Excel/CSV...")
try:
    try:
        acc_df = pd.read_excel('Hasil_CleanerACC.xlsx')
    except Exception:
        acc_df = pd.read_csv('Hasil_CleanerACC.xlsx - Sheet1.csv')

    try:
        hma_df = pd.read_excel('HMA_csv2ex.xlsx')
    except Exception:
        hma_df = pd.read_csv('HMA_csv2ex.xlsx - Data.csv')

    print(f"--> Data Accurate dimuat: {len(acc_df)} baris")
    print(f"--> Data Coretax dimuat: {len(hma_df)} baris")

except Exception as e:
    print(f"--> [ERROR] Gagal memuat file: {e}")
    exit()

print("--> Melakukan pembersihan data (Cleaning)...")

def clean_money(val):
    try:
        return float(val)
    except Exception:
        return 0.0

def clean_tin(val):
    s = str(val)
    s = re.sub(r'\D', '', s)
    if s:
        try:
            s = str(int(s))
        except Exception:
            pass
    if len(s) > 1:
        return s[:-1]
    return s

def clean_name(val):
    if pd.isna(val):
        return ""
    s = str(val).lower()
    prefixes = ['pt.', 'pt ', 'cv.', 'cv ', 'ud.', 'ud ', 'tb.', 'tb ', 'toko ', 'bengkel ']
    for p in prefixes:
        s = s.replace(p, '')
    s = re.sub(r'[,.\-]', ' ', s)
    s = re.sub(r'[^a-z0-9\s]', '', s)
    s = re.sub(r'\s+', ' ', s)
    return s.strip()

month_map = {
    'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'Mei': '05', 'Jun': '06', 
    'Jul': '07', 'Agu': '08', 'Sep': '09', 'Okt': '10', 'Nop': '11', 'Des': '12'
}

def parse_indo_date(date_str):
    try:
        parts = str(date_str).split()
        if len(parts) == 3:
            d, m, y = parts
            if m in month_map:
                return datetime(int(y), int(month_map[m]), int(d))
    except Exception:
        pass
    return pd.NaT

month_map_rev = {int(v): k for k, v in month_map.items()}

def format_indo_date(dt):
    if pd.isna(dt):
        return ""
    try:
        return f"{dt.day:02d} {month_map_rev[dt.month]} {dt.year}"
    except Exception:
        return ""

col_tgl = 'Tgl. Pajak' if 'Tgl. Pajak' in acc_df.columns else 'Tanggal'
acc_df['Date_Obj'] = acc_df[col_tgl].apply(parse_indo_date)
acc_df['Amount_Clean'] = acc_df['Jumlah Pajak'].apply(clean_money)
acc_df['TIN_Clean'] = acc_df['Nomor Pajak'].apply(clean_tin)
acc_df['Name_Clean'] = acc_df['Nama Pelanggan'].apply(clean_name)

hma_df['Date_Obj'] = pd.to_datetime(hma_df['DocumentDate']).dt.tz_localize(None)
hma_df['Amount_Clean'] = hma_df['VAT'].apply(clean_money)
hma_df['TIN_Clean'] = hma_df['TIN'].apply(clean_tin)
hma_df['Name_Clean'] = hma_df['Name'].apply(clean_name)

print("--> Standardisasi Nama, Tanggal, dan NPWP selesai.")

print("--> Mendeteksi Bulan Laporan...")
try:
    valid_dates = acc_df['Date_Obj'].dropna()
    if not valid_dates.empty:
        dominant_month_num = valid_dates.dt.month.mode()[0]
        map_bulan_indo = {
            1: 'JANUARI', 2: 'FEBRUARI', 3: 'MARET', 4: 'APRIL', 5: 'MEI', 6: 'JUNI',
            7: 'JULI', 8: 'AGUSTUS', 9: 'SEPTEMBER', 10: 'OKTOBER', 11: 'NOPEMBER', 12: 'DESEMBER'
        }
        detected_month_name = map_bulan_indo.get(int(dominant_month_num), '')
        print(f"--> Bulan terdeteksi dominan: {detected_month_name} (Bulan ke-{int(dominant_month_num)})")
        output_file = f'Laporan_Analisa_Pajak_{detected_month_name}.xlsx'
    else:
        print("--> Tidak ada tanggal valid ditemukan. Menggunakan nama default.")
        output_file = 'Laporan_Analisa_Pajak_General.xlsx'
except Exception as e:
    print(f"--> Gagal mendeteksi bulan ({e}). Menggunakan nama default.")
    output_file = 'Laporan_Analisa_Pajak_General.xlsx'

print(f"--> Nama file output diset: {output_file}")

print("--> Memulai proses pencocokan data (Matching 6 Tahap)...")

acc_df['Match_Status'] = 'Not in Coretax'
acc_df['Match_Index_HMA'] = -1
hma_df['Match_Status'] = 'Not in Accurate'
hma_df['Match_Index_ACC'] = -1

hma_by_tin = {}
for i, r in hma_df.iterrows():
    t = r['TIN_Clean']
    if t not in hma_by_tin:
        hma_by_tin[t] = []
    hma_by_tin[t].append(i)

hma_by_amt = {}
for i, r in hma_df.iterrows():
    amt = int(r['Amount_Clean'])
    if amt not in hma_by_amt:
        hma_by_amt[amt] = []
    hma_by_amt[amt].append(i)

used_hma = set()

match_count_1 = 0
for i, row in acc_df.iterrows():
    tin = row['TIN_Clean']
    amt = row['Amount_Clean']
    date_acc = row['Date_Obj']
    
    candidates = hma_by_tin.get(tin, [])
    candidates = [c for c in candidates if c not in used_hma]
    for c_idx in candidates:
        date_hma = hma_df.at[c_idx, 'Date_Obj']
        if (abs(hma_df.at[c_idx, 'Amount_Clean'] - amt) <= 5) and (date_acc == date_hma):
            acc_df.at[i, 'Match_Status'] = 'Matched'
            acc_df.at[i, 'Match_Index_HMA'] = c_idx
            hma_df.at[c_idx, 'Match_Status'] = 'Matched'
            hma_df.at[c_idx, 'Match_Index_ACC'] = i
            used_hma.add(c_idx)
            match_count_1 += 1
            break

print(f"--> Tahap 1 (Exact TIN+Amt+Tgl) selesai. Matches: {match_count_1}")

match_count_2 = 0
for i, row in acc_df.iterrows():
    if row['Match_Status'] == 'Matched':
        continue
    amt = int(row['Amount_Clean'])
    name = row['Name_Clean']
    date_acc = row['Date_Obj']
    
    candidates = []
    for delta in range(-5, 6):
        candidates.extend(hma_by_amt.get(amt + delta, []))
    candidates = [c for c in candidates if c not in used_hma]
    
    for c_idx in candidates:
        c_name = hma_df.at[c_idx, 'Name_Clean']
        date_hma = hma_df.at[c_idx, 'Date_Obj']
        name_match = (name == c_name) or (name in c_name) or (c_name in name) or (not name and not c_name)
        if name_match and (date_acc == date_hma):
            acc_df.at[i, 'Match_Status'] = 'Matched'
            acc_df.at[i, 'Match_Index_HMA'] = c_idx
            hma_df.at[c_idx, 'Match_Status'] = 'Matched'
            hma_df.at[c_idx, 'Match_Index_ACC'] = i
            used_hma.add(c_idx)
            match_count_2 += 1
            break

print(f"--> Tahap 2 (Exact Name+Amt+Tgl) selesai. Matches tambahan: {match_count_2}")

match_count_3 = 0
for i, row in acc_df[acc_df['Match_Status'] == 'Not in Coretax'].iterrows():
    if pd.isna(row['Date_Obj']):
        continue
    tin = row['TIN_Clean']
    amt = row['Amount_Clean']
    
    candidates = hma_by_tin.get(tin, [])
    candidates = [c for c in candidates if c not in used_hma]
    for c_idx in candidates:
        date_hma = hma_df.at[c_idx, 'Date_Obj']
        if pd.isna(date_hma):
            continue
        
        day_diff = abs((row['Date_Obj'] - date_hma).days)
        if (abs(hma_df.at[c_idx, 'Amount_Clean'] - amt) <= 5) and (day_diff <= 3):
            acc_df.at[i, 'Match_Status'] = 'Matched'
            acc_df.at[i, 'Match_Index_HMA'] = c_idx
            hma_df.at[c_idx, 'Match_Status'] = 'Matched'
            hma_df.at[c_idx, 'Match_Index_ACC'] = i
            used_hma.add(c_idx)
            match_count_3 += 1
            break

print(f"--> Tahap 3 (Date Window +/- 3 Hari) selesai. Matches tambahan: {match_count_3}")

match_count_4 = 0
for i, row in acc_df[acc_df['Match_Status'] == 'Not in Coretax'].iterrows():
    amt = int(row['Amount_Clean'])
    name_acc = row['Name_Clean']
    
    candidates = []
    for delta in range(-5, 6):
        candidates.extend(hma_by_amt.get(amt + delta, []))
    candidates = [c for c in candidates if c not in used_hma]
    
    for c_idx in candidates:
        name_hma = hma_df.at[c_idx, 'Name_Clean']
        score = difflib.SequenceMatcher(None, name_acc, name_hma).ratio()
        
        if score >= 0.8:
            date_hma = hma_df.at[c_idx, 'Date_Obj']
            day_diff = abs((row['Date_Obj'] - date_hma).days) if not pd.isna(date_hma) else 999
            
            if (row['Date_Obj'] == date_hma) or (day_diff <= 3):
                acc_df.at[i, 'Match_Status'] = 'Matched'
                acc_df.at[i, 'Match_Index_HMA'] = c_idx
                hma_df.at[c_idx, 'Match_Status'] = 'Matched'
                hma_df.at[c_idx, 'Match_Index_ACC'] = i
                used_hma.add(c_idx)
                match_count_4 += 1
                break

print(f"--> Tahap 4 (Fuzzy Name Score > 80%) selesai. Matches tambahan: {match_count_4}")

match_count_5 = 0
for i, row in acc_df[acc_df['Match_Status'] == 'Not in Coretax'].iterrows():
    amt = int(row['Amount_Clean'])
    date_acc = row['Date_Obj']
    if pd.isna(date_acc):
        continue

    candidates = []
    for delta in range(-5, 6):
        candidates.extend(hma_by_amt.get(amt + delta, []))
    candidates = [c for c in candidates if c not in used_hma]

    for c_idx in candidates:
        date_hma = hma_df.at[c_idx, 'Date_Obj']
        if date_acc == date_hma:
            acc_df.at[i, 'Match_Status'] = 'Matched'
            acc_df.at[i, 'Match_Index_HMA'] = c_idx
            hma_df.at[c_idx, 'Match_Status'] = 'Matched'
            hma_df.at[c_idx, 'Match_Index_ACC'] = i
            used_hma.add(c_idx)
            match_count_5 += 1
            break

print(f"--> Tahap 5 (Timeline: Tgl + Amt) selesai. Matches tambahan: {match_count_5}")

match_count_6 = 0
for i, row in acc_df[acc_df['Match_Status'] == 'Not in Coretax'].iterrows():
    amt = int(row['Amount_Clean'])
    name_prefix = row['Name_Clean'][:5]
    if len(name_prefix) < 3:
        continue 

    candidates = []
    for delta in range(-5, 6):
        candidates.extend(hma_by_amt.get(amt + delta, []))
    candidates = [c for c in candidates if c not in used_hma]

    for c_idx in candidates:
        hma_name_prefix = hma_df.at[c_idx, 'Name_Clean'][:5]
        if name_prefix == hma_name_prefix:
            acc_df.at[i, 'Match_Status'] = 'Matched'
            acc_df.at[i, 'Match_Index_HMA'] = c_idx
            hma_df.at[c_idx, 'Match_Status'] = 'Matched'
            hma_df.at[c_idx, 'Match_Index_ACC'] = i
            used_hma.add(c_idx)
            match_count_6 += 1
            break

print(f"--> Tahap 6 (Prefix: 5 Karakter Nama + Amt) selesai. Matches tambahan: {match_count_6}")
print(f"--> Total Data Matched: {len(used_hma)}")

print("--> Menyusun data laporan akhir...")

col_map = {
    'Date_ACC': 'Tanggal Accurate', 'Date_HMA': 'Tanggal Coretax',
    'Name_ACC': 'Nama Accurate', 'Name_HMA': 'Nama Coretax',
    'TIN_ACC': 'Nomor Pajak Accurate', 'TIN_HMA': 'Nomor Pajak Coretax',
    'Amount_ACC': 'Jumlah Accurate', 'Amount_HMA': 'Jumlah Coretax',
    'Difference': 'Perbedaan/Selisih', 'Notes': 'Catatan'
}

def add_total(df, label):
    if df.empty:
        return df
    total_acc = df['Jumlah Accurate'].sum() if 'Jumlah Accurate' in df else 0
    total_hma = df['Jumlah Coretax'].sum() if 'Jumlah Coretax' in df else 0
    total_diff = df['Perbedaan/Selisih'].sum() if 'Perbedaan/Selisih' in df else 0
    
    row = {c: '' for c in df.columns}
    row[df.columns[0]] = label
    row['Jumlah Accurate'] = total_acc
    row['Jumlah Coretax'] = total_hma
    row['Perbedaan/Selisih'] = total_diff
    return pd.concat([df, pd.DataFrame([row])], ignore_index=True)

list_a = []
for i, row in hma_df[hma_df['Match_Status'] == 'Not in Accurate'].iterrows():
    list_a.append({
        'Date_ACC': None, 'Date_HMA': format_indo_date(row['Date_Obj']),
        'Name_ACC': None, 'Name_HMA': row['Name'],
        'TIN_ACC': None, 'TIN_HMA': str(row['TIN']),
        'Amount_ACC': 0, 'Amount_HMA': row['VAT'],
        'Difference': -row['VAT'], 'Notes': 'Hanya di Coretax'
    })
df_a = pd.DataFrame(list_a).rename(columns=col_map)
df_a = add_total(df_a, "TOTAL A")

list_b = []
for i, row in acc_df[acc_df['Match_Status'] == 'Not in Coretax'].iterrows():
    list_b.append({
        'Date_ACC': row['Tgl. Pajak'], 'Date_HMA': None,
        'Name_ACC': row['Nama Pelanggan'], 'Name_HMA': None,
        'TIN_ACC': str(row['Nomor Pajak']), 'TIN_HMA': None,
        'Amount_ACC': row['Jumlah Pajak'], 'Amount_HMA': 0,
        'Difference': row['Jumlah Pajak'], 'Notes': 'Hanya di Accurate'
    })
df_b = pd.DataFrame(list_b).rename(columns=col_map)
df_b = add_total(df_b, "TOTAL B")

list_c = []
for i, row in acc_df[acc_df['Match_Status'] == 'Matched'].iterrows():
    hma_row = hma_df.loc[row['Match_Index_HMA']]
    notes = []
    
    if row['Name_Clean'] != hma_row['Name_Clean']:
        if (row['Name_Clean'] not in hma_row['Name_Clean']) and (hma_row['Name_Clean'] not in row['Name_Clean']):
            notes.append("Nama Beda")
             
    if row['TIN_Clean'] != hma_row['TIN_Clean']:
        notes.append("Format NPWP Beda")
    if row['Date_Obj'] != hma_row['Date_Obj']:
        notes.append("Beda Tanggal")
    
    diff = row['Jumlah Pajak'] - hma_row['VAT']
    if diff != 0:
        notes.append(f"Selisih: {diff}")
    
    list_c.append({
        'Date_ACC': row['Tgl. Pajak'], 'Date_HMA': format_indo_date(hma_row['Date_Obj']),
        'Name_ACC': row['Nama Pelanggan'], 'Name_HMA': hma_row['Name'],
        'TIN_ACC': str(row['Nomor Pajak']), 'TIN_HMA': str(hma_row['TIN']),
        'Amount_ACC': row['Jumlah Pajak'], 'Amount_HMA': hma_row['VAT'],
        'Difference': diff, 'Notes': "; ".join(notes)
    })
df_c = pd.DataFrame(list_c).rename(columns=col_map)
if not df_c.empty:
    df_c['helper'] = df_c['Catatan'].apply(lambda x: 0 if x else 1)
    df_c.sort_values(by=['helper', 'Tanggal Accurate'], inplace=True)
    df_c.drop(columns=['helper'], inplace=True)
df_c = add_total(df_c, "TOTAL C")

summ_data = {
    'Uraian': [
        'Total Pajak Accurate', 'Total Pajak Coretax', 'Selisih Total',
        'Total Tidak Ada di Accurate', 'Total Tidak Ada di Coretax'
    ],
    'Nilai': [
        acc_df['Jumlah Pajak'].sum(), hma_df['VAT'].sum(),
        acc_df['Jumlah Pajak'].sum() - hma_df['VAT'].sum(),
        df_a.iloc[-1]['Jumlah Coretax'] if not df_a.empty else 0,
        df_b.iloc[-1]['Jumlah Accurate'] if not df_b.empty else 0
    ]
}
df_summ = pd.DataFrame(summ_data)

print(f"--> Menulis file Excel: {output_file} ...")
writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
wb = writer.book

fmt_head = wb.add_format({'bold': True, 'align': 'center', 'bg_color': '#D3D3D3', 'border': 1})
fmt_curr = wb.add_format({'num_format': '#,##0', 'border': 1})
fmt_norm = wb.add_format({'border': 1})
fmt_tin = wb.add_format({'num_format': '@', 'border': 1})
fmt_title = wb.add_format({'bold': True, 'font_size': 12})
fmt_tot_bg = wb.add_format({'bold': True, 'bg_color': '#FFFF00', 'border': 1})
fmt_tot_curr = wb.add_format({'bold': True, 'bg_color': '#FFFF00', 'num_format': '#,##0', 'border': 1})

ws1 = wb.add_worksheet('Ringkasan')
ws1.write(0, 0, "RINGKASAN", fmt_title)
ws1.write(1, 0, "Uraian", fmt_head)
ws1.write(1, 1, "Nilai (Rp)", fmt_head)
for i, r in df_summ.iterrows():
    ws1.write(i+2, 0, r['Uraian'], fmt_norm)
    ws1.write(i+2, 1, r['Nilai'], fmt_curr)
ws1.set_column(0, 0, 40)
ws1.set_column(1, 1, 20)

ws2 = wb.add_worksheet('Rincian')
cursor = 0

def write_block(title, df, row_idx):
    ws2.write(row_idx, 0, title, fmt_title)
    if df.empty:
        ws2.write(row_idx+1, 0, "Nihil", fmt_norm)
        return row_idx + 3
    
    cols = list(df.columns)
    for c, col in enumerate(cols):
        ws2.write(row_idx+1, c, col, fmt_head)
    
    for r, row_data in df.iterrows():
        is_last = (r == len(df)-1)
        f_norm = fmt_tot_bg if is_last else fmt_norm
        f_curr = fmt_tot_curr if is_last else fmt_curr
        f_tin = fmt_tot_bg if is_last else fmt_tin
        
        for c, col in enumerate(cols):
            val = row_data[col]
            if pd.isna(val):
                val = ""
            ft = f_norm
            if col in ['Jumlah Accurate', 'Jumlah Coretax', 'Perbedaan/Selisih']:
                ft = f_curr
            elif 'Nomor Pajak' in col:
                ft = f_tin
            ws2.write(row_idx+2+r, c, val, ft)
            
    return row_idx + 2 + len(df) + 2

cursor = write_block("A. DATA YANG TIDAK ADA DI ACCURATE", df_a, cursor)
cursor = write_block("B. DATA YANG TIDAK ADA DI CORETAX", df_b, cursor)
cursor = write_block("C. DATA MATCHED", df_c, cursor)

widths = [18, 18, 30, 30, 25, 25, 20, 20, 20, 40]
for i, w in enumerate(widths):
    ws2.set_column(i, i, w)

print("--> Memproses Data Asli agar format angka (NPWP/ID) aman...")

def fix_scientific_notation(df_src, keywords):
    df_fix = df_src.copy()
    for col in df_fix.columns:
        if any(k.lower() in col.lower() for k in keywords):
            df_fix[col] = df_fix[col].astype(str).str.replace(r'\.0$', '', regex=True)
            df_fix[col] = df_fix[col].replace('nan', '')
    return df_fix

keywords_id = ['Nomor Pajak', 'NPWP', 'NIK', 'TIN', 'Tax', 'ID']
acc_source = fix_scientific_notation(acc_df, keywords_id)
hma_source = fix_scientific_notation(hma_df, keywords_id)

acc_source.to_excel(writer, sheet_name='Data Asli Accurate', index=False)
hma_source.to_excel(writer, sheet_name='Data Asli Coretax', index=False)

writer.sheets['Data Asli Accurate'].set_column(0, 20, 20)
writer.sheets['Data Asli Coretax'].set_column(0, 20, 20)

writer.close()
print("--> Selesai! File laporan telah siap!")