import pandas as pd
import os
import glob

def hitung_total_kolom_g(folder_path):
    files = glob.glob(os.path.join(folder_path, "*.xlsx")) + glob.glob(os.path.join(folder_path, "*.xls"))
    
    total_keseluruhan = 0
    
    print(f"--> Ditemukan {len(files)} file Excel di folder '{folder_path}'.\n")

    for file in files:
        try:
            df = pd.read_excel(file, header=None)
            
            if df.shape[1] > 6:
                data_g = df.iloc[:, 6]
                data_g_numeric = pd.to_numeric(data_g, errors='coerce').fillna(0)
                subtotal = data_g_numeric.sum()
                print(f"--> OK File: {os.path.basename(file)} -> Subtotal Kolom G: {subtotal}")
                total_keseluruhan += subtotal
            else:
                print(f"--> Skip File: {os.path.basename(file)} -> Tidak memiliki Kolom G.")
                
        except Exception as e:
            print(f"--> Error gagal membaca file {os.path.basename(file)}: {e}")

    print("-" * 50)
    print(f"--> TOTAL JUMLAH SELURUHNYA: {total_keseluruhan}")

if __name__ == "__main__":
    lokasi_folder = "." 
    
    hitung_total_kolom_g(lokasi_folder)