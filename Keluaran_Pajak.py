import os
import shutil
import glob
import sys
import subprocess

def main():
    try:
        current_dir = os.getcwd()
        dapur_dir = os.path.join(current_dir, "Dapur")
        addon_dir = os.path.join(current_dir, "Addon")

        safi_exists = os.path.exists("Safi.xls")
        fella_exists = os.path.exists("Fella.xls")
        
        data_exports = glob.glob("data_export*.csv") + glob.glob("data_export*.xlsx")

        if not (safi_exists or fella_exists):
            print("Gagal: File Safi.xls atau Fella.xls tidak ditemukan.")
            input("Tekan enter untuk keluar")
            return

        if not data_exports:
            print("Gagal: File data_export (csv/xlsx) tidak ditemukan.")
            input("Tekan enter untuk keluar")
            return

        if not os.path.exists(dapur_dir) or not os.path.exists(addon_dir):
            print("Gagal: Folder Dapur atau Addon tidak ditemukan.")
            input("Tekan enter untuk keluar")
            return

        dapur_files = [
            "1_Cleaner&MergerACC.py",
            "2_HMA_ex2ex_analytics_third.py",
            "2_HMA_csv2ex_analytics_third.py",
            "3_Rincian2.py",
            "4_rincian3.py"
        ]

        addon_files = [
            "audit_data_gabungan_keluaran_ex.py",
            "cleaner&merger.py",
            "cleaner_acc_keluaran_1612.py",
            "csv2ex.py",
            "hitung_ppn_all_file_keluaran.py",
            "ma_ex2ex.py",
            "tang2ddmmmyy_ex2ex_keluaran.py"
        ]

        missing_files = []

        for f in dapur_files:
            if not os.path.exists(os.path.join(dapur_dir, f)):
                missing_files.append(f"Dapur/{f}")

        for f in addon_files:
            if not os.path.exists(os.path.join(addon_dir, f)):
                missing_files.append(f"Addon/{f}")

        if missing_files:
            print("Gagal: File berikut tidak ditemukan:")
            for mf in missing_files:
                print(f"- {mf}")
            input("Tekan enter untuk keluar")
            return

        print("Pengecekan selesai. Memulai proses penyalinan file...")

        if safi_exists:
            shutil.copy("Safi.xls", dapur_dir)
        elif fella_exists:
            shutil.copy("Fella.xls", dapur_dir)

        for export_file in data_exports:
            shutil.copy(export_file, dapur_dir)

        print("Menjalankan 1_Cleaner&MergerACC.py...")
        subprocess.run([sys.executable, "1_Cleaner&MergerACC.py"], cwd=dapur_dir, check=True)

        hma_ex2ex_path = os.path.join(dapur_dir, "HMA_ex2ex.xlsx")
        hma_csv2ex_path = os.path.join(dapur_dir, "HMA_csv2ex.xlsx")

        if os.path.exists(hma_ex2ex_path):
            print("Menjalankan 2_HMA_ex2ex_analytics_third.py...")
            subprocess.run([sys.executable, "2_HMA_ex2ex_analytics_third.py"], cwd=dapur_dir, check=True)
        elif os.path.exists(hma_csv2ex_path):
            print("Menjalankan 2_HMA_csv2ex_analytics_third.py...")
            subprocess.run([sys.executable, "2_HMA_csv2ex_analytics_third.py"], cwd=dapur_dir, check=True)
        else:
            print("Gagal: File hasil HMA_ex2ex.xlsx atau HMA_csv2ex.xlsx tidak ditemukan di folder Dapur.")
            input("Tekan enter untuk keluar")
            return

        print("Menjalankan 3_Rincian2.py...")
        subprocess.run([sys.executable, "3_Rincian2.py"], cwd=dapur_dir, check=True)
        
        print("Menjalankan 4_rincian3.py...")
        subprocess.run([sys.executable, "4_rincian3.py"], cwd=dapur_dir, check=True)

        result_files = glob.glob(os.path.join(dapur_dir, "Laporan_Analisa_Pajak*.xlsx"))

        if result_files:
            target_file = result_files[0]
            filename = os.path.basename(target_file)
            destination = os.path.join(current_dir, filename)
            
            shutil.copy(target_file, destination)
            os.remove(target_file)
            print(f"Selesai. File {filename} telah disalin ke folder utama.")
        else:
            print("Gagal: File Laporan_Analisa_Pajak tidak ditemukan setelah proses selesai.")
            input("Tekan enter untuk keluar")
            return

    except subprocess.CalledProcessError as e:
        print(f"Terjadi kesalahan saat menjalankan script: {e}")
        input("Tekan enter untuk keluar")
    except Exception as e:
        print(f"Terjadi kesalahan: {e}")
        input("Tekan enter untuk keluar")

if __name__ == "__main__":
    main()
