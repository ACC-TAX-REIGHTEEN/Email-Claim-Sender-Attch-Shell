import os
import glob
import shutil
import subprocess
import sys

def jalankan_otomatisasi():
    folder_utama = os.path.dirname(os.path.abspath(__file__))
    folder_dapur = os.path.join(folder_utama, "Dapur")
    file_syarat = ["1_EkstrakData.py", "2_GmailSender.py", "config.conf", "__init__.py"]

    if not os.path.exists(folder_dapur) or not os.path.isdir(folder_dapur):
        print("--> Folder Dapur tidak ditemukan.")
        input("--> Tekan enter untuk keluar.")
        return

    for file in file_syarat:
        jalur_file = os.path.join(folder_dapur, file)
        if not os.path.isfile(jalur_file):
            print("--> File " + file + " tidak ditemukan di dalam folder Dapur.")
            input("--> Tekan enter untuk keluar.")
            return

    ekstensi_hapus_dapur = ["*.xls*", "*.pdf", "*.PDF"]
    for ekstensi in ekstensi_hapus_dapur:
        for file in glob.glob(os.path.join(folder_dapur, ekstensi)):
            if os.path.isfile(file):
                os.remove(file)

    file_sumber_pola = ["Surat Saffiela - 130326.xlsm", "*.pdf", "*.PDF"]
    ada_file_dipindah = False
    
    for pola in file_sumber_pola:
        for file in glob.glob(os.path.join(folder_utama, pola)):
            if os.path.isfile(file):
                shutil.copy2(file, os.path.join(folder_dapur, os.path.basename(file)))
                ada_file_dipindah = True

    if not ada_file_dipindah:
        print("--> File sumber tidak ditemukan untuk diproses.")
        input("--> Tekan enter untuk keluar.")
        return

    print("--> Memulai eksekusi 1_EkstrakData.py")
    subprocess.run([sys.executable, "1_EkstrakData.py"], cwd=folder_dapur)

    print("--> Memulai eksekusi 2_GmailSender.py")
    subprocess.run([sys.executable, "2_GmailSender.py"], cwd=folder_dapur)

    for ekstensi in ekstensi_hapus_dapur:
        for file in glob.glob(os.path.join(folder_dapur, ekstensi)):
            if os.path.isfile(file):
                os.remove(file)

    ekstensi_hapus_utama = ["*.pdf", "*.PDF"]
    for ekstensi in ekstensi_hapus_utama:
        for file in glob.glob(os.path.join(folder_utama, ekstensi)):
            if os.path.isfile(file):
                os.remove(file)

    print("--> Semua proses telah selesai dijalankan.")
    input("--> Tekan enter untuk keluar.")

if __name__ == "__main__":
    jalankan_otomatisasi()