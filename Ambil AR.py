import sys
import shutil
import subprocess
from pathlib import Path

FOLDER_DAPUR = Path("Dapur")

FILE_SYARAT = [
    "__init__.py",
    "1_CleanerAcc.py",
    "1B_DownloaderMasterData.py",
    "1C_MergedMaster2Main.py",
    "1D_CleanZeroAR.py",
    "2_FilterAR.py",
    "3_CalculateAR.py",
    "4_HelperCleaningData.py",
    "5_InjectDataToSS.py",
    "credentials.json",
    "piutang.conf",
]

ALUR_EKSEKUSI = [
    ("1_CleanerAcc.py", "Memulai eksekusi pembersihan data utama"),
    ("1B_DownloaderMasterData.py", "Memulai pengunduhan data master"),
    ("1C_MergedMaster2Main.py", "Memulai penggabungan data master ke data utama"),
    ("1D_CleanZeroAR.py", "Memulai pembersihan saldo piutang nol (Zero AR)"),
    ("2_FilterAR.py", "Memulai eksekusi filter data sementara"),
    ("3_CalculateAR.py", "Memulai eksekusi menyalin dan menyusun data pada template"),
    ("4_HelperCleaningData.py", "Memulai persiapan data untuk disusun ke Spreadsheets"),
    ("5_InjectDataToSS.py", "Memulai unggah data ke Spreadsheets"),
]

def bersihkan_file_sementara(folder: Path, pola_file: list[str]) -> None:
    for pola in pola_file:
        for file_path in folder.glob(pola):
            try:
                file_path.unlink()
            except Exception:
                pass


def salin_laporan_ar(folder_sumber: Path) -> None:
    for file_laporan in folder_sumber.glob("*AR.xlsm"):
        shutil.copy2(file_laporan, file_laporan.name)

def jalankan_otomatisasi():
    if not FOLDER_DAPUR.is_dir():
        print("--> Folder Dapur tidak ditemukan.")
        input("--> Tekan enter untuk keluar.")
        return

    for nama_file in FILE_SYARAT:
        jalur_file = FOLDER_DAPUR / nama_file
        if not jalur_file.is_file():
            print(f"--> File {nama_file} tidak ditemukan di dalam folder Dapur.")
            input("--> Tekan enter untuk keluar.")
            return

    bersihkan_file_sementara(FOLDER_DAPUR, ["*temp.xlsx", "ExportFile.xls"])

    file_export_sumber = Path("ExportFile.xls")
    if file_export_sumber.is_file():
        shutil.copy2(file_export_sumber, FOLDER_DAPUR / file_export_sumber.name)
    else:
        print("--> File ExportFile.xls tidak ditemukan untuk diproses.")
        input("--> Tekan enter untuk keluar.")
        return

    try:
        for skrip, pesan in ALUR_EKSEKUSI:
            print(f"--> {pesan}")
            subprocess.run([sys.executable, skrip], cwd=FOLDER_DAPUR, check=True)

            if skrip == "3_CalculateAR.py":
                salin_laporan_ar(FOLDER_DAPUR)

    except subprocess.CalledProcessError as err:
        print(f"\n[ERROR] Terjadi kesalahan saat menjalankan skrip.")
        input("--> Tekan enter untuk keluar.")
        return

    bersihkan_file_sementara(FOLDER_DAPUR, ["*temp.xlsx", "ExportFile.xls", "Print_AR.xlsm"])

    print("--> Semua proses telah selesai dijalankan.")
    input("--> Tekan enter untuk keluar.")


if __name__ == "__main__":
    jalankan_otomatisasi()