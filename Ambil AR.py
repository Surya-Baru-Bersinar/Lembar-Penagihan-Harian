import sys
import shutil
import subprocess
from pathlib import Path

FOLDER_DAPUR = Path("Dapur")

FILE_SYARAT = [
    "__init__.py",
    "1_CleanerAcc.py",
    "1_CleanerAccGiro.py",
    "1B_DownloaderMasterData.py",
    "1C_MergedMaster2Main.py",
    "1D_CleanZeroAR.py",
    "1E_DownloaderPendingData.py",
    "2_CompareGiro.py",
    "2_ComparePending.py",
    "2_FilterAR.py",
    "3_CalculateAR.py",
    "3_CalculateARPurePython.py",
    "4_HelperCleaningData.py",
    "5_InjectDataToSS.py",
    "6_PrintByPython.py",
    "credentials.json",
    "piutang.conf",
]

ALUR_EKSEKUSI = [
    ("1_CleanerAcc.py", "Memulai eksekusi pembersihan data utama"),
    ("1_CleanerAccGiro.py", "Memulai eksekusi pembersihan data utama"),
    ("1B_DownloaderMasterData.py", "Memulai pengunduhan data master"),
    ("1C_MergedMaster2Main.py", "Memulai penggabungan data master ke data utama"),
    ("1D_CleanZeroAR.py", "Memulai pembersihan saldo piutang nol"),
    ("1E_DownloaderPendingData.py", "Memulai pengunduhan data pendingan"),
    ("2_CompareGiro.py", "Memulai komparasi data giro"),
    ("2_ComparePending.py", "Memulai komparasi data pendingan"),
    ("2_FilterAR.py", "Memulai eksekusi filter data sementara"),
    ("3_CalculateAR.py", "Memulai eksekusi menyalin dan menyusun data pada template"),
    ("3_CalculateARPurePython.py", "Menggunakan metode pure code yang ringan"),
    ("4_HelperCleaningData.py", "Memulai persiapan data untuk disusun ke Spreadsheets"),
    ("5_InjectDataToSS.py", "Memulai unggah data ke Spreadsheets"),
    ("6_PrintByPython.py", "Memulai menjalankan print otomatis dengan konfigurasi"),
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

    bersihkan_file_sementara(FOLDER_DAPUR, ["*temp.xlsx", "Giro.xls", "Piutang.xls"])

    file_piutang = Path("Piutang.xls")
    if file_piutang.is_file():
        shutil.copy2(file_piutang, FOLDER_DAPUR / file_piutang.name)
    else:
        print("--> File Piutang.xls tidak ditemukan untuk diproses.")
        input("--> Tekan enter untuk keluar.")
        return

    file_giro = Path("Giro.xls")
    if file_giro.is_file():
        shutil.copy2(file_giro, FOLDER_DAPUR / file_giro.name)

    try:
        for skrip, pesan in ALUR_EKSEKUSI:
            print(f"--> {pesan}")
            subprocess.run([sys.executable, skrip], cwd=FOLDER_DAPUR, check=True)

            if skrip == "3_CalculateAR.py":
                salin_laporan_ar(FOLDER_DAPUR)

    except subprocess.CalledProcessError:
        print("\nError terjadi kesalahan saat menjalankan skrip.")
        input("--> Tekan enter untuk keluar.")
        return

    bersihkan_file_sementara(
        FOLDER_DAPUR, ["*temp.xlsx", "Giro.xls", "Piutang.xls", "Print_AR.xlsm"]
    )

    print("--> Semua proses telah selesai dijalankan.")
    input("--> Tekan enter untuk keluar.")


if __name__ == "__main__":
    jalankan_otomatisasi()
