import configparser
import os
import pandas as pd


def cek_konfigurasi(config_path="piutang.conf"):
    if not os.path.exists(config_path):
        print(f"--> File konfigurasi '{config_path}' tidak ditemukan.")
        return False, 0.0

    config = configparser.ConfigParser(allow_no_value=True, strict=False)
    try:
        config.read(config_path)
        if "GIRO" in config and "giro_stats" in config["GIRO"]:
            status = config["GIRO"]["giro_stats"].strip()
            if status.lower() == "ya":
                giro_cut_raw = config["GIRO"].get("giro_cut", "0")
                try:
                    giro_cut = float(str(giro_cut_raw).strip())
                except ValueError:
                    giro_cut = 0.0

                return True, giro_cut
            else:
                print(
                    f"--> Status giro_stats adalah '{status}'. Script dilewati (skip)."
                )
                return False, 0.0
        else:
            print(
                "--> Section [GIRO] atau key 'giro_stats' tidak ditemukan di piutang.conf."
            )
            return False, 0.0
    except Exception as e:
        print(f"--> Gagal membaca file konfigurasi: {e}")
        return False, 0.0


def proses_pencocokan_giro():
    status_aktif, giro_cut = cek_konfigurasi("piutang.conf")
    if not status_aktif:
        return

    print(
        f"--> Fitur GIRO aktif. Batas toleransi selisih (giro_cut): Rp {giro_cut:,.0f}"
    )

    export_path = "Piutang_clean_temp.xlsx"
    giro_path = "Giro_temp.xlsx"

    if not os.path.exists(export_path) or not os.path.exists(giro_path):
        print(
            f"--> File '{export_path}' atau '{giro_path}' tidak ditemukan!"
        )
        return

    print("--> Membaca file Excel...")

    df_export = pd.read_excel(export_path)
    df_giro = pd.read_excel(giro_path)

    df_export.columns = df_export.columns.str.strip()
    df_giro.columns = df_giro.columns.str.strip()

    df_export["_kode_pelanggan_clean"] = (
        df_export["Kode Pelanggan"].astype(str).str.strip()
    )
    df_export["_no_faktur_clean"] = (
        df_export["No. Faktur"].astype(str).str.strip()
    )

    df_giro["_no_pelanggan_clean"] = (
        df_giro["No. Pelanggan"].astype(str).str.strip()
    )
    df_giro["_no_faktur_so_clean"] = (
        df_giro["No. Faktur. (SO)"].astype(str).str.strip()
    )

    df_export["Nilai Faktur"] = pd.to_numeric(
        df_export["Nilai Faktur"], errors="coerce"
    ).fillna(0)
    df_giro["Nilai terima"] = pd.to_numeric(
        df_giro["Nilai terima"], errors="coerce"
    ).fillna(0)

    print(
        "--> Menghitung akumulasi penerimaan Giro per Nomor Faktur..."
    )
    giro_totals = (
        df_giro.groupby(["_no_pelanggan_clean", "_no_faktur_so_clean"])[
            "Nilai terima"
        ]
        .sum()
        .reset_index()
    )
    giro_totals.rename(
        columns={"Nilai terima": "_total_nilai_terima"}, inplace=True
    )

    df_merged = pd.merge(
        df_export,
        giro_totals,
        left_on=["_kode_pelanggan_clean", "_no_faktur_clean"],
        right_on=["_no_pelanggan_clean", "_no_faktur_so_clean"],
        how="left",
    )

    is_klop = df_merged["_total_nilai_terima"].notna() & (
        (df_merged["Nilai Faktur"] - df_merged["_total_nilai_terima"]).abs()
        <= giro_cut
    )

    total_dihapus = is_klop.sum()

    if total_dihapus == 0:
        print("--> Tidak ditemukan data yang klop untuk dihapus.")
        return

    df_hasil = df_export[~is_klop].copy()

    df_hasil.drop(
        columns=["_kode_pelanggan_clean", "_no_faktur_clean"], inplace=True
    )

    df_hasil.to_excel(export_path, index=False)

    print(
        f"--> Berhasil menghapus {total_dihapus} data yang klop (dengan toleransi <= Rp {giro_cut:,.0f}) dari '{export_path}'."
    )


if __name__ == "__main__":
    proses_pencocokan_giro()