import configparser
import os
import pandas as pd

def cek_konfigurasi(config_path="piutang.conf"):
    if not os.path.exists(config_path):
        print(f"--> File konfigurasi '{config_path}' tidak ditemukan.")
        return False, {}

    config = configparser.ConfigParser(allow_no_value=True, strict=False)
    try:
        config.read(config_path)
        if "PENDING" in config and "pend_stats" in config["PENDING"]:
            status = config["PENDING"]["pend_stats"].strip()
            if status.lower() == "ya":
                cfg_pending = {
                    "sheet": config["PENDING"].get("pend_sheet", "").strip(),
                    "col_key": config["PENDING"]
                    .get("pend_col_key", "")
                    .strip(),
                    "col_ret": config["PENDING"]
                    .get("pend_col_ret", "")
                    .strip(),
                }
                return True, cfg_pending
            else:
                print(
                    f"--> Status pend_stats adalah '{status}'. Script dilewati (skip)."
                )
                return False, {}
        else:
            print(
                "--> Section [PENDING] atau key 'pend_stats' tidak ditemukan di piutang.conf."
            )
            return False, {}
    except Exception as e:
        print(f"--> Gagal membaca file konfigurasi: {e}")
        return False, {}

def proses_pencocokan_pending():
    status_aktif, cfg = cek_konfigurasi("piutang.conf")
    if not status_aktif:
        return

    export_path = "Piutang_clean_temp.xlsx"
    pending_path = "Pending_temp.xlsx"

    if not os.path.exists(export_path) or not os.path.exists(pending_path):
        print(
            f"--> File '{export_path}' atau '{pending_path}' tidak ditemukan!"
        )
        return

    print("--> Membaca file Excel...")

    try:
        df_export = pd.read_excel(export_path)
        df_pending = pd.read_excel(pending_path, sheet_name=cfg["sheet"])
    except Exception as e:
        print(f"--> Gagal membaca file Excel atau sheet '{cfg['sheet']}': {e}")
        return

    df_export.columns = df_export.columns.str.strip()
    df_pending.columns = df_pending.columns.str.strip()

    col_key = cfg["col_key"]
    col_ret = cfg["col_ret"]

    if col_key not in df_pending.columns or col_ret not in df_pending.columns:
        print(
            f"--> Kolom '{col_key}' atau '{col_ret}' tidak ditemukan di sheet '{cfg['sheet']}'!"
        )
        return

    if "No. Faktur" not in df_export.columns:
        print(
            "--> Kolom 'No. Faktur' tidak ditemukan di Piutang_clean_temp.xlsx!"
        )
        return

    df_export["_no_faktur_clean"] = (
        df_export["No. Faktur"].astype(str).str.strip()
    )
    df_pending["_key_clean"] = df_pending[col_key].astype(str).str.strip()

    is_ret_empty = (
        df_pending[col_ret].isna()
        | (df_pending[col_ret].astype(str).str.strip() == "")
        | (df_pending[col_ret].astype(str).str.strip().str.lower() == "nan")
        | (df_pending[col_ret].astype(str).str.strip().str.lower() == "nat")
    )

    keys_to_remove = set(df_pending.loc[is_ret_empty, "_key_clean"].unique())

    is_in_remove_list = df_export["_no_faktur_clean"].isin(keys_to_remove)
    total_dihapus = is_in_remove_list.sum()

    if total_dihapus == 0:
        print(
            "--> Tidak ditemukan data pending tanpa tanggal yang cocok untuk dihapus."
        )
        return

    df_hasil = df_export[~is_in_remove_list].copy()
    df_hasil.drop(columns=["_no_faktur_clean"], inplace=True)

    df_hasil.to_excel(export_path, index=False)

    print(
        f"--> Berhasil menghapus {total_dihapus} baris dari '{export_path}'."
    )

if __name__ == "__main__":
    proses_pencocokan_pending()
