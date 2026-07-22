import configparser
import os
import pandas as pd

def find_header_and_read(file_path, sheet_name, required_cols_options):
    sheet_target = sheet_name if sheet_name else 0
    df_raw = pd.read_excel(file_path, sheet_name=sheet_target, header=None)
    
    header_idx = None
    matched_cols = []
    
    for idx, row in df_raw.iterrows():
        row_values = [str(val).strip() for val in row.dropna().values]
        
        temp_matched = []
        for col_opts in required_cols_options:
            if isinstance(col_opts, list):
                found = next((c for c in col_opts if c in row_values), None)
                if found:
                    temp_matched.append(found)
            else:
                if col_opts in row_values:
                    temp_matched.append(col_opts)
                    
        if len(temp_matched) == len(required_cols_options):
            header_idx = idx
            matched_cols = temp_matched
            break
            
    if header_idx is None:
        raise ValueError(
            f"Kolom {required_cols_options} tidak ditemukan pada sheet '{sheet_target}' di file {file_path}."
        )
    
    df = pd.read_excel(file_path, sheet_name=sheet_target, header=header_idx)
    df.columns = df.columns.astype(str).str.strip()
    return df, matched_cols

def process_piutang():
    config_file = 'piutang.conf'
    
    if not os.path.exists(config_file):
        print(f"--> File konfigurasi '{config_file}' tidak ditemukan.")
        return

    config = configparser.ConfigParser(allow_no_value=True, strict=False)
    config.read(config_file)
    
    if 'MASTER' not in config:
        print("--> Seksi [MASTER] tidak ditemukan dalam file piutang.conf.")
        return

    master_cfg = config['MASTER']
    masdatus = master_cfg.get('masdatus', 'No').strip()
    
    if masdatus.lower() != 'ya':
        print(f"--> Status masdatus adalah '{masdatus}'. Seluruh proses di-skip.")
        return

    print("--> Status masdatus = Ya. Memulai proses sinkronisasi data...")

    mas_sheet = master_cfg.get('mas_sheet', '').strip()
    mas_col_key = master_cfg.get('mas_col_key', '').strip()
    mas_col_ret = master_cfg.get('mas_col_ret', '').strip()

    master_file = "Master_temp.xlsx"
    export_file = "ExportFile_clean_temp.xlsx"

    if not os.path.exists(master_file):
        print(f"--> File master '{master_file}' tidak ditemukan.")
        return
    if not os.path.exists(export_file):
        print(f"--> File target '{export_file}' tidak ditemukan.")
        return

    print(f"--> Membaca {master_file}...")
    df_master, _ = find_header_and_read(
        file_path=master_file,
        sheet_name=mas_sheet,
        required_cols_options=[mas_col_key, mas_col_ret]
    )

    lookup_dict = {}
    for _, row in df_master.iterrows():
        key_val = row[mas_col_key]
        ret_val = row[mas_col_ret]
        
        if pd.notna(key_val) and pd.notna(ret_val):
            clean_key = str(key_val).strip().upper()
            clean_ret = str(ret_val).strip()
            lookup_dict[clean_key] = clean_ret

    print(f"--> Berhasil memuat {len(lookup_dict)} data acuan dari Master.")

    print(f"--> Membaca {export_file}...")
    
    key_options = ["Kode Pelanggan", "Kode"]
    val_options = "Nama Pelanggan"

    df_target, matched_cols = find_header_and_read(
        file_path=export_file,
        sheet_name=0,
        required_cols_options=[key_options, val_options]
    )

    target_key_col = matched_cols[0]
    target_val_col = matched_cols[1]

    updated_count = 0
    
    def update_row(row):
        nonlocal updated_count
        kode_raw = row[target_key_col]
        
        if pd.notna(kode_raw):
            clean_kode = str(kode_raw).strip().upper()
            if clean_kode in lookup_dict:
                updated_count += 1
                return lookup_dict[clean_kode]
                
        return row[target_val_col]

    df_target[target_val_col] = df_target.apply(update_row, axis=1)

    df_target.to_excel(export_file, index=False)
    print(f"--> Selesai! Berhasil memperbarui {updated_count} baris data 'Nama Pelanggan' pada '{export_file}'.")

if __name__ == "__main__":
    process_piutang()