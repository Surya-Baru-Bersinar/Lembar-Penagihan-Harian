import pandas as pd
import numpy as np

def clean_data_autofit(input_file, output_file):
    print(f"--> Sedang memproses file: {input_file}...")
    
    try:
        df = pd.read_excel(input_file, header=None)
    except Exception as e:
        print(f"--> Error membaca file: {e}")
        return
        
    target_headers = [
        "No. Faktur", "Tgl Faktur", "Kode", "Nama Pelanggan", "Nilai Faktur", "Sisa Piutang", "Umur JT"
    ]
    
    header_map = {}
    start_row = 0
    
    for i in range(min(150, len(df))):
        for j in range(len(df.columns)):
            val = str(df.iat[i, j]).strip()
            if val in target_headers and val not in header_map:
                header_map[val] = j
                if val == "No. Faktur":
                    start_row = i
                    
    if "No. Faktur" not in header_map:
        print("--> Error: Kolom No. Faktur tidak ditemukan.")
        return

    col_faktur = header_map["No. Faktur"]
    df_data = df.iloc[start_row + 1:].copy()
    
    df_data[col_faktur] = df_data[col_faktur].astype(str).str.strip()
    kondisi_kosong_faktur = df_data[col_faktur].str.lower().isin(['nan', 'none', ''])
    df_data.loc[kondisi_kosong_faktur, col_faktur] = np.nan
    
    df_clean = df_data.dropna(subset=[col_faktur]).copy()
    
    header_labels = ["No. Faktur", "Faktur", "No.", "Total", "Halaman", "Page", "Tanggal"]
    df_clean = df_clean[~df_clean[col_faktur].astype(str).str.contains('|'.join(header_labels), case=False, na=False)]
    
    if "Kode" in header_map:
        col_kode = header_map["Kode"]
        df_clean[col_kode] = df_clean[col_kode].astype(str).str.strip()
        kondisi_kosong_kode = df_clean[col_kode].str.lower().isin(['nan', 'none', ''])
        df_clean.loc[kondisi_kosong_kode, col_kode] = np.nan
        df_clean = df_clean.dropna(subset=[col_kode]).copy()
        
    final_cols = []
    for h in target_headers:
        if h in header_map:
            df_clean[h] = df_clean[header_map[h]]
            final_cols.append(h)
            
    df_final = df_clean[final_cols].copy()
    
    def parse_to_float(val):
        if pd.isna(val) or str(val).strip() == "":
            return np.nan
        if isinstance(val, (int, float)):
            return float(val)
        s = str(val).strip()
        
        if '.' in s and ',' in s:
            if s.rfind(',') > s.rfind('.'):
                s = s.replace('.', '').replace(',', '.')
            else:
                s = s.replace(',', '')
        elif ',' in s:
            if len(s) - s.rfind(',') <= 3:
                s = s.replace(',', '.')
            else:
                s = s.replace(',', '')
        elif '.' in s:
            if s.count('.') > 1 or (len(s) - s.rfind('.') == 4):
                s = s.replace('.', '')
        try:
            return float(s)
        except:
            return np.nan

    cols_to_clean = ['Nilai Faktur', 'Sisa Piutang']
    for col in cols_to_clean:
        if col in df_final.columns:
            df_final[col] = df_final[col].apply(parse_to_float)

    df_final.reset_index(drop=True, inplace=True)
    
    try:
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Data Bersih')
            workbook = writer.book
            worksheet = writer.sheets['Data Bersih']
            
            format_angka = workbook.add_format({'num_format': '#,##0.00'})
            
            for i, col in enumerate(df_final.columns):
                panjang_maksimal = df_final[col].apply(lambda x: len(str(x))).max() if not df_final.empty else 0
                max_len = max(panjang_maksimal, len(col)) + 2
                
                if col in cols_to_clean:
                    worksheet.set_column(i, i, max_len, format_angka)
                else:
                    worksheet.set_column(i, i, max_len)
                    
        print(f"--> SUKSES! File tersimpan rapi di: {output_file}")
        
    except Exception as e:
        print(f"--> Error saat menyimpan file: {e}")

    return df_final

input_filename = 'ExportFile.xls' 
output_filename = 'ExportFile_clean_temp.xlsx'

if __name__ == "__main__":
    clean_data_autofit(input_filename, output_filename)