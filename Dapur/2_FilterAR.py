import pandas as pd
import re

map_penagih = {}
current_sales = None
mode = None

with open('piutang.conf', 'r') as f:
    for line in f:
        line = line.strip()
        if not line:
            continue
        if line == '[NAMA SALES]':
            mode = 'sales'
        elif line == '[KODE PELANGGAN]':
            mode = 'kode'
        else:
            if mode == 'sales':
                current_sales = line
            elif mode == 'kode' and current_sales:
                map_penagih[line] = current_sales

df = pd.read_excel('ExportFile_clean_temp.xlsx')

if 'Kode Pelanggan' in df.columns and 'Kode' not in df.columns:
    df.rename(columns={'Kode Pelanggan': 'Kode'}, inplace=True)

kolom_diambil = ['Kode', 'Nama Pelanggan', 'Umur JT', 'No. Faktur', 'Tgl Faktur', 'Nilai Faktur', 'Sisa Piutang']
df = df[kolom_diambil].copy()
df = df[df['Kode'].isin(map_penagih.keys())].copy()

indo_months_in = {
    'Jan': 'Jan', 'Feb': 'Feb', 'Mar': 'Mar', 'Apr': 'Apr', 'Mei': 'May', 'Jun': 'Jun',
    'Jul': 'Jul', 'Agu': 'Aug', 'Sep': 'Sep', 'Okt': 'Oct', 'Nop': 'Nov', 'Des': 'Dec',
    'Peb': 'Feb', 'Ags': 'Aug', 'Agt': 'Aug',
    
    'jan': 'Jan', 'feb': 'Feb', 'mar': 'Mar', 'apr': 'Apr', 'mei': 'May', 'jun': 'Jun',
    'jul': 'Jul', 'agu': 'Aug', 'ags': 'Aug', 'agt': 'Aug', 'sep': 'Sep', 'okt': 'Oct', 'nop': 'Nov', 
    'nov': 'Nov', 'des': 'Dec'
}

def konversi_bulan_indo(teks_tanggal):
    if pd.isna(teks_tanggal):
        return teks_tanggal
    teks_tanggal = str(teks_tanggal)
    for indo, eng in indo_months_in.items():
        teks_tanggal = re.sub(r'\b' + indo + r'\b', eng, teks_tanggal)
    return teks_tanggal

tgl_bersih = df['Tgl Faktur'].apply(konversi_bulan_indo)

df['Tgl_Faktur_Parsed'] = pd.to_datetime(tgl_bersih, format='mixed', errors='coerce')

hari_ini = pd.Timestamp.now().normalize()
df['Umur JT'] = (hari_ini - df['Tgl_Faktur_Parsed']).dt.days
df['Tgl Faktur'] = df['Tgl_Faktur_Parsed'].dt.strftime('%d/%m/%Y')
df = df.drop(columns=['Tgl_Faktur_Parsed'])
df['Nilai Faktur'] = pd.to_numeric(df['Nilai Faktur'], errors='coerce').fillna(0)
df['Sisa Piutang'] = pd.to_numeric(df['Sisa Piutang'], errors='coerce').fillna(0)
df['Terbayar'] = df['Nilai Faktur'] - df['Sisa Piutang']

kolom_baru = ['Kode', 'Nama Pelanggan', 'Umur JT', 'No. Faktur', 'Tgl Faktur', 'Nilai Faktur', 'Terbayar', 'Sisa Piutang']
df = df[kolom_baru].copy()
df['Penagih'] = df['Kode'].map(map_penagih)
df = df.sort_values(by=['Penagih', 'Nama Pelanggan', 'Kode', 'No. Faktur'])

data_akhir = []
penagih_sebelumnya = None
sub_nilai = 0
sub_terbayar = 0
sub_sisa = 0

for idx, row in df.iterrows():
    current_penagih = row['Penagih']
    
    if penagih_sebelumnya is not None and current_penagih != penagih_sebelumnya:
        baris_total = {col: None for col in df.columns}
        baris_total['Tgl Faktur'] = 'TOTAL ' + str(penagih_sebelumnya)
        baris_total['Nilai Faktur'] = sub_nilai
        baris_total['Terbayar'] = sub_terbayar if sub_terbayar != 0 else None
        baris_total['Sisa Piutang'] = sub_sisa
        data_akhir.append(baris_total)
        
        baris_pemisah = {col: None for col in df.columns}
        data_akhir.append(baris_pemisah)
        
        sub_nilai = 0
        sub_terbayar = 0
        sub_sisa = 0
        
    row_dict = row.to_dict()
    sub_nilai += row['Nilai Faktur']
    sub_terbayar += row['Terbayar']
    sub_sisa += row['Sisa Piutang']
    
    row_dict['Terbayar'] = row['Terbayar'] if row['Terbayar'] != 0 else None
    
    data_akhir.append(row_dict)
    penagih_sebelumnya = current_penagih

if penagih_sebelumnya is not None:
    baris_total = {col: None for col in df.columns}
    baris_total['Tgl Faktur'] = 'TOTAL ' + str(penagih_sebelumnya)
    baris_total['Nilai Faktur'] = sub_nilai
    baris_total['Terbayar'] = sub_terbayar if sub_terbayar != 0 else None
    baris_total['Sisa Piutang'] = sub_sisa
    data_akhir.append(baris_total)

df_hasil = pd.DataFrame(data_akhir)
df_hasil = df_hasil[['Penagih'] + kolom_baru]

writer = pd.ExcelWriter('Laporan_Piutang_Penagih_temp.xlsx', engine='xlsxwriter')
df_hasil.to_excel(writer, index=False, sheet_name='Sheet1')

workbook = writer.book
worksheet = writer.sheets['Sheet1']

format_uang = workbook.add_format({'num_format': '#,##0.00'})

for i, col in enumerate(df_hasil.columns):
    panjang_maksimal = df_hasil[col].apply(lambda x: len(str(x))).max()
    column_len = max(panjang_maksimal, len(col)) + 2
    
    if col in ['Nilai Faktur', 'Terbayar', 'Sisa Piutang']:
        worksheet.set_column(i, i, column_len, format_uang)
    else:
        worksheet.set_column(i, i, column_len)

writer.close()

print("--> Proses selesai!")