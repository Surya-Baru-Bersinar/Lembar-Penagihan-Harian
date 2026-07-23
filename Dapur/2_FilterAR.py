import pandas as pd
import re

daftar_instruksi = []
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
                daftar_instruksi.append({
                    'sales': current_sales,
                    'kode': line
                })

df = pd.read_excel('ExportFile_clean_temp.xlsx')

if 'Kode Pelanggan' in df.columns and 'Kode' not in df.columns:
    df.rename(columns={'Kode Pelanggan': 'Kode'}, inplace=True)

kolom_diambil = ['Kode', 'Nama Pelanggan', 'Umur JT', 'No. Faktur', 'Tgl Faktur', 'Nilai Faktur', 'Sisa Piutang']
df = df[kolom_diambil].copy()

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

data_terproses = []

for item in daftar_instruksi:
    sales = item['sales']
    kode = item['kode']
    
    if kode.upper() == 'KOSONG':
        baris_kosong = {
            'Penagih': sales,
            'Kode': '',
            'Nama Pelanggan': '',
            'Umur JT': None,
            'No. Faktur': '',
            'Tgl Faktur': '',
            'Nilai Faktur': 0,
            'Terbayar': 0,
            'Sisa Piutang': 0,
            'Is_Kosong': True
        }
        data_terproses.append(baris_kosong)
    else:
        df_match = df[df['Kode'] == kode].copy()
        for _, row in df_match.iterrows():
            row_dict = row.to_dict()
            row_dict['Penagih'] = sales
            row_dict['Is_Kosong'] = False
            data_terproses.append(row_dict)

df_hasil_raw = pd.DataFrame(data_terproses)

KAPASITAS_HALAMAN = 17
data_final = []

for sales_name, group_data in df_hasil_raw.groupby('Penagih', sort=False):
    total_rows = len(group_data)
    nomor_urut = 1
    
    sub_nilai = group_data['Nilai Faktur'].sum()
    sub_terbayar = group_data['Terbayar'].sum()
    sub_sisa = group_data['Sisa Piutang'].sum()
    
    for i in range(0, total_rows, KAPASITAS_HALAMAN):
        chunk = group_data.iloc[i : i + KAPASITAS_HALAMAN]
        halaman_ke = (i // KAPASITAS_HALAMAN) + 1
        
        for _, row in chunk.iterrows():
            r = row.to_dict()
            r['No'] = nomor_urut
            r['Halaman'] = f"Halaman {halaman_ke}"
            
            if r.get('Is_Kosong'):
                r['Nilai Faktur'] = None
                r['Terbayar'] = None
                r['Sisa Piutang'] = None
            else:
                r['Terbayar'] = r['Terbayar'] if r['Terbayar'] != 0 else None
                
            data_final.append(r)
            nomor_urut += 1

    total_halaman_sales = ((total_rows - 1) // KAPASITAS_HALAMAN) + 1
    baris_total = {
        'Halaman': f"Halaman {total_halaman_sales}",
        'No': None,
        'Penagih': sales_name,
        'Kode': None,
        'Nama Pelanggan': None,
        'Umur JT': None,
        'No. Faktur': None,
        'Tgl Faktur': f"TOTAL {sales_name}",
        'Nilai Faktur': sub_nilai,
        'Terbayar': sub_terbayar if sub_terbayar != 0 else None,
        'Sisa Piutang': sub_sisa
    }
    data_final.append(baris_total)

    baris_pemisah = {col: None for col in ['Halaman', 'No', 'Penagih', 'Kode', 'Nama Pelanggan', 'Umur JT', 'No. Faktur', 'Tgl Faktur', 'Nilai Faktur', 'Terbayar', 'Sisa Piutang']}
    data_final.append(baris_pemisah)

df_final = pd.DataFrame(data_final)

kolom_output = ['Halaman', 'No', 'Penagih', 'Kode', 'Nama Pelanggan', 'Umur JT', 'No. Faktur', 'Tgl Faktur', 'Nilai Faktur', 'Terbayar', 'Sisa Piutang']
df_final = df_final[kolom_output]

writer = pd.ExcelWriter('Laporan_Piutang_Penagih_temp.xlsx', engine='xlsxwriter')
df_final.to_excel(writer, index=False, sheet_name='Sheet1')

workbook = writer.book
worksheet = writer.sheets['Sheet1']
format_uang = workbook.add_format({'num_format': '#,##0.00'})

for i, col in enumerate(df_final.columns):
    panjang_maksimal = df_final[col].apply(lambda x: len(str(x))).max()
    column_len = max(panjang_maksimal, len(col)) + 2
    
    if col in ['Nilai Faktur', 'Terbayar', 'Sisa Piutang']:
        worksheet.set_column(i, i, column_len, format_uang)
    else:
        worksheet.set_column(i, i, column_len)

writer.close()
print("--> Proses berhasil! Data lengkap dengan baris kosong, paginasi 17 baris, dan Rangkuman Total.")