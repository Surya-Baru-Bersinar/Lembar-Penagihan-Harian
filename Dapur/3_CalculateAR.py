import pandas as pd
import xlwings as xw

config = {}
current_key = None
try:
    with open('piutang.conf', 'r') as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            if line.startswith('[') and line.endswith(']'):
                current_key = line[1:-1]
            else:
                if current_key and current_key not in config:
                    config[current_key] = line
except Exception as e:
    print(f"--> Gagal membaca piutang.conf: {e}")

df = pd.read_excel('Laporan_Piutang_Penagih_temp.xlsx')

df = df[df['Penagih'].notna()]
df = df[~df['Penagih'].astype(str).str.startswith('TOTAL', na=False)]
df = df[df['Kode'].notna()]
df = df[df['Kode'].astype(str).str.strip() != '']

app = xw.App(visible=False)

try:
    wb = app.books.open('TEMPLATE.xlsm')
    
    ws_temp = wb.sheets.active
    ws_temp.name = "TEMP_DESIGN"
    
    max_row_temp = ws_temp.range('A' + str(ws_temp.cells.last_cell.row)).end('up').row
    if max_row_temp < 7:
        max_row_temp = 7

    ws_out = wb.sheets.add(name="Print AR", after=ws_temp)

    current_out_row = 1
    groups = df.groupby('Penagih', sort=False)

    for penagih, group_df in groups:
        start_row_for_group = current_out_row
        
        ws_temp.range('1:4').copy(ws_out.range(f'{current_out_row}:{current_out_row + 3}'))
        
        ws_out.range((start_row_for_group + 1, 4)).value = penagih
        ws_out.range((start_row_for_group + 1, 8)).value = config.get('PERUSAHAAN', '')
        ws_out.range((start_row_for_group + 1, 11)).value = config.get('DIVISI', '')
        ws_out.range((start_row_for_group + 1, 15)).value = config.get('TANGGAL', '')
        ws_out.range((start_row_for_group + 2, 15)).value = config.get('INPUT', '')
        
        current_out_row += 4
        data_start_row = current_out_row
        nomor = 1
        
        for _, row_data in group_df.iterrows():
            ws_temp.range('5:5').copy(ws_out.range(f'{current_out_row}:{current_out_row}'))
            
            ws_out.range((current_out_row, 2)).value = nomor
            ws_out.range((current_out_row, 3)).value = row_data.get('Kode', '')
            ws_out.range((current_out_row, 4)).value = row_data.get('Nama Pelanggan', '')
            ws_out.range((current_out_row, 5)).value = row_data.get('Umur JT', '')
            ws_out.range((current_out_row, 6)).value = row_data.get('No. Faktur', '')
            ws_out.range((current_out_row, 7)).value = row_data.get('Tgl Faktur', '')
            
            val_faktur = row_data.get('Nilai Faktur', None)
            val_terbayar = row_data.get('Terbayar', None)
            val_sisa = row_data.get('Sisa Piutang', None)
            
            ws_out.range((current_out_row, 8)).value = val_faktur if pd.notna(val_faktur) and val_faktur != "" else None
            ws_out.range((current_out_row, 9)).value = val_terbayar if pd.notna(val_terbayar) and val_terbayar != "" else None
            ws_out.range((current_out_row, 10)).value = val_sisa if pd.notna(val_sisa) and val_sisa != "" else None

            current_out_row += 1
            nomor += 1
            
        data_end_row = current_out_row - 1
        
        ws_temp.range('6:6').copy(ws_out.range(f'{current_out_row}:{current_out_row}'))
        
        ws_out.range((current_out_row, 2)).value = "TOTAL TAGIHAN"
        ws_out.range((current_out_row, 8)).formula = f"=SUM(H{data_start_row}:H{data_end_row})"
        ws_out.range((current_out_row, 9)).formula = f"=SUM(I{data_start_row}:I{data_end_row})"
        ws_out.range((current_out_row, 10)).formula = f"=SUM(J{data_start_row}:J{data_end_row})"
        
        current_out_row += 1
        
        if max_row_temp >= 7:
            jumlah_baris_footer = max_row_temp - 7 + 1
            ws_temp.range(f'7:{max_row_temp}').copy(ws_out.range(f'{current_out_row}:{current_out_row + jumlah_baris_footer - 1}'))
            current_out_row += jumlah_baris_footer
            
        current_out_row += 2

    ws_out.autofit('c')
    
    lebar_spesifik = {
        'F': 30,
        'G': 30,
        'H': 37,        
        'I': 37,
        'J': 37,
        'K': 28,
        'L': 34,
        'M': 28,
        'N': 37,
        'O': 37,
        'P': 30
    }
    for col_letter, width in lebar_spesifik.items():
        ws_out.range(f'{col_letter}1').column_width = width

    for r in range(1, current_out_row + 1):
        vals = ws_out.range((r, 1), (r, 16)).value
        if vals:
            for val in vals:
                if isinstance(val, str) and "TTD SALES & COLLECTOR" in val:
                    ws_out.range(f'{r}:{r}').row_height = 115
                    break

    if len(ws_temp.shapes) > 0:
        for shape in ws_temp.shapes:
            orig_top = shape.top
            orig_left = shape.left
            
            shape.api.Copy()
            ws_out.activate()
            ws_out.api.Paste()
            
            new_shape = ws_out.shapes[-1]
            new_shape.top = orig_top
            new_shape.left = orig_left

    app.display_alerts = False
    ws_temp.delete()
    app.display_alerts = True

    wb.save('Print_AR.xlsm')
    wb.close()
    print("--> Proses selesai, file telah disimpan sebagai Print_AR.xlsm")

finally:
    app.quit()