from datetime import datetime
import sys
import time
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

pure_setting = config.get('PURE', '')
if '=' in pure_setting:
    key_name, val_name = pure_setting.split('=', 1)
    if key_name.strip().lower() == 'pr_process' and val_name.strip().lower() == 'ya':
        print("--> Proses di-skip karena nilai pr_process = Ya pada piutang.conf.")
        sys.exit(0)
tgl_config_str = config.get('TANGGAL', '')
tgl_config_val = tgl_config_str

if tgl_config_str:
    try:
        tgl_config_val = datetime.strptime(tgl_config_str, '%d/%m/%Y').date()
    except ValueError:
        tgl_config_val = tgl_config_str

df = pd.read_excel('Laporan_Piutang_Penagih_temp.xlsx')

if 'Tgl Faktur' in df.columns:
    df['Tgl Faktur'] = pd.to_datetime(df['Tgl Faktur'], dayfirst=True, errors='coerce')

df_data = df[df['No'].notna()].copy()

app = xw.App(visible=False)

try:
    app.display_alerts = False
    app.screen_updating = False

    wb = app.books.open('TEMPLATE.xlsm')
    
    ws_temp = wb.sheets.active
    ws_temp.name = "TEMP_DESIGN"
    
    max_row_temp = ws_temp.range('A' + str(ws_temp.cells.last_cell.row)).end('up').row
    if max_row_temp < 7:
        max_row_temp = 7

    ws_out = wb.sheets.add(name="Print AR", after=ws_temp)

    current_out_row = 1

    if 'Halaman' in df_data.columns:
        groups = df_data.groupby(['Penagih', 'Halaman'], sort=False)
    else:
        groups = df_data.groupby('Penagih', sort=False)

    for group_keys, group_df in groups:
        penagih = group_keys[0] if isinstance(group_keys, tuple) else group_keys
        
        start_row_for_group = current_out_row
        
        ws_temp.range('1:4').copy(ws_out.range(f'{current_out_row}:{current_out_row + 3}'))
        
        ws_out.range((start_row_for_group + 1, 4)).value = penagih
        ws_out.range((start_row_for_group + 1, 8)).value = config.get('PERUSAHAAN', '')
        ws_out.range((start_row_for_group + 1, 11)).value = config.get('DIVISI', '')
        
        ws_out.range((start_row_for_group + 1, 15)).value = tgl_config_val
        
        ws_out.range((start_row_for_group + 2, 15)).value = config.get('INPUT', '')
        
        current_out_row += 4
        data_start_row = current_out_row
        
        for _, row_data in group_df.iterrows():
            ws_temp.range('5:5').copy(ws_out.range(f'{current_out_row}:{current_out_row}'))
            
            ws_out.range((current_out_row, 2)).value = int(row_data['No']) if pd.notna(row_data['No']) else ''
            ws_out.range((current_out_row, 3)).value = row_data.get('Kode', '') if pd.notna(row_data.get('Kode')) else ''
            ws_out.range((current_out_row, 4)).value = row_data.get('Nama Pelanggan', '') if pd.notna(row_data.get('Nama Pelanggan')) else ''
            ws_out.range((current_out_row, 5)).value = row_data.get('Umur JT', '') if pd.notna(row_data.get('Umur JT')) else ''
            ws_out.range((current_out_row, 6)).value = row_data.get('No. Faktur', '') if pd.notna(row_data.get('No. Faktur')) else ''
            
            tgl_faktur = row_data.get('Tgl Faktur', None)
            if pd.notna(tgl_faktur):
                ws_out.range((current_out_row, 7)).value = tgl_faktur
            else:
                ws_out.range((current_out_row, 7)).value = ''
            
            val_faktur = row_data.get('Nilai Faktur', None)
            val_terbayar = row_data.get('Terbayar', None)
            val_sisa = row_data.get('Sisa Piutang', None)
            
            ws_out.range((current_out_row, 8)).value = val_faktur if pd.notna(val_faktur) and val_faktur != "" else None
            ws_out.range((current_out_row, 9)).value = val_terbayar if pd.notna(val_terbayar) and val_terbayar != "" else None
            ws_out.range((current_out_row, 10)).value = val_sisa if pd.notna(val_sisa) and val_sisa != "" else None

            current_out_row += 1
            
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
        'D': 75, 'E': 15, 'F': 30, 'G': 32, 'H': 35, 'I': 35, 'J': 35,
        'K': 25, 'L': 35, 'M': 15, 'N': 37, 'O': 37, 'P': 30
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

    app.api.CutCopyMode = False

    if len(ws_temp.shapes) > 0:
        for shape in ws_temp.shapes:
            try:
                orig_top = shape.top
                orig_left = shape.left
                
                macro_action = None
                try:
                    macro_action = shape.api.OnAction
                except Exception:
                    pass

                shape.api.Copy()
                time.sleep(0.15)
                ws_out.activate()
                ws_out.api.Paste()
                time.sleep(0.15)
                
                new_shape = ws_out.shapes[-1]
                new_shape.top = orig_top
                new_shape.left = orig_left
                
                if macro_action:
                    try:
                        new_shape.api.OnAction = macro_action
                    except Exception:
                        pass

            except Exception as e:
                print(f"--> Gagal menyalin shape/tombol: {e}")

    app.api.CutCopyMode = False
    time.sleep(0.2)

    try:
        ws_temp.delete()
    except Exception as e:
        print(f"--> Gagal menghapus ws_temp: {e}")
        try:
            ws_temp.visible = False
        except Exception:
            pass

    app.screen_updating = True
    time.sleep(0.2)

    wb.save('Print_AR.xlsm')
    wb.close()
    print("--> Proses ekspor berhasil! File disimpan sebagai Print_AR.xlsm")

finally:
    try:
        app.screen_updating = True
        app.display_alerts = True
    except Exception:
        pass
    try:
        app.quit()
    except Exception:
        pass
    try:
        app.kill()
    except Exception:
        pass