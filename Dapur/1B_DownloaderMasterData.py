import configparser
import requests
import os

def download_sheet(url, output_name):
    if not url or str(url).strip() == "":
        print(f"--> Peringatan: URL untuk {output_name} kosong di piutang.conf! Proses unduh dilewati.")
        return

    if '/d/' in url:
        base_id = url.split('/d/')[1].split('/')[0]
        download_url = f"https://docs.google.com/spreadsheets/d/{base_id}/export?format=xlsx"
    else:
        print(f"--> URL tidak valid untuk {output_name}: {url}")
        return

    try:
        print(f"--> Sedang mengunduh {output_name}...")
        response = requests.get(download_url)
        response.raise_for_status()
        with open(output_name, 'wb') as f:
            f.write(response.content)
        print(f"--> Berhasil! File disimpan sebagai {output_name}")
    except Exception as e:
        print(f"--> Gagal mengunduh {output_name}: {e}")

def main():
    piutang = configparser.ConfigParser(allow_no_value=True, strict=False)
    piutang_file = 'piutang.conf'
    
    if not os.path.exists(piutang_file):
        print("--> File piutang.conf tidak ditemukan!")
        return

    piutang.read(piutang_file)

    mapping = {
        'MASTER': 'Master_temp.xlsx'
    }

    for section, output_name in mapping.items():
        if piutang.has_section(section):
            masdatus = piutang.get(section, 'masdatus', fallback='')
            
            if str(masdatus).strip().lower() == 'ya':
                url = piutang.get(section, 'url', fallback='')
                download_sheet(url, output_name)
            else:
                print(f"--> Status masdatus pada [{section}] bernilai '{masdatus}'. Proses untuk {output_name} dilewati.")
        else:
            print(f"--> Section [{section}] tidak ditemukan di piutang.conf")

if __name__ == "__main__":
    main()