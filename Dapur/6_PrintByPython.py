import configparser
import ctypes
import os
import tkinter as tk
from tkinter import ttk
import win32print
import xlwings as xw

def cek_status_pprint():
    nama_conf = "piutang.conf"
    path_conf = os.path.abspath(nama_conf)

    if not os.path.exists(path_conf):
        print(f"--> Peringatan: File '{nama_conf}' tidak ditemukan!")
        return False

    config = configparser.ConfigParser(allow_no_value=True, strict=False)
    try:
        config.read(path_conf, encoding="utf-8")

        if "PPRINT" in config and "status" in config["PPRINT"]:
            status_val = config["PPRINT"]["status"].strip().upper()

            return status_val == "YA"
        else:
            print(
                "--> Peringatan: Section [PPRINT] atau key 'status' tidak ditemukan di piutang.conf."
            )
            return False

    except Exception as e:
        print(f"--> Gagal membaca file konfigurasi: {e}")
        return False

def tampilkan_dialog_printer():
    daftar_printer = [
        p[2]
        for p in win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        )
    ]
    printer_default = win32print.GetDefaultPrinter()

    printer_terpilih = [None]

    root = tk.Tk()
    root.title("Pilih Printer")
    root.geometry("380x160")
    root.resizable(False, False)
    root.attributes("-topmost", True)

    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (380 // 2)
    y = (root.winfo_screenheight() // 2) - (160 // 2)
    root.geometry(f"380x160+{x}+{y}")

    label = ttk.Label(
        root, text="Pilih Printer untuk Mencetak:", font=("Segoe UI", 10)
    )
    label.pack(pady=(15, 5))

    combo = ttk.Combobox(
        root, values=daftar_printer, state="readonly", width=40
    )
    if printer_default in daftar_printer:
        combo.set(printer_default)
    elif daftar_printer:
        combo.set(daftar_printer[0])
    combo.pack(pady=5)

    def aksi_ok():
        printer_terpilih[0] = combo.get()
        root.destroy()

    def aksi_batal():
        root.destroy()

    frame_tombol = ttk.Frame(root)
    frame_tombol.pack(pady=15)

    btn_ok = ttk.Button(frame_tombol, text="OK", command=aksi_ok, width=10)
    btn_ok.pack(side=tk.LEFT, padx=5)

    btn_cancel = ttk.Button(
        frame_tombol, text="Cancel", command=aksi_batal, width=10
    )
    btn_cancel.pack(side=tk.LEFT, padx=5)

    root.mainloop()

    return printer_terpilih[0]


def cetak_laporan_ar_xlwings():
    if not cek_status_pprint():
        msg_skip = (
            "Status PPRINT = No. Proses dicetak DILEWATI."
        )
        print(f"--> {msg_skip}")
        return

    nama_file = "Print_AR.xlsm"
    path_file = os.path.abspath(nama_file)

    if not os.path.exists(path_file):
        msg = f"File '{nama_file}' tidak ditemukan di folder!"
        print(f"--> {msg}")
        ctypes.windll.user32.MessageBoxW(0, msg, "Peringatan", 48)
        return

    printer_dipilih = tampilkan_dialog_printer()

    if not printer_dipilih:
        msg = "Proses pencetakan dibatalkan."
        print(f"--> {msg}")
        ctypes.windll.user32.MessageBoxW(0, msg, "Batal", 64)
        return

    app = xw.App(visible=False, add_book=False)

    try:
        wb = app.books.open(path_file)
        ws = wb.sheets.active

        app.screen_updating = False

        try:
            app.api.ActivePrinter = printer_dipilih
        except Exception:
            pass

        ws.api.ResetAllPageBreaks()

        ps = ws.api.PageSetup

        try:
            ps.Orientation = 2
        except Exception:
            pass

        try:
            ps.Zoom = False
        except Exception:
            pass

        try:
            ps.FitToPagesWide = 1
        except Exception:
            pass

        try:
            ps.FitToPagesTall = 1
        except Exception:
            pass

        try:
            ps.LeftMargin = app.api.InchesToPoints(0.25)
            ps.RightMargin = app.api.InchesToPoints(0.25)
            ps.TopMargin = app.api.InchesToPoints(0.25)
            ps.BottomMargin = app.api.InchesToPoints(0.25)
        except Exception:
            pass

        try:
            ps.PaperSize = 2
        except Exception:
            pass

        last_row = ws.used_range.last_cell.row
        start_row = 0
        jumlah_terhitung = 0

        for r in range(1, last_row + 1):
            ada_pembuka = False
            ada_penutup = False

            for c in range(2, 17):
                nilai_sel = str(ws.cells(r, c).value or "").upper()
                if "LAPORAN HASIL TAGIHAN" in nilai_sel:
                    ada_pembuka = True
                if "TTD SALES & COLLECTOR" in nilai_sel:
                    ada_penutup = True

            if ada_pembuka and start_row == 0:
                start_row = r

            if ada_penutup and start_row > 0:
                end_row = r

                ps.PrintArea = f"B{start_row}:P{end_row}"

                ws.api.PrintOut(
                    From=1, To=1, Copies=1, ActivePrinter=printer_dipilih
                )

                jumlah_terhitung += 1
                start_row = 0

        app.screen_updating = True

        if jumlah_terhitung > 0:
            msg = f"Selesai! Total ada {jumlah_terhitung} kelompok laporan yang dicetak."
            print(f"--> {msg}")
            ctypes.windll.user32.MessageBoxW(0, msg, "Sukses", 64)
        else:
            msg = "Tidak ditemukan blok data dengan kata kunci yang sesuai."
            print(f"--> {msg}")
            ctypes.windll.user32.MessageBoxW(0, msg, "Peringatan", 48)

        wb.close()

    finally:
        app.quit()


if __name__ == "__main__":
    cetak_laporan_ar_xlwings()