import os
import win32com.client
import win32print

printers = win32print.EnumPrinters(
    win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
)
print("--> Daftar Printer Tersedia:")
for i, printer in enumerate(printers):
    print(f"--> {i + 1}. {printer[2]}")

choice = int(input("--> Masukkan nomor printer yang ingin digunakan: ")) - 1
chosen_printer = printers[choice][2]

old_printer = win32print.GetDefaultPrinter()
win32print.SetDefaultPrinter(chosen_printer)

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

file_path = os.path.abspath("Print_AR.xlsx")
workbook = excel.Workbooks.Open(file_path)
sheet = workbook.Worksheets(1)

max_row = sheet.UsedRange.Rows.Count
current_row = 1

while current_row <= max_row:
    while (
        current_row <= max_row
        and sheet.Cells(current_row, 2).Value is None
    ):
        current_row += 1

    if current_row > max_row:
        break

    start_row = current_row

    while (
        current_row <= max_row
        and sheet.Cells(current_row, 2).Value is not None
    ):
        current_row += 1

    end_row = current_row - 1

    sheet.PageSetup.PrintArea = f"B{start_row}:P{end_row}"
    sheet.PageSetup.Orientation = 2
    sheet.PageSetup.Zoom = False
    sheet.PageSetup.FitToPagesWide = 1
    sheet.PageSetup.FitToPagesTall = 1

    sheet.PrintOut()

workbook.Close(False)
excel.Quit()

win32print.SetDefaultPrinter(old_printer)
print("--> Proses cetak selesai.")
