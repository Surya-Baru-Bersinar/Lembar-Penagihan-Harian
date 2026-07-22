Sub CetakLaporanARPerBlok()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long, c As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim printerDipilih As Boolean
    Dim jumlahTerhitung As Long
    Dim teksSel As String
    Dim adaPembuka As Boolean, adaPenutup As Boolean
    
    Set ws = ActiveSheet
    
    printerDipilih = Application.Dialogs(xlDialogPrinterSetup).Show
    If Not printerDipilih Then
        MsgBox "Proses pencetakan dibatalkan.", vbInformation, "Batal"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ws.ResetAllPageBreaks
    
    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperLetter
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    
    lastRow = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1
    startRow = 0
    jumlahTerhitung = 0
    
    For r = 1 To lastRow
        adaPembuka = False
        adaPenutup = False
        
        For c = 1 To 16
            teksSel = CStr(ws.Cells(r, c).Value)
            
            If InStr(1, teksSel, "LAPORAN HASIL TAGIHAN", vbTextCompare) > 0 Then
                adaPembuka = True
            End If
            
            If InStr(1, teksSel, "TTD SALES & COLLECTOR", vbTextCompare) > 0 Then
                adaPenutup = True
            End If
        Next c
        
        If adaPembuka And startRow = 0 Then
            startRow = r
        End If
        
        If adaPenutup And startRow > 0 Then
            endRow = r
            
            ws.PageSetup.PrintArea = "A" & startRow & ":P" & endRow
            ws.PrintOut From:=1, To:=1, Copies:=1
            
            jumlahTerhitung = jumlahTerhitung + 1
            startRow = 0 '
        End If
    Next r
    
    Application.ScreenUpdating = True
    
    If jumlahTerhitung > 0 Then
        MsgBox "Selesai! Total ada " & jumlahTerhitung & " kelompok laporan Penagih yang dicetak.", vbInformation, "Sukses"
    Else
        MsgBox "Tidak ditemukan blok data dengan kata kunci yang sesuai.", vbExclamation, "Peringatan"
    End If
End Sub
