Attribute VB_Name = "ModulePrint"
Option Explicit

Sub PrintTotalBarangMasuk()
    Call SetWorksheets
    
    Dim saveLocation As String
    Dim fileName As String
    
    fileName = "Laporan-Total-Barang-Masuk_" & Format(Now, "DD-MM-YYYY_HH-MM_") & ".pdf"
    saveLocation = getPath("\Laporan Data\Total Barang Masuk\") + fileName
    
    ' Set Header
    Dim headerText As String
    headerText = "Toko Alat Olahraga - Laporan Total Barang Masuk"
    
    Dim wsToExport As Worksheet
    Set wsToExport = wsDataModel
    
    Dim targetExportRange As Range
    Set targetExportRange = wsToExport.Range("B2").CurrentRegion
    
    Dim answer As Integer
    answer = MsgBox("Apakah anda yakin ingin membuatnya?", vbQuestion + vbYesNo + vbDefaultButton2, "Buat Laporan Baru Total Barang Masuk")
    
    If answer = vbYes Then
        ' Set Page Setup Options
        With wsToExport.PageSetup
            .LeftHeader = "&""Arial,Bold""&12" & headerText
            .RightHeader = "&""Arial,Regular""&12" & convertHariIndonesia(Format(Now, "DDDD")) & ", " & Format(Now, "DD ") & convertBulanIndonesia(Format(Now, "DD/MM/YYYY")) & " " & Format(Now, "YYYY - HH::MM")
            .CenterHorizontally = True
            .Orientation = xlPortrait
            .FitToPagesWide = 1
            .Zoom = False
        End With
    
        targetExportRange.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            fileName:=saveLocation, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False
'            IgnorePrintAreas:=False, _
'            OpenAfterPublish:=True

        MsgBox """" & fileName & """ Berhasil dibuat", vbInformation
    End If
End Sub

Sub PrintTotalPenjualanBarang()
    Call SetWorksheets
    
    Dim saveLocation As String
    Dim fileName As String
    
    fileName = "Laporan-Total-Penjualan-Barang_" & Format(Now, "DD-MM-YYYY_HH-MM_") & ".pdf"
    saveLocation = getPath("\Laporan Data\Total Penjualan Barang\") + fileName
    
    ' Set Header
    Dim headerText As String
    headerText = "Toko Alat Olahraga - Laporan Total Penjualan Barang"
    
    Dim wsToExport As Worksheet
    Set wsToExport = wsDataModel
    
    Dim targetExportRange As Range
    Set targetExportRange = wsToExport.Range("G2").CurrentRegion
    
    Dim answer As Integer
    answer = MsgBox("Apakah anda yakin ingin membuatnya?", vbQuestion + vbYesNo + vbDefaultButton2, "Buat Laporan Baru Total Penjualan Barang")
    
    If answer = vbYes Then
        ' Set Page Setup Options
        With wsToExport.PageSetup
            .LeftHeader = "&""Arial,Bold""&12" & headerText
            .RightHeader = "&""Arial,Regular""&12" & convertHariIndonesia(Format(Now, "DDDD")) & ", " & Format(Now, "DD ") & convertBulanIndonesia(Format(Now, "DD/MM/YYYY")) & " " & Format(Now, "YYYY - HH::MM")
            .CenterHorizontally = True
            .Orientation = xlPortrait
            .FitToPagesWide = 1
            .Zoom = False
        End With
    
        targetExportRange.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            fileName:=saveLocation, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False
'            IgnorePrintAreas:=False, _
'            OpenAfterPublish:=True

        MsgBox """" & fileName & """ Berhasil dibuat", vbInformation
    End If
End Sub

Sub PrintTotalHargaBeli()
    Call SetWorksheets
    
    Dim saveLocation As String
    Dim fileName As String
    
    fileName = "Laporan-Total-Harga-Beli_" & Format(Now, "DD-MM-YYYY_HH-MM_") & ".pdf"
    saveLocation = getPath("\Laporan Data\Total Harga Beli\") + fileName
    
    ' Set Header
    Dim headerText As String
    headerText = "Toko Alat Olahraga - Laporan Total Harga Beli"
    
    Dim wsToExport As Worksheet
    Set wsToExport = wsDataModel
    
    Dim targetExportRange As Range
    Set targetExportRange = wsToExport.Range("S2").CurrentRegion
    
    Dim answer As Integer
    answer = MsgBox("Apakah anda yakin ingin membuatnya?", vbQuestion + vbYesNo + vbDefaultButton2, "Buat Laporan Baru Total Harga Beli")
    
    If answer = vbYes Then
        ' Set Page Setup Options
        With wsToExport.PageSetup
            .LeftHeader = "&""Arial,Bold""&12" & headerText
            .RightHeader = "&""Arial,Regular""&12" & convertHariIndonesia(Format(Now, "DDDD")) & ", " & Format(Now, "DD ") & convertBulanIndonesia(Format(Now, "DD/MM/YYYY")) & " " & Format(Now, "YYYY - HH::MM")
            .CenterHorizontally = True
            .Orientation = xlPortrait
            .FitToPagesWide = 1
            .Zoom = False
        End With
    
        targetExportRange.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            fileName:=saveLocation, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False
'            IgnorePrintAreas:=False, _
'            OpenAfterPublish:=True

        MsgBox """" & fileName & """ Berhasil dibuat", vbInformation
    End If
End Sub

Sub PrintTotalHargaJual()
    Call SetWorksheets
    
    Dim saveLocation As String
    Dim fileName As String
    
    fileName = "Laporan-Total-Harga-Jual_" & Format(Now, "DD-MM-YYYY_HH-MM_") & ".pdf"
    saveLocation = getPath("\Laporan Data\Total Harga Jual\") + fileName
    
    ' Set Header
    Dim headerText As String
    headerText = "Toko Alat Olahraga - Laporan Total Harga Jual"
    
    Dim wsToExport As Worksheet
    Set wsToExport = wsDataModel
    
    Dim targetExportRange As Range
    Set targetExportRange = wsToExport.Range("W2").CurrentRegion
    
    Dim answer As Integer
    answer = MsgBox("Apakah anda yakin ingin membuatnya?", vbQuestion + vbYesNo + vbDefaultButton2, "Buat Laporan Baru Total Harga Jual")
    
    If answer = vbYes Then
        ' Set Page Setup Options
        With wsToExport.PageSetup
            .LeftHeader = "&""Arial,Bold""&12" & headerText
            .RightHeader = "&""Arial,Regular""&12" & convertHariIndonesia(Format(Now, "DDDD")) & ", " & Format(Now, "DD ") & convertBulanIndonesia(Format(Now, "DD/MM/YYYY")) & " " & Format(Now, "YYYY - HH::MM")
            .CenterHorizontally = True
            .Orientation = xlPortrait
            .FitToPagesWide = 1
            .Zoom = False
        End With
    
        targetExportRange.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            fileName:=saveLocation, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False
'            IgnorePrintAreas:=False, _
'            OpenAfterPublish:=True

        MsgBox """" & fileName & """ Berhasil dibuat", vbInformation
    End If
End Sub

Sub PrintTotalKeuntungan()
    Call SetWorksheets
    
    Dim saveLocation As String
    Dim fileName As String
    
    fileName = "Laporan-Total-Keuntungan_" & Format(Now, "DD-MM-YYYY_HH-MM_") & ".pdf"
    saveLocation = getPath("\Laporan Data\Total Keuntungan\") + fileName
    
    ' Set Header
    Dim headerText As String
    headerText = "Toko Alat Olahraga - Laporan Total Keuntungan"
    
    Dim wsToExport As Worksheet
    Set wsToExport = wsDataModel
    
    Dim targetExportRange As Range
    Set targetExportRange = wsToExport.Range("AA2").CurrentRegion
    
    Dim answer As Integer
    answer = MsgBox("Apakah anda yakin ingin membuatnya?", vbQuestion + vbYesNo + vbDefaultButton2, "Buat Laporan Baru Total Keuntungan")
    
    If answer = vbYes Then
        ' Set Page Setup Options
        With wsToExport.PageSetup
            .LeftHeader = "&""Arial,Bold""&12" & headerText
            .RightHeader = "&""Arial,Regular""&12" & convertHariIndonesia(Format(Now, "DDDD")) & ", " & Format(Now, "DD ") & convertBulanIndonesia(Format(Now, "DD/MM/YYYY")) & " " & Format(Now, "YYYY - HH::MM")
            .CenterHorizontally = True
            .Orientation = xlPortrait
            .FitToPagesWide = 1
            .Zoom = False
        End With
    
        targetExportRange.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            fileName:=saveLocation, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False
'            IgnorePrintAreas:=False, _
'            OpenAfterPublish:=True

        MsgBox """" & fileName & """ Berhasil dibuat", vbInformation
    End If
End Sub
