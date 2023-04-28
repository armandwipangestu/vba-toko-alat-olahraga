Attribute VB_Name = "ModulePrint"
Option Explicit

Sub PrintTotalPembelian()
    Call SetWorksheets
    
    makeDirectory
    Dim saveLocation As String
    Dim fileName As String
    
    fileName = "\Laporan Total Pembelian_" & Format(Now, "DD-MM-YYYY_HH-MM_") & ".pdf"
    saveLocation = getPath("\Laporan Data\Total Pembelian") + fileName
    
    ' Set Header
    Dim headerText As String
    headerText = "Toko Alat Olahraga - Laporan Total Pembelian"
    
    Dim wsToExport As Worksheet
    Set wsToExport = wsDataModel
    
    Dim targetExportRange As Range
    Set targetExportRange = wsToExport.Range("N2").CurrentRegion
    
    Dim answer As Integer
    answer = MsgBox("Apakah anda yakin?", vbQuestion + vbYesNo + vbDefaultButton2, "Print Total Pembelian")
    
    If answer = vbYes Then
        ' Set Page Setup Options
        With wsToExport.PageSetup
            .LeftHeader = "&""Arial,Bold""&14" & headerText
            .RightHeader = "&""Arial,Regular""&14" & convertHariIndonesia(Format(Now, "DDDD")) & ", " & Format(Now, "DD MMMM YYYY - H:M")
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
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=True
    End If
End Sub

