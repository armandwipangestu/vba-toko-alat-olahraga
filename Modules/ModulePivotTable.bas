Attribute VB_Name = "ModulePivotTable"
Option Explicit

Sub RefreshPivotTable()
    'Sheets("Data Model").Activate
    
    Dim Pivot1 As PivotTable
    Dim Pivot2 As PivotTable
    Dim Pivot3 As PivotTable
    Dim Pivot4 As PivotTable
    Dim Pivot5 As PivotTable
    Dim Pivot6 As PivotTable
    Dim Pivot7 As PivotTable

    Set Pivot1 = wsDataModel.PivotTables("PivotBulanBarangMasuk")
    Set Pivot2 = wsDataModel.PivotTables("PivotBulanPenjualanBarang")
    Set Pivot3 = wsDataModel.PivotTables("PivotMerekBarangMasuk")
    Set Pivot4 = wsDataModel.PivotTables("PivotMerekPenjualanBarang")
    Set Pivot5 = wsDataModel.PivotTables("PivotMerekTotalPembelian")
    Set Pivot6 = wsDataModel.PivotTables("PivotMerekTotalPenjualan")
    Set Pivot7 = wsDataModel.PivotTables("PivotBulanTotalKeuntungan")
    
    'Melakukan refresh data pada PivotTable
    Pivot1.RefreshTable
    Pivot2.RefreshTable
    Pivot3.RefreshTable
    Pivot4.RefreshTable
    Pivot5.RefreshTable
    Pivot6.RefreshTable
    Pivot7.RefreshTable
End Sub

Sub LegacyPrintTotalPembelian()
    Call SetWorksheets
    Dim PivotPembelian As PivotTable
    Set PivotPembelian = wsDataModel.PivotTables("PivotMerekTotalPembelian")
    'PivotPembelian.TableRange1.PrintOut
    
    makeDirectory
    Dim saveLocation As String
    Dim fileName As String
    
    fileName = "\Laporan Total Pembelian_" & Format(Now, "DD-MM-YYYY_HH-MM_") & ".pdf"
    saveLocation = getPath("\Laporan Data\Total Pembelian") + fileName
    'saveLocation = getPath & "\Laporan Data\Total Pembelian\Laporan Total Pembelian_" & Format(Now, "DD-MM-YYYY_HH-MM_") & ".pdf"
    
    ' Set Header
    Dim headerText As String
    headerText = "Laporan Total Pembelian"
    
    Dim wsToExport As Worksheet
    Set wsToExport = wsDataModel
    
    Dim targetExportRange As Range
    'Set targetExportRange = PivotPembelian.TableRange1 'range includes the page field
    Set targetExportRange = wsToExport.Range("N2").CurrentRegion 'wsToExport.Range("N2:O8")
    
    Dim answer As Integer
    answer = MsgBox("Apakah anda yakin?", vbQuestion + vbYesNo + vbDefaultButton2, "Print Total Pembelian")
    
    If answer = vbYes Then
        ' Set Page Setup Options
        With wsToExport.PageSetup
            .LeftHeader = "&""Arial,Bold""&14" & headerText
            .RightHeader = "&""Arial,Regular""&14" & convertHariIndonesia(Format(Now, "DDDD")) & ", " & Format(Now, "DD MMMM YYYY")
            .CenterHorizontally = True
            '.CenterVertically = True
            .Orientation = xlPortrait
            .FitToPagesWide = 1
            '.FitToPagesTall = False
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
