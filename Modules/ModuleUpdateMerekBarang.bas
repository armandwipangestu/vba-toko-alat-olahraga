Attribute VB_Name = "ModuleUpdateMerekBarang"
Option Explicit

Sub updateMerekMasterBarang(idMerekBarang As String, merekBarang As String)
    ' Master Barang
    ' Call SetWorksheets
    Dim lastRow As Long
    Dim i As Long
    lastRow = wsMasterBarang.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        If wsMasterBarang.Cells(i, 3).Value = idMerekBarang Then
            wsMasterBarang.Cells(i, 4).Value = merekBarang
        End If
    Next i
End Sub
