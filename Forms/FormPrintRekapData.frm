VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPrintRekapData 
   Caption         =   "Form Print Rekap Data"
   ClientHeight    =   9900.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14625
   OleObjectBlob   =   "FormPrintRekapData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormPrintRekapData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    BackColor = RGB(29, 29, 66)
    LabelPilihData.BackColor = RGB(29, 29, 66)
    FrameNamaFile.BackColor = RGB(29, 29, 66)
    FrameNamaFile.ForeColor = RGB(255, 255, 255)
    FramePreview.BackColor = RGB(29, 29, 66)
    FramePreview.ForeColor = RGB(255, 255, 255)
    CmdBtnExportPdf.BackColor = RGB(37, 215, 152)
    CmdBtnHapus.BackColor = RGB(231, 21, 86)
    
    ComboBoxData.List = Array("Total Barang Masuk", "Total Penjualan Barang", _
                              "Total Harga Beli", "Total Harga Jual", _
                              "Total Keuntungan")
    ComboBoxData.Value = ComboBoxData.List(0)
End Sub

Private Sub ComboBoxData_Change()
    'MsgBox ComboBoxData.Value
    ListBox1.Clear
    viewFileName (ComboBoxData.Value)
End Sub

Private Sub viewFileName(data As String)
    makeDirectory
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim i As Integer
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    Set oFolder = oFSO.GetFolder(Application.ActiveWorkbook.Path & "\Laporan Data\" & data)

'    If data = "Total Pembelian" Then
        For Each oFile In oFolder.Files
            ListBox1.AddItem oFile.Name
        Next oFile
'    End If
End Sub

Private Sub ListBox1_Click()
    Dim x As Long
    Dim html As String
    Dim fullPath As String
    
    fullPath = getPath("\Laporan Data\")
    fullPath = fullPath + ComboBoxData.Value + "\"
    
    With ListBox1
        For x = 0 To .ListCount - 1
            If .Selected(x) Then
                DoEvents
                html = "<HTML><BODY><embed src=""" & fullPath & .List(x, 0) & """ height=""100%"" width=""100%"" /></body></html>"
                
                'Debug.Print html
                WebBrowserPdfView.Navigate "about:blank"
                Do While WebBrowserPdfView.ReadyState <> READYSTATE_COMPLETE
                    DoEvents
                Loop
                WebBrowserPdfView.Document.write html
            End If
        Next x
    End With
End Sub

Private Sub CmdBtnExportPdf_Click()
    If ComboBoxData.Value = "Total Barang Masuk" Then
        PrintTotalBarangMasuk
        ListBox1.Clear
        viewFileName (ComboBoxData.Value)
    End If
    
    If ComboBoxData.Value = "Total Penjualan Barang" Then
        PrintTotalPenjualanBarang
        ListBox1.Clear
        viewFileName (ComboBoxData.Value)
    End If

    If ComboBoxData.Value = "Total Harga Beli" Then
        PrintTotalHargaBeli
        ListBox1.Clear
        viewFileName (ComboBoxData.Value)
    End If
    
    If ComboBoxData.Value = "Total Harga Jual" Then
        PrintTotalHargaJual
        ListBox1.Clear
        viewFileName (ComboBoxData.Value)
    End If
    
    If ComboBoxData.Value = "Total Keuntungan" Then
        PrintTotalKeuntungan
        ListBox1.Clear
        viewFileName (ComboBoxData.Value)
    End If
End Sub

Private Sub CmdBtnHapus_Click()
    Dim x As Long
    Dim fullPath As String
    
    fullPath = getPath("\Laporan Data\")
    fullPath = fullPath + ComboBoxData.Value + "\"

    If ComboBoxData.Value = "Total Barang Masuk" Then
        With ListBox1
            For x = 0 To .ListCount - 1
                If .Selected(x) Then
                    DoEvents
                    Dim answerTotalBarangMasuk As Integer
                    answerTotalBarangMasuk = MsgBox("Apakah anda yakin ingin menghapus laporan ini?", vbQuestion + vbYesNo + vbDefaultButton2, .List(x, 0))
                    
                    If answerTotalBarangMasuk = vbYes Then
                        Kill (fullPath & .List(x, 0))
                        ListBox1.Clear
                        viewFileName (ComboBoxData.Value)
                        MsgBox "Laporan Berhasil Dihapus!", vbInformation
                    End If
                End If
            Next x
        End With
    End If
    
    If ComboBoxData.Value = "Total Penjualan Barang" Then
        With ListBox1
            For x = 0 To .ListCount - 1
                If .Selected(x) Then
                    DoEvents
                    Dim answerTotalPenjualanBarang As Integer
                    answerTotalPenjualanBarang = MsgBox("Apakah anda yakin ingin menghapus laporan ini?", vbQuestion + vbYesNo + vbDefaultButton2, .List(x, 0))
                    
                    If answerTotalPenjualanBarang = vbYes Then
                        Kill (fullPath & .List(x, 0))
                        ListBox1.Clear
                        viewFileName (ComboBoxData.Value)
                        MsgBox "Laporan Berhasil Dihapus!", vbInformation
                    End If
                End If
            Next x
        End With
    End If

    If ComboBoxData.Value = "Total Harga Beli" Then
        With ListBox1
            For x = 0 To .ListCount - 1
                If .Selected(x) Then
                    DoEvents
                    Dim answerTotalHargaBeli As Integer
                    answerTotalHargaBeli = MsgBox("Apakah anda yakin ingin menghapus laporan ini?", vbQuestion + vbYesNo + vbDefaultButton2, .List(x, 0))
                    
                    If answerTotalHargaBeli = vbYes Then
                        Kill (fullPath & .List(x, 0))
                        ListBox1.Clear
                        viewFileName (ComboBoxData.Value)
                        MsgBox "Laporan Berhasil Dihapus!", vbInformation
                    End If
                End If
            Next x
        End With
    End If
    
    If ComboBoxData.Value = "Total Harga Jual" Then
        With ListBox1
            For x = 0 To .ListCount - 1
                If .Selected(x) Then
                    DoEvents
                    Dim answerTotalHargaJual As Integer
                    answerTotalHargaJual = MsgBox("Apakah anda yakin ingin menghapus laporan ini?", vbQuestion + vbYesNo + vbDefaultButton2, .List(x, 0))
                    
                    If answerTotalHargaJual = vbYes Then
                        Kill (fullPath & .List(x, 0))
                        ListBox1.Clear
                        viewFileName (ComboBoxData.Value)
                        MsgBox "Laporan Berhasil Dihapus!", vbInformation
                    End If
                End If
            Next x
        End With
    End If
    
    If ComboBoxData.Value = "Total Keuntungan" Then
        With ListBox1
            For x = 0 To .ListCount - 1
                If .Selected(x) Then
                    DoEvents
                    Dim answerTotalKeuntungan As Integer
                    answerTotalKeuntungan = MsgBox("Apakah anda yakin ingin menghapus laporan ini?", vbQuestion + vbYesNo + vbDefaultButton2, .List(x, 0))
                    
                    If answerTotalKeuntungan = vbYes Then
                        Kill (fullPath & .List(x, 0))
                        ListBox1.Clear
                        viewFileName (ComboBoxData.Value)
                        MsgBox "Laporan Berhasil Dihapus!", vbInformation
                    End If
                End If
            Next x
        End With
    End If
End Sub
