VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormKategoriBarang 
   Caption         =   "Form Kategori Barang"
   ClientHeight    =   8265.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5490
   OleObjectBlob   =   "FormKategoriBarang.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormKategoriBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cariById As Range

Private Sub UserForm_Initialize()
    BackColor = RGB(29, 29, 66)
    LabelIdKategoriBarang.BackColor = RGB(29, 29, 66)
    LabelKategoriBarang.BackColor = RGB(29, 29, 66)
    LabelCari.BackColor = RGB(29, 29, 66)
    CmdBtnSimpan.BackColor = RGB(37, 215, 152)
    CmdBtnBatal.BackColor = RGB(255, 192, 0)
    CmdBtnHapus.BackColor = RGB(231, 21, 86)
    CmdBtnKeluar.BackColor = RGB(113, 59, 219)
    
    TextBoxIdKategoriBarang.Value = buatIdKategoriBarang
    TextBoxIdKategoriBarang.Enabled = False
    tampilDataListBoxInit
End Sub

Private Sub bersihForm()
    TextBoxIdKategoriBarang.Text = vbNullString
    TextBoxKategoriBarang.Text = vbNullString
    TextBoxCari.Text = vbNullString
End Sub

Private Sub CmdBtnSimpan_Click()
    Set cariById = cariKategoriBarang("A", TextBoxIdKategoriBarang.Value)
    Dim baris As Long
    Dim trigger As Boolean
    Dim kategoriBarangBefore As String
    
    If cariById Is Nothing Then
        baris = getBarisKategoriBarang + 1
    Else
        trigger = True
        baris = cariById.Row
        kategoriBarangBefore = cariById.Offset(0, 1).Value
    End If
    
    Dim isiData As Variant
    isiData = Array(TextBoxIdKategoriBarang.Value, TextBoxKategoriBarang.Value)
    
    wsKategoriBarang.Range("A" & baris).Resize(1, 2).Value = isiData
    
    If trigger Then
        If baris = cariById.Row Then
            If kategoriBarangBefore <> TextBoxKategoriBarang.Value Then
                ' sheet master barang
                Dim lastRow As Long
                Dim i As Long
                lastRow = wsMasterBarang.Cells(Rows.Count, 1).End(xlUp).Row
                For i = 2 To lastRow
                    If wsMasterBarang.Cells(i, 5).Value = cariById.Value Then
                        wsMasterBarang.Cells(i, 6).Value = cariById.Offset(0, 1).Value
                    End If
                Next i
                
                ' sheet barang masuk
                Dim lastRowBarangMasuk As Long
                Dim iBarangMasuk As Long
                lastRowBarangMasuk = wsBarangMasuk.Cells(Rows.Count, 1).End(xlUp).Row
                For iBarangMasuk = 2 To lastRowBarangMasuk
                    If wsBarangMasuk.Cells(iBarangMasuk, 7).Value = cariById.Value Then
                        wsBarangMasuk.Cells(iBarangMasuk, 8).Value = cariById.Offset(0, 1).Value
                    End If
                Next iBarangMasuk
                
                ' sheet penjualan barang
                Dim lastRowPenjualanBarang As Long
                Dim iPenjualanBarang As Long
                lastRowPenjualanBarang = wsPenjualanBarang.Cells(Rows.Count, 1).End(xlUp).Row
                For iPenjualanBarang = 2 To lastRowPenjualanBarang
                    If wsPenjualanBarang.Cells(iPenjualanBarang, 7).Value = cariById.Value Then
                        wsPenjualanBarang.Cells(iPenjualanBarang, 8).Value = cariById.Offset(0, 1).Value
                    End If
                Next iPenjualanBarang
            End If
            
            RefreshPivotTable
        End If
    End If
    
    MsgBox "Data berhasil disimpan!", vbInformation
    Call bersihForm
    TextBoxIdKategoriBarang.Text = buatIdKategoriBarang
    ListBoxKategoriBarang.Clear
    tampilDataListBoxInit
    TextBoxIdKategoriBarang.Enabled = False
End Sub

Private Sub CmdBtnBatal_Click()
    Call bersihForm
    TextBoxIdKategoriBarang.Text = buatIdKategoriBarang
    TextBoxIdKategoriBarang.Enabled = False
    ListBoxKategoriBarang.Clear
    tampilDataListBoxInit
End Sub

Private Sub CmdBtnHapus_Click()
    Set cariById = cariKategoriBarang("A", TextBoxIdKategoriBarang.Value)

    If TextBoxIdKategoriBarang.Text = "" Then
        MsgBox "Silahkan Isi ID Barang!", vbCritical
        Exit Sub
    End If
    
    If cariById Is Nothing Then
        MsgBox "Data ID Barang Tidak Ditemukan!", vbInformation
    Else
        cariById.EntireRow.Delete
        MsgBox "Data Berhasil Di Hapus!", vbInformation
    End If
    
    Call bersihForm
    TextBoxIdKategoriBarang.Text = buatIdKategoriBarang
    ListBoxKategoriBarang.Clear
    tampilDataListBoxInit
    TextBoxIdKategoriBarang.Enabled = False
End Sub

Private Sub CmdBtnKeluar_Click()
    Unload Me
End Sub

Private Sub TextBoxCari_Change()
    tampilDataListBoxCari
End Sub

Private Sub tampilDataListBoxInit()
    With ListBoxKategoriBarang
        .ColumnCount = 2
        .List = wsKategoriBarang.Range("A2:B2").CurrentRegion.Value
        .ColumnWidths = "100;200"
        
        '.ColumnHeads = True
        ''Data Model'!$G$20
        '.RowSource = "'Merek Barang'!B2:B" & getBarisMerekBarang
        '.ColumnHeads = True
        '.ColumnWidths = "100;150"
    End With
End Sub

Private Sub tampilDataListBoxCari()
    Dim i As Integer

    ListBoxKategoriBarang.Clear
    With ListBoxKategoriBarang
        .ColumnCount = 2
        .AddItem
        .List(.ListCount - 1, 0) = "ID Kategori Barang"
        .List(.ListCount - 1, 1) = "Kategori Barang"
        .ColumnWidths = "100;200"
        .ForeColor = vbBlack
    End With
    
    For i = 2 To getBarisKategoriBarang
        If LCase(wsKategoriBarang.Cells(i, 2)) Like "*" & LCase(TextBoxCari) & "*" Then
            With ListBoxKategoriBarang
                .AddItem
                .List(.ListCount - 1, 0) = wsKategoriBarang.Cells(i, 1)
                .List(.ListCount - 1, 1) = wsKategoriBarang.Cells(i, 2)
            End With
        End If
    Next
End Sub

Private Sub ListBoxKategoriBarang_Click()
    Dim x As Long
    
    With ListBoxKategoriBarang
        For x = .ListCount - 1 To 1 Step -1
            If .Selected(x) Then
                DoEvents
                TextBoxIdKategoriBarang.Value = .List(x, 0)
                TextBoxKategoriBarang.Value = .List(x, 1)
            End If
        Next x
    End With
    
    TextBoxIdKategoriBarang.Enabled = True
End Sub

