VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormMasterBarang 
   Caption         =   "Form Master Barang"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8595.001
   OleObjectBlob   =   "FormMasterBarang.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormMasterBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cariById As Range
Dim cariByIdMerekBarang As Range
Dim cariByIdKategoriBarang As Range

Private Sub UserForm_Initialize()
    BackColor = RGB(29, 29, 66)
    LabelIdMasteriBarang.BackColor = RGB(29, 29, 66)
    LabelNamaBarang.BackColor = RGB(29, 29, 66)
    LabelMerekBarang.BackColor = RGB(29, 29, 66)
    LabelKategoriBarang.BackColor = RGB(29, 29, 66)
    LabelHargaBeli.BackColor = RGB(29, 29, 66)
    LabelHargaJual.BackColor = RGB(29, 29, 66)
    LabelStok.BackColor = RGB(29, 29, 66)
    LabelCari.BackColor = RGB(29, 29, 66)
    CmdBtnSimpan.BackColor = RGB(37, 215, 152)
    CmdBtnBatal.BackColor = RGB(255, 192, 0)
    CmdBtnHapus.BackColor = RGB(231, 21, 86)
    CmdBtnKeluar.BackColor = RGB(113, 59, 219)
    
    TextBoxIdMasterBarang.Value = buatIdMasterBarang
    TextBoxIdMasterBarang.Enabled = False
    ComboBoxMerekBarang.List = wsMerekBarang.Range("B2:B" & getBarisMerekBarang).Value
    ComboBoxKategoriBarang.List = wsKategoriBarang.Range("B2:B" & getBarisKategoriBarang).Value
    TextBoxStok.Value = 0
    tampilDataListBoxInit
End Sub

Private Sub bersihForm()
    TextBoxIdMasterBarang.Text = vbNullString
    TextBoxNamaBarang.Text = vbNullString
    ComboBoxMerekBarang.Text = vbNullString
    ComboBoxKategoriBarang.Text = vbNullString
    TextBoxHargaBeli.Text = vbNullString
    TextBoxHargaJual.Text = vbNullString
    TextBoxStok.Value = 0
    TextBoxCari.Text = vbNullString
End Sub

Private Sub CmdBtnSimpan_Click()
    Set cariById = cariMasterBarang("A", TextBoxIdMasterBarang.Value)
    Set cariByIdMerekBarang = cariMerekBarang("B", ComboBoxMerekBarang.Value)
    Set cariByIdKategoriBarang = cariKategoriBarang("B", ComboBoxKategoriBarang.Value)
    
    Dim idMerekBarang As String
    Dim idKategoriBarang As String
    
    idMerekBarang = cariByIdMerekBarang.Offset(0, -1).Value
    idKategoriBarang = cariByIdKategoriBarang.Offset(0, -1).Value
    
    Dim baris As Long
    Dim trigger As Boolean
    Dim idMerekBarangBefore As String
    Dim namaMerekBefore As String
    Dim idKategoriBarangBefore As String
    Dim namaKategoriBefore As String
    Dim namaBarangBefore As String
    
    If cariById Is Nothing Then
        baris = getBarisMasterBarang + 1
    Else
        trigger = True
        baris = cariById.Row
        idMerekBarangBefore = cariById.Offset(0, 2).Value
        namaMerekBefore = cariById.Offset(0, 3).Value
        idKategoriBarangBefore = cariById.Offset(0, 4).Value
        namaKategoriBefore = cariById.Offset(0, 5).Value
        namaBarangBefore = cariById.Offset(0, 1).Value
    End If
    
    Dim isiData As Variant
    isiData = Array(TextBoxIdMasterBarang.Value, TextBoxNamaBarang.Value, _
                    idMerekBarang, ComboBoxMerekBarang.Value, _
                    idKategoriBarang, ComboBoxKategoriBarang.Value, _
                    TextBoxHargaBeli.Value, TextBoxHargaJual.Value, _
                    TextBoxStok.Value)
                    
    wsMasterBarang.Range("A" & baris).Resize(1, 9).Value = isiData
    
    If trigger Then
        If baris = cariById.Row Then
        
            ' merek barang
            If namaMerekBefore <> ComboBoxMerekBarang.Value Then
    '            MsgBox namaMerekBefore & " = " & ComboBoxMerekBarang.Value & " " & cariById.Offset(0, 2).Value
                Dim lastRow As Long
                Dim i As Long
                lastRow = wsBarangMasuk.Cells(Rows.Count, 1).End(xlUp).Row
                For i = 2 To lastRow
    '                MsgBox wsBarangMasuk.Cells(i, 5).Value & " = " & idMerekBarangBefore
                    If wsBarangMasuk.Cells(i, 5).Value = idMerekBarangBefore Then
                        wsBarangMasuk.Cells(i, 5).Value = cariById.Offset(0, 2).Value
                        wsBarangMasuk.Cells(i, 6).Value = cariById.Offset(0, 3).Value
                    End If
                Next i
                
                lastRow = wsPenjualanBarang.Cells(Rows.Count, 1).End(xlUp).Row
                For i = 2 To lastRow
    '                MsgBox wsBarangMasuk.Cells(i, 5).Value & " = " & idMerekBarangBefore
                    If wsPenjualanBarang.Cells(i, 5).Value = idMerekBarangBefore Then
                        wsPenjualanBarang.Cells(i, 5).Value = cariById.Offset(0, 2).Value
                        wsPenjualanBarang.Cells(i, 6).Value = cariById.Offset(0, 3).Value
                    End If
                Next i
            End If
            
            ' kategori barang
            If namaKategoriBefore <> ComboBoxKategoriBarang.Value Then
    '            MsgBox namaKategoriBefore & " = " & ComboBoxKategoriBarang.Value & " " & cariById.Offset(0, 4).Value
                Dim lastRowKategoriBarang As Long
                Dim iKategoriBarang As Long
                lastRowKategoriBarang = wsBarangMasuk.Cells(Rows.Count, 1).End(xlUp).Row
                For iKategoriBarang = 2 To lastRowKategoriBarang
    '                MsgBox wsBarangMasuk.Cells(iKategoriBarang, 7).Value & " = " & idKategoriBarangBefore
                    If wsBarangMasuk.Cells(iKategoriBarang, 7).Value = idKategoriBarangBefore Then
                        wsBarangMasuk.Cells(iKategoriBarang, 7).Value = cariById.Offset(0, 4).Value
                        wsBarangMasuk.Cells(iKategoriBarang, 8).Value = cariById.Offset(0, 5).Value
                    End If
                Next iKategoriBarang
                
                lastRowKategoriBarang = wsPenjualanBarang.Cells(Rows.Count, 1).End(xlUp).Row
                For iKategoriBarang = 2 To lastRowKategoriBarang
    '                MsgBox wsBarangMasuk.Cells(iKategoriBarang, 7).Value & " = " & idKategoriBarangBefore
                    If wsPenjualanBarang.Cells(iKategoriBarang, 7).Value = idKategoriBarangBefore Then
                        wsPenjualanBarang.Cells(iKategoriBarang, 7).Value = cariById.Offset(0, 4).Value
                        wsPenjualanBarang.Cells(iKategoriBarang, 8).Value = cariById.Offset(0, 5).Value
                    End If
                Next iKategoriBarang
            End If
        
            ' nama barang
            If namaBarangBefore <> TextBoxNamaBarang.Value Then
    '            MsgBox namaBarangBefore & " = " & TextBoxNamaBarang.Value
                Dim lastRowNamaBarang As Long
                Dim iNamaBarang As Long
                lastRowNamaBarang = wsBarangMasuk.Cells(Rows.Count, 1).End(xlUp).Row
                For iNamaBarang = 2 To lastRowNamaBarang
                    If wsBarangMasuk.Cells(iNamaBarang, 9).Value = cariById.Value Then
                        wsBarangMasuk.Cells(iNamaBarang, 10).Value = cariById.Offset(0, 1).Value
                    End If
                Next iNamaBarang
                
                lastRowNamaBarang = wsPenjualanBarang.Cells(Rows.Count, 1).End(xlUp).Row
                For iNamaBarang = 2 To lastRowNamaBarang
                    If wsPenjualanBarang.Cells(iNamaBarang, 9).Value = cariById.Value Then
                        wsPenjualanBarang.Cells(iNamaBarang, 10).Value = cariById.Offset(0, 1).Value
                    End If
                Next iNamaBarang
            End If
            
            RefreshPivotTable
        End If
    End If
    
    MsgBox "Data berhasil disimpan!", vbInformation
    Call bersihForm
    TextBoxIdMasterBarang.Text = buatIdMasterBarang
    ListBoxMasterBarang.Clear
    tampilDataListBoxInit
    TextBoxIdMasterBarang.Enabled = False
End Sub

Private Sub CmdBtnBatal_Click()
    Call bersihForm
    TextBoxIdMasterBarang.Text = buatIdMasterBarang
    TextBoxIdMasterBarang.Enabled = False
    ListBoxMasterBarang.Clear
    tampilDataListBoxInit
End Sub

Private Sub CmdBtnHapus_Click()
    Set cariById = cariMasterBarang("A", TextBoxIdMasterBarang.Value)

    If TextBoxIdMasterBarang.Text = "" Then
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
    TextBoxIdMasterBarang.Text = buatIdMasterBarang
    ListBoxMasterBarang.Clear
    tampilDataListBoxInit
    TextBoxIdMasterBarang.Enabled = False
End Sub

Private Sub CmdBtnKeluar_Click()
    Unload Me
End Sub

Private Sub TextBoxCari_Change()
    tampilDataListBoxCari
End Sub

Private Sub tampilDataListBoxInit()
    With ListBoxMasterBarang
        .ColumnCount = 9
        .List = wsMasterBarang.Range("A2:B2").CurrentRegion.Value
        .ColumnWidths = "100;200;100;100;100;100;100;100;100"
        '.ColumnHeads = True
        
        ''Data Model'!$G$20
        '.RowSource = "'Merek Barang'!B2:B" & getBarisMerekBarang
        '.ColumnHeads = True
        '.ColumnWidths = "100;150"
    End With
End Sub

Private Sub tampilDataListBoxCari()
    Dim i As Integer

    ListBoxMasterBarang.Clear
    With ListBoxMasterBarang
        .ColumnCount = 9
        .AddItem
        .List(.ListCount - 1, 0) = "ID Barang"
        .List(.ListCount - 1, 1) = "Nama Barang"
        .List(.ListCount - 1, 2) = "ID Merek Barang"
        .List(.ListCount - 1, 3) = "Merek Barang"
        .List(.ListCount - 1, 4) = "ID Kategori Barang"
        .List(.ListCount - 1, 5) = "Kategori Barang"
        .List(.ListCount - 1, 6) = "Harga Beli"
        .List(.ListCount - 1, 7) = "Harga Jual"
        .List(.ListCount - 1, 8) = "Stok"
        .ColumnWidths = "100;200;100;100;100;100;100;100;100"
        .ForeColor = vbBlack
    End With
    
    For i = 2 To getBarisMasterBarang
        If LCase(wsMasterBarang.Cells(i, 2)) Like "*" & LCase(TextBoxCari) & "*" Then
            With ListBoxMasterBarang
                .AddItem
                .List(.ListCount - 1, 0) = wsMasterBarang.Cells(i, 1)
                .List(.ListCount - 1, 1) = wsMasterBarang.Cells(i, 2)
                .List(.ListCount - 1, 2) = wsMasterBarang.Cells(i, 3)
                .List(.ListCount - 1, 3) = wsMasterBarang.Cells(i, 4)
                .List(.ListCount - 1, 4) = wsMasterBarang.Cells(i, 5)
                .List(.ListCount - 1, 5) = wsMasterBarang.Cells(i, 6)
                .List(.ListCount - 1, 6) = wsMasterBarang.Cells(i, 7)
                .List(.ListCount - 1, 7) = wsMasterBarang.Cells(i, 8)
                .List(.ListCount - 1, 8) = wsMasterBarang.Cells(i, 9)
            End With
        End If
    Next
End Sub

Private Sub ListBoxMasterBarang_Click()
    Dim x As Long
    
    With ListBoxMasterBarang
        For x = .ListCount - 1 To 1 Step -1
            If .Selected(x) Then
                DoEvents
                TextBoxIdMasterBarang.Value = .List(x, 0)
                TextBoxNamaBarang.Value = .List(x, 1)
                ComboBoxMerekBarang.Value = .List(x, 3)
                ComboBoxKategoriBarang.Value = .List(x, 5)
                TextBoxHargaBeli.Value = .List(x, 6)
                TextBoxHargaJual.Value = .List(x, 7)
                TextBoxStok.Value = .List(x, 8)
            End If
        Next x
    End With
    
    TextBoxIdMasterBarang.Enabled = True
End Sub
