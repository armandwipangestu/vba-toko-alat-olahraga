VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPenjualanBarang 
   Caption         =   "Form Penjualan Barang"
   ClientHeight    =   10110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8640.001
   OleObjectBlob   =   "FormPenjualanBarang.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormPenjualanBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cariById As Range
Dim cariByIdMasterBarang As Range

Private Sub UserForm_Initialize()
    BackColor = RGB(29, 29, 66)
    LabelIdPenjualanBarang.BackColor = RGB(29, 29, 66)
    LabelNamaBarang.BackColor = RGB(29, 29, 66)
    LabelTanggalTerjual.BackColor = RGB(29, 29, 66)
    LabelFormatTanggal.BackColor = RGB(29, 29, 66)
    LabelJumlahPenjualan.BackColor = RGB(29, 29, 66)
    LabelStok.BackColor = RGB(29, 29, 66)
    LabelHargaBeli.BackColor = RGB(29, 29, 66)
    LabelHargaJual.BackColor = RGB(29, 29, 66)
    LabelKeuntungan.BackColor = RGB(29, 29, 66)
    LabelCari.BackColor = RGB(29, 29, 66)
    CmdBtnSimpan.BackColor = RGB(37, 215, 152)
    CmdBtnBatal.BackColor = RGB(255, 192, 0)
    CmdBtnHapus.BackColor = RGB(231, 21, 86)
    CmdBtnKeluar.BackColor = RGB(113, 59, 219)

    TextBoxIdPenjualanBarang.Value = buatIdPenjualanBarang
    TextBoxIdPenjualanBarang.Enabled = False
    TextBoxStok.Enabled = False
    TextBoxHargaBeli.Enabled = False
    TextBoxHargaJual.Enabled = False
    TextBoxKeuntungan.Enabled = False
    ComboBoxNamaBarang.List = wsMasterBarang.Range("B2:B" & getBarisMasterBarang).Value
    TextBoxTanggalTerjual.Value = Format(Now, "DD/MM/YYYY")
    tampilDataListBoxInit
End Sub

Private Sub bersihForm()
    TextBoxIdPenjualanBarang.Text = vbNullString
    ComboBoxNamaBarang.Text = vbNullString
    TextBoxTanggalTerjual.Text = vbNullString
    TextBoxJumlahPenjualan.Text = vbNullString
    TextBoxHargaBeli.Text = vbNullString
    TextBoxHargaJual.Text = vbNullString
    TextBoxKeuntungan.Text = vbNullString
    TextBoxStok.Text = vbNullString
    TextBoxCari.Text = vbNullString
End Sub

Private Sub CmdBtnSimpan_Click()
    Set cariById = cariPenjualanBarang("A", TextBoxIdPenjualanBarang.Value)
    Set cariByIdMasterBarang = cariMasterBarang("B", ComboBoxNamaBarang.Value)

    Dim tanggalTerjual As Date
    Dim bulan As String
    Dim tahun As String
    Dim idMerekBarang As String
    Dim merekBarang As String
    Dim idKategoriBarang As String
    Dim kategoriBarang As String
    Dim idBarang As String
    Dim namaBarang As String
    Dim hargaBeli As Double
    Dim hargaJual As Double
    Dim jumlahPenjualan As Integer

    tanggalTerjual = TextBoxTanggalTerjual.Value
    bulan = convertBulanIndonesia(tanggalTerjual)
    tahun = Format(tanggalTerjual, "yyyy")
    idMerekBarang = cariByIdMasterBarang.Offset(0, 1).Value
    merekBarang = cariByIdMasterBarang.Offset(0, 2).Value
    idKategoriBarang = cariByIdMasterBarang.Offset(0, 3).Value
    kategoriBarang = cariByIdMasterBarang.Offset(0, 4).Value
    idBarang = cariByIdMasterBarang.Offset(0, -1).Value
    namaBarang = cariByIdMasterBarang.Offset(0, 0).Value
    hargaBeli = cariByIdMasterBarang.Offset(0, 5).Value
    hargaJual = cariByIdMasterBarang.Offset(0, 6).Value
    jumlahPenjualan = TextBoxJumlahPenjualan.Value

    Dim baris As Long
    
'    If TextBoxJumlahPenjualan.Value <> "" Then
'        Set cariByIdMasterBarang = cariMasterBarang("B", ComboBoxNamaBarang.Value)
'        Dim stok As Integer
'        Dim jumlahPenjualanCek As Integer
'
'        stok = cariByIdMasterBarang.Offset(0, 7).Value
'        jumlahPenjualanCek = TextBoxJumlahPenjualan.Value
'
'        If stok - jumlahPenjualanCek < 0 Then
'            MsgBox "Stok barang hanya tersisa " & stok, vbCritical
'            Exit Sub
'        End If
'    End If

    If cariById Is Nothing Then
        baris = getBarisPenjualanBarang + 1
        
        If TextBoxJumlahPenjualan.Value <> "" Then
            Dim stok As Integer
            
            stok = cariByIdMasterBarang.Offset(0, 7).Value
            
            If stok - jumlahPenjualan < 0 Then
                MsgBox "Stok barang hanya tersisa " & stok, vbCritical
                Exit Sub
            End If
        End If
        
        cariByIdMasterBarang.Offset(0, 7).Value = cariByIdMasterBarang.Offset(0, 7).Value - jumlahPenjualan
    Else
        baris = cariById.Row
        
        If cariById.Offset(0, 12).Value - TextBoxJumlahPenjualan.Value + cariByIdMasterBarang.Offset(0, 7).Value < 0 Then
            MsgBox "Pengubahan Jumlah Penjualan tidak bisa kurang dari 0", vbCritical
            Exit Sub
        End If
        
        If cariById.Offset(0, 9).Value <> ComboBoxNamaBarang.Value Then
            cariMasterBarang("B", cariById.Offset(0, 9).Value).Offset(0, 7).Value = cariMasterBarang("B", cariById.Offset(0, 9).Value).Offset(0, 7).Value + TextBoxJumlahPenjualan.Value
            cariMasterBarang("B", ComboBoxNamaBarang.Value).Offset(0, 7).Value = cariMasterBarang("B", ComboBoxNamaBarang.Value).Offset(0, 7).Value - TextBoxJumlahPenjualan.Value
        End If
        
        If cariById.Offset(0, 12).Value > jumlahPenjualan Then
            cariByIdMasterBarang.Offset(0, 7).Value = cariByIdMasterBarang.Offset(0, 7).Value + (cariById.Offset(0, 12) - jumlahPenjualan)
        End If
        
        If cariById.Offset(0, 12).Value < jumlahPenjualan Then
            cariByIdMasterBarang.Offset(0, 7).Value = cariByIdMasterBarang.Offset(0, 7).Value - (jumlahPenjualan - cariById.Offset(0, 12).Value)
        End If
    End If

    Dim isiData As Variant
    isiData = Array(TextBoxIdPenjualanBarang.Value, tanggalTerjual, _
                    bulan, tahun, idMerekBarang, merekBarang, _
                    idKategoriBarang, kategoriBarang, _
                    idBarang, namaBarang, hargaBeli, hargaJual, _
                    jumlahPenjualan)
                    
    wsPenjualanBarang.Range("A" & baris).Resize(1, 13).Value = isiData
    MsgBox "Data berhasil disimpan!", vbInformation
    Call bersihForm
    TextBoxIdPenjualanBarang.Text = buatIdPenjualanBarang
    TextBoxTanggalTerjual.Value = Format(Now, "DD/MM/YYYY")
    ListBoxPenjualanBarang.Clear
    tampilDataListBoxInit
    TextBoxIdPenjualanBarang.Enabled = False
    RefreshPivotTable
End Sub

Private Sub CmdBtnBatal_Click()
    Call bersihForm
    TextBoxIdPenjualanBarang.Text = buatIdPenjualanBarang
    TextBoxTanggalTerjual.Value = Format(Now, "DD/MM/YYYY")
    TextBoxIdPenjualanBarang.Enabled = False
    ListBoxPenjualanBarang.Clear
    tampilDataListBoxInit
End Sub

Private Sub CmdBtnHapus_Click()
    Set cariById = cariPenjualanBarang("A", TextBoxIdPenjualanBarang.Value)
    Set cariByIdMasterBarang = cariMasterBarang("B", cariById.Offset(0, 9))
    
    If TextBoxIdPenjualanBarang.Text = "" Then
        MsgBox "Silahkan ISI ID Penjualan Barang!", vbCritical
        Exit Sub
    End If
    
    If cariById Is Nothing Then
        MsgBox "Data ID Penjualan Barang Tidak Ditemukan!", vbInformation
    Else
        cariByIdMasterBarang.Offset(0, 7).Value = cariByIdMasterBarang.Offset(0, 7).Value + cariById.Offset(0, 12).Value
        cariById.EntireRow.Delete
        MsgBox "Data Berhasil Di Hapus!", vbInformation
    End If
    
    Call bersihForm
    TextBoxIdPenjualanBarang.Text = buatIdPenjualanBarang
    TextBoxTanggalTerjual.Value = Format(Now, "DD/MM/YYYY")
    ListBoxPenjualanBarang.Clear
    tampilDataListBoxInit
    TextBoxIdPenjualanBarang.Enabled = False
    RefreshPivotTable
End Sub

Private Sub CmdBtnKeluar_Click()
    Unload Me
End Sub

Private Sub ComboBoxNamaBarang_Change()
    tampilHarga
    tampilKeuntungan
    tampilStok
End Sub

Private Sub tampilStok()
    If ComboBoxNamaBarang.Value <> "" Then
        Set cariByIdMasterBarang = cariMasterBarang("B", ComboBoxNamaBarang.Value)
        TextBoxStok.Value = cariByIdMasterBarang.Offset(0, 7).Value
    End If
End Sub

Private Sub TextBoxJumlahPenjualan_Change()
    tampilKeuntungan
End Sub

Private Sub tampilHarga()
    If ComboBoxNamaBarang.Value <> "" Then
        Set cariByIdMasterBarang = cariMasterBarang("B", ComboBoxNamaBarang.Value)
        TextBoxHargaBeli.Value = cariByIdMasterBarang.Offset(0, 5).Value
        TextBoxHargaJual.Value = cariByIdMasterBarang.Offset(0, 6).Value
    End If
End Sub

Private Sub tampilKeuntungan()
    If ComboBoxNamaBarang.Value <> "" Then
        Set cariByIdMasterBarang = cariMasterBarang("B", ComboBoxNamaBarang.Value)
        Dim hargaJual As Double
        Dim hargaBeli As Double
        Dim keuntungan As Double

        If TextBoxJumlahPenjualan.Value = "" Then
            TextBoxKeuntungan.Value = vbNullString
            Exit Sub
        End If

        hargaJual = cariByIdMasterBarang.Offset(0, 6).Value
        hargaBeli = cariByIdMasterBarang.Offset(0, 5).Value
        keuntungan = (hargaJual - hargaBeli) * TextBoxJumlahPenjualan.Value
        TextBoxKeuntungan.Value = keuntungan
    End If
End Sub

Private Sub TextBoxCari_Change()
    tampilDataListBoxCari
End Sub

Private Sub tampilDataListBoxInit()
    Dim i As Long

    With ListBoxPenjualanBarang
        .ColumnCount = 14
        .List = wsPenjualanBarang.Range("A2:B2").CurrentRegion.Value
        .ColumnWidths = "100;100;100;100;100;100;100;100;100;200;100;100;100;100"
        'Update the date format for column 2
        For i = 0 To .ListCount - 1
            .List(i, 1) = Format(.List(i, 1), "dd/mm/yyyy")
        Next i
    End With
End Sub

Private Sub tampilDataListBoxCari()
    Dim i As Integer

    ListBoxPenjualanBarang.Clear
    With ListBoxPenjualanBarang
        .ColumnCount = 14
        .AddItem
        .List(.ListCount - 1, 0) = "ID Penjualan Barang"
        .List(.ListCount - 1, 1) = "Tanggal Terjual"
        .List(.ListCount - 1, 2) = "Bulan"
        .List(.ListCount - 1, 3) = "Tahun"
        .List(.ListCount - 1, 4) = "ID Merek Barang"
        .List(.ListCount - 1, 5) = "Merek Barang"
        .List(.ListCount - 1, 6) = "ID Kategori Barang"
        .List(.ListCount - 1, 7) = "Kategori Barang"
        .List(.ListCount - 1, 8) = "ID Barang"
        .List(.ListCount - 1, 9) = "Nama Barang"
        .List(.ListCount - 1, 10) = "Harga Beli"
        .List(.ListCount - 1, 11) = "Harga Jual"
        .List(.ListCount - 1, 12) = "Jumlah Penjualan"
        .List(.ListCount - 1, 13) = "Keuntungan"
        .ColumnWidths = "100;100;100;100;100;100;100;100;100;200;100;100;100;100"
        .ForeColor = vbBlack
    End With

    For i = 2 To getBarisPenjualanBarang
        If LCase(wsPenjualanBarang.Cells(i, 10)) Like "*" & LCase(TextBoxCari) & "*" Then
            With ListBoxPenjualanBarang
                .AddItem
                .List(.ListCount - 1, 0) = wsPenjualanBarang.Cells(i, 1)
                .List(.ListCount - 1, 1) = Format(wsPenjualanBarang.Cells(i, 2), "dd/mm/yyyy")
                .List(.ListCount - 1, 2) = wsPenjualanBarang.Cells(i, 3)
                .List(.ListCount - 1, 3) = wsPenjualanBarang.Cells(i, 4)
                .List(.ListCount - 1, 4) = wsPenjualanBarang.Cells(i, 5)
                .List(.ListCount - 1, 5) = wsPenjualanBarang.Cells(i, 6)
                .List(.ListCount - 1, 6) = wsPenjualanBarang.Cells(i, 7)
                .List(.ListCount - 1, 7) = wsPenjualanBarang.Cells(i, 8)
                .List(.ListCount - 1, 8) = wsPenjualanBarang.Cells(i, 9)
                .List(.ListCount - 1, 9) = wsPenjualanBarang.Cells(i, 10)
                .List(.ListCount - 1, 10) = wsPenjualanBarang.Cells(i, 11)
                .List(.ListCount - 1, 11) = wsPenjualanBarang.Cells(i, 12)
                .List(.ListCount - 1, 12) = wsPenjualanBarang.Cells(i, 13)
                .List(.ListCount - 1, 13) = wsPenjualanBarang.Cells(i, 14)
            End With
        End If
    Next
End Sub

Private Sub ListBoxPenjualanBarang_Click()
    Dim x As Long

    With ListBoxPenjualanBarang
        For x = .ListCount - 1 To 1 Step -1
            If .Selected(x) Then
                DoEvents
                TextBoxIdPenjualanBarang.Value = .List(x, 0)
                ComboBoxNamaBarang.Value = .List(x, 9)
                TextBoxTanggalTerjual.Value = .List(x, 1)
                TextBoxJumlahPenjualan.Value = .List(x, 12)
                TextBoxHargaBeli.Value = .List(x, 10)
                TextBoxHargaJual.Value = .List(x, 11)
                TextBoxKeuntungan.Value = .List(x, 13)
            End If
        Next x
    End With

    TextBoxIdPenjualanBarang.Enabled = True
End Sub
