VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormBarangMasuk 
   Caption         =   "Form Barang Masuk"
   ClientHeight    =   9015.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8595.001
   OleObjectBlob   =   "FormBarangMasuk.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormBarangMasuk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cariById As Range
Dim cariByIdMasterBarang As Range

Private Sub UserForm_Initialize()
    BackColor = RGB(29, 29, 66)
    LabelIdBarangMasuk.BackColor = RGB(29, 29, 66)
    LabelNamaBarang.BackColor = RGB(29, 29, 66)
    LabelTanggalMasuk.BackColor = RGB(29, 29, 66)
    LabelFormatTanggal.BackColor = RGB(29, 29, 66)
    LabelJumlahMasuk.BackColor = RGB(29, 29, 66)
    LabelCari.BackColor = RGB(29, 29, 66)
    CmdBtnSimpan.BackColor = RGB(37, 215, 152)
    CmdBtnBatal.BackColor = RGB(255, 192, 0)
    CmdBtnHapus.BackColor = RGB(231, 21, 86)
    CmdBtnKeluar.BackColor = RGB(113, 59, 219)
    
    TextBoxIdBarangMasuk.Value = buatIdBarangMasuk
    TextBoxIdBarangMasuk.Enabled = False
    ComboBoxNamaBarang.List = wsMasterBarang.Range("B2:B" & getBarisMasterBarang).Value
    TextBoxTanggalMasuk.Value = Format(Now, "DD/MM/YYYY")
    tampilDataListBoxInit
End Sub

Private Sub bersihForm()
    TextBoxIdBarangMasuk.Text = vbNullString
    ComboBoxNamaBarang.Text = vbNullString
    TextBoxTanggalMasuk.Text = vbNullString
    TextBoxJumlahMasuk.Text = vbNullString
    TextBoxCari.Text = vbNullString
End Sub

Private Sub CmdBtnSimpan_Click()
    Set cariById = cariBarangMasuk("A", TextBoxIdBarangMasuk.Value)
    Set cariByIdMasterBarang = cariMasterBarang("B", ComboBoxNamaBarang.Value)

    Dim tanggalMasuk As Date
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
    Dim jumlahMasuk As Integer

    tanggalMasuk = TextBoxTanggalMasuk.Value
    bulan = convertBulanIndonesia(tanggalMasuk)
    tahun = Format(tanggalMasuk, "yyyy")
    idMerekBarang = cariByIdMasterBarang.Offset(0, 1).Value
    merekBarang = cariByIdMasterBarang.Offset(0, 2).Value
    idKategoriBarang = cariByIdMasterBarang.Offset(0, 3).Value
    kategoriBarang = cariByIdMasterBarang.Offset(0, 4).Value
    idBarang = cariByIdMasterBarang.Offset(0, -1).Value
    namaBarang = cariByIdMasterBarang.Offset(0, 0).Value
    hargaBeli = cariByIdMasterBarang.Offset(0, 5).Value
    hargaJual = cariByIdMasterBarang.Offset(0, 6).Value
    jumlahMasuk = TextBoxJumlahMasuk.Value

    Dim baris As Long

    If cariById Is Nothing Then
        baris = getBarisBarangMasuk + 1
        cariByIdMasterBarang.Offset(0, 7).Value = cariByIdMasterBarang.Offset(0, 7).Value + jumlahMasuk
    Else
        baris = cariById.Row
        ' PR apabila nama barang dirubah, harus mengurangi total stok barang sebelumnya dan menambahkan total stok ke barang yang dipilih
        'If namaBarang <> cariById.Offset(0, 9).Value Then
        '    cariByIdMasterBarang.Offset(0, 7).Value = cariByIdMasterBarang.Offset(0, 7).Value - cariById.Offset(0, 12).Value
        'End If
        'cariById.Offset(0, 12).Value
        If cariById.Offset(0, 9).Value <> ComboBoxNamaBarang.Value Then
            cariMasterBarang("B", cariById.Offset(0, 9).Value).Offset(0, 7).Value = cariMasterBarang("B", cariById.Offset(0, 9).Value).Offset(0, 7).Value - TextBoxJumlahMasuk.Value
            cariMasterBarang("B", ComboBoxNamaBarang.Value).Offset(0, 7).Value = cariMasterBarang("B", ComboBoxNamaBarang.Value).Offset(0, 7).Value + TextBoxJumlahMasuk.Value
        End If
        
        ' bawah done
        ' PR pivot table data yang berhubungan dengan harga harus nya dikalikan dengan jumlah barang nya
        If cariById.Offset(0, 12).Value > jumlahMasuk Then
            cariByIdMasterBarang.Offset(0, 7).Value = cariByIdMasterBarang.Offset(0, 7).Value - (cariById.Offset(0, 12).Value - jumlahMasuk)
        End If
        
        If cariById.Offset(0, 12).Value < jumlahMasuk Then
            cariByIdMasterBarang.Offset(0, 7).Value = cariByIdMasterBarang.Offset(0, 7).Value + (jumlahMasuk - cariById.Offset(0, 12).Value)
        End If
    End If

    Dim isiData As Variant
    isiData = Array(TextBoxIdBarangMasuk.Value, tanggalMasuk, _
                    bulan, tahun, idMerekBarang, merekBarang, _
                    idKategoriBarang, kategoriBarang, _
                    idBarang, namaBarang, hargaBeli, hargaJual, _
                    jumlahMasuk)

    wsBarangMasuk.Range("A" & baris).Resize(1, 13).Value = isiData
    MsgBox "Data berhasil disimpan!", vbInformation
    Call bersihForm
    TextBoxIdBarangMasuk.Text = buatIdBarangMasuk
    TextBoxTanggalMasuk.Value = Format(Now, "DD/MM/YYYY")
    ListBoxBarangMasuk.Clear
    tampilDataListBoxInit
    TextBoxIdBarangMasuk.Enabled = False
    RefreshPivotTable
End Sub

Private Sub CmdBtnBatal_Click()
    Call bersihForm
    TextBoxIdBarangMasuk.Text = buatIdBarangMasuk
    TextBoxTanggalMasuk.Value = Format(Now, "DD/MM/YYYY")
    TextBoxIdBarangMasuk.Enabled = False
    ListBoxBarangMasuk.Clear
    tampilDataListBoxInit
End Sub

Private Sub CmdBtnHapus_Click()
    Set cariById = cariBarangMasuk("A", TextBoxIdBarangMasuk.Value)
    Set cariByIdMasterBarang = cariMasterBarang("B", cariById.Offset(0, 9))
    
    If TextBoxIdBarangMasuk.Text = "" Then
        MsgBox "Silahkan Isi ID Barang Masuk!", vbCritical
        Exit Sub
    End If
    
    If cariById Is Nothing Then
        MsgBox "Data ID Barang Masuk Tidak Ditemukan!", vbInformation
    Else
        cariByIdMasterBarang.Offset(0, 7).Value = cariByIdMasterBarang.Offset(0, 7).Value - cariById.Offset(0, 12).Value
        cariById.EntireRow.Delete
        MsgBox "Data Berhasil Di Hapus!", vbInformation
    End If
    
    Call bersihForm
    TextBoxIdBarangMasuk.Text = buatIdBarangMasuk
    TextBoxTanggalMasuk.Value = Format(Now, "DD/MM/YYYY")
    TextBoxIdBarangMasuk.Enabled = False
    ListBoxBarangMasuk.Clear
    tampilDataListBoxInit
    RefreshPivotTable
End Sub

Private Sub CmdBtnKeluar_Click()
    Unload Me
End Sub

Private Sub TextBoxCari_Change()
    tampilDataListBoxCari
End Sub

Private Sub tampilDataListBoxInit()
    Dim i As Long

    With ListBoxBarangMasuk
        .ColumnCount = 13
        .List = wsBarangMasuk.Range("A2:B2").CurrentRegion.Value
        .ColumnWidths = "100;100;100;100;100;100;100;100;100;200;100;100;100"
        'Update the date format for column 2
        For i = 0 To .ListCount - 1
            .List(i, 1) = Format(.List(i, 1), "dd/mm/yyyy")
        Next i
    End With
End Sub

Private Sub tampilDataListBoxCari()
    Dim i As Integer

    ListBoxBarangMasuk.Clear
    With ListBoxBarangMasuk
        .ColumnCount = 13
        .AddItem
        .List(.ListCount - 1, 0) = "ID Barang Masuk"
        .List(.ListCount - 1, 1) = "Tanggal Masuk"
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
        .List(.ListCount - 1, 12) = "Jumlah Masuk"
        .ColumnWidths = "100;100;100;100;100;100;100;100;100;200;100;100;100"
        .ForeColor = vbBlack
    End With
    
    For i = 2 To getBarisBarangMasuk
        If LCase(wsBarangMasuk.Cells(i, 10)) Like "*" & LCase(TextBoxCari) & "*" Then
            With ListBoxBarangMasuk
                .AddItem
                .List(.ListCount - 1, 0) = wsBarangMasuk.Cells(i, 1)
                .List(.ListCount - 1, 1) = Format(wsBarangMasuk.Cells(i, 2), "dd/mm/yyyy")
                .List(.ListCount - 1, 2) = wsBarangMasuk.Cells(i, 3)
                .List(.ListCount - 1, 3) = wsBarangMasuk.Cells(i, 4)
                .List(.ListCount - 1, 4) = wsBarangMasuk.Cells(i, 5)
                .List(.ListCount - 1, 5) = wsBarangMasuk.Cells(i, 6)
                .List(.ListCount - 1, 6) = wsBarangMasuk.Cells(i, 7)
                .List(.ListCount - 1, 7) = wsBarangMasuk.Cells(i, 8)
                .List(.ListCount - 1, 8) = wsBarangMasuk.Cells(i, 9)
                .List(.ListCount - 1, 9) = wsBarangMasuk.Cells(i, 10)
                .List(.ListCount - 1, 10) = wsBarangMasuk.Cells(i, 11)
                .List(.ListCount - 1, 11) = wsBarangMasuk.Cells(i, 12)
                .List(.ListCount - 1, 12) = wsBarangMasuk.Cells(i, 13)
            End With
        End If
    Next
End Sub

Private Sub ListBoxBarangMasuk_Click()
    Dim x As Long
    
    With ListBoxBarangMasuk
        For x = .ListCount - 1 To 1 Step -1
            If .Selected(x) Then
                DoEvents
                TextBoxIdBarangMasuk.Value = .List(x, 0)
                ComboBoxNamaBarang.Value = .List(x, 9)
                TextBoxTanggalMasuk.Value = .List(x, 1)
                TextBoxJumlahMasuk.Value = .List(x, 12)
            End If
        Next x
    End With
    
    TextBoxIdBarangMasuk.Enabled = True
End Sub
