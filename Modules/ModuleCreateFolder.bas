Attribute VB_Name = "ModuleCreateFolder"
Option Explicit

Sub makeRootDirectory()
    Dim pathLaporanData As String
    Dim fdObj As Object
    Application.ScreenUpdating = False
    Set fdObj = CreateObject("Scripting.FileSystemObject")
    
    pathLaporanData = getPath("\Laporan Data")
    
    If Not fdObj.FolderExists(pathLaporanData) Then
        fdObj.CreateFolder (pathLaporanData)
    End If
    
    Application.ScreenUpdating = True
    
'    If Dir(pathLaporanData, vbDirectory) = vbNullString Then
'        MsgBox "Tidak ada folder " & pathLaporanData
'    End If
End Sub

Sub makeDirectoryTotalBarangMasuk()
    Dim pathTotalBarangMasuk As String
    Dim fdObj As Object
    Application.ScreenUpdating = False
    Set fdObj = CreateObject("Scripting.FileSystemObject")
    
    pathTotalBarangMasuk = getPath("\Laporan Data\Total Barang Masuk")
    
    If Not fdObj.FolderExists(pathTotalBarangMasuk) Then
        fdObj.CreateFolder (pathTotalBarangMasuk)
    End If
    
    Application.ScreenUpdating = True

End Sub

Sub makeDirectoryTotalPenjualanBarang()
    Dim pathTotalPenjualanBarang As String
    Dim fdObj As Object
    Application.ScreenUpdating = False
    Set fdObj = CreateObject("Scripting.FileSystemObject")
    
    pathTotalPenjualanBarang = getPath("\Laporan Data\Total Penjualan Barang")
    
    If Not fdObj.FolderExists(pathTotalPenjualanBarang) Then
        fdObj.CreateFolder (pathTotalPenjualanBarang)
    End If
    
    Application.ScreenUpdating = True

End Sub

Sub makeDirectoryTotalHargaBeli()
    Dim pathTotalHargaBeli As String
    Dim fdObj As Object
    Application.ScreenUpdating = False
    Set fdObj = CreateObject("Scripting.FileSystemObject")
    
    pathTotalHargaBeli = getPath("\Laporan Data\Total Harga Beli")
    
    If Not fdObj.FolderExists(pathTotalHargaBeli) Then
        fdObj.CreateFolder (pathTotalHargaBeli)
    End If
    
    Application.ScreenUpdating = True

End Sub

Sub makeDirectoryTotalHargaJual()
    Dim pathToHargaJual As String
    Dim fdObj As Object
    Application.ScreenUpdating = False
    Set fdObj = CreateObject("Scripting.FileSystemObject")
    
    pathToHargaJual = getPath("\Laporan Data\Total Harga Jual")
    
    If Not fdObj.FolderExists(pathToHargaJual) Then
        fdObj.CreateFolder (pathToHargaJual)
    End If
    
    Application.ScreenUpdating = True

End Sub

Sub makeDirectoryTotalKeuntungan()
    Dim pathTotalKeuntungan As String
    Dim fdObj As Object
    Application.ScreenUpdating = False
    Set fdObj = CreateObject("Scripting.FileSystemObject")
    
    pathTotalKeuntungan = getPath("\Laporan Data\Total Keuntungan")
    
    If Not fdObj.FolderExists(pathTotalKeuntungan) Then
        fdObj.CreateFolder (pathTotalKeuntungan)
    End If
    
    Application.ScreenUpdating = True

End Sub

Sub makeDirectory()
    makeRootDirectory
    makeDirectoryTotalBarangMasuk
    makeDirectoryTotalPenjualanBarang
    makeDirectoryTotalHargaBeli
    makeDirectoryTotalHargaJual
    makeDirectoryTotalKeuntungan
End Sub
