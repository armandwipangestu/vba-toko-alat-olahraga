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

Sub makeDirectoryTotalPembelian()
    Dim pathTotalPembelian As String
    Dim fdObj As Object
    Application.ScreenUpdating = False
    Set fdObj = CreateObject("Scripting.FileSystemObject")
    
    pathTotalPembelian = getPath("\Laporan Data\Total Pembelian")
    
    If Not fdObj.FolderExists(pathTotalPembelian) Then
        fdObj.CreateFolder (pathTotalPembelian)
    End If
    
    Application.ScreenUpdating = True

End Sub

Sub makeDirectory()
    makeRootDirectory
    makeDirectoryTotalPembelian
End Sub
