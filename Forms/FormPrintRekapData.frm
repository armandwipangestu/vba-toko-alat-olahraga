VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPrintRekapData 
   Caption         =   "Form Print Rekap Data"
   ClientHeight    =   9165.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13485
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
    FrameNamaFile.BackColor = RGB(29, 29, 66)
    FrameNamaFile.ForeColor = RGB(255, 255, 255)
    FramePreview.BackColor = RGB(29, 29, 66)
    FramePreview.ForeColor = RGB(255, 255, 255)
    ComboBoxData.List = Array("Total Pembelian")
    ComboBoxData.Value = ComboBoxData.List(0)
End Sub

Private Sub ComboBoxData_Change()
    'MsgBox ComboBoxData.Value
    viewFileName (ComboBoxData.Value)
End Sub

Private Sub viewFileName(data As String)
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim i As Integer
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    Set oFolder = oFSO.GetFolder(Application.ActiveWorkbook.path & "\Laporan Data\" & data)

    If data = "Total Pembelian" Then
        For Each oFile In oFolder.Files
            ListBox1.AddItem oFile.Name
        Next oFile
    End If
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
