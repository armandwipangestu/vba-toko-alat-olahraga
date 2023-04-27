Attribute VB_Name = "ModuleCariData"
Option Explicit

Public Function cariMerekBarang(rowName As String, keywordName As String) As Range
    Dim baris As Long
    Dim dbArea As Range
    Dim cariKeyword As Range
    
    Set dbArea = wsMerekBarang.Range(rowName & "2:" & rowName & getBarisMerekBarang)
    Set cariKeyword = dbArea.Find(keywordName, , xlValues, xlWhole)
    Set cariMerekBarang = cariKeyword
End Function

Public Function cariKategoriBarang(rowName As String, keywordName As String) As Range
    Dim baris As Long
    Dim dbArea As Range
    Dim cariKeyword As Range
    
    Set dbArea = wsKategoriBarang.Range(rowName & "2:" & rowName & getBarisKategoriBarang)
    Set cariKeyword = dbArea.Find(keywordName, , xlValues, xlWhole)
    Set cariKategoriBarang = cariKeyword
End Function

Public Function cariMasterBarang(rowName As String, keywordName As String) As Range
    Dim baris As Long
    Dim dbArea As Range
    Dim cariKeyword As Range
    
    Set dbArea = wsMasterBarang.Range(rowName & "2:" & rowName & getBarisMasterBarang)
    Set cariKeyword = dbArea.Find(keywordName, , xlValues, xlWhole)
    Set cariMasterBarang = cariKeyword
End Function

Public Function cariBarangMasuk(rowName As String, keywordName As String) As Range
    Dim baris As Long
    Dim dbArea As Range
    Dim cariKeyword As Range
    
    Set dbArea = wsBarangMasuk.Range(rowName & "2:" & rowName & getBarisBarangMasuk)
    Set cariKeyword = dbArea.Find(keywordName, , xlValues, xlWhole)
    Set cariBarangMasuk = cariKeyword
End Function

Public Function cariPenjualanBarang(rowName As String, keywordName As String) As Range
    Dim baris As Long
    Dim dbArea As Range
    Dim cariKeyword As Range
    
    Set dbArea = wsPenjualanBarang.Range(rowName & "2:" & rowName & getBarisPenjualanBarang)
    Set cariKeyword = dbArea.Find(keywordName, , xlValues, xlWhole)
    Set cariPenjualanBarang = cariKeyword
End Function
