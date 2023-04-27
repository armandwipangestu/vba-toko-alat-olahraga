Attribute VB_Name = "WorksheetAlias"
Option Explicit

Public wsMenu As Worksheet
Public wsDataModel As Worksheet
Public wsMerekBarang As Worksheet
Public wsKategoriBarang As Worksheet
Public wsMasterBarang As Worksheet
Public wsBarangMasuk As Worksheet
Public wsPenjualanBarang As Worksheet
'Public wsRekapPenjualan As Worksheet

Sub SetWorksheets()
    Set wsMenu = ThisWorkbook.Worksheets("Menu")
    Set wsDataModel = ThisWorkbook.Worksheets("Data Model")
    Set wsMerekBarang = ThisWorkbook.Worksheets("Merek Barang")
    Set wsKategoriBarang = ThisWorkbook.Worksheets("Kategori Barang")
    Set wsMasterBarang = ThisWorkbook.Worksheets("Master Barang")
    Set wsBarangMasuk = ThisWorkbook.Worksheets("Barang Masuk")
    Set wsPenjualanBarang = ThisWorkbook.Worksheets("Penjualan Barang")
    'Set wsRekapPenjualan = ThisWorkbook.Worksheets("Rekap Penjualan")
End Sub
