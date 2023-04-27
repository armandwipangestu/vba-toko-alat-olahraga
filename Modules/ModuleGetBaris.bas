Attribute VB_Name = "ModuleGetBaris"
Option Explicit

Public Function getBarisMerekBarang() As Long
    Call SetWorksheets
    getBarisMerekBarang = wsMerekBarang.Range("A" & wsMerekBarang.Rows.Count).End(xlUp).Row
End Function

Public Function getBarisKategoriBarang() As Long
    Call SetWorksheets
    getBarisKategoriBarang = wsKategoriBarang.Range("A" & wsKategoriBarang.Rows.Count).End(xlUp).Row
End Function

Public Function getBarisMasterBarang() As Long
    Call SetWorksheets
    getBarisMasterBarang = wsMasterBarang.Range("A" & wsMasterBarang.Rows.Count).End(xlUp).Row
End Function

Public Function getBarisBarangMasuk() As Long
    Call SetWorksheets
    getBarisBarangMasuk = wsBarangMasuk.Range("A" & wsBarangMasuk.Rows.Count).End(xlUp).Row
End Function

Public Function getBarisPenjualanBarang() As Long
    Call SetWorksheets
    getBarisPenjualanBarang = wsPenjualanBarang.Range("A" & wsPenjualanBarang.Rows.Count).End(xlUp).Row
End Function
