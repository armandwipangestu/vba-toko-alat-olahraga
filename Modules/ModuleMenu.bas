Attribute VB_Name = "ModuleMenu"
Option Explicit

Sub menuMerekBarang()
    Call SetWorksheets
    wsMenu.Shapes.Range(Array("shape_active_merek_barang")).Visible = msoTrue
    wsMenu.Shapes.Range(Array("shape_active_kategori_barang")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_master_barang")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_barang_masuk")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_penjualan_barang")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_rekap_penjualan")).Visible = msoFalse
    
    FormMerekBarang.Show
    wsMenu.Range("A1").Select
End Sub

Sub menuKategoriBarang()
    Call SetWorksheets
    wsMenu.Shapes.Range(Array("shape_active_merek_barang")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_kategori_barang")).Visible = msoTrue
    wsMenu.Shapes.Range(Array("shape_active_master_barang")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_barang_masuk")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_penjualan_barang")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_rekap_penjualan")).Visible = msoFalse
    
    'wsKategoriBarang.Select
    FormKategoriBarang.Show
    wsMenu.Range("A1").Select
End Sub

Sub menuMasterBarang()
    Call SetWorksheets
    wsMenu.Shapes.Range(Array("shape_active_merek_barang")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_kategori_barang")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_master_barang")).Visible = msoTrue
    wsMenu.Shapes.Range(Array("shape_active_barang_masuk")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_penjualan_barang")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_rekap_penjualan")).Visible = msoFalse
    
    FormMasterBarang.Show
    wsMenu.Range("A1").Select
End Sub

Sub menuBarangMasuk()
    Call SetWorksheets
    wsMenu.Shapes.Range(Array("shape_active_merek_barang")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_kategori_barang")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_master_barang")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_barang_masuk")).Visible = msoTrue
    wsMenu.Shapes.Range(Array("shape_active_penjualan_barang")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_rekap_penjualan")).Visible = msoFalse
    
    FormBarangMasuk.Show
    wsMenu.Range("A1").Select
End Sub

Sub menuPenjualanBarang()
    Call SetWorksheets
    wsMenu.Shapes.Range(Array("shape_active_merek_barang")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_kategori_barang")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_master_barang")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_barang_masuk")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_penjualan_barang")).Visible = msoTrue
    wsMenu.Shapes.Range(Array("shape_active_rekap_penjualan")).Visible = msoFalse
    
    FormPenjualanBarang.Show
    wsMenu.Range("A1").Select
End Sub

Sub menuPrintRekapData()
    Call SetWorksheets
    wsMenu.Shapes.Range(Array("shape_active_merek_barang")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_kategori_barang")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_master_barang")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_barang_masuk")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_penjualan_barang")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_rekap_penjualan")).Visible = msoTrue
    
    FormPrintRekapData.Show
    wsMenu.Range("A1").Select
End Sub

Sub resetActive()
    Call SetWorksheets
    wsMenu.Shapes.Range(Array("shape_active_merek_barang")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_kategori_barang")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_master_barang")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_barang_masuk")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_penjualan_barang")).Visible = msoFalse
    wsMenu.Shapes.Range(Array("shape_active_rekap_penjualan")).Visible = msoFalse
End Sub
