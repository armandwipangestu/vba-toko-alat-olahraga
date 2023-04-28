Attribute VB_Name = "ModuleTema"
Option Explicit

Sub temaBiru()
    Call SetWorksheets
    wsMenu.Shapes.Range(Array("shape_dashboard")).Select
    
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(52, 56, 205)
        .Transparency = 0
        .Solid
    End With
    
    wsMenu.Shapes.Range(Array("shape_tanggal")).Select
    
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(52, 56, 205)
        .Transparency = 0
        .Solid
    End With
    
    wsMenu.Shapes.Range(Array("shape_total_barang_masuk")).Select
    
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(52, 56, 205)
        .Transparency = 0
        .Solid
    End With
    
    wsMenu.Shapes.Range(Array("shape_total_penjualan_barang")).Select
    
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(52, 56, 205)
        .Transparency = 0
        .Solid
    End With
    
    wsMenu.Range("A1").Select
End Sub

Sub temaUngu()
    Call SetWorksheets
    wsMenu.Shapes.Range(Array("shape_dashboard")).Select
    
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(105, 68, 198)
        .Transparency = 0
        .Solid
    End With
    
    wsMenu.Shapes.Range(Array("shape_tanggal")).Select
    
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(105, 68, 198)
        .Transparency = 0
        .Solid
    End With
    
    wsMenu.Shapes.Range(Array("shape_total_barang_masuk")).Select
    
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(105, 68, 198)
        .Transparency = 0
        .Solid
    End With
    
    wsMenu.Shapes.Range(Array("shape_total_penjualan_barang")).Select
    
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(105, 68, 198)
        .Transparency = 0
        .Solid
    End With
    
    wsMenu.Range("A1").Select
End Sub

Sub temaHitam()
    Call SetWorksheets
    wsMenu.Shapes.Range(Array("shape_dashboard")).Select
    
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        '.ForeColor.RGB = RGB(16, 23, 39)
        '.ForeColor.RGB = RGB(39, 40, 45)
        .ForeColor.RGB = RGB(29, 29, 66)
        .Transparency = 0
        .Solid
    End With
    
    wsMenu.Shapes.Range(Array("shape_tanggal")).Select
    
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        '.ForeColor.RGB = RGB(16, 23, 39)
        '.ForeColor.RGB = RGB(39, 40, 45)
        .ForeColor.RGB = RGB(29, 29, 66)
        .Transparency = 0
        .Solid
    End With
    
    wsMenu.Shapes.Range(Array("shape_total_barang_masuk")).Select
    
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        '.ForeColor.RGB = RGB(16, 23, 39)
        '.ForeColor.RGB = RGB(39, 40, 45)
        .ForeColor.RGB = RGB(29, 29, 66)
        .Transparency = 0
        .Solid
    End With
    
    wsMenu.Shapes.Range(Array("shape_total_penjualan_barang")).Select
    
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        '.ForeColor.RGB = RGB(16, 23, 39)
        '.ForeColor.RGB = RGB(39, 40, 45)
        .ForeColor.RGB = RGB(29, 29, 66)
        .Transparency = 0
        .Solid
    End With
    
    wsMenu.Range("A1").Select
End Sub
