Attribute VB_Name = "ModuleTampilData"
Private Sub ContohTampilData()
    With ListBoxMerekBarang
        .ColumnCount = 2
        .List = wsMerekBarang.Range("A2:B2").CurrentRegion.Value
        .ColumnWidths = "100;150"
        '.ColumnHeads = True
        
        ''Data Model'!$G$20
        '.RowSource = "'Merek Barang'!B2:B" & getBarisMerekBarang
        '.ColumnHeads = True
        '.ColumnWidths = "100;150"
    End With
End Sub

Private Sub ContohTampilData2()
    Dim i As Integer

    ListBoxMerekBarang.Clear
    With ListBoxMerekBarang
        .ColumnCount = 2
        .AddItem
        .List(.ListCount - 1, 0) = "ID Merek Barang"
        .List(.ListCount - 1, 1) = "Merek Barang"
        .ColumnWidths = "100;150"
        .ForeColor = vbBlack
    End With
    
    For i = 2 To getBarisMerekBarang
        If LCase(wsMerekBarang.Cells(i, 2)) Like "*" & LCase(TextBoxCari) & "*" Then
            With ListBoxMerekBarang
                .AddItem
                .List(.ListCount - 1, 0) = wsMerekBarang.Cells(i, 1)
                .List(.ListCount - 1, 1) = wsMerekBarang.Cells(i, 2)
            End With
        End If
    Next
End Sub
