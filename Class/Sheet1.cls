Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Not Intersect(Range("E1:W50"), Target) Is Nothing Then
        Call resetActive
    End If
End Sub
