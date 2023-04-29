Attribute VB_Name = "ModulePath"
Option Explicit

Public Function getPath(endPoint As String) As String
    getPath = Application.ActiveWorkbook.Path + endPoint
End Function
