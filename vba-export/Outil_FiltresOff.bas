Attribute VB_Name = "Outil_FiltresOff"
' Cellule Office

Option Explicit

Sub FiltresOff(Optional fichier As Variant)
    If IsMissing(fichier) = True Then
        If sheetExists("DATA") = True Then
        With ThisWorkbook.Worksheets("DATA").Rows(1)
            .AutoFilter Field:=3
            .AutoFilter Field:=4
            .AutoFilter Field:=5
'            .AutoFilter Field:=64
'            .AutoFilter Field:=62
'            .AutoFilter Field:=66
        End With
        End If
    ElseIf IsMissing(fichier) = False Then
        With Workbooks(fichier).Worksheets("TRIE").Cells
            .AutoFilter Field:=3
            .AutoFilter Field:=4
            .AutoFilter Field:=5
'            .AutoFilter Field:=62
'            .AutoFilter Field:=64
'            .AutoFilter Field:=66
        End With
    End If
End Sub


