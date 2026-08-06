Attribute VB_Name = "Outil_FiltrerPopul"
' Cellule Office

Option Explicit

Sub PopulFilter(ByVal onglet As String)

    Dim lcol As Integer

    lcol = ThisWorkbook.sheets(onglet).Range("A6:BA6").Find("*", , , , xlByRows, xlPrevious).Column
    ThisWorkbook.sheets(onglet).Range(Cells(6, 13 - 1), Cells(6, lcol)).AutoFilter

End Sub

Sub PopulFilterDyn(ByVal onglet As String)

    Dim lcol As Integer
    lcol = ThisWorkbook.sheets(onglet).Range("BH6:GG6").Find("*", , , , xlByRows, xlPrevious).Column
    ThisWorkbook.sheets(onglet).Range(Cells(6, 72 - 1), Cells(6, lcol)).AutoFilter

End Sub



