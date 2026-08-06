Attribute VB_Name = "SortDEFS"
Public LeverCount() As String
Sub AdsS()
    DataADD.Show
End Sub
Sub Eds()
    DataLIST.CommandButton2.Visible = False
    DataLIST.Show
End Sub
Sub DeLC()
    DataLIST.CommandButton1.Visible = False
    DataLIST.Show
End Sub
Function SortField()
     With ThisWorkbook.Worksheets("DEFINITION SDV")
        .Sort.SortFields.Clear
        .Sort.SortFields.Add key:=.Range("A2:A500") _
        , SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortNormal
    End With
    With ThisWorkbook.Worksheets("DEFINITION SDV").Sort
        .SetRange ActiveWorkbook.Worksheets("Feuil2").Range("A1:E500")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Function

Function getLever(onglet As String) As Boolean
        Dim r As Range
        Dim i As Long
        i = 1
        Erase LeverCount
        getLever = False
        Set r = ThisWorkbook.sheets("structure").Columns(2).Find(What:=onglet, LookIn:=xlFormulas, _
                    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                    MatchCase:=False, SearchFormat:=False)
                    Set r = r.Offset(1, 2)
                    If Not r Is Nothing Then
                            While Len(r.Value) > 0
                                If r.Value = "New Selector_Lever_Position" Or r.Value = "Old Selector_Lever_Position" Or r.Value = "Selector_Lever_Position" Then
                                       ReDim Preserve LeverCount(i)
                                       LeverCount(i) = r.Value
                                       i = i + 1
                                       getLever = True
                                End If
                                Set r = r.Offset(1, 0)
                            Wend
                    End If
End Function





