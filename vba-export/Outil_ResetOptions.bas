Attribute VB_Name = "Outil_ResetOptions"
' Cellule Office

Option Explicit

Sub ResetOptions(Optional sd As String)
    Dim v
    Dim i As Long
    
'    If Len(sd) > 0 Then
'        ThisWorkbook.sheets(sd).Shapes("AllShifts").ControlFormat.Value = xlOn
'        ThisWorkbook.sheets(sd).Shapes("ColorScale").ControlFormat.Value = xlOn
'    Else
'        V = ThisWorkbook.sheets("structure").UsedRange.Columns(2).Value
'        For i = 2 To UBound(V, 1)
'            If Len(V(i, 1)) > 0 And sheetExists(V(i, 1)) = True Then
'            ThisWorkbook.sheets(V(i, 1)).Shapes("AllShifts").ControlFormat.Value = xlOn
'            ThisWorkbook.sheets(V(i, 1)).Shapes("ColorScale").ControlFormat.Value = xlOn
'            End If
'        Next i
'        Erase V
'    End If

End Sub


