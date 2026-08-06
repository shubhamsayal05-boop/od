Attribute VB_Name = "Feuil4"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Deactivate()
     On Error Resume Next
    Application.EnableEvents = False
    checkChange
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim tempCell As Range
    Dim val As Variant

    On Error GoTo Fin
    Application.EnableEvents = False
    
    Set tempCell = Me.Range("Temperature").MergeArea.Cells(1, 1).Offset(1, 0)
    
    
    If Not Intersect(Target, tempCell) Is Nothing Then

        
        val = tempCell.Value

        With ThisWorkbook.Worksheets("GRAPHIQUES")
            .Range("C2").Value = "Cold(<" & val & "°C)"
            .Range("D2").Value = "Hot(>=" & val & "°C)"
            .Range("O2").Value = "Cold(<" & val & "°C)"
            .Range("P2").Value = "Hot(>=" & val & "°C)"
        End With

    End If
    
Fin:
    Application.EnableEvents = True

End Sub



