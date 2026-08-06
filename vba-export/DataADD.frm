Attribute VB_Name = "DataADD"
Attribute VB_Base = "0{FE6BE862-DC3F-41FD-94CF-59AF768BE014}{34182397-8D1E-4214-A4A7-F2DAD6DC98C9}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub CommandButton6_Click()
    Dim i As Long
    
    If Len(Me.POrdre) > 0 And Len(Me.Pcolonne) > 0 Then
         With ThisWorkbook.sheets("DEFINITION SDV")
                .Range("A2:E3").Copy Destination:=ThisWorkbook.Worksheets("DEFINITION SDV").Cells(ThisWorkbook.Worksheets("DEFINITION SDV").Range("A65000").End(xlUp).row + 1, 1)
                
                i = ThisWorkbook.Worksheets("DEFINITION SDV").Range("A65000").End(xlUp).row
                .Range("A" & i - 1 & ":A" & i) = Me.POrdre
                .Range("B" & i - 1) = Me.Pcolonne
                DataLoad.code.Caption = Me.POrdre & "--" & Me.Pcolonne
                Unload Me
                DataLoad.Show
                
        End With
    End If
End Sub

Private Sub UserForm_Initialize()
   Me.POrdre = _
  ThisWorkbook.Worksheets("DEFINITION SDV").Range("A" & ThisWorkbook.Worksheets("DEFINITION SDV").Range("A65000").End(xlUp).row) + 1
  addSDVList
End Sub

Function addSDVList()
  
    Dim v
    Dim i As Long
    
    v = ThisWorkbook.sheets("structure").UsedRange.Columns(2).Value
    For i = 2 To UBound(v, 1)
        If Len(v(i, 1)) > 0 Then
          Me.Pcolonne.AddItem UCase(v(i, 1))
        End If
    Next i
    Erase v
  
   
End Function
