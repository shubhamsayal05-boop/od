Attribute VB_Name = "DataLoad"
Attribute VB_Base = "0{3C6EF337-6D16-49B6-9A51-EAEE3B4500C2}{A57F0047-1342-4D0D-ADAD-47F9FF464463}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub CommandButton1_Click()
      Dim j As Long
      Dim v
      If Me.ListView11.ListItems.Count = 0 Then
          MsgBox "PARAMETRES VIDES", vbCritical, "ODRIV"
      Else
            v = ThisWorkbook.sheets("DEFINITION SDV").UsedRange.Columns("A:E").Value
            With ThisWorkbook.Worksheets("DEFINITION SDV")
'                   .Outline.ShowLevels RowLevels:=2
                    For j = 1 To UBound(v, 1)
                          If IsNumeric(v(j, 1)) = True And Len(v(j, 1)) > 0 And Len(v(j, 3)) = 0 Then
                             If v(j, 1) & "--" & v(j, 2) = Me.code.Caption Then
                                getLINE (j + 2)
                                Unload Me
                                Application.CutCopyMode = False
                                Exit Sub
                              End If
                          End If
                  Next j
              End With
     End If
End Sub

Private Sub CommandButton3_Click()
        Dim v As Integer
        v = SelA
         If v <> 999 Then
            Me.ListView11.ListItems.Remove (v)
         End If
End Sub

Private Sub CommandButton4_Click()
        Me.ids.Caption = 0
        DataEdit.Show
End Sub

Private Sub ListView11_DblClick()
        Dim v As Integer
        v = SelA
         If v <> 999 Then
            Me.ids.Caption = v
            DataEdit.Pcolonne.Value = Me.ListView11.ListItems(v)
            DataEdit.ComboBox2.Value = Me.ListView11.ListItems(v).SubItems(1)
            DataEdit.PValeur.Value = Me.ListView11.ListItems(v).SubItems(2)
            DataEdit.POrdre.Value = Me.ListView11.ListItems(v).SubItems(3)
            DataEdit.Show
         End If
End Sub

Private Sub ListView11_ItemClick(ByVal Item As MSComctlLib.ListItem)
        If SelA <> 999 Then Me.ids.Caption = SelA
End Sub

Private Sub UserForm_Initialize()
        Call getDatas
End Sub

Function getDatas()
      Dim j As Long
      Dim i As Integer
      Dim li As Variant
      Dim v
      
       v = ThisWorkbook.sheets("DEFINITION SDV").UsedRange.Columns("A:E").Value
      With ThisWorkbook.Worksheets("DEFINITION SDV")
'             .Outline.ShowLevels RowLevels:=2
              For j = 1 To UBound(v, 1)
                    If IsNumeric(v(j, 1)) = True And Len(v(j, 1)) > 0 And Len(v(j, 3)) = 0 Then
                       If v(j, 1) & "--" & v(j, 2) = DataLIST.code.Caption Then
                            Me.code.Caption = DataLIST.code.Caption
                            i = j + 2
                            While Len(.Range("C" & i)) > 0
                                Set li = Me.ListView11.ListItems.Add(, , v(i, 2))
                                li.SubItems(1) = v(i, 3)
                                li.SubItems(2) = v(i, 4)
                                li.SubItems(3) = v(i, 5)
                                i = i + 1
                            Wend
                            With Me.ListView11
                                .FullRowSelect = True
                                .View = lvwReport
                                .Gridlines = True
                                .LabelEdit = lvwManual
                            End With
                            Unload DataLIST
                            Exit Function
                       End If
                   End If
              Next j

     End With

End Function
Function SelA() As Integer
Dim i As Long
 SelA = 999

        For i = 1 To Me.ListView11.ListItems.Count
                         If Me.ListView11.ListItems(i).Selected = True Then
                            SelA = i
                            Exit Function
                        End If
      Next i

End Function

Function getLINE(Racine As Integer)
    Dim i As Long
    Dim r As Range
     i = Racine + 1
    With ThisWorkbook.Worksheets("DEFINITION SDV")
            If Len(.Range("A" & Racine)) > 0 Then
                While .Range("A" & i) = .Range("A" & Racine)
                     If r Is Nothing Then
                        Set r = .Range("A" & i)
                     Else
                        Set r = Union(r, .Range("A" & i))
                     End If
                     i = i + 1
                Wend
                If Not r Is Nothing Then
                     r.EntireRow.Delete
                End If
          Else
                .Range("B" & Racine & ":E" & Racine).Borders(1).LineStyle = xlContinuous
                .Range("B" & Racine & ":E" & Racine).Borders(2).LineStyle = xlContinuous
                .Range("B" & Racine & ":E" & Racine).Borders(3).LineStyle = xlContinuous
                .Range("B" & Racine & ":E" & Racine).Borders(4).LineStyle = xlContinuous
          End If
          For i = Me.ListView11.ListItems.Count To 1 Step -1
                   If i <> Me.ListView11.ListItems.Count Then
                        .Rows(Racine).Copy
                        .Rows(Racine).Insert Shift:=xlDown
                  End If
                   .Range("A" & Racine) = .Range("A" & Racine - 1)
                   .Range("B" & Racine) = Me.ListView11.ListItems(i).text
                   .Range("C" & Racine) = Me.ListView11.ListItems(i).ListSubItems(1)
                   .Range("D" & Racine) = Me.ListView11.ListItems(i).ListSubItems(2)
                   .Range("E" & Racine) = Me.ListView11.ListItems(i).ListSubItems(3)
                   .Rows(Racine).OutlineLevel = 2
          Next i
    End With
    

End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
            If Me.ListView11.ListItems.Count = 0 Then
                    Cancel = True
                    MsgBox "PARAMETRES VIDES", vbCritical, "ODRIV"
            End If
    End If
End Sub
