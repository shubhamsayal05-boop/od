Attribute VB_Name = "Feuil1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Private Sub Worksheet_Change(ByVal Target As Range)

      Dim saveValeur As String
      Dim celActu As Range
      Dim cColRow(2) As Integer
      Dim Lastval, LastCell As Range
      Dim Lst
      Dim Tst
      Dim cm

    If Cells(Rows.Count, 1).End(xlUp).row < 2 Then Exit Sub


    If Target.Cells.Count > 1 Then
          On Error Resume Next
          Application.EnableEvents = False
          Application.Undo
          Application.EnableEvents = True
          Exit Sub
    End If

    If Application.Intersect(Range("A1:A2000"), Target) Is Nothing Then
         On Error Resume Next
         Application.EnableEvents = False
          Application.Undo
          Application.EnableEvents = True

          Exit Sub
    End If

    If InStr(1, Cells(Target.row, 2), ".") <> 0 Then
         On Error Resume Next
         Application.EnableEvents = False
          Application.Undo
          Application.EnableEvents = True
          Exit Sub
    End If



    Lst = 2
    For cm = 1 To Lst
       If Cells(Rows.Count, cm).End(xlUp).row > Tst Then Tst = Cells(Rows.Count, cm).End(xlUp).row
    Next cm

    If Tst = 1 Then
        Exit Sub
    ElseIf Cells(Rows.Count, 1).End(xlUp).row = 1 Then
        Exit Sub
    End If

    If Target.row >= 1 And Target.row <= DernLigne And Target.Column < 2 Then
           cColRow(1) = dernColonne
           cColRow(2) = DernLigne
           Application.ScreenUpdating = False
           Application.EnableEvents = False
           saveValeur = Target.Value
           Set celActu = ActiveCell
           On Error Resume Next
           Application.Undo
         If cColRow(1) <> dernColonne Or cColRow(2) <> DernLigne Then
                 Application.EnableEvents = True
                 Application.ScreenUpdating = True

                 Exit Sub
          End If

          If Application.CountA(Range("A" & Target.row & ":B" & Target.row)) = 0 Then
                Application.EnableEvents = True
                Application.ScreenUpdating = True
                Exit Sub
          End If

           Lastval = Target.Value
           Target.Value = saveValeur
           Set LastCell = Target
           celActu.Select
           Application.ScreenUpdating = True
    Else
          Application.EnableEvents = False
          Application.Undo
          Application.EnableEvents = True

          Exit Sub
    End If


On Error GoTo geserreur

        If LastCell.Value <> Lastval Then
          Application.EnableEvents = False
          If Application.WorksheetFunction.CountIf(Columns(LastCell.Column), LastCell.Value) = 1 Then
                   Call renameRatingGroup(CStr(Lastval), LastCell.Value)
          Else
                MsgBox "Valeur Existe Déjà", vbCritical, "ODRIV"
                LastCell.Value = Lastval
          End If
          Application.EnableEvents = True

      End If
      Application.EnableEvents = True

geserreur:

      If ERR.Number <> 0 Then
           MsgBox ERR.description, vbCritical, "BILIQUE"
           Application.EnableEvents = True
           Application.Cursor = xlDefault
      End If

End Sub

Function dernColonne() As Long
    dernColonne = 8
End Function
Function DernLigne() As Long
    DernLigne = Cells(Rows.Count, 1).End(xlUp).row
End Function
Function renameRatingGroup(Lasts As String, News As String)
    Dim i As Long
    Dim j As Long
    i = 23
    Dim x As Integer
    
   x = getLastRowRating
    With ThisWorkbook.sheets("RATING")
            While i <= x
                If .Cells(i, 2).Value = Lasts Then
                    .Cells(i, 2).Value = News
                    Exit Function
                End If
                i = i + 1
            Wend
    End With

End Function









