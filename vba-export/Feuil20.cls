Attribute VB_Name = "Feuil20"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Private Sub Worksheet_Activate()
    On Error Resume Next
    Application.EnableEvents = False
    checkChange
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

      Dim cColRow(2) As Integer
      Dim saveValeur As String
      Dim celActu As Range
      Exit Sub
      On Error GoTo geserreur


                          If Target.Cells.Count = 1 And Target.row > 2 And Target.row <= DernLigne And UCase(Cells(Target.row, 1)) <> "SOMME" And Target.Column <= 9 And Target.Interior.color = RGB(255, 255, 255) Then
                                 cColRow(1) = Target.Column
                                 cColRow(2) = DernLigne
                                 Application.ScreenUpdating = False
                                 Application.EnableEvents = False
                                 saveValeur = Target.Value

                                 Set celActu = ActiveCell
                                 On Error Resume Next
                                 Application.Undo
                                 On Error GoTo geserreur
                                  If cColRow(1) > 9 Or cColRow(2) <> DernLigne Or Len(saveValeur) = 0 Then
                                       Application.EnableEvents = True
                                       Application.ScreenUpdating = True
                                       Exit Sub
                                End If

                                 Target.Value = saveValeur
                                 celActu.Select
                                 Application.ScreenUpdating = True
                                 Application.EnableEvents = True
                          Else
                                 If Target.Cells.Count > 1 Then MsgBox "La modification de plusieur cellules n'est pas autorisée ", vbCritical, "ODRIV"
                                Application.EnableEvents = False
                                Application.Undo
                                Application.EnableEvents = True
                                Exit Sub
                          End If




geserreur:

      If ERR.Number <> 0 Then
           MsgBox ERR.description, vbCritical, "ODRIV"
           Application.EnableEvents = True
      End If

End Sub
Function DernLigne() As Long
    DernLigne = Cells(Rows.Count, 1).End(xlUp).row
End Function



