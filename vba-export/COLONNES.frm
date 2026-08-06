Attribute VB_Name = "COLONNES"
Attribute VB_Base = "0{8D68D071-0382-4F22-901C-EE9DC59A28A8}{A784C77C-FC2F-4C26-965B-0518A06809EB}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private colon As Object
Private SELE As Object
Function loadME()
  
  Dim v
  Dim i As Integer
  'Trier
    ThisWorkbook.Worksheets("ENTETE_COLONNE").Sort.SortFields.Clear
    ThisWorkbook.Worksheets("ENTETE_COLONNE").Sort.SortFields.Add key:=Range( _
        "A2:A3000"), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:= _
        xlSortNormal
    With ThisWorkbook.Worksheets("ENTETE_COLONNE").Sort
        .SetRange Range("A1:A3000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
   v = ThisWorkbook.sheets("ENTETE_COLONNE").UsedRange.Columns(1).Value
    For i = 2 To UBound(v, 1)
        If Len(v(i, 1)) > 0 And ExistsSCrip(CStr(v(i, 1))) = False Then
            Me.ListeValeur.AddItem v(i, 1)
        End If
    Next i
    Erase v
    
End Function

Private Sub CommandButton1_Click()
        Dim i As Integer
        Dim j As Integer
        j = 1
        If TCount = 0 Then
            MsgBox "Aucune Selection", vbCritical, "ODRIV"
       Else
            If DataEdit.code.Caption = "1" Then
                For i = 0 To Me.ListeValeur.ListCount - 1
                    If Me.ListeValeur.Selected(i) = True Then
                        DataEdit.Pcolonne.Value = Me.ListeValeur.list(i)
                    End If
                Next i
            ElseIf DataEdit.code.Caption = "2" Then
                For i = 0 To Me.ListeValeur.ListCount - 1
                    If Me.ListeValeur.Selected(i) = True Then
                        DataEdit.PValeur.Value = Me.ListeValeur.list(i)
                    End If
                Next i
            End If
            Unload Me
        End If
End Sub

Private Sub UserForm_Activate()
loadME
End Sub

Function ExistsSCrip(onglet As String) As Boolean
        If colon Is Nothing Then Set colon = CreateObject("Scripting.Dictionary")
        If Not colon.Exists(UCase(onglet)) Then
            colon.Add key:=UCase(onglet), Item:=UCase(onglet)
            ExistsSCrip = False
        Else
            ExistsSCrip = True
        End If
End Function
Function ExistsSELE(onglet As String) As Boolean
        If SELE Is Nothing Then
            ExistsSELE = False
            Exit Function
        End If
        If Not SELE.Exists(UCase(onglet)) Then
            ExistsSELE = False
        Else
            ExistsSELE = True
        End If
End Function
Function TCount() As Integer
     Dim i As Integer
     TCount = 0
     For i = 0 To Me.ListeValeur.ListCount - 1
                If Me.ListeValeur.Selected(i) = True Then
                    TCount = TCount + 1
                End If
    Next i
End Function
