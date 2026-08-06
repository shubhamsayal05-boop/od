Attribute VB_Name = "SELECTFIELD"
Attribute VB_Base = "0{A659A89C-C66C-4A5F-8E99-9BA6D0469606}{3726BB98-F649-435A-8BBA-F1F2A0E3DE2D}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private colon As Object
Private SELE As Object

'Function loadMe()
'
'  Dim v
'  Dim i As Integer
'
'
'   v = ThisWorkbook.sheets("structure").UsedRange.Columns(4).Value
'    For i = 2 To UBound(v, 1)
'        If Len(v(i, 1)) > 0 And ExistsSCrip(CStr(v(i, 1))) = False Then
'            Me.ListeValeur.AddItem v(i, 1)
'            If InStr(1, Me.TextBox2.Value & ";", ";" & v(i, 1) & ";") <> 0 Then
'                Me.ListeValeur.Selected(Me.ListeValeur.ListCount - 1) = True
'            End If
'        End If
'    Next i
'    Erase v
'
'End Function

Private Sub CommandButton1_Click()
        Dim i As Integer
        Dim j As Integer
        j = 1
       
       If TCount > 1 Then
            MsgBox "Attention une Selection autorisée ", vbCritical, "ODRIV"
        Else
'             LView.ListItems.Clear
             Call EraseFL
             For i = 0 To Me.ListeValeur.ListCount - 1
                If Me.ListeValeur.Selected(i) = True Then
                         LView.text = Me.ListeValeur.list(i)
'                         PARAMETRAGES.SetEditing (True)
'                         j = j + 1
                End If
            Next i
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
Function EraseFL()
    LView.text = ""
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


Function SeleC(v As String)
        Me.TextBox2.Value = Me.TextBox2.Value & ";" & v
End Function

Function SetListeS(vT As String)
      Me.TextBox1.Value = vT
End Function
Function LView() As Object
    If Me.TextBox1.Value = "TextBox24" Then
       Set LView = ConfigSetting.TextBox24
    ElseIf Me.TextBox1.Value = "TextBox25" Then
     Set LView = ConfigSetting.TextBox25
    ElseIf Me.TextBox1.Value = "TextBox26" Then
     Set LView = ConfigSetting.TextBox26
    End If
End Function

Function loadME()
    Dim DernL As Long
    Dim v
    Dim i As Integer
    Dim comparVal As String
    
    Dim valSDV
    IsCreate = False
    
    valSDV = UCase(ConfigSetting.TextBox2.Value)
    DernL = DernLigne
    ThisWorkbook.sheets("structure").Outline.ShowLevels RowLevels:=2
    v = ThisWorkbook.sheets("structure").Range("A2:AM" & DernL).Value
    
    For i = 2 To UBound(v, 1)
        If Len(v(i, 2)) > 0 And UCase(CStr(v(i, 2))) = valSDV Then
                i = i + 1
                While i <= UBound(v, 1)
                       If Len(v(i, 4)) > 0 Then
                            Me.ListeValeur.AddItem v(i, 4)
                            If InStr(1, Me.TextBox2.Value & ";", ";" & v(i, 4) & ";") <> 0 Then
                                Me.ListeValeur.Selected(Me.ListeValeur.ListCount - 1) = True
                            End If
                        End If
                       i = i + 1
                 Wend
                Exit Function
        End If
    Next i
     ThisWorkbook.sheets("structure").Outline.ShowLevels RowLevels:=1
    Erase v
End Function

Function DernLigne()
        Dim lastr As Long
        Dim derniereColonne As Integer
        Dim cm As Integer
        With ThisWorkbook.sheets("structure")
            .Outline.ShowLevels RowLevels:=2
            lastr = 0
            derniereColonne = 5
            For cm = 1 To derniereColonne
                If .Cells(.Rows.Count, cm).End(xlUp).row > lastr Then lastr = .Cells(.Rows.Count, cm).End(xlUp).row
            Next cm
            DernLigne = lastr
           .Outline.ShowLevels RowLevels:=1
        End With
End Function


