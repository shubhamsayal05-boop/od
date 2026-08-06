Attribute VB_Name = "ConfigSetting"
Attribute VB_Base = "0{A4A99893-BD25-47DC-B78E-7AE72AD24F73}{1AD69F07-E921-4F31-B44B-CA20ABFAFBC2}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private lineCheck As Boolean
Private ligneSDV As Long
Private ligneConfig As Long
Private Sub ComboBox19_Enter()
    If Len(Me.TextBox24) = 0 Then
                MsgBox "Choisir d'abord une reference de ligne", vbCritical, "ODRIV"
                Me.ComboBox19.Locked = True
    Else
                Me.ComboBox19.Locked = False
    End If
End Sub

Private Sub ComboBox20_Enter()
   If Len(Me.TextBox25) = 0 Then
                MsgBox "Choisir d'abord une reference de colonne gauche", vbCritical, "ODRIV"
                Me.ComboBox20.Locked = True
    Else
                Me.ComboBox20.Locked = False
    End If
End Sub

Private Sub ComboBox21_Enter()
        If Len(Me.TextBox26) = 0 Then
                MsgBox "Choisir d'abord une reference de colonne droite", vbCritical, "ODRIV"
                Me.ComboBox21.Locked = True
    Else
                Me.ComboBox21.Locked = False
    End If
End Sub

Private Sub CommandButton2_Click()
        If Len(Me.TextBox2.Value) = 0 Then
            MsgBox "REMPLIR SDV", vbCritical, "ODRIV"
         ElseIf Len(Me.TextBox22.Value) = 0 Then
            MsgBox "REMPLIR NOM DE CONFIGURATION", vbCritical, "ODRIV"
       
        Else
            TraitementOperation
         End If
End Sub

Private Sub CommandButton24_Click()
    Call VCA(Me.TextBox24)
    SELECTFIELD.Show
End Sub

Private Sub CommandButton25_Click()
    SELECTFIELD.Label2.Caption = "Selectionner Colonne Gauche"
    Call VCA(Me.TextBox25)
    SELECTFIELD.Show
End Sub

Private Sub CommandButton26_Click()
    SELECTFIELD.Label2.Caption = "Selectionner Colonne Droite"
    Call VCA(Me.TextBox26)
    SELECTFIELD.Show
End Sub



Private Sub UserForm_Initialize()
      InitListe
End Sub

Function InitListe()
Dim i As Integer
        
Dim TableField(4) As Object
Dim tabV(4) As String
Set TableField(1) = Me.ListView15
Set TableField(2) = Me.ListView16
Set TableField(3) = Me.ListView17
Set TableField(4) = Me.ListView18

tabV(1) = "ENGINE"
tabV(2) = "GEARBOX"
tabV(3) = "NBGEAR"
tabV(4) = "AREA"
For i = 1 To UBound(TableField)
        Call RemplilListe(TableField(i), tabV(i))
Next i

Call RemplissageColonne(Me.ComboBox20)
Call RemplissageColonne(Me.ComboBox21)
Call RemplissageLigne(Me.ComboBox19)

End Function
Function RemplilListe(ListV As Object, Rg As String)
    Dim c As Range
   
   
     With ThisWorkbook.Worksheets("CONFIGURATIONS")
         Set c = .Range(Rg)
         Set c = c.Offset(1, 0)
         While c.Value <> ""
             If Rg = "VERSION" Then
                  ListV.AddItem "V" & c.Value
             ElseIf Rg = "VEHICLE" Then
                  ListV.AddItem c.Value
             ElseIf Rg = "MILESTONE" Then
                  ListV.AddItem c.Value
             ElseIf Rg = "AREA" Then
                  ListV.AddItem c.Value
             ElseIf Rg = "ENGINE" Then
                  ListV.AddItem c.Value
             ElseIf Rg = "GEARBOX" Then
                  ListV.AddItem c.Value
             ElseIf Rg = "NBGEAR" Then
                  ListV.AddItem c.Value
             End If
             
             Set c = c.Offset(1, 0)
             
         Wend
    End With
End Function




Function VCA(o As Object)
        
        With o
              SELECTFIELD.SeleC (o.text)
       End With
   
       SELECTFIELD.SetListeS (o.Name)
End Function
 Function TraitementOperation()
            lineCheck = False
            If IsCreate = True And Me.editing.Value <> "OK" Then
                    MsgBox "Nom de Configuration Déjà Donné", vbCritical, "ODRIV"
            Else
                    If Me.editing.Value <> "OK" Then createSDV
                    RemplissageDonnee
            End If
 End Function

Function IsCreate() As Boolean
    Dim DernL As Long
    Dim v
    Dim i As Integer
    Dim comparVal As String
    IsCreate = False
    
    DernL = DernLigne
    ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Outline.ShowLevels RowLevels:=2
    v = ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Range("A2:AM" & DernL).Value
    
    For i = 2 To UBound(v, 1)
        If Len(v(i, 1)) > 0 And UCase(CStr(v(i, 1))) = UCase(Me.TextBox2) Then
                ligneSDV = i + 1
                If ligneSDV = DernL Then
                    lineCheck = True
                    Exit Function
                End If
                If i + 1 <= UBound(v, 1) Then i = i + 1
                If i <= UBound(v, 1) And Len(v(i, 1)) = 0 Then
                    comparVal = v(i, 1)
                    While Len(comparVal) = 0 And i <= UBound(v, 1)
                          If Len(v(i, 2)) > 0 Then
                             If (Right(v(i, 2), ((Len(v(i, 2)) - 1) - InStr(1, v(i, 2), ": ")))) = UCase(Me.TextBox22.Value) Then
                                IsCreate = True
                                Exit Function
                             End If
                          End If
                          i = i + 1
                          If i <= UBound(v, 1) Then comparVal = v(i, 1)
                    Wend
                Else
                    lineCheck = True
                End If
            
        End If
    Next i
     ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Outline.ShowLevels RowLevels:=1
    Erase v
End Function

Sub createSDV()
        Dim i As Integer
        Dim v As String
        Dim LastRS As Long
        
        Application.EnableEvents = False
        With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
                .Outline.ShowLevels RowLevels:=2
                If lineCheck = True Then
                        Call GroupConfig(False)
                        ThisWorkbook.sheets("CONFIGURATIONS ARRAY").Rows("2:32").Copy
                        .Rows(ligneSDV + 1).Insert Shift:=xlDown
'                        With ActiveSheet.Outline
'                            .AutomaticStyles = False
'                            .SummaryRow = xlAbove
'                            .SummaryColumn = xlRight
'                        End With
                        .Cells(ligneSDV + 1, 2) = "Config n°1 : " & replace(UCase(Me.TextBox22), ":", " ")
                        Me.TextBox22 = replace(UCase(Me.TextBox22), ":", " ")
                        ligneConfig = ligneSDV + 1
                        Call GroupConfig(True)
              Else
                    i = ligneSDV + 1
                    
'                    While Len(.Cells(i, 1)) = 0 And i <= .Range("B65000").End(xlUp).Row
                    LastRS = DernLigne
                    While Len(.Cells(i, 1)) = 0 And i <= LastRS
                             If Left(.Cells(i, 2), 9) = "Config n°" Then
                                v = Left(.Cells(i, 2), InStr(1, .Cells(i, 2), ":"))
                                v = replace(v, "Config n°", "")
                                v = replace(v, " ", "")
                                v = replace(v, ":", "")
                             End If
                             i = i + 1
                    Wend
                    Call GroupConfig(False)
                    If i >= LastRS Then
                        ThisWorkbook.sheets("CONFIGURATIONS ARRAY").Rows("1:32").Copy
                    Else
                        ThisWorkbook.sheets("CONFIGURATIONS ARRAY").Rows("2:32").Copy
                    End If
                    .Rows(i).Insert Shift:=xlDown
                    If i >= LastRS Then i = i + 1
                    .Cells(i, 2) = "Config n°" & val(v) + 1 & " : " & replace(UCase(Me.TextBox22), ":", " ")
                    Me.TextBox22 = replace(UCase(Me.TextBox22), ":", " ")
                    ligneConfig = i + 1
                    Call GroupConfig(True)
              End If

                .Outline.ShowLevels RowLevels:=1
        End With
        Application.EnableEvents = True

End Sub
Function FTNum() As Long
 
    Dim v
    Dim i As Integer
    Dim comparVal As String
    Dim DernL As Long
    FTNum = 0
       
    DernL = DernLigne
    ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Outline.ShowLevels RowLevels:=2
    v = ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Range("A2:AM" & DernL).Value

    For i = 2 To UBound(v, 1)
        If Len(v(i, 1)) > 0 And UCase(CStr(v(i, 1))) = UCase(Me.TextBox2) Then
                If i + 1 <= UBound(v, 1) Then i = i + 1
                If i <= UBound(v, 1) And Len(v(i, 1)) = 0 Then
                    comparVal = v(i, 1)
                    While Len(comparVal) = 0 And i <= UBound(v, 1)
                          If Len(v(i, 2)) > 0 Then
                             If UCase((Right(v(i, 2), ((Len(v(i, 2)) - 1) - InStr(1, v(i, 2), ": "))))) = UCase(Me.TextBox22.Value) Then
                                FTNum = i + 1
                                Exit Function
                             End If
                          End If
                          i = i + 1
                          If i <= UBound(v, 1) Then comparVal = v(i, 1)
                    Wend
               
                End If
            
        End If
    Next i
     ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Outline.ShowLevels RowLevels:=1
    Erase v
End Function
Function RemplissageDonnee()
        Dim r As Long
        Dim i As Integer
        Dim j As Integer
        With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
             .Outline.ShowLevels RowLevels:=2
'             r = .Range("A65000").End(xlUp).Row + 1
             Application.EnableEvents = False
             r = FTNum
'            For i = 1 To 7 Step 2
'                       For j = 3 To .Cells(r + i, .Columns.Count).End(xlToLeft).Column
'                           CompareV (.Cells(r + i, j))
'                     Next j
'            Next i
          i = r + 1
          While Application.CountA(.Range("B" & i & ":G" & i)) > 0 Or .Range("C" & i).Interior.color = 855309
                If UCase(.Range("B" & i).Value) = "ENGINE TYPE" Then
                    Call checkComparaison(i + 1, 2, 1)
                End If
                
                If UCase(.Range("E" & i).Value) = "GEARBOX TYPE" Then
                     Call checkComparaison(i + 1, 5, 2)
                End If
                
                If UCase(.Range("B" & i).Value) = "NUMBER OF GEARS" Then
                    Call checkComparaison(i + 1, 2, 3)
                End If
                
                If UCase(.Range("E" & i).Value) = "AREA" Then
                    Call checkComparaison(i + 1, 5, 4)
                End If
               
                i = i + 1
            Wend
            
            .Cells(r + 27, 3) = IIf(Len(Me.TextBox25) > 0, Me.TextBox25, "X")
            .Cells(r + 27, 4) = IIf(Len(Me.TextBox26) > 0, Me.TextBox26, "X")
            .Cells(r + 27, 2) = IIf(Len(Me.TextBox24) > 0, Me.TextBox24, "X")
            
            .Cells(r + 29, 3) = IIf(Len(Me.ComboBox20) > 0, Me.ComboBox20, "X")
            .Cells(r + 29, 2) = IIf(Len(Me.ComboBox19) > 0, Me.ComboBox19, "X")
            .Cells(r + 29, 4) = IIf(Len(Me.ComboBox21) > 0, Me.ComboBox21, "X")
            
            Application.EnableEvents = True
            .Outline.ShowLevels RowLevels:=1
            .Select
            MsgBox "Paramètres Ajoutés", vbInformation, "ODRIV"
            Unload Me
            Exit Function

        End With
End Function
Function checkComparaison(i As Long, col As Integer, id As Integer)
    Dim r As Range
   
    With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
        Set r = .Cells(i, col + 1)
        
        While r.Interior.color = 855309
            If Len(r.Value) >= 0 Then
                Call CompareV(r, id)
            End If
            Set r = r.Offset(1, 0)
        Wend
        
    End With
End Function

Function CompareV(r As Range, id As Integer)
        Dim i As Integer

        With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
              If id = 1 Then
                       For i = 0 To Me.ListView15.ListCount - 1
                            If Me.ListView15.Selected(i) = False And UCase(Me.ListView15.list(i)) = UCase(r.Value) Then
                                 r.Offset(0, 1) = ""
                            ElseIf Me.ListView15.Selected(i) = True And UCase(Me.ListView15.list(i)) = UCase(r.Value) Then
                                r.Offset(0, 1) = "X"
                            End If
                       Next i
                       
            ElseIf id = 2 Then
                       For i = 0 To Me.ListView16.ListCount - 1
                            If Me.ListView16.Selected(i) = False And UCase(Me.ListView16.list(i)) = UCase(r.Value) Then
                                 r.Offset(0, 1) = ""
                            ElseIf Me.ListView16.Selected(i) = True And UCase(Me.ListView16.list(i)) = UCase(r.Value) Then
                                r.Offset(0, 1) = "X"
                            End If
                       Next i
                       
             ElseIf id = 3 Then
                       For i = 0 To Me.ListView17.ListCount - 1
                            If Me.ListView17.Selected(i) = False And UCase(Me.ListView17.list(i)) = UCase(r.Value) Then
                                 r.Offset(0, 1) = ""
                            ElseIf Me.ListView17.Selected(i) = True And UCase(Me.ListView17.list(i)) = UCase(r.Value) Then
                                r.Offset(0, 1) = "X"
                            End If
                       Next i
                       
            ElseIf id = 4 Then
                       For i = 0 To Me.ListView18.ListCount - 1
                            If Me.ListView18.Selected(i) = False And UCase(Me.ListView18.list(i)) = UCase(r.Value) Then
                                 r.Offset(0, 1) = ""
                            ElseIf Me.ListView18.Selected(i) = True And UCase(Me.ListView18.list(i)) = UCase(r.Value) Then
                                r.Offset(0, 1) = "X"
                            End If
                       Next i
                       
            End If
        End With
End Function

Function RemplissageColonne(TableField As Object)
    TableField.AddItem "INTERVALLE"
    TableField.AddItem "INFERIORITE OU EGALITE"
    TableField.AddItem "EGALITE"
End Function
Function RemplissageLigne(TableField As Object)
    TableField.AddItem "INTERVALLE"
    TableField.AddItem "SUPERIORITE"
    TableField.AddItem "SUPERIORITE OU EGALITE"
    TableField.AddItem "INFERIORITE"
    TableField.AddItem "INFERIORITE OU EGALITE"
    TableField.AddItem "EGALITE"
End Function

Function GroupConfig(action As Boolean) As Long
        Dim i As Integer
        Dim n As Integer
        Dim LastRS As Long
        
        LastRS = DernLigne
        With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
                 i = ligneSDV
                 n = i + 1
'                While Len(.Cells(n, 1)) = 0 And n <= .Range("B65000").End(xlUp).Row
                While Len(.Cells(n, 1)) = 0 And n <= LastRS
                         n = n + 1
                Wend
                If n <= LastRS Then n = n - 1
                On Error Resume Next
                If n > i Then
'                    n = n - 1
                    If action = False Then
                        .Outline.ShowLevels RowLevels:=2
                        .Rows(i + 1 & ":" & n).Ungroup
                    ElseIf action = True Then
                        .Rows(i + 1 & ":" & n).Group
                     End If
                End If
                On Error GoTo 0

     End With
End Function
Function DernLigne()
        Dim lastr As Long
        Dim derniereColonne As Integer
        Dim cm As Integer
        
        With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
            .Outline.ShowLevels RowLevels:=2
            lastr = 0
            derniereColonne = 30
            For cm = 1 To derniereColonne
                If .Cells(.Rows.Count, cm).End(xlUp).row > lastr Then lastr = .Cells(.Rows.Count, cm).End(xlUp).row
            Next cm
            DernLigne = lastr
           .Outline.ShowLevels RowLevels:=1
        End With
        
End Function















