Attribute VB_Name = "AddPowertrain"
Attribute VB_Base = "0{8B09D429-D88A-437E-9BB7-ADA0AA8D4B86}{3A6E9A07-15FC-4E13-9285-36FD65A441EE}"
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

Private Sub CommandButton2_Click()
          
        If Len(Me.TextBox22.Value) = 0 Then
            MsgBox "REMPLIR NOM DE CONFIGURATION", vbCritical, "ODRIV"
        Else
            TraitementOperation
         End If
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
    Dim v
    Dim i As Integer
    IsCreate = False
    
    ligneConfig = 0
    ligneSDV = 0
   v = ThisWorkbook.sheets("POWERTRAIN").UsedRange.Value
    For i = 3 To UBound(v, 1)
        If UCase(v(i, 1)) = "SOMME" And ligneSDV = 0 Then ligneSDV = i
        
        If v(i, 1) = "Titre config" And UCase(CStr(v(i, 2))) = UCase(Me.TextBox22) Then
                ligneConfig = i
                IsCreate = True
                Exit Function
        End If
        
    Next i
    
    Erase v
End Function

Sub createSDV()
     
        Dim i As Integer
      
        Application.EnableEvents = False
        With ThisWorkbook.sheets("POWERTRAIN")
                    i = .Range("A65000").End(xlUp).row
                    ThisWorkbook.sheets("POWERTRAIN").Rows("3:" & ligneSDV).Copy Destination:=.Range("A" & i + 1)
                    .Range("B" & i + 1) = Me.TextBox22
                    .Range("B" & i + 11 & ":E" & (.Range("A65000").End(xlUp).row - 1)) = 0
                    .Range("G" & i + 11 & ":I" & (.Range("A65000").End(xlUp).row - 1)) = 0
                    .Range("B" & .Range("A65000").End(xlUp).row).formula = _
                            "=powerSummCells(" & .Range("A" & .Range("A65000").End(xlUp).row).Address & ", NOW())"
               
                    ligneConfig = i + 1
        End With
        Application.EnableEvents = True

End Sub

Function RemplissageDonnee()
        Dim r As Long
        Dim i As Integer
        Dim j As Integer
        
        With ThisWorkbook.sheets("POWERTRAIN")
'             r = .Range("A65000").End(xlUp).Row + 1
             Application.EnableEvents = False
             r = ligneConfig
            For i = 1 To 7 Step 2
                    For j = 2 To .Cells(r + i, .Columns.Count).End(xlToLeft).Column
                           CompareV (.Cells(r + i, j))
                    Next j
            Next i
            
            Application.EnableEvents = True
            MsgBox "Paramètres Ajoutés", vbInformation, "ODRIV"
            Unload Me
            Exit Function

        End With
End Function

Function CompareV(r As Range)
        Dim i As Integer

        With ThisWorkbook.sheets("POWERTRAIN")
            
              If .Cells(r.row, 1) = "Engine type" Then
                       For i = 0 To Me.ListView15.ListCount - 1
                            If Me.ListView15.Selected(i) = False And UCase(Me.ListView15.list(i)) = UCase(r.Value) Then
                                 r.Offset(1, 0) = ""
                            ElseIf Me.ListView15.Selected(i) = True And UCase(Me.ListView15.list(i)) = UCase(r.Value) Then
                                r.Offset(1, 0) = "X"
                            End If
                       Next i
                       
            ElseIf .Cells(r.row, 1) = "Gearbox type" Then
                       For i = 0 To Me.ListView16.ListCount - 1
                            If Me.ListView16.Selected(i) = False And UCase(Me.ListView16.list(i)) = UCase(r.Value) Then
                                 r.Offset(1, 0) = ""
                            ElseIf Me.ListView16.Selected(i) = True And UCase(Me.ListView16.list(i)) = UCase(r.Value) Then
                                r.Offset(1, 0) = "X"
                            End If
                       Next i
                       
             ElseIf .Cells(r.row, 1) = "Number of gears" Then
                       For i = 0 To Me.ListView17.ListCount - 1
                            If Me.ListView17.Selected(i) = False And UCase(Me.ListView17.list(i)) = UCase(r.Value) Then
                                 r.Offset(1, 0) = ""
                            ElseIf Me.ListView17.Selected(i) = True And UCase(Me.ListView17.list(i)) = UCase(r.Value) Then
                                r.Offset(1, 0) = "X"
                            End If
                       Next i
                       
            ElseIf .Cells(r.row, 1) = "Area" Then
                       For i = 0 To Me.ListView18.ListCount - 1
                            If Me.ListView18.Selected(i) = False And UCase(Me.ListView18.list(i)) = UCase(r.Value) Then
                                 r.Offset(1, 0) = ""
                            ElseIf Me.ListView18.Selected(i) = True And UCase(Me.ListView18.list(i)) = UCase(r.Value) Then
                                r.Offset(1, 0) = "X"
                            End If
                       Next i
                       
            End If
        End With
End Function










