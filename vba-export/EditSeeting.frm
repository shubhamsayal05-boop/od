Attribute VB_Name = "EditSeeting"
Attribute VB_Base = "0{E94D88F3-A251-48E2-972A-B3A6F072DBEB}{FB1BEE73-3835-4D46-BB75-12E9BD4FE313}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub ComboBox2_Change()
        ChargeConfig
End Sub


Private Sub CommandButton2_Click()
        If Len(Me.ComboBox3) > 0 And Len(Me.ComboBox2.Value) > 0 Then
           Call ReplaceValue
           Unload Me
           ConfigSetting.editing = "OK"
           ConfigSetting.Show
           
        Else
            MsgBox "REMPLIR", vbCritical, "ODRIV"
        End If
End Sub

Private Sub UserForm_Initialize()
    initList
End Sub
Sub initList()
        Dim r As Long
        Dim i As Integer
      
        Application.EnableEvents = False
        With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
                r = .Range("A65000").End(xlUp).row
                 For i = 3 To r
                            If Len(.Cells(i, 1)) > 0 Then
                               Me.ComboBox2.AddItem .Cells(i, 1)
                            End If
                 Next i
        End With
        
        
        Application.EnableEvents = True

End Sub


Function ChargeConfig()
    Dim v
    Dim i As Integer
    Dim comparVal As String
    Me.ComboBox3.Clear
   ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Outline.ShowLevels RowLevels:=2
   v = ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").UsedRange.Value
   For i = 2 To UBound(v, 1)
        If Len(v(i, 1)) > 0 And UCase(CStr(v(i, 1))) = UCase(Me.ComboBox2) Then
                If i + 1 <= UBound(v, 1) Then i = i + 1
                If i <= UBound(v, 1) And Len(v(i, 1)) = 0 Then
                    comparVal = v(i, 1)
                    While Len(comparVal) = 0 And i <= UBound(v, 1)
                          If Left(v(i, 2), 9) = "Config n°" Then Me.ComboBox3.AddItem v(i, 2)
                          i = i + 1
                          If i <= UBound(v, 1) Then comparVal = v(i, 1)
                    Wend
               
                End If
        End If
    Next i
    
    ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Outline.ShowLevels RowLevels:=1
    Erase v
End Function
Function ReplaceValue()
        Dim r As Long
        Dim i As Long
       
       
        With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
                 .Outline.ShowLevels RowLevels:=2
                r = DernLigne
                For i = 3 To r
                            If (.Cells(i, 1)) = Me.ComboBox2.Value Then
                                 ConfigSetting.TextBox2 = .Cells(i, 1)
                                 If Len(.Cells(i + 1, 1)) = 0 And i + 1 <= r Then
                                     i = i + 1
                                     While Len(.Cells(i, 1)) = 0 And i <= r
                                             If .Cells(i, 2) = Me.ComboBox3 Then
                                                Call RemplissageDonnee(i)
                                                Exit Function
                                             End If
                                             i = i + 1
                                    Wend
                                End If
                            End If
                 Next i
                 .Outline.ShowLevels RowLevels:=1
        End With
End Function
Function RemplissageDonnee(FTNum As Long)
        Dim r As Long
        Dim i As Integer
        Dim j As Integer
        
        With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
             .Outline.ShowLevels RowLevels:=2
'             r = .Range("A65000").End(xlUp).Row + 1
             Application.EnableEvents = False
             r = FTNum
             
             
             ConfigSetting.TextBox22 = (Right(.Cells(r, 2), ((Len(.Cells(r, 2)) - 1) - InStr(1, .Cells(r, 2), ": "))))
             ConfigSetting.TextBox2.Locked = True
             ConfigSetting.TextBox22.Locked = True
             
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
'            For i = 1 To 7 Step 2
'                       For j = 3 To .Cells(r + i, .Columns.Count).End(xlToLeft).Column
'                           CompareV (.Cells(r + i, j))
'                     Next j
'            Next i
            
'            ConfigSetting.TextBox23 = .Cells(r + 12, 2) & "-" & .Cells(r + 10, 2)

            
            ConfigSetting.TextBox25 = IIf(Len(.Cells(r + 27, 3)) > 1, .Cells(r + 27, 3), "")
            ConfigSetting.TextBox26 = IIf(Len(.Cells(r + 27, 4)) > 1, .Cells(r + 27, 4), "")
            ConfigSetting.TextBox24 = IIf(Len(.Cells(r + 27, 2)) > 1, .Cells(r + 27, 2), "")
            ConfigSetting.ComboBox20 = IIf(Len(.Cells(r + 29, 3)) > 1, .Cells(r + 29, 3), "")
            ConfigSetting.ComboBox19 = IIf(Len(.Cells(r + 29, 2)) > 1, .Cells(r + 29, 2), "")
            ConfigSetting.ComboBox21 = IIf(Len(.Cells(r + 29, 4)) > 1, .Cells(r + 29, 4), "")
            
            .Outline.ShowLevels RowLevels:=1
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
                       For i = 0 To ConfigSetting.ListView15.ListCount - 1
                            If r.Offset(0, 1) = "X" And UCase(ConfigSetting.ListView15.list(i)) = UCase(r.Value) Then
                                 ConfigSetting.ListView15.Selected(i) = True
                            ElseIf r.Offset(0, 1) = "" And UCase(ConfigSetting.ListView15.list(i)) = UCase(r.Value) Then
                                ConfigSetting.ListView15.Selected(i) = False
                            End If
                       Next i
                       
            ElseIf id = 2 Then
                       For i = 0 To ConfigSetting.ListView16.ListCount - 1
                            If r.Offset(0, 1) = "X" And UCase(ConfigSetting.ListView16.list(i)) = UCase(r.Value) Then
                                 ConfigSetting.ListView16.Selected(i) = True
                            ElseIf r.Offset(0, 1) = "" And UCase(ConfigSetting.ListView16.list(i)) = UCase(r.Value) Then
                                ConfigSetting.ListView16.Selected(i) = False
                            End If
                       Next i
                       
                       
             ElseIf id = 3 Then
                       For i = 0 To ConfigSetting.ListView17.ListCount - 1
                            If r.Offset(0, 1) = "X" And UCase(ConfigSetting.ListView17.list(i)) = UCase(r.Value) Then
                                 ConfigSetting.ListView17.Selected(i) = True
                            ElseIf r.Offset(0, 1) = "" And UCase(ConfigSetting.ListView17.list(i)) = UCase(r.Value) Then
                                ConfigSetting.ListView17.Selected(i) = False
                            End If
                       Next i
                       
                       
            ElseIf id = 4 Then
                       For i = 0 To ConfigSetting.ListView18.ListCount - 1
                            If r.Offset(0, 1) = "X" And UCase(ConfigSetting.ListView18.list(i)) = UCase(r.Value) Then
                                 ConfigSetting.ListView18.Selected(i) = True
                            ElseIf r.Offset(0, 1) = "" And UCase(ConfigSetting.ListView18.list(i)) = UCase(r.Value) Then
                                ConfigSetting.ListView18.Selected(i) = False
                            End If
                       Next i
                       
                       
            End If
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









