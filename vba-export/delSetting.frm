Attribute VB_Name = "delSetting"
Attribute VB_Base = "0{B11A382C-896A-44C8-AEE8-030CBD0952AD}{D7E936FC-8807-4561-BDCC-A405E4B96140}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private RSdv As Long

Private Sub ComboBox4_Change()
 ChargeConfig
End Sub

Private Sub CommandButton1_Click()
        If Len(Me.ComboBox2.Value) > 0 Then
            Call Dels
            MsgBox "Suppression Réussie", vbInformation, "ODRIV"
            Unload Me
        Else
            MsgBox "REMPLIR ", vbCritical, "ODRIV"
        End If
End Sub

Private Sub CommandButton3_Click()
        Dim r As Long, n As Long
        Dim i As Integer
        Dim lasM As Long
        If Len(Me.ComboBox4.Value) = 0 Or Len(Me.ComboBox3.Value) = 0 Then
                MsgBox "Choisir", vbCritical, "ODRIV"
                Exit Sub
        End If
        Application.EnableEvents = False
        With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
                r = DernLigne
                i = FTNum
                n = i + 2
'                While Len(.Cells(n, 1)) = 0 And n <= .Range("B65000").End(xlUp).Row
                While Len(.Cells(n, 1)) = 0 And n <= r
                         If UCase(Left(.Cells(n, 2), 6)) = "CONFIG" Then
                             If i - 1 = RSdv Then lasM = n - 1 Else lasM = n - 2
                            n = r + 1
                         End If
                         n = n + 1
                Wend
                If lasM = 0 Then
                    If n > r Then
                        If i - 1 = RSdv Then lasM = n Else lasM = n - 1
                     Else
                         If i - 1 = RSdv Then
                           lasM = n - 1
                         Else
                            lasM = n - 2
                         End If
                    
                    End If
              End If
                If lasM > 0 And lasM > i Then
                      If i - 1 = RSdv Then .Rows(i & ":" & lasM).Delete Shift:=xlUp Else .Rows(i - 1 & ":" & lasM).Delete Shift:=xlUp
                    
                     MsgBox "Suppression Réussie", vbInformation, "ODRIV"
                      
'                      If Application.CountA(.Rows(RSdv + 1).EntireRow) = 0 And Application.CountA(.Rows(RSdv + 2).EntireRow) > 0 Then
'                          .Rows(RSdv + 1).Delete Shift:=xlUp
'                      End If
                      Unload Me
                Else
                    MsgBox "N"
                End If
         End With
        Application.EnableEvents = True
        ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Outline.ShowLevels RowLevels:=1
End Sub


Private Sub UserForm_Initialize()
        Dim r As Long
        Dim i As Integer
      
        Application.EnableEvents = False
        With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
                r = .Range("A65000").End(xlUp).row
                 For i = 3 To r
                            If Len(.Cells(i, 1)) > 0 Then
                               Me.ComboBox2.AddItem .Cells(i, 1)
                               Me.ComboBox4.AddItem .Cells(i, 1)
                            End If
                 Next i
                 Call ChargeConfig
                 Me.MultiPage1.Value = 0
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
    For i = 1 To UBound(v, 1)
        If Len(v(i, 1)) > 0 And UCase(CStr(v(i, 1))) = UCase(Me.ComboBox4) Then
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
Function Dels() As Long
        Dim r As Long
        Dim i As Integer
        Dim n As Integer
        Dim IsConfig As Boolean
        Dim lastr As Long
        IsConfig = False
        
        Application.EnableEvents = False
         ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Outline.ShowLevels RowLevels:=2
        With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
                lastr = DernLigne
                r = .Range("A65000").End(xlUp).row
                For i = 3 To r
                            If (.Cells(i, 1)) = Me.ComboBox2.Value Then
                                    If Len(.Cells(i + 1, 1)) = 0 And i + 1 <= lastr Then n = i + 1 Else n = i
                                    While Len(.Cells(n, 1)) = 0 And n <= lastr
                                             n = n + 1
                                             IsConfig = True
                                    Wend
                                    
                                    If IsConfig = True Then .Rows(i & ":" & n - 1).Delete Shift:=xlUp Else .Rows(i & ":" & n).Delete Shift:=xlUp
                            End If
               Next i
           ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Outline.ShowLevels RowLevels:=1
           Application.EnableEvents = True
     End With
End Function

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
        If Len(v(i, 1)) > 0 And UCase(CStr(v(i, 1))) = UCase(Me.ComboBox4) Then
                RSdv = i + 1
                If i + 1 <= UBound(v, 1) Then i = i + 1
                
                
                If i <= UBound(v, 1) And Len(v(i, 1)) = 0 Then
                    comparVal = v(i, 1)
                    While Len(comparVal) = 0 And i <= UBound(v, 1)
                          If Len(v(i, 2)) > 0 Then
                             If UCase(v(i, 2)) = UCase(Me.ComboBox3.Value) Then
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







