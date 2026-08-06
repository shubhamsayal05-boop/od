Attribute VB_Name = "Rating_NoteGlobale"
Option Explicit

Sub NoteGlobale()

    Dim R_Note As Boolean
    Dim O_Note As Boolean
    Dim G_Note As Boolean
    Dim R_Pred As Boolean
    Dim O_Pred As Boolean
    Dim G_Pred As Boolean
    Dim i As Integer
    Dim SommeIndices As Double
    Dim SommePonderee As Double
    Dim lastRD As Long
    Dim v
    Dim lt As Long
    Dim r As Range
    
    R_Note = False
    O_Note = False
    G_Note = False
    R_Pred = False
    O_Pred = False
    G_Pred = False
    lastRD = getLastRowRating - 22
   If colorGlobalDriv.Count = 0 Then Exit Sub
    With ThisWorkbook.sheets("RATING")
         lt = getLastRowRating
        Set r = .Range("D" & 23)
        While r.row <= lt
                If colorGlobalDriv(UCase(r)) = "High Risk" Then
                    R_Note = True
                ElseIf colorGlobalDriv(UCase(r)) = "Medium Risk" Then
                    O_Note = True
                ElseIf colorGlobalDriv(UCase(r)) = "Low Risk" Then
                    G_Note = True
                End If

                If colorGlobalDrivPred(UCase(r)) = "High Risk" Then
                    R_Pred = True
                ElseIf colorGlobalDrivPred(UCase(r)) = "Medium Risk" Then
                    O_Pred = True
                ElseIf colorGlobalDrivPred(UCase(r)) = "Low Risk" Then
                    G_Pred = True
                End If


               Set r = r.Offset(1, 0)
        Wend
         

        'NOTE

'        If R_Note = True Then
''            .Range("E11") = "RED"
''            .Range("K3") = "At least one case is red."
'        ElseIf O_Note = True Then
''            .Range("E11") = "YELLOW"
''            .Range("K3") = "At least one case is YELLOW."
'        ElseIf G_Note = True Then
'            .Range("E11") = "GREEN"
''            .Range("K3") = "All cases are green."
'        End If
'
'       ' prediction
'       If sheets("HOME").Range("Milestone") <> 4 Then
'            If R_Pred = True Then
'                .Range("F11") = "RED"
'            ElseIf O_Pred = True Then
'                .Range("F11") = "YELLOW"
'            ElseIf G_Pred = True Then
'                .Range("F11") = "GREEN"
'            End If
'       End If
'
        
       
        'Indice Qualité
        SommeIndices = 0
        SommePonderee = 0
        
        v = ThisWorkbook.sheets("structure").Range("B1").CurrentRegion.Columns(1).Value
        For i = 2 To UBound(v, 1)
           
            If Len(v(i, 1)) > 0 And sheetExists(v(i, 1)) Then
                If ThisWorkbook.sheets(v(i, 1)).Range("J5") > 0 Then
                        ThisWorkbook.sheets("RATING").Range("M" & ThisWorkbook.sheets("RATING").Columns("B:F").Find(What:=v(i, 1), LookIn:=xlValues, LookAt:=xlWhole).row) = ThisWorkbook.sheets(v(i, 1)).Range("J5")
                      If ThisWorkbook.sheets(v(i, 1)).Range("G8") >= OvMinPts(v(i, 1)) Then
                        SommePonderee = SommePonderee + Weight(v(i, 1)) * ThisWorkbook.sheets(v(i, 1)).Range("J5")
                        SommeIndices = SommeIndices + Weight(v(i, 1))
                      End If
                End If
            End If
            
        Next i
        Erase v
   

        If SommeIndices <> 0 Then
            ThisWorkbook.sheets("RATING").Range("RESULTATGLOBAL1") = Round(SommePonderee / SommeIndices, 1)
        Else
            ThisWorkbook.sheets("RATING").Range("RESULTATGLOBAL1") = ""
        End If

    End With
End Sub

Sub hyperlinkAdd()
        Dim v
        Dim tabFormat(6)
        
        Dim i As Integer
        v = ThisWorkbook.sheets("structure").Range("B1").CurrentRegion.Columns(1).Value
        With ThisWorkbook.sheets("RATING")
        For i = 2 To UBound(v, 1)
                If Len(v(i, 1)) > 0 And sheetExists(v(i, 1)) Then
                    If ThisWorkbook.sheets(v(i, 1)).Visible = xlSheetVisible Then
                        If .Cells(.Range(order(v(i, 1), 0)).row, 3).Font.Size = 2 Then
                              tabFormat(0) = .Cells(.Range(order(v(i, 1), 0)).row, 4).Font.Size
                              tabFormat(1) = .Cells(.Range(order(v(i, 1), 0)).row, 4).Interior.color
                              tabFormat(2) = .Cells(.Range(order(v(i, 1), 0)).row, 4).Borders(1).LineStyle
                              tabFormat(3) = .Cells(.Range(order(v(i, 1), 0)).row, 4).Borders(2).LineStyle
                              tabFormat(4) = .Cells(.Range(order(v(i, 1), 0)).row, 4).Borders(3).LineStyle
                              tabFormat(5) = .Cells(.Range(order(v(i, 1), 0)).row, 4).Borders(4).LineStyle
                              tabFormat(6) = .Cells(.Range(order(v(i, 1), 0)).row, 4).Font.color
                          
                            ThisWorkbook.sheets("RATING").Hyperlinks.Add anchor:=.Cells(.Range(order(v(i, 1), 0)).row, 4) _
                            , Address:="", SubAddress:="'" & v(i, 1) & "'" & "!A1", TextToDisplay:=v(i, 1)
                            
                            .Cells(.Range(order(v(i, 1), 0)).row, 4).Font.Size = tabFormat(0)
                            .Cells(.Range(order(v(i, 1), 0)).row, 4).Interior.color = tabFormat(1)
                            .Cells(.Range(order(v(i, 1), 0)).row, 4).Borders(1).LineStyle = tabFormat(2)
                            .Cells(.Range(order(v(i, 1), 0)).row, 4).Borders(2).LineStyle = tabFormat(3)
                            .Cells(.Range(order(v(i, 1), 0)).row, 4).Borders(3).LineStyle = tabFormat(4)
                            .Cells(.Range(order(v(i, 1), 0)).row, 4).Borders(4).LineStyle = tabFormat(5)
                            .Cells(.Range(order(v(i, 1), 0)).row, 4).Font.color = tabFormat(6)
                        End If
                    End If
                End If
        Next i
        End With
        Erase v
End Sub





Sub NoteGlobaleDyn()

    Dim R_Note As Boolean
    Dim O_Note As Boolean
    Dim G_Note As Boolean
    Dim R_Pred As Boolean
    Dim O_Pred As Boolean
    Dim G_Pred As Boolean
    Dim i As Integer
    Dim SommeIndices As Double
    Dim SommePonderee As Double
    Dim lastRD As Long
    Dim v
    Dim lt As Long
    Dim r As Range
    
    R_Note = False
    O_Note = False
    G_Note = False
    R_Pred = False
    O_Pred = False
    G_Pred = False
    lastRD = getLastRowRating - 22
    If colorGlobalDyn.Count = 0 Then Exit Sub
    With ThisWorkbook.sheets("RATING")
         lt = .Range("D65000").End(xlUp).row
        Set r = .Range("D" & 23)
        While r.row <= lt
                If colorGlobalDyn(UCase(r)) = "High Risk" Then
                    R_Note = True
                ElseIf colorGlobalDyn(UCase(r)) = "Medium Risk" Then
                    O_Note = True
                ElseIf colorGlobalDyn(UCase(r)) = "Low Risk" Then
                    G_Note = True
                End If
                
                If colorGlobalDynPred(UCase(r)) = "High Risk" Then
                    R_Pred = True
                ElseIf colorGlobalDynPred(UCase(r)) = "Medium Risk" Then
                    O_Pred = True
                ElseIf colorGlobalDynPred(UCase(r)) = "Low Risk" Then
                    G_Pred = True
                End If


               Set r = r.Offset(1, 0)
        Wend
         

        'NOTE
'        If R_Note = True Then
'            .Range("E17") = "RED"
'        ElseIf O_Note = True Then
'            .Range("E17") = "YELLOW"
'        ElseIf G_Note = True Then
'            .Range("E17") = "GREEN"
'        End If
'
'       ' prediction
'        If sheets("HOME").Range("Milestone") <> 4 Then
'            If R_Pred = True Then
'                .Range("F17") = "RED"
'            ElseIf O_Pred = True Then
'                .Range("F17") = "YELLOW"
'            ElseIf G_Pred = True Then
'                .Range("F17") = "GREEN"
'            End If
'        End If
        
        
       
        'Indice Qualité
        SommeIndices = 0
        SommePonderee = 0
        
        v = ThisWorkbook.sheets("structure").Range("B1").CurrentRegion.Columns(1).Value
        For i = 2 To UBound(v, 1)
           
            If Len(v(i, 1)) > 0 And sheetExists(v(i, 1)) Then
                If ThisWorkbook.sheets(v(i, 1)).Range("BQ5") > 0 Then
                  
                        ThisWorkbook.sheets("RATING").Cells(ThisWorkbook.sheets("RATING").Columns("B:F").Find(What:=v(i, 1), LookIn:=xlValues, LookAt:=xlWhole).row, ThisWorkbook.sheets("RATING").Range("DynIndex").Column) = ThisWorkbook.sheets(v(i, 1)).Range("BQ5")
                   
                      If ThisWorkbook.sheets(v(i, 1)).Range("BN8") >= OvMinPts(v(i, 1)) Then
                        SommePonderee = SommePonderee + Weight(v(i, 1)) * ThisWorkbook.sheets(v(i, 1)).Range("BQ5")
                        SommeIndices = SommeIndices + Weight(v(i, 1))
                      End If
                End If
            End If
            
        Next i
        Erase v
   

        If SommeIndices <> 0 Then
            ThisWorkbook.sheets("RATING").Range("RESULTATGLOBAL2") = Round(SommePonderee / SommeIndices, 1)
        Else
            ThisWorkbook.sheets("RATING").Range("RESULTATGLOBAL2") = ""
        End If

    End With
End Sub









