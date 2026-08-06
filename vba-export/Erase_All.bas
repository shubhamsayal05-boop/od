Attribute VB_Name = "Erase_All"
Option Explicit
Sub EraseAll()
    RAZ_SDV_Sheets
    
End Sub
Sub RAZ_onglet(onglet As String)
    Dim NEVENTS As Variant
    Application.DisplayAlerts = False
    NEVENTS = TotEventSheet(onglet)
    ThisWorkbook.sheets("RATING").Rows("23:600").EntireRow.Hidden = False
    With ThisWorkbook.sheets(onglet)
        .Visible = -1
        If NEVENTS > 2 Then
            If .AutoFilterMode Then
                .Cells.AutoFilter
            End If
            .Range("L7:AW" & NEVENTS + 4).Clear
           
            .Range("D4:I5").ClearContents
            
            .Range("J5").ClearContents
            
            .Range("G8:G19").ClearContents
            
            .Range("K8:K19").Value = ""
            
            .Range("I14:I19").ClearContents
             
            .Range("I8:I10").ClearContents
            
            .Range("J14:J19").ClearContents
            
            .Range("J8:J10").ClearContents
            
            .Range("H20:H22").ClearContents
            
            If .Name = "Lever change" Or .Name = "Auto start" Or .Name = "Auto stop" Then
                .Range("G26:I55").ClearContents
            End If

            .ChartObjects("Graphique P1").Chart.ChartTitle.Characters.Font.ColorIndex = 1
            .ChartObjects("Graphique P2").Chart.ChartTitle.Characters.Font.ColorIndex = 1
            .ChartObjects("Graphique P3").Chart.ChartTitle.Characters.Font.ColorIndex = 1
            
            
        End If
    End With
    Call eraseRating(onglet)
    Call eraseGraphStatus
    Application.DisplayAlerts = True

End Sub
Sub RAZ_ongletDyn(onglet As String)
    Dim NEVENTS As Variant
    Application.DisplayAlerts = False
    NEVENTS = TotEventSheet(onglet)
    ThisWorkbook.sheets("RATING").Rows("22:600").EntireRow.Hidden = False
    With ThisWorkbook.sheets(onglet)
        .Visible = -1
        If NEVENTS > 2 Then
            If .AutoFilterMode Then
                .Cells.AutoFilter
            End If
            .Range("BS7:GG" & NEVENTS + 4).Clear
            
            .Range("BK4:BP5").ClearContents
            
            .Range("BQ5").ClearContents
            
            .Range("BN8:BN19").ClearContents
            
            .Range("BR8:BR19").Value = ""
            
            .Range("BP14:BP19").ClearContents
             
            .Range("BP8:BP10").ClearContents
            
            .Range("BQ14:BQ19").ClearContents
            
            .Range("BQ8:BQ10").ClearContents
            
            .Range("BO20:BO22").ClearContents
            
            If .Name = "Lever change" Or .Name = "Auto start" Or .Name = "Auto stop" Then
                .Range("BN26:BP55").ClearContents
            End If
            
            .ChartObjects("Graphique P11").Chart.ChartTitle.Characters.Font.ColorIndex = 1
            .ChartObjects("Graphique P12").Chart.ChartTitle.Characters.Font.ColorIndex = 1
            .ChartObjects("Graphique P13").Chart.ChartTitle.Characters.Font.ColorIndex = 1

            
        End If
    End With
    Call eraseRating(onglet)
    Call eraseGraphStatus
    Application.DisplayAlerts = True

End Sub
Sub Erase_All2(Optional eff As Boolean)                                       '[Effacer toutes les données]
    
    If eff = True Then
        Call eraseRating
        Call eraseGraphStatus
    Else
        Application.DisplayAlerts = False
        If sheetExists("DATA") Then
            ThisWorkbook.sheets("DATA").Delete
            If sheetExists("GRILLE") Then ThisWorkbook.sheets("GRILLE").Delete
        End If
        Application.DisplayAlerts = True
        Call EraseAll
        Call ResetOptions
        ThisWorkbook.sheets("structure").Range("N1") = 2
        Call eraseRating
        Call eraseGraphStatus
        With ThisWorkbook.sheets("HOME")
            .Range("idProjects") = ""
            .Range("Project") = ""
            .Range("Moniteur") = ""
            .Range("Gears") = ""
            .Range("Fuel") = ""
            .Range("Mode") = ""
            .Range("Milestone") = ""
            .Range("Area") = ""
            .Range("Prestation") = ""
            .Range("Software") = ""
            .Range("DriveVersion") = ""
            .Range("C24") = ""
            .Range("C23") = ""
            .Range("H23") = ""
            .Range("Moniteur").Interior.color = RGB(255, 255, 255)
        End With
    End If
    
    

End Sub
Sub RAZ_SDV_Sheets(Optional SDV As String, Optional hide_sdv As Boolean)
    Dim v
    Dim i As Long
    Application.DisplayAlerts = False
    
    If hide_sdv = True Then
         v = ThisWorkbook.sheets("structure").Range("B1").CurrentRegion.Columns(1).Value
        For i = 2 To UBound(v, 1)
            If Len(v(i, 1)) > 0 And sheetExists(v(i, 1)) Then
                ThisWorkbook.sheets(v(i, 1)).Visible = False
            End If
        Next i
        Erase v
        
    ElseIf Len(SDV) > 2 Then
        If sheetExists(SDV) Then
             ThisWorkbook.sheets(SDV).Visible = -1
            ThisWorkbook.sheets(SDV).Delete
        End If
    Else
        v = ThisWorkbook.sheets("structure").Range("B1").CurrentRegion.Columns(1).Value
        For i = 2 To UBound(v, 1)
    
            If Len(v(i, 1)) > 0 And sheetExists(v(i, 1)) Then
                 ThisWorkbook.sheets(v(i, 1)).Visible = -1
                ThisWorkbook.sheets(v(i, 1)).Delete
            End If
        Next i
        Erase v
    End If
    
    
    Application.DisplayAlerts = True
End Sub

Function eraseRating(Optional k As String)
        Dim v As Long
        Dim r As Range
        Dim DernC As Integer
        
        If Len(k) > 0 Then
            Call delByKeyRating(k)
        Else
            With ThisWorkbook.sheets("RATING")
'                    .Range("B12:AA50").Interior.Pattern = xlSolid
                    .Range("E" & 11) = ""
                    .Range("F" & 11) = ""
                    .Range("E" & 11).Interior.color = RGB(242, 242, 242)
                    .Range("F" & 11).Interior.color = RGB(242, 242, 242)
                    
                    .Range("E" & 17) = ""
                    .Range("F" & 17) = ""
                    .Range("E" & 17).Interior.color = RGB(242, 242, 242)
                    .Range("F" & 17).Interior.color = RGB(242, 242, 242)
                    
                    .Range("RESULTATGLOBAL1") = ""
                    .Range("RESULTATGLOBAL2") = ""
                    v = getLastRowRating
                    
                    DernC = .Cells(21, .Columns.Count).End(xlToLeft).Column
                    Set r = .Range("D" & 23)
                     While r.row <= v
                            If Len(r.Value) > 0 Then
                                Call delByKeyRating(r.Value)
                            Else
                                .Range(.Cells(r.row, 7), .Cells(r.row, DernC)).ClearContents
                                .Range(.Cells(r.row, 7), .Cells(r.row, DernC)).Font.Size = 15
                                .Range(.Cells(r.row, 7), .Cells(r.row, DernC)).Font.Bold = True
                                .Range(.Cells(r.row, 7), .Cells(r.row, DernC)).HorizontalAlignment = xlCenter
                                .Range(.Cells(r.row, 7), .Cells(r.row, DernC)).VerticalAlignment = xlBottom
                               
                            End If
                            .Rows(r.row).EntireRow.Hidden = False
                             Set r = r.Offset(1, 0)
                     Wend
                    .Shapes("UpdateTargetButton").Visible = False
                    '.Range("l15:l" & V).Value = ""
             End With
             
        End If
        
End Function

Function eraseGraphStatus()
        Dim v As Long
        Dim r As Range
        Dim DernL As Integer
        Dim getVeh As String
        Dim i As Integer
        
        getVeh = ""
        With ThisWorkbook.Worksheets("CONFIGURATIONS")
             Set r = .Range("VEHICLE")
             Set r = r.Offset(1, 0)
             While r.Value <> ""
                 getVeh = IIf(getVeh = "", r.Value, getVeh & "," & r.Value)
                 Set r = r.Offset(1, 0)
             Wend
        End With
    
        With ThisWorkbook.sheets("Graph_status")
                  DernL = .Cells(.Rows.Count, 1).End(xlUp).row
                  For i = 1 To DernL
                        If InStr(1, "," & getVeh & ",", "," & .Cells(i, 1) & ",") <> 0 Then
                                .Cells(i, 2).ClearContents
                        End If
                  Next i
        End With
        
End Function

Function delByKeyRating(sheets As String)
      Dim i As Integer
      With ThisWorkbook.Worksheets("RATING")
         i = rowStartSDV("RATING", sheets, 4)
         If i <> 0 Then
            .Range(.Cells(i, 13), .Cells(i, .Range("colPD1").Column - 1)) = ""
            .Range(.Cells(i, .Range("colPDD3").Column + 1), .Cells(i, .Cells(21, .Columns.Count).End(xlToLeft).Column)) = ""
            
'            .Range("M" & i & ":O" & i) = ""
'            .Range("V" & i & ":Y" & i) = ""

            .Range("G" & i & ":L" & i).Font.color = RGB(255, 255, 255)
            
            .Range(.Cells(i, .Range("colPD1").Column), .Cells(i, .Range("colPDD3").Column)).Font.color = RGB(255, 255, 255)
            
'            .Range("P" & i & ":U" & i).Font.color = RGB(255, 255, 255)
            .Range("C" & i).Font.Size = 12
            Call hyperlinkRemove(.Range("D" & i))
            Call hideShowTarget(False)
         End If
        
    End With
End Function


Function hyperlinkRemove(RsB As Range)
      Dim tabFormat(6)
      
        
        tabFormat(0) = RsB.Font.Size
        tabFormat(1) = RsB.Interior.color
        tabFormat(6) = RsB.Font.color
        
        RsB.Hyperlinks.Delete
        
        RsB.Font.Size = tabFormat(0)
        RsB.Interior.color = tabFormat(1)
        With ThisWorkbook.Worksheets("RATING").Range("D" & RsB.row & ":F" & RsB.row)
                .Borders(1).LineStyle = xlContinuous
                .Borders(2).LineStyle = xlContinuous
                .Borders(3).LineStyle = xlContinuous
                .Borders(4).LineStyle = xlContinuous
        End With
        RsB.Font.color = tabFormat(6)
        
        RsB.Font.Underline = xlUnderlineStyleNone
        
End Function







