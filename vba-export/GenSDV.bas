Attribute VB_Name = "GenSDV"
Option Explicit
Function GetNoteGlobalTarget(part As String, vehCible As String) As Double
    Dim i As Integer
    Dim nom As String
    Dim poids As Double
    Dim Target As Double
    Dim NB As Double
    Dim n As Integer
    Dim LR As Integer
    Dim colDR As Integer, colDY As Integer
    NB = 0
    Target = 0
    poids = 0
    i = 23
    LR = getLastRowRating
    
    
    With ThisWorkbook.sheets("RATING")
            While i <= LR
                    nom = ThisWorkbook.sheets("RATING").Cells(i, 4).Value
                    If part = "driv" Then
                          colDR = getTargetColonnePage(vehCible, "DRIV")
                           If checkEmptyRating(i, colDR, part) = True Then
                                    GetNoteGlobalTarget = -555
                                    Exit Function
                            End If
                    Else
                          colDY = getTargetColonnePage(vehCible, "DYN")
                          If checkEmptyRating(i, colDY, part) = True Then
                                    GetNoteGlobalTarget = -555
                                    Exit Function
                            End If
                    End If
                    
                    poids = checkPoidsName(nom)
                    Dim r As Range
                    
                    With ThisWorkbook.sheets("RATING")
                            If .Range("C" & i).Font.Size = 2 Then
                               If part = "driv" Then
                                    If Len(.Cells(i, colDR).Value) > 0 Then
                                        If ThisWorkbook.sheets(.Cells(i, 4).Value).Range("G8") >= OvMinPts(.Cells(i, 4).Value) Then
                                           NB = NB + poids
                                           Target = Target + (.Cells(i, colDR).Value * poids)
                                        End If
                                   End If
                                Else
                                   If Len(.Cells(i, colDY).Value) > 0 Then
                                        If ThisWorkbook.sheets(.Cells(i, 4).Value).Range("G8") >= OvMinPts(.Cells(i, 4).Value) Then
                                           NB = NB + poids
                                           Target = Target + (.Cells(i, colDY).Value * poids)
                                         End If
                                   End If
                                  
                                End If
                            End If
                    End With
                     i = i + 1
                     n = n + 1
           Wend
      End With
    
    If NB > 0 Then GetNoteGlobalTarget = Target / NB
End Function
Function checkEmptyRating(i As Integer, col As Integer, part As String) As Boolean
    checkEmptyRating = False
    With ThisWorkbook.sheets("RATING")
        If part = "driv" Then
            If .Range("M" & i) <> .Cells(i, col) Then
                 If Len(.Range("M" & i)) = 0 Or Len(.Cells(i, col)) = 0 Then
                     checkEmptyRating = True
                 End If
            End If
        Else
             If .Cells(i, .Range("DynIndex").Column) <> .Cells(i, col) Then
                  If Len(.Cells(i, .Range("DynIndex").Column)) = 0 Or Len(.Cells(i, col)) = 0 Then
                       checkEmptyRating = True
                 End If
            End If
       End If
    End With
End Function
Function checkPoidsName(nom As String)
        Dim r As Range
        Dim cel As Range
        Dim col As Integer
        
        checkPoidsName = 0
        Set r = ThisWorkbook.sheets("SETTINGS").UsedRange.Find(What:=nom, LookIn:=xlValues, LookAt:=xlWhole)
        If Not r Is Nothing Then
            With ThisWorkbook.sheets("SETTINGS")
'                    Set cel = .Rows("14").Find(What:="WEIGHT", LookIn:= _
'                                  xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:= _
'                                  xlNext, MatchCase:=False, SearchFormat:=False)
'                                  If Not cel Is Nothing Then Col = cel.Column + 2 Else Exit Function
                      checkPoidsName = .Cells(r.row + 11, 3).Value
             End With
         End If
End Function
Function sheetExists(ByVal sh As String) As Boolean
    On Error Resume Next
    sheetExists = CBool(Not ThisWorkbook.sheets(sh) Is Nothing)
    If ERR.Number = 9 Then ERR.Clear
End Function


Function GenSdVSheets(sdv1 As String)
    Dim r As Range
    Dim col As Integer
    Dim newCritere As Boolean
    Dim o As Integer
    Dim n As Integer
    Dim p As Integer
    Dim q As Integer
    Dim Indice As Boolean
    Dim st As String
  
    Indice = False
    newCritere = False
    If Len(sdv1) > 0 Then
       
        Set r = ThisWorkbook.sheets("structure").Range("B2")
        st = getNumberRow(sdv1)
        If st = "" Then Exit Function
        Call createSheetSDv(sdv1)
        p = val(Split(st, ";")(1))
        q = val(Split(st, ";")(0)) + 1
        o = p
        n = q
        Call GenSdVDrivability(sdv1, o, n)
        o = p
        n = q
        Call GenSdVDinamyc(sdv1, o, n)
    End If
   
End Function

Function GenSdVDrivability(sdv1 As String, o As Integer, n As Integer)
    Dim r As Range
    Dim col As Integer
    Dim newCritere As Boolean
    Dim Indice As Boolean
    Dim st As String
  
    Indice = False
    newCritere = False
   
        col = 13

        Do While o >= n
            With ThisWorkbook.Worksheets("structure")
                ThisWorkbook.Worksheets(sdv1).Cells(6, col).Value = .Range("D" & n)
                Call colorCell(ThisWorkbook.Worksheets(sdv1).Cells(6, col))
                If .Range("C" & n) = "criteria" Then
                    If newCritere = False Then
                        Call newCreteria(col - 1, sdv1)
                        newCritere = True
                    End If
                     Call insertCreteria(col, sdv1)
                End If
            End With
            n = n + 1
            col = col + 1
        Loop
        If col > 13 Then
                ThisWorkbook.Worksheets(sdv1).Cells(6, col).Value = "Indice occurrencé"
                colorCell ThisWorkbook.Worksheets(sdv1).Cells(6, col)
                ThisWorkbook.Worksheets(sdv1).Cells(6, col + 1).Value = "Id BdD"
                colorCell ThisWorkbook.Worksheets(sdv1).Cells(6, col + 1)
        End If
   
    
   
End Function

Function GenSdVDinamyc(sdv1 As String, o As Integer, n As Integer)
     Dim r As Range
    Dim col As Integer
    Dim newCritere As Boolean
    Dim Indice As Boolean
    Dim st As String
  
    Indice = False
    newCritere = False
   
        col = 72

        Do While o >= n
            With ThisWorkbook.Worksheets("structure")
                ThisWorkbook.Worksheets(sdv1).Cells(6, col).Value = .Range("D" & n)
                Call colorCell(ThisWorkbook.Worksheets(sdv1).Cells(6, col))
                If .Range("C" & n) = "criteria" Then
                    If newCritere = False Then
                        Call newCreteria(col - 1, sdv1)
                        newCritere = True
                    End If
                     Call insertCreteria(col, sdv1)
                End If
            End With
            n = n + 1
            col = col + 1
        Loop
        If col > 72 Then
                ThisWorkbook.Worksheets(sdv1).Cells(6, col).Value = "Indice occurrencé"
                colorCell ThisWorkbook.Worksheets(sdv1).Cells(6, col)
                ThisWorkbook.Worksheets(sdv1).Cells(6, col + 1).Value = "Id BdD"
                colorCell ThisWorkbook.Worksheets(sdv1).Cells(6, col + 1)
        End If
   
   
End Function

Sub colorCell(ByVal c As Range)
    With c.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With c.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With c.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With c.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    c.Borders(xlInsideVertical).LineStyle = xlNone
    c.Borders(xlInsideHorizontal).LineStyle = xlNone

    With c
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    c.Font.Bold = True

    With c.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
End Sub

Function RemplirCoverage(onglet As String)
            Dim x As Long, D As Long
            
            With ThisWorkbook.Worksheets("POWERTRAIN")
               x = Formules.StartConFig
               If x = 0 Then Exit Function
               D = x
               While .Cells(D, 1) <> "SOMME"
                    If UCase(.Cells(D, 1)) = UCase(onglet) Then
                        ThisWorkbook.Worksheets(onglet).Range("G20").formula = "=IFERROR(I8 / 'POWERTRAIN'!" & .Cells(D, 7).Address & ", 0)"
                        ThisWorkbook.Worksheets(onglet).Range("G21").formula = "=IFERROR(I9 / 'POWERTRAIN'!" & .Cells(D, 8).Address & ", 0)"
                        ThisWorkbook.Worksheets(onglet).Range("G22").formula = "=IFERROR(I10 / 'POWERTRAIN'!" & .Cells(D, 9).Address & ", 0)"
                        Exit Function
                    End If
                    D = D + 1
              Wend
               
        End With
        
End Function

Function getNumberRow(SDV As String) As String
    Dim v
    Dim i As Integer
    Dim p As Integer
    getNumberRow = ""
    v = ThisWorkbook.sheets("structure").UsedRange.Value
    For i = 2 To UBound(v, 1)
        If UCase(v(i, 2)) = UCase(SDV) Then
                getNumberRow = i
                i = i + 1
                p = i
                While i <= UBound(v, 1) And Len(v(p, 2)) = 0
                    If i <= UBound(v, 1) Then p = i
                    i = i + 1
                Wend
                If p < UBound(v, 1) Then p = p - 1
                getNumberRow = getNumberRow & ";" & p
                Erase v
                Exit Function
       End If
    Next i
    
    Erase v
    
End Function


Function createSheetSDv(SDV As String)
            If Not sheetExists(SDV) Then
                Application.DisplayAlerts = False
                ThisWorkbook.sheets("VIERGE").Copy After:=ThisWorkbook.sheets("VIERGE")
                Application.DisplayAlerts = True
                With ActiveSheet
                    .Name = SDV
                    .Tab.color = vbGreen
                    .Range("B1") = ActiveSheet.Name
                    .Range("BI1") = ActiveSheet.Name
                    
                End With
                
            End If
                       
End Function


Function newCreteria(col As Integer, sht As String)
        With ThisWorkbook.Worksheets(sht)
                .Cells(3, col).Value = "Waterline"
                .Cells(4, col).Value = "Target"
                .Cells(5, col).Value = "Criticity"
                
                .Cells(3, col).Font.color = 255
                .Cells(4, col).Font.color = 5287936
                .Cells(5, col).Font.color = 0
                   
        End With
        
        With ThisWorkbook.Worksheets(sht).Range(ThisWorkbook.Worksheets(sht).Cells(3, col), ThisWorkbook.Worksheets(sht).Cells(5, col))
                .Borders(1).LineStyle = xlContinuous
                .Borders(2).LineStyle = xlContinuous
                .Borders(3).LineStyle = xlContinuous
                .Borders(4).LineStyle = xlContinuous
                .Interior.color = 14277081
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
                .Font.Size = 12
        End With
End Function

Function insertCreteria(col As Integer, sht As String)
        With ThisWorkbook.Worksheets(sht).Range(ThisWorkbook.Worksheets(sht).Cells(3, col), ThisWorkbook.Worksheets(sht).Cells(4, col))
                .Borders(4).LineStyle = xlContinuous
                .HorizontalAlignment = xlCenter
                .Font.Size = 10
                .NumberFormat = "0.0"
                .Font.Bold = True
        End With
        ThisWorkbook.Worksheets(sht).Cells(3, col).Font.color = 255
        ThisWorkbook.Worksheets(sht).Cells(4, col).Font.color = 5287936
        With ThisWorkbook.Worksheets(sht).Cells(5, col)
                .Borders(1).LineStyle = xlContinuous
                .Borders(2).LineStyle = xlContinuous
                .Borders(3).LineStyle = xlContinuous
                .Borders(4).LineStyle = xlContinuous
                .Value = 3
                .Font.Bold = True
                .Font.Size = 12
                .HorizontalAlignment = xlCenter
        End With
        
         Call formatConditionColor(ThisWorkbook.Worksheets(sht).Cells(5, col))
End Function

Function formatConditionColor(r As Range)
    With r
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=1"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).Font.color = -16383844
        .FormatConditions(1).Interior.color = 13551615
        .FormatConditions(1).StopIfTrue = False
        
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=2"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).Font.color = -16751204
        .FormatConditions(1).Interior.color = 10284031
        .FormatConditions(1).StopIfTrue = False
        
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=3"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).Font.color = -16752384
        .FormatConditions(1).Interior.color = 13561798
        .FormatConditions(1).StopIfTrue = False
        
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).Font.color = -16752384
        .FormatConditions(1).Interior.color = 13561798
        .FormatConditions(1).StopIfTrue = False
    End With
End Function

Function checkCriteria(ByVal onglet As String)
   Dim i As Long
   Dim j As Long
   
   checkCriteria = False
   With ThisWorkbook.sheets(onglet)
        For i = 13 To .Cells(6, .Columns.Count).End(xlToLeft).Column
             If Len(.Cells(5, i).Value) > 0 Then Exit For
        Next i
        i = i + 1
        j = i
        While Len(.Cells(5, j)) > 0
                If CStr(.Cells(5, j).Value) <> "3" Then
                    checkCriteria = True
                    Exit Function
                End If
                j = j + 1
        Wend
'        For j = i To .Cells(5, .Columns.Count).End(xlToLeft).Column
'            If CStr(.Cells(5, j).Value) <> "3" And Len(.Cells(5, j).Value) > 0 Then
'                checkCriteria = True
'                Exit Function
'            End If
'        Next j
       
   End With
End Function

Function checkCriteriaDyn(ByVal onglet As String)
   Dim i As Long
   Dim j As Long
   
   checkCriteriaDyn = False
   With ThisWorkbook.sheets(onglet)
        i = 72
        While Len(.Cells(6, i)) <> 0 And Len(.Cells(5, i).Value) = 0
            i = i + 1
        Wend
        i = i + 1
        
    
         j = i
        While Len(.Cells(5, j)) <> 0
            If CStr(.Cells(5, j).Value) <> "3" And Len(.Cells(5, j).Value) > 0 Then
                checkCriteriaDyn = True
                Exit Function
            End If
            j = j + 1
        Wend
       
   End With
End Function

Function checkCorrespondancePriority(ByVal onglet As String)
        Dim LigneDebut As Long
        Dim Types As Long
        checkCorrespondancePriority = False
        LigneDebut = Rating_Priorisation.LigneSeeting(onglet)
        
        If LigneDebut = 0 Then Exit Function
        Types = Rating_Priorisation.StartConFig(LigneDebut)
        If Types = 0 Then Exit Function
        checkCorrespondancePriority = True
End Function

Function checkCorrespondancePriorityDyn(ByVal onglet As String)
        Dim LigneDebut As Long
        Dim Types As Long
        checkCorrespondancePriorityDyn = False
        LigneDebut = Rating_PriorisationDyn.LigneSeeting(onglet)
        
        If LigneDebut = 0 Then Exit Function
        Types = Rating_PriorisationDyn.StartConFig(LigneDebut)
        If Types = 0 Then Exit Function
        checkCorrespondancePriorityDyn = True
End Function























