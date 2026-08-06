Attribute VB_Name = "convertGraph"
Option Explicit

Function poliGraphDyn(Feuil As String)
     Dim p As shape
     Dim i As Long
     Dim j As Long
     Dim Term As String
     
     For Each p In ThisWorkbook.sheets(Feuil).Shapes
        If p.Name = "Graphique_00" Or p.Name = "Graphique_11" Then
            If InStr(1, ThisWorkbook.sheets(Feuil).ChartObjects(p.Name).Chart.ChartTitle.text, "New Selector_Lever_Position") <> 0 _
            And InStr(1, ThisWorkbook.sheets(Feuil).ChartObjects(p.Name).Chart.ChartTitle.text, "Old Selector_Lever_Position") <> 0 Then
                If ThisWorkbook.sheets(Feuil).ChartObjects(p.Name).Chart.Axes(xlCategory).AxisTitle.text = "Old Selector_Lever_Position" Then
                      i = ThisWorkbook.sheets(Feuil).Range("BH6:GG6").Find("New Selector_Lever_Position", , , xlPart).Column
                      j = ThisWorkbook.sheets(Feuil).Range("BH6:GG6").Find("Old Selector_Lever_Position", , , xlPart).Column
                Else
                      i = ThisWorkbook.sheets(Feuil).Range("BH6:GG6").Find("Old Selector_Lever_Position", , , xlPart).Column
                      j = ThisWorkbook.sheets(Feuil).Range("BH6:GG6").Find("New Selector_Lever_Position", , , xlPart).Column
                 End If
                    ThisWorkbook.sheets(Feuil).ChartObjects(p.Name).Chart.FullSeriesCollection(1).XValues = "={" & Split(allSelector_Lever_PositionConfig, "#")(0) & "}"
                    ThisWorkbook.sheets(Feuil).ChartObjects(p.Name).Chart.FullSeriesCollection(1).Values = "={" & Split(allSelector_Lever_PositionConfig, "#")(1) & "}"
                    ThisWorkbook.sheets(Feuil).ChartObjects(p.Name).Chart.Axes(xlValue).MaximumScaleIsAuto = True
          Else
                Term = ""
                If InStr(1, ThisWorkbook.sheets(Feuil).ChartObjects(p.Name).Chart.ChartTitle.text, "New Selector_Lever_Position") <> 0 Then
                    Term = "New Selector_Lever_Position"
                    i = ThisWorkbook.sheets(Feuil).Range("BH6:GG6").Find("New Selector_Lever_Position", , , xlPart).Column
                ElseIf InStr(1, ThisWorkbook.sheets(Feuil).ChartObjects(p.Name).Chart.ChartTitle.text, "Old Selector_Lever_Position") <> 0 Then
                   Term = "Old Selector_Lever_Position"
                    i = ThisWorkbook.sheets(Feuil).Range("BH6:GG6").Find("Old Selector_Lever_Position", , , xlPart).Column
                 ElseIf InStr(1, ThisWorkbook.sheets(Feuil).ChartObjects(p.Name).Chart.ChartTitle.text, "Selector_Lever_Position") <> 0 Then
                   Term = "Selector_Lever_Position"
                    i = ThisWorkbook.sheets(Feuil).Range("BH6:GG6").Find("Selector_Lever_Position", , , xlPart).Column
                 End If
                 If Term <> "" Then
                     If ThisWorkbook.sheets(Feuil).ChartObjects(p.Name).Chart.Axes(xlCategory).AxisTitle.text = Term Then
                        ThisWorkbook.sheets(Feuil).ChartObjects(p.Name).Chart.FullSeriesCollection(1).XValues = "={" & CompteNew(ThisWorkbook.sheets(Feuil).Name, i) & "}"
                    Else
                        ThisWorkbook.sheets(Feuil).ChartObjects(p.Name).Chart.FullSeriesCollection(1).Values = "={" & CompteNew(ThisWorkbook.sheets(Feuil).Name, i) & "}"
                     End If
                  End If
            End If
        End If
     Next p
    
End Function
Function poliGraph(Feuil As String)
     Dim p As shape
     Dim i As Long
     Dim j As Long
     Dim Term As String
     
     For Each p In ThisWorkbook.sheets(Feuil).Shapes
        If p.Name = "Graphique_0" Or p.Name = "Graphique_1" Then
            If InStr(1, ThisWorkbook.sheets(Feuil).ChartObjects(p.Name).Chart.ChartTitle.text, "New Selector_Lever_Position") <> 0 _
            And InStr(1, ThisWorkbook.sheets(Feuil).ChartObjects(p.Name).Chart.ChartTitle.text, "Old Selector_Lever_Position") <> 0 Then
                If ThisWorkbook.sheets(Feuil).ChartObjects(p.Name).Chart.Axes(xlCategory).AxisTitle.text = "Old Selector_Lever_Position" Then
                      i = ThisWorkbook.sheets(Feuil).Rows(6).Find("New Selector_Lever_Position", , , xlPart).Column
                      j = ThisWorkbook.sheets(Feuil).Rows(6).Find("Old Selector_Lever_Position", , , xlPart).Column
                Else
                      i = ThisWorkbook.sheets(Feuil).Rows(6).Find("Old Selector_Lever_Position", , , xlPart).Column
                      j = ThisWorkbook.sheets(Feuil).Rows(6).Find("New Selector_Lever_Position", , , xlPart).Column
                 End If
                    ThisWorkbook.sheets(Feuil).ChartObjects(p.Name).Chart.FullSeriesCollection(1).XValues = "={" & Split(allSelector_Lever_PositionConfig, "#")(0) & "}"
                    ThisWorkbook.sheets(Feuil).ChartObjects(p.Name).Chart.FullSeriesCollection(1).Values = "={" & Split(allSelector_Lever_PositionConfig, "#")(1) & "}"
                    ThisWorkbook.sheets(Feuil).ChartObjects(p.Name).Chart.Axes(xlValue).MaximumScaleIsAuto = True
          Else
                Term = ""
                If InStr(1, ThisWorkbook.sheets(Feuil).ChartObjects(p.Name).Chart.ChartTitle.text, "New Selector_Lever_Position") <> 0 Then
                    Term = "New Selector_Lever_Position"
                    i = ThisWorkbook.sheets(Feuil).Rows(6).Find("New Selector_Lever_Position", , , xlPart).Column
                ElseIf InStr(1, ThisWorkbook.sheets(Feuil).ChartObjects(p.Name).Chart.ChartTitle.text, "Old Selector_Lever_Position") <> 0 Then
                   Term = "Old Selector_Lever_Position"
                    i = ThisWorkbook.sheets(Feuil).Rows(6).Find("Old Selector_Lever_Position", , , xlPart).Column
                 ElseIf InStr(1, ThisWorkbook.sheets(Feuil).ChartObjects(p.Name).Chart.ChartTitle.text, "Selector_Lever_Position") <> 0 Then
                   Term = "Selector_Lever_Position"
                    i = ThisWorkbook.sheets(Feuil).Rows(6).Find("Selector_Lever_Position", , , xlPart).Column
                 End If
                 If Term <> "" Then
                     If ThisWorkbook.sheets(Feuil).ChartObjects(p.Name).Chart.Axes(xlCategory).AxisTitle.text = Term Then
                        ThisWorkbook.sheets(Feuil).ChartObjects(p.Name).Chart.FullSeriesCollection(1).XValues = "={" & CompteNew(ThisWorkbook.sheets(Feuil).Name, i) & "}"
                    Else
                        ThisWorkbook.sheets(Feuil).ChartObjects(p.Name).Chart.FullSeriesCollection(1).Values = "={" & CompteNew(ThisWorkbook.sheets(Feuil).Name, i) & "}"
                     End If
                  End If
            End If
        End If
     Next p
    
End Function
Function CompteNew(Feuil As String, Cible As Long) As String
     Dim n As String
     Dim i As Long
      With ThisWorkbook.Worksheets(Feuil)
        If .Cells(.Rows.Count, Cible).End(xlUp).row = 6 Then CompteNew = 0
          For i = 7 To .Cells(.Rows.Count, Cible).End(xlUp).row
                   n = getSelector_Lever_PositionConfig(.Cells(i, Cible))
                    If CompteNew = "" Then
                       CompteNew = IIf(Len(n) > 0, n, "0")
                    Else
                       CompteNew = CompteNew & "," & IIf(Len(n) > 0, n, "0")
                    End If
           Next i
      End With
End Function


Public Function GetTaux(vehCible As String)
    Dim v
    Dim i As Long
    GetTaux = 0
    v = ThisWorkbook.sheets("TARGET VEHICLE").UsedRange.Value
    For i = 2 To UBound(v, 1)
    
        If StrComp("Rate of low points", v(i, 1), vbTextCompare) = 0 And _
           StrComp(ThisWorkbook.sheets("HOME").Range("DriveVersion").Value, v(i, 2), vbTextCompare) = 0 And _
           StrComp(vehCible, v(i, 3), vbTextCompare) = 0 And _
           StrComp(ThisWorkbook.sheets("HOME").Range("Mode").Value, v(i, 4), vbTextCompare) = 0 Then

                 
                          If IsNumeric(v(i, 5)) Then
                                GetTaux = v(i, 5)
                                Erase v
                                Exit Function
                            End If
                    
                        
        End If
    Next i
    Erase v
  
End Function
Public Function GetTauxDyn(vehCible As String)
    Dim v
    Dim i As Long
    GetTauxDyn = 0
    v = ThisWorkbook.sheets("TARGET VEHICLE").UsedRange.Value
    For i = 2 To UBound(v, 1)
    
        If StrComp("Rate of low points", v(i, 1), vbTextCompare) = 0 And _
           StrComp(ThisWorkbook.sheets("HOME").Range("DriveVersion").Value, v(i, 2), vbTextCompare) = 0 And _
           StrComp(vehCible, v(i, 3), vbTextCompare) = 0 And _
           StrComp(ThisWorkbook.sheets("HOME").Range("Mode").Value, v(i, 4), vbTextCompare) = 0 Then

                   
                            If IsNumeric(v(i, 6)) Then
                               GetTauxDyn = v(i, 6)
                                Erase v
                                Exit Function
                            End If
                     
                        
        End If
    Next i
    Erase v
  
End Function
Function GraphToValueDyn(Feuil As String)
      Dim p As shape
     Dim i As Integer
     Dim j As Long
     Dim Term As String
     Dim Adds As String
     Dim VSa As String
     Dim tabDCheck() As String
     Dim tSave() As Variant
     With ThisWorkbook.sheets(Feuil)
             For Each p In .Shapes
                If (p.Name = "Graphique_00" Or p.Name = "Graphique_11") And p.Visible = True Then
                 Call imageToBackgroud(p.Name, Feuil)
                    VSa = p.Chart.FullSeriesCollection(1).formula
                    For i = 1 To 2
                        Term = Split(VSa, ",")(i)
                        If InStr(1, Term, "!") <> 0 Then
                             If i = 1 Then
                                Adds = Join(WorksheetFunction.Transpose(.Range(Right(Term, Len(Term) - InStr(1, Term, "!")))), ";")
                            Else
                                Adds = Join(WorksheetFunction.Transpose(.Range(Term)), ";")
                            End If
                             For j = 1 To Len(Adds)
                                     Adds = replace(Adds, ",,", ",0,")
                             Next j
                             
                             
                            If InStr(1, Adds, ";") = 0 Then
                                ReDim tabDCheck(0)
                                tabDCheck(0) = Adds
                            Else
                                tabDCheck = Split(Adds, ";")
                            End If
                            
                           
                            ReDim tSave(UBound(tabDCheck))
                            For j = 0 To UBound(tabDCheck)
                                    If IsLong(tabDCheck(j)) = True Then
                                            If CLng(tabDCheck(j)) <> tabDCheck(j) Then
                                                    If Len(tabDCheck(j)) > 5 Then
                                                         tSave(j) = CDbl(Left(tabDCheck(j), 5))
                                                    Else
                                                         tSave(j) = CDbl(tabDCheck(j))
                                                    End If
                                            Else
                                                    tSave(j) = val(tabDCheck(j))
                                            End If
                                    Else
                                        tSave(j) = 0
                                    End If
                             Next j
                             If UBound(tSave) > 0 Then
                                     If i = 1 Then
                                          On Error Resume Next
                                          p.Chart.FullSeriesCollection(1).XValues = tSave
                                          ERR.Clear
                                          On Error GoTo 0
                                    Else
                                         On Error Resume Next
                                         p.Chart.FullSeriesCollection(1).Values = tSave
                                         ERR.Clear
                                         On Error GoTo 0
                                    End If
                             End If
                             
                        End If
                    Next i
                End If
            Next p
    End With
End Function

Function GraphToValue(Feuil As String)
      Dim p As shape
     Dim i As Integer
     Dim j As Long
     Dim Term As String
     Dim Adds As String
     Dim VSa As String
     Dim tabDCheck() As String
     Dim tSave() As Variant
     With ThisWorkbook.sheets(Feuil)
             For Each p In .Shapes
                If (p.Name = "Graphique_0" Or p.Name = "Graphique_1") And p.Visible = True Then
                    Call imageToBackgroud(p.Name, Feuil)
                    VSa = p.Chart.FullSeriesCollection(1).formula
                    For i = 1 To 2
                        Term = Split(VSa, ",")(i)
                        If InStr(1, Term, "!") <> 0 Then
                             If i = 1 Then
                                Adds = Join(WorksheetFunction.Transpose(.Range(Right(Term, Len(Term) - InStr(1, Term, "!")))), ";")
                            Else
                                Adds = Join(WorksheetFunction.Transpose(.Range(Term)), ";")
                            End If
                             For j = 1 To Len(Adds)
                                     Adds = replace(Adds, ",,", ",0,")
                             Next j
                             
                             
                            If InStr(1, Adds, ";") = 0 Then
                                ReDim tabDCheck(0)
                                tabDCheck(0) = Adds
                            Else
                                tabDCheck = Split(Adds, ";")
                            End If
                            
                           
                            ReDim tSave(UBound(tabDCheck))
                            For j = 0 To UBound(tabDCheck)
                                    If IsLong(tabDCheck(j)) = True Then
                                            If CLng(tabDCheck(j)) <> tabDCheck(j) Then
                                                    If Len(tabDCheck(j)) > 5 Then
                                                         tSave(j) = CDbl(Left(tabDCheck(j), 5))
                                                    Else
                                                         tSave(j) = CDbl(tabDCheck(j))
                                                    End If
                                            Else
                                                    tSave(j) = val(tabDCheck(j))
                                            End If
                                    Else
                                        tSave(j) = 0
                                    End If
                             Next j
                             If UBound(tSave) > 0 Then
                                     If i = 1 Then
                                          On Error Resume Next
                                          p.Chart.FullSeriesCollection(1).XValues = tSave
                                          ERR.Clear
                                          On Error GoTo 0
                                    Else
                                         On Error Resume Next
                                         p.Chart.FullSeriesCollection(1).Values = tSave
                                         ERR.Clear
                                         On Error GoTo 0
                                    End If
                             End If
                             
                        End If
                    Next i
                End If
            Next p
    End With
End Function

Function getSelector_Lever_PositionConfig(configName As String) As String
    Dim r As Range
    Set r = ThisWorkbook.sheets("Configurations").Range("DMU")
    getSelector_Lever_PositionConfig = ""
    While Len(r.Value) > 0
            If UCase(r.Offset(0, 1)) = UCase(configName) Then
                If IsNumeric(r.Value) Then
                    getSelector_Lever_PositionConfig = r.Value
                    Exit Function
                End If
            End If
            Set r = r.Offset(1, 0)
    Wend
End Function

Function allSelector_Lever_PositionConfig() As String
    Dim r As Range
    Dim s1 As String, s2 As String
    Set r = ThisWorkbook.sheets("Configurations").Range("DMU")
    allSelector_Lever_PositionConfig = ""
    While Len(r.Value) > 0
           If s1 = "" Then s1 = r.Offset(0, 1) Else s1 = s1 & ", " & r.Offset(0, 1)
           If s2 = "" Then s2 = r.Value Else s2 = s2 & ", " & r.Value
            Set r = r.Offset(1, 0)
    Wend
    allSelector_Lever_PositionConfig = s2 & "#" & s1
End Function

Function CreateNew(titres As String)
    Dim lastRow As Long
    Dim tableL(13) As Integer
    Dim i As Integer
    Dim j As Integer
    
    tableL(1) = 3
    tableL(2) = 5
    tableL(3) = 7
    tableL(4) = 10
    tableL(5) = 12
    tableL(6) = 14
    tableL(7) = 17
    tableL(8) = 19
    tableL(9) = 22
     tableL(10) = 24
    tableL(11) = 27
    tableL(12) = 29
    tableL(13) = 31
    
    If Len(titres) = 0 Then
        MsgBox "Aucune Valeur", vbCritical, "Odriv"
    ElseIf InStr(1, ";" & getSDV1 & ";", ";" & UCase(titres) & ";") <> 0 Then
         MsgBox "Existe Déjà", vbCritical, "Odriv"
    Else
       
        lastRow = ThisWorkbook.Worksheets("PARAMETRES GRAPH").UsedRange.Rows.Count + 1
        With ThisWorkbook.Worksheets("PARAMETRES GRAPH")
                .Rows(lastRow - 32 & ":" & lastRow - 1).Copy Destination:=.Cells(lastRow, 1)
                .Cells(lastRow, 1) = titres
                
                .Cells(lastRow + 1, 4) = "Désactivè"
                .Cells(lastRow + 8, 4) = "Désactivè"
                .Cells(lastRow + 15, 4) = "Désactivè"
                .Cells(lastRow + 18, 4) = "Désactivè"
                .Cells(lastRow + 21, 4) = "Désactivè"
              
                For j = 1 To UBound(tableL)
                    For i = 2 To 4
                        .Cells(lastRow + tableL(j), i) = ""
                    Next i
                Next j
                
                
        End With
     
    End If
End Function

Function getSDV1()
    Dim v As String
    Dim c As Range
    Dim i As Long
    Dim j As Long
    
    i = getLastRowRating
    v = ""

    With ThisWorkbook.Worksheets("RATING")
        j = 23
        While j <= i
            Set c = .Cells(j, 4)
            If Len(c.Value) > 0 And sheetExists(c.Value) = True And .Rows(j).Hidden = False Then
                If v = "" Then v = c.Value Else v = v & ";" & c.Value
            End If
            j = j + 1
        Wend
    End With
    getSDV1 = v
End Function

Function IsLong(v As Variant) As Long
         Dim checkLong As Long
         IsLong = False
       
        If IsNumeric(v) = True Then
            On Error Resume Next
            checkLong = CLng(v)
            If ERR.Number <> 0 Then
                ERR.Clear
            Else
                IsLong = True
            End If
            
        End If
End Function

Function TransPart(partEng As String)
    Dim startIndex As Long
    Dim endIndex As Long
    
    startIndex = InStr(partEng, "{")
    endIndex = InStr(partEng, "}")
    
    If InStr(1, partEng, "{") <> 0 Then
        TransPart = Mid(partEng, 1, startIndex - 1) & "``" & _
        Mid(partEng, startIndex + 1, endIndex - startIndex - 1) & "``" & _
        Split(Mid(partEng, endIndex + 2), ",")(0)
    Else
        TransPart = replace(partEng, ",", "``")
    End If
End Function
Function replaceTerm(Term As String, onglet As String)
    With ThisWorkbook.sheets(onglet)
            If InStr(1, Term, "!") <> 0 Then
                replaceTerm = .Range(Right(Term, Len(Term) - InStr(1, Term, "!")))
            Else
                replaceTerm = .Range(Term)
            End If
      End With
End Function
Sub TESTE()
    Call imageToBackgroud("Graphique_0", "(TO) KD - tip in downshift")
End Sub
Function imageToBackgroud(NomShape As String, onglet As String)
    Exit Function
    Dim cheminDossier As String, cheminBase As String
    Dim nomFichier As String
    Dim ws As Worksheet
    Dim cheminComplet As String
    Dim plage As Range
    Dim plageCopy As Range
    Dim sh As Variant
    Dim LigneDebut As Long, Types As Long
    Dim colStart As Long, rowStart As Long
    Dim ongletShapes As Chart
    Dim nbIntervalles
    Dim limitLigne
    On Error GoTo Ers
    
    
    LigneDebut = Rating_Priorisation.LigneSeeting(onglet)
    If LigneDebut = 0 Then Exit Function
    Types = Rating_Priorisation.StartConFig(LigneDebut)
    If Types = 0 Then Exit Function
    colStart = 10
    rowStart = Types - 23
    
    With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
        Set plage = .Range(.Cells(rowStart, colStart), .Cells(PDrow(.Cells(rowStart, colStart)), PDcol(.Cells(rowStart, colStart))))
    End With
  
    cheminBase = "C:\ODRIV\"
    cheminDossier = cheminBase & "background\"
    nomFichier = "background.PNG"
    cheminComplet = cheminDossier & nomFichier
    If (Dir(cheminBase, vbDirectory) = vbNullString) Then MkDir cheminBase
    If (Dir(cheminDossier, vbDirectory) = vbNullString) Then MkDir cheminDossier
    
     Set ws = ThisWorkbook.sheets("Résultats")
    ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Outline.ShowLevels RowLevels:=2
    plage.Copy Destination:=ws.Range("AB1")
    ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Outline.ShowLevels RowLevels:=1
    
    
    Set plage = ws.Range(ws.Cells(1, 28), ws.Cells(ws.Cells(ws.Rows.Count, 28).End(xlUp).row, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column))
    Set ongletShapes = ThisWorkbook.sheets(onglet).ChartObjects(NomShape).Chart
    Call AjusterHauteurLignesSelonIntervalleAxe(ongletShapes, "Résultats", plage)
    
    With ongletShapes.Axes(xlValue)
         nbIntervalles = ((.MaximumScale - .MinimumScale) / .MajorUnit)
    End With
    If nbIntervalles < ws.Cells(ws.Rows.Count, 28).End(xlUp).row Then
        limitLigne = nbIntervalles
    Else
         limitLigne = ws.Cells(ws.Rows.Count, 28).End(xlUp).row
    End If
   
    Set plageCopy = ws.Range(ws.Cells(1, 28), ws.Cells(limitLigne, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column))
    plageCopy.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    ws.Visible = xlSheetVisible
    If ws.Shapes.Count <> 0 Then
        For Each sh In ws.Shapes
            If sh.Name <> "xCopy" Then sh.Delete
        Next sh
    End If
     ws.Activate
     ws.ChartObjects("xCopy").Activate
    ws.Shapes("xCopy").Chart.Paste
    plage.Clear
    ws.Visible = xlSheetVeryHidden
    ThisWorkbook.sheets(onglet).Activate
   
  
   ws.Shapes("xCopy").Height = ongletShapes.PlotArea.InsideHeight
   ws.Shapes("xCopy").Chart.Shapes.Range(Array(1)).Width = ws.Shapes("xCopy").Width
   ws.Shapes("xCopy").Chart.Shapes.Range(Array(1)).Flip msoFlipVertical
   ws.Shapes("xCopy").Chart.Export Filename:=cheminComplet, filtername:="PNG"
   
    
     Application.DisplayAlerts = False
     ws.Shapes("xCopy").Chart.Shapes.Range(Array(1)).Delete
    Application.DisplayAlerts = True

    Set ws = ThisWorkbook.sheets(onglet)
    Dim shp As shape
    
    
    Set shp = ws.Shapes(NomShape)
    If Not shp Is Nothing Then
        shp.Chart.PlotArea.Format.Fill.UserPicture cheminComplet
        shp.Chart.PlotArea.Format.Fill.Transparency = 0.5
    End If
'    ws.Range("A1").Select
    On Error GoTo 0
    Kill cheminComplet
    
Ers:
    If ERR.Number <> 0 Then
'        MsgBox ERR.description, vbCritical, "ODRIV"
    End If
End Function

Function AjusterHauteurLignesSelonIntervalleAxe(GraphChart As Chart, feuilleTampon As String, plage As Range)
    Dim ws As Worksheet
    Dim nbIntervalles As Long
    Dim hauteurTotaleGraphique As Double
    Dim hauteurParIntervalle As Double
    Dim hauteurLigne As Double
    Dim i As Integer


    Set ws = ThisWorkbook.sheets(feuilleTampon)
'    Set plage = ws.Range("A1:D5")
    With GraphChart.Axes(xlValue)
        nbIntervalles = (.MaximumScale - .MinimumScale) / .MajorUnit
    End With
    hauteurTotaleGraphique = GraphChart.PlotArea.InsideHeight
    hauteurParIntervalle = hauteurTotaleGraphique / nbIntervalles
   
'    hauteurLigne = (hauteurParIntervalle * nbIntervalles) / plage.Rows.Count
  
    For i = 1 To plage.Rows.Count
        plage.Rows(i).RowHeight = hauteurParIntervalle
    Next i
    
    If plage.Rows.Count < nbIntervalles Then
        plage.Rows(i - 1).RowHeight = hauteurParIntervalle * ((nbIntervalles - plage.Rows.Count) + 1)
    End If
    
End Function


Function PDrow(r As Range) As Integer
  Dim dcol As Integer
  dcol = r.row
  Set r = r.Offset(1, 0)
  While r.Value <> "" And Len(r.Value) > 0
      dcol = dcol + 1
      Set r = r.Offset(1, 0)
  Wend
  
  PDrow = dcol
End Function



Function PDcol(r As Range) As Integer
 Dim dcol As Integer
  dcol = r.Column
  Set r = r.Offset(0, 1)
  While r.Value <> "" And Len(r.Value) > 0
      dcol = dcol + 1
      Set r = r.Offset(0, 1)
  Wend
  PDcol = dcol
End Function






