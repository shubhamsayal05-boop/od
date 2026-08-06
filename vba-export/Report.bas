Attribute VB_Name = "Report"
Option Explicit
Private chemin As String
Private getLower As String
Private getLowerDyn As String
Private SDVListe() As String
Sub FormReport()
'        MsgBox "En Cours de Dev Bientot Disponible", vbCritical, "Odriv"
        Preremplissage.Show
End Sub
Function REPORTC(fieldR As Variant)
    Dim objword As Object
    Dim objDoc As Object
    Dim FieldSrp As Variant
    Dim c As Object
    Dim nbpages As Integer
    Dim msg_err
    
    On Error GoTo Ers
    Application.EnableEvents = False
    FieldSrp = fieldR
    Unload Preremplissage
    ProgressLoad
    ProgressTitle ("Creation Document")
    Call getSDV
    Call CreatedOC
    Set objword = CreateObject("Word.Application")
    'a change
    objword.Visible = True
    Set objDoc = objword.Documents.Open(chemin)
    Application.DisplayStatusBar = False
   getLower = ""
   getLowerDyn = ""
   ProgressTitle ("Copy Rating")
    Call CopyHome(objword, objDoc)
    Call CopyVersion(objword, objDoc)
    Call CopyRating1(objword, objDoc)
    Call CopyRating2(objword, objDoc)
    'Call recopyVal(objword, objDoc)
    ProgressTitle ("MAJ Titres")
    Call InsertPic(objword, objDoc)
    Call VerifierPageAvecEnteteVide(objword, objDoc)
    ProgressTitle ("Update Summary")
    Call updateSummary(objDoc)
    ProgressTitle ("Preremplissage")
    Call Remplissage(FieldSrp, objword, objDoc)
'    Call RedLowPoint(objWord, objDoc)
'    Call RedLowPointDyn(objWord, objDoc)
    Call UpdateFieldsDoc(objDoc, objword)
    
    'Mode View Print
    objword.ActiveWindow.ActivePane.View.Type = 3
    objDoc.Save
     Application.DisplayStatusBar = True
     ProgressExit
    objword.Visible = True
    objword.Activate
    'Call VerifierPageAvecEnteteVide(objword, objDoc)
    
'    ObjDoc.Close False
'    objWord.Quit
    
    
    objDoc.ActiveWindow.VERTICALPERCENTSCROLLED = 0
    If Not objword Is Nothing Then Set objword = Nothing
    If Not objDoc Is Nothing Then Set objDoc = Nothing
    
    
     Application.EnableEvents = True

Ers:
    If ERR.Number <> 0 Then
        msg_err = ERR.description
        If Not objDoc Is Nothing Then objDoc.Close False
        Application.EnableEvents = True
        Application.DisplayStatusBar = True
        ProgressExit
        MsgBox msg_err, vbCritical, "ODRIV"
        Unload PleaseWait
    End If
End Function
Sub VerifierPageAvecEnteteVide(ByRef objword As Object, objDoc As Object)
    
   Dim nbpages As Long
   Dim i As Long
    nbpages = objDoc.ActiveWindow.Panes(1).pages.Count
    'MsgBox nbpages
   For i = 11 To nbpages
     objword.Selection.GoTo What:=1, Which:=1, Count:=i
     objword.Selection.GoTo What:=-1, Name:="\page"
     
     'MsgBox Len(objword.Selection.Paragraphs(1).Range.Text)
    
     If Len(objword.Selection.Paragraphs(1).Range.text) = 1 Then
        
         objword.Selection.Bookmarks("\page").Range.Delete
         i = i - 1
         nbpages = nbpages - 1
     End If
   Next i
   'MsgBox nbpages
   
End Sub

Function UpdateFieldsDoc(objDoc As Object, objword As Object)
    Dim objField As Object
    For Each objField In objDoc.Fields
        objField.Update
    Next objField
    
End Function
Sub CreatedOC()
     Dim WDObj As Object
    Dim wordApp As Object
    Dim wordDoc As Object
    Application.ScreenUpdating = False
    Set WDObj = sheets("DNT").OLEObjects("REPORT")
    WDObj.Activate
    WDObj.Object.Application.Visible = True
    Set wordApp = WDObj.Object.Application
    Set wordDoc = wordApp.ActiveDocument
    chemin = ThisWorkbook.Path & "/Report_" & ThisWorkbook.Worksheets("HOME").Range("Project") & "_" & replace(Time, ":", "_") & ".docx"
    wordDoc.SaveAs (chemin)
    wordDoc.Close
    ThisWorkbook.Worksheets("HOME").Select
    Application.ScreenUpdating = True
    Load PleaseWait
    
    PleaseWait.Show vbModeless
    DoEvents
End Sub

'Sub test()
'    Dim c As Word.Document
'    Dim NbPages As Integer
'    Dim objword As Word.Application
'    Set objword = GetObject(, "Word.Application")
'    Set c = objword.ActiveDocument
'    NbPages = c.ActiveWindow.Panes(1).Pages.Count
'    MsgBox NbPages
'    objword.Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=NbPages - 1
'    objword.Selection.Delete
'    MsgBox NbPages
'End Sub

Function newSdvPage(objword As Object, objDoc As Object, nameSdv As String, Numb As String)

        With objword.Selection
          
'            .Style = objDoc.Styles("Titre 1")
'            .Range.ListFormat.RemoveNumbers NumberType:=1
            .typeText text:=Numb & " " & nameSdv
'            .Shading.Texture = 0
'            .Shading.ForegroundPatternColor = -16777216
'            .Shading.BackgroundPatternColor = -16777216
'            .Borders(-1).LineStyle = 0
'            .Borders(-2).LineStyle = 0
'            .Borders(-4).LineStyle = 0
            
            .HomeKey Unit:=5
            .Expand 5
            On Error Resume Next
            .style = objDoc.Styles("Heading 2")
            .style = objDoc.Styles("Title 2")
            .style = objDoc.Styles("Titre 2")
            
            On Error GoTo 0
            
            .Range.ListFormat.RemoveNumbers NumberType:=1
            .Font.ColorIndex = 1
            .Font.Bold = True
            .Font.Size = 14
            .EndKey Unit:=5
             .TypeParagraph
            .HomeKey Unit:=5
            .ParagraphFormat.Alignment = 0
           
        End With
        
End Function
Function insertPart(objword As Object, i As Integer)
  
    Dim j As Long
    Dim TaBS(6) As String
    Dim t As Integer
   
   TaBS(1) = "Synthesis"
   TaBS(2) = "Points visualisation 1"
   TaBS(3) = "Points visualisation 2"
   TaBS(4) = "Highest Criticality to improve:"
   TaBS(5) = "Lowest Criticality to improve:"

'   TaBS(5) = "->In coherence with the synthesis, show the adequate bad priorities"
'   TabS(6) = "->In coherence with the synthesis, show the adequate bad priorities"
 
   With objword.Selection

                .Font.ColorIndex = 6
                .typeText text:=TaBS(i)
                objword.ListGalleries(1).ListTemplates(1).ListLevels(1).NumberFormat = ChrW(61656)
                objword.ListGalleries(1).ListTemplates(1).ListLevels(1).Font.Name = "Wingdings"
                .Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
                objword.ListGalleries(1).ListTemplates(1), ContinuePreviousList:= _
                False, ApplyTo:=0, DefaultListBehavior:=2
                .HomeKey Unit:=5
                .Expand 5
                .Font.ColorIndex = 1
                .Font.Bold = True
                .Font.Size = 12
                .EndKey Unit:=5
                .TypeParagraph
                .Range.ListFormat.RemoveNumbers NumberType:=1
                
                 'objDoc.Bookmarks.Add Range:=.Range, Name:="S" & j & i & t
                .TypeParagraph
                .TypeParagraph
                .MoveUp Unit:=5, Count:=1
   End With


End Function
Function updateSummary(ByRef objword As Object)
    objword.tablesOfContents(1).Update
End Function

Function InsertPic(ByRef objword As Object, objDoc As Object)
    Dim i As Integer
    Dim j As Long
    Dim h As Integer
    Dim t As Integer
    Dim o As String
    Dim x As String
    Dim getParamSdv As String
    Dim v() As String
    Dim c As Object
    Dim nbpages As Integer
    
    objDoc.Bookmarks("StartOne").Select
   
    For j = 0 To UBound(SDVListe)
        o = SDVListe(j)
        ThisWorkbook.sheets(o).Cells.EntireColumn.Hidden = False
        For t = 1 To 2
            ProgressTitle ("Copie des données : " & o)
            If t = 1 Or (t = 2 And checkCriteriaDyn(o) = True And checkCorrespondancePriorityDyn(o) = True) Then
                If j = 0 And t = 1 Then
                Else
                   objword.Selection.InsertNewPage
                End If
               If t = 1 Then Call newSdvPage(objword, objDoc, UCase(o) & " DRIVABILITY", "2." & j + 1 & "." & t) Else Call newSdvPage(objword, objDoc, UCase(o) & " RESPONSIVENESS", "2." & j + 1 & "." & t)
               
                
                For i = 1 To 5
                    x = j & i & t
                    
                        If i = 1 Then
                            Call insertPart(objword, i)
                            Call CopySummary(objword, objDoc, o, t)
                           
                        ElseIf i = 2 Or i = 3 Then
                                If t = 1 Then
                                    getParamSdv = UpdateGraph.checkGraphEnable(o)
                                Else
                                    getParamSdv = UpdateGraphDyn.checkGraphEnable(o)
                                End If
                                If getParamSdv <> "" Then
                                        If i = 2 Then
                                                v = Split(getParamSdv, ";")
                                                For h = 0 To UBound(v)
                                                      If Split(v(h), ":")(0) = "Graphique_0" Or Split(v(h), ":")(0) = "Graphique_00" Then
                                                           If ThisWorkbook.Worksheets(o).Shapes(Split(v(h), ":")(0)).Visible = True Then
                                                                Call insertPart(objword, i)
                                                                If t = 1 Then
                                                                    Call CopyGraph0(objword, objDoc, o, t)
                                                                Else
                                                                    Call CopyGraph0(objword, objDoc, o, t)
                                                                End If
                                                           End If
                                                           Exit For
                                                      Else
                                                            If UpdateGraph.checkObject(CStr(Split(v(h), ":")(0)), o) <> "" Then
                                                                Call insertPart(objword, i)
                                                                If t = 1 Then
                                                                    
                                                                    Call CopyLeverAS(objword, objDoc, o, UpdateGraph.checkObject(CStr(Split(v(h), ":")(0)), o))
                                                                Else
                                                
                                                                     Call CopyLeverAS(objword, objDoc, o, UpdateGraphDyn.checkObject(CStr(Split(v(h), ":")(0)), o))
                                                                End If
                                                                
                                                            End If
                                                            Exit For
                                                      End If
                                                Next h
                                         Else
                                                If t = 1 Then
                                                    getParamSdv = UpdateGraph.checkGraphEnable(o)
                                                Else
                                                    getParamSdv = UpdateGraphDyn.checkGraphEnable(o)
                                                End If
                                               v = Split(getParamSdv, ";")
                                               If InStr(1, getParamSdv, "Graphique_1") <> 0 Or InStr(1, getParamSdv, "Graphique_11") Then
                                                    For h = 0 To UBound(v)
                                                         If Split(v(1), ":")(0) = "Graphique_1" Or InStr(1, getParamSdv, "Graphique_11") Then
                                                              If ThisWorkbook.Worksheets(o).Shapes(Split(v(1), ":")(0)).Visible = True Then
                                                                 Call insertPart(objword, i)
                                                                 Call CopyGraph1(objword, objDoc, o, t)
                                                                 
                                                              End If
                                                              Exit For
                                                          End If
                                                     Next h
                                                End If
                                         End If
                                 End If
                        ElseIf i = 4 Then
                            If t = 1 Then
                                Call CopyPriorityPoints(objword, objDoc, "Hight", ThisWorkbook.Worksheets(o), i)
                            Else
                                Call CopyPriorityPointsDyn(objword, objDoc, "Hight", ThisWorkbook.Worksheets(o), i)
                            End If
                        ElseIf i = 5 Then
                            If t = 1 Then
                                 Call CopyPriorityPoints(objword, objDoc, "Low", ThisWorkbook.Worksheets(o), i)
                            Else
                                Call CopyPriorityPointsDyn(objword, objDoc, "Low", ThisWorkbook.Worksheets(o), i)
                            End If
                        End If
                
                 
        '             If i <= 3 Then
        '               objDoc.InlineShapes(LastPicNumber(objDoc)).Width = 520
        '               objDoc.InlineShapes(LastPicNumber(objDoc)).LockAspectRatio = 0
        '               objDoc.InlineShapes(LastPicNumber(objDoc)).Height = 283.5
        '            End If
        '              x = x + 1
                
                Next i
            End If
      Next t
    Next j
    
    'Call VerifierPageAvecEnteteVide(objword, objDoc)
    
End Function

 
 
Function TakePic(objword As Object, Optional objDoc As Object, Optional r As Range, Optional s As shape, Optional x As String)
    
     On Error Resume Next
    Dim i As Integer
    
    For i = 1 To 10
        ERR.Clear
         If Not r Is Nothing Then Call COPYp(r) Else Call COPYp(, s)
        If x <> "" Then
            If Left(x, 1) = "A" Or UCase(Left(x, 3)) = "SYS" Then
                objDoc.Bookmarks(x).Select
            Else
                objDoc.Bookmarks("S" & x).Select
            End If
        End If
        With objword.Selection
               .PasteSpecial Link:=False, DataType:=4, Placement:=0, DisplayAsIcon:=False
        End With
        If ERR.Number = 0 Then Exit Function Else Application.Wait Now + TimeValue("0:00:02")
    Next i

End Function
Function TakePicSelection(objword As Object, Optional objDoc As Object, Optional r As Range, Optional s As shape)
    
     On Error Resume Next
    Dim i As Integer
    
    For i = 1 To 10
        ERR.Clear
         If Not r Is Nothing Then Call COPYp(r) Else Call COPYp(, s)
        
        With objword.Selection

               .PasteSpecial Link:=False, DataType:=4, Placement:=0, DisplayAsIcon:=False
        End With
        
        If ERR.Number = 0 Then
            objword.Selection.TypeParagraph
            objword.Selection.TypeParagraph
            Exit Function
        Else
            Application.Wait Now + TimeValue("0:00:02")
        End If
    Next i
End Function
Function CopyTable(objword As Object, Optional objDoc As Object, Optional r As Range, Optional s As shape, Optional x As String)
    
    On Error Resume Next
    Dim i As Integer
    
    For i = 1 To 10
        ERR.Clear
         If Not r Is Nothing Then Call COPYp(r) Else Call COPYp(, s)
        If x <> "" Then
            If Left(x, 1) = "A" Or UCase(Left(x, 3)) = "SYS" Then
                objDoc.Bookmarks(x).Select
            Else
                objDoc.Bookmarks("S" & x).Select
            End If
        End If
        With objword.Selection
              .PasteExcelTable False, False, False
              With objDoc.tables(4)
                    .Select
'                    .Columns.Width = 10
                    .Columns(3).Width = 50
                    .Columns(1).Width = 100
                    .Rows.Height = 5
               End With
        End With
        If ERR.Number = 0 Then Exit Function Else Application.Wait Now + TimeValue("0:00:02")
    Next i
    
    
End Function

Function CopyVersion(objword As Object, objDoc As Object)
    Dim plage As Range
    Set plage = ThisWorkbook.sheets("DocVersions").Range("B6:D11")
    Call CopyTable(objword, objDoc, plage, , "SysLink")
End Function
Function CopyRating1(objword As Object, objDoc As Object)
    Dim plage As Range
    Dim sh As Worksheet
    Dim colD As Integer, colC As Integer, totC As Integer
'    Set Plage = ThisWorkbook.sheets("RATING").Range("A1:V10")
'
'    Call TakePic(objWord, ObjDoc, Plage, , "ACRating2")
'    ObjDoc.InlineShapes(1).Width = 500
   

    Application.DisplayAlerts = False
    sheets("RATING").Copy before:=sheets(ThisWorkbook.sheets.Count)
    Application.DisplayAlerts = True
    Set sh = sheets(ThisWorkbook.sheets.Count - 1)
    With sh
'        .Rows("1:8").EntireRow.Hidden = True
'        .Rows("15:19").EntireRow.Hidden = True
         totC = .Cells(10, .Columns.Count).End(xlToLeft).Column
        .Shapes.Range(Array("Image 11")).Delete
         Set plage = .Range("B1:" & .Cells(18, totC).Address)
         Call TakePic(objword, objDoc, plage, , "SysDr")
         
       
        
        colD = .Rows("21:22").Find(What:="Responsiveness Lowest Events", LookAt:=xlWhole).Column + 1
        colC = .Range("colPD1").Column
       .Columns(colLettre(colC) & ":" & colLettre(colD)).EntireColumn.Hidden = True
       colD = .Rows("21:22").Find(What:="Drivability Lowest Events", LookAt:=xlWhole).Column
        Set plage = .Range("B20:" & .Cells(100, colD).Address)
        Call TakePic(objword, objDoc, plage, , "SysDynT")
       
    End With
    Application.DisplayAlerts = False
    sh.Delete
    Application.DisplayAlerts = True
End Function
Function CopySummary(objword As Object, objDoc As Object, SDV As String, x As Integer)
    Dim plage As Range
    
    If x = 1 Then
        Set plage = ThisWorkbook.sheets(SDV).Range("B4:K22")
        Call TakePicSelection(objword, objDoc, plage)
    ElseIf x = 2 Then
        Set plage = ThisWorkbook.sheets(SDV).Range("BI4:BR22")
        Call TakePicSelection(objword, objDoc, plage)
    End If
    
End Function
Function CopyGraph0(objword As Object, objDoc As Object, SDV As String, x As Integer)
           If x = 1 Then
                ThisWorkbook.sheets(SDV).Shapes.Range(Array("Keys", "Graphique_0")).Group.Name = "Groupage"
                Call TakePicSelection(objword, objDoc, , ThisWorkbook.sheets(SDV).Shapes("Groupage"))
                ThisWorkbook.sheets(SDV).Shapes.Range(Array("Groupage")).Ungroup
            ElseIf x = 2 Then
                ThisWorkbook.sheets(SDV).Shapes.Range(Array("Keyss", "Graphique_00")).Group.Name = "Groupage"
                Call TakePicSelection(objword, objDoc, , ThisWorkbook.sheets(SDV).Shapes("Groupage"))
                ThisWorkbook.sheets(SDV).Shapes.Range(Array("Groupage")).Ungroup
            End If
            
'            ThisWorkbook.sheets(sdv).Shapes.Range(Array("Keys", "Graphique_0")).Group.Name = "Groupage"
'            Call TakePic(objWord, ObjDoc, , ThisWorkbook.sheets(sdv).Shapes("Groupage"), x)
'            ThisWorkbook.sheets(sdv).Shapes.Range(Array("Groupage")).Ungroup
End Function
Function CopyGraph1(objword As Object, objDoc As Object, SDV As String, x As Integer)
            
            If x = 1 Then
                Call TakePicSelection(objword, objDoc, , ThisWorkbook.sheets(SDV).Shapes("Graphique_1"))
            ElseIf x = 2 Then
                Call TakePicSelection(objword, objDoc, , ThisWorkbook.sheets(SDV).Shapes("Graphique_11"))
            End If
End Function
Function CopyLeverAS(objword As Object, objDoc As Object, SDV As String, tables As String)
      
            Call TakePicSelection(objword, objDoc, getRangeTable(SDV, tables))
       
End Function

Function CopyHome(objword As Object, objDoc As Object)
    Dim plage As Range
    Set plage = ThisWorkbook.sheets("HOME").Range("B6:L24")

    Call TakePic(objword, objDoc, plage, , "SysHome")
End Function
Function CopyPriorityPoints(objword As Object, objDoc As Object, Filt As String, Shts As Worksheet, Pos As Integer)
    Dim plage As Range
    Dim colonne As Integer
    Dim lastRow As Long
    Dim TotalRow As Long
    Dim x As Long
    Dim i As Integer
    Dim j As Integer
    Dim cD
    Dim trouveCol As Boolean
    Dim cG As String
    Dim tabPS() As String
    Dim colPS As String
    With Shts
         For x = 13 To 15
            If .Cells(.Rows.Count, x).End(xlUp).row > TotalRow Then TotalRow = .Cells(.Rows.Count, x).End(xlUp).row
        Next x
        
       colonne = getLastColumnDrivability(Shts.Name) - 1
       
        Call HideC3(Shts.Name, "driv")
        If FilterPriority(Shts, Filt) = True And colonne <> 0 Then
            For x = 13 To 15
                If .Cells(.Rows.Count, x).End(xlUp).row > lastRow Then lastRow = .Cells(.Rows.Count, x).End(xlUp).row
            Next x
            Set plage = .Range(.Cells(3, 13), .Cells(lastRow, colonne))
'            Plage.Columns.AutoFit
            'Rec Point Bas_________________________
            If UCase(Filt) = "HIGHT" Then
                    cG = getColumnPoint(Shts.Name)
                    trouveCol = False
                    If cG <> "" Then
                            If InStr(1, cG, ";") = 0 Then
                                ReDim tabPS(0)
                                tabPS(0) = cG
                            Else
                                tabPS = Split(cG, ";")
                            End If
                            colPS = Shts.Name & "#"
                           
                            For Each cD In plage.Rows
                               
                                If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                                        colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                                        For j = 0 To UBound(tabPS)
                                               If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                                     colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                                     trouveCol = True
                                               End If
                                        Next j
                                        If cD.row < (plage.Rows.Count + plage.row) - 1 Then
                                              colPS = colPS & ";"
                                        End If
                                  End If
                            Next cD
                            If trouveCol = True Then
                               
                                If getLower = "" Then getLower = colPS Else getLower = getLower & "||" & colPS
                            End If
                    End If
            End If
            '_____________________________________
            Call insertPart(objword, Pos)
            Call TakePicSelection(objword, objDoc, plage)
             objDoc.InlineShapes(objDoc.InlineShapes.Count).Width = 500
            Call UnFilterPriority(Shts, TotalRow)
       
        End If
    End With
End Function
Function CopyPriorityPointsDyn(objword As Object, objDoc As Object, Filt As String, Shts As Worksheet, Pos As Integer)
    Dim plage As Range
    Dim colonne As Integer
    Dim lastRow As Long
    Dim TotalRow As Long
    Dim x As Long
    Dim i As Integer
    Dim j As Integer
    Dim cD
    Dim trouveCol As Boolean
    Dim cG As String
    Dim tabPS() As String
    Dim colPS As String
    With Shts
         For x = 72 To 74
            If .Cells(.Rows.Count, x).End(xlUp).row > TotalRow Then TotalRow = .Cells(.Rows.Count, x).End(xlUp).row
        Next x
      
       colonne = getLastColumnDinamyc(Shts.Name) - 1
      
        Call HideC3(Shts.Name, "dyn")
        If FilterPriorityDyn(Shts, Filt) = True And colonne <> 0 Then
            For x = 72 To 74
                If .Cells(.Rows.Count, x).End(xlUp).row > lastRow Then lastRow = .Cells(.Rows.Count, x).End(xlUp).row
            Next x
            Set plage = .Range(.Cells(3, 72), .Cells(lastRow, colonne))
'            Plage.Columns.AutoFit
            'Rec Point Bas_________________________
            If UCase(Filt) = "HIGHT" Then
                    cG = getColumnPoint(Shts.Name)
                    trouveCol = False
                    If cG <> "" Then
                            If InStr(1, cG, ";") = 0 Then
                                ReDim tabPS(0)
                                tabPS(0) = cG
                            Else
                                tabPS = Split(cG, ";")
                            End If
                            colPS = Shts.Name & "#"
                           
                            For Each cD In plage.Rows
                               
                                If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                                        colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 73)
                                        For j = 0 To UBound(tabPS)
                                               If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                                     colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                                     trouveCol = True
                                               End If
                                        Next j
                                        If cD.row < (plage.Rows.Count + plage.row) - 1 Then
                                              colPS = colPS & ";"
                                        End If
                                  End If
                            Next cD
                            If trouveCol = True Then
                               
                                If getLowerDyn = "" Then getLowerDyn = colPS Else getLowerDyn = getLowerDyn & "||" & colPS
                            End If
                    End If
            End If
            '_____________________________________
            Call insertPart(objword, Pos)
            Call TakePicSelection(objword, objDoc, plage)
             objDoc.InlineShapes(objDoc.InlineShapes.Count).Width = 500
            Call UnFilterPriority(Shts, TotalRow)
       
        End If
    End With
End Function

Function COPYp(Optional plage As Range, Optional s As shape)
   ERR.Clear
'    On Error Resume Next
    If Not plage Is Nothing Then plage.Copy Else s.Copy
'    While Err.Number <> 0
'        If Not Plage Is Nothing Then Plage.CopyPicture Appearance:=xlScreen, Format:=xlPicture Else S.CopyPicture Appearance:=xlScreen, Format:=xlPicture
'        If Err.Number = 0 Then Exit Function Else Err.Clear
'    Wend
End Function

Function LastPicNumber(objDoc As Object) As Integer
    Dim p
    Dim i As Integer
    For Each p In objDoc.InlineShapes
        i = i + 1
    Next p
    LastPicNumber = i
End Function

Function FilterPriority(Shts As Worksheet, Filt As String) As Boolean
    Dim r As Range
    Dim lastRow As Long
    Dim colonne As Integer
    Dim i As Long
    Dim x As Integer
    Dim TabCrit As String
    Dim VCrit(1) As String
    Dim p As Integer
   x = 0
  
    VCrit(1) = ""
    FilterPriority = False
    With Shts
                     colonne = 13
                     lastRow = .Cells(.Rows.Count, colonne).End(xlUp).row
                     If lastRow > 7 Then
                        TabCrit = ";;" & Join(WorksheetFunction.Transpose(.Range(.Cells(7, colonne), .Cells(lastRow, colonne))), ";;") & ";;"
                    Else
                         TabCrit = ";;" & .Cells(7, colonne) & ";;"
                    End If
                     
                             For p = 1 To 4
                                If Filt = "Hight" Then
                                    If InStr(1, TabCrit, ";;" & CStr(p) & ";;") <> 0 And x < 1 Then
                                              VCrit(1) = CStr(p)
                                              x = x + 1
                                    End If
                                ElseIf Filt = "Low" Then
                                     If InStr(1, TabCrit, ";;" & CStr(p) & ";;") <> 0 And x < 2 Then
                                          
                                          If x = 1 Then
                                             VCrit(1) = CStr(p)
                                         End If
                                          x = x + 1
                                    End If
                                End If
                            Next p
                    
                             
                    
                   If VCrit(1) = "" Then Exit Function
                   
                
                      For i = 7 To lastRow
                          If (.Cells(i, colonne)) <> VCrit(1) Then
                              If r Is Nothing Then Set r = .Cells(i, colonne) Else Set r = Union(r, .Cells(i, colonne))
                          Else
                             FilterPriority = True
                          End If
                     Next i
                     If Not r Is Nothing And FilterPriority = True Then r.EntireRow.Hidden = True
                     
              
       
'         Call UnFilterPriority(Shts)
    End With
    
End Function
Function FilterPriorityDyn(Shts As Worksheet, Filt As String) As Boolean
    Dim r As Range
    Dim lastRow As Long
    Dim colonne As Integer
    Dim i As Long
    Dim x As Integer
    Dim TabCrit As String
    Dim VCrit(1) As String
    Dim p As Integer
   x = 0
  
    VCrit(1) = ""
    FilterPriorityDyn = False
    With Shts
       
             
             
                     colonne = 72
                     lastRow = .Cells(.Rows.Count, colonne).End(xlUp).row
                     If lastRow > 7 Then
                        TabCrit = ";;" & Join(WorksheetFunction.Transpose(.Range(.Cells(7, colonne), .Cells(lastRow, colonne))), ";;") & ";;"
                    Else
                         TabCrit = ";;" & .Cells(7, colonne) & ";;"
                    End If
                     
                             For p = 1 To 4
                                If Filt = "Hight" Then
                                    If InStr(1, TabCrit, ";;" & CStr(p) & ";;") <> 0 And x < 1 Then
                                              VCrit(1) = CStr(p)
                                              x = x + 1
                                    End If
                                ElseIf Filt = "Low" Then
                                     If InStr(1, TabCrit, ";;" & CStr(p) & ";;") <> 0 And x < 2 Then
                                          
                                          If x = 1 Then
                                             VCrit(1) = CStr(p)
                                         End If
                                          x = x + 1
                                    End If
                                End If
                            Next p
                    
                             
                    
                   If VCrit(1) = "" Then Exit Function
                   
                      For i = 7 To lastRow
                          If (.Cells(i, colonne)) <> VCrit(1) Then
                              If r Is Nothing Then Set r = .Cells(i, colonne) Else Set r = Union(r, .Cells(i, colonne))
                          Else
                             FilterPriorityDyn = True
                          End If
                     Next i
                     If Not r Is Nothing And FilterPriorityDyn = True Then r.EntireRow.Hidden = True
                     
               
       
'         Call UnFilterPriorityDyn(Shts)
    End With
    
End Function

Function UnFilterPriority(Shts As Worksheet, TotalRow As Long)
    With Shts
        .Rows("7:" & TotalRow).EntireRow.Hidden = False
    End With
End Function

Function getSDV()
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
                If v = "" Then v = c.Value Else v = v & "#" & c.Value
            End If
            j = j + 1
          Wend
    End With

    If InStr(1, v, "#") = 0 Then
        ReDim SDVListe(0)
        SDVListe(0) = v
    Else
        SDVListe = Split(v, "#")
    End If
   
End Function


Function CopyRating2(objword As Object, objDoc As Object)
   Dim plage As Range
   Dim k As Integer
   Dim i As Integer
   Dim n As Long
   Dim lastRD As String
   Dim sh As Worksheet
   Dim colD As Integer, colC As Integer, totC As Integer
   
    Application.DisplayAlerts = False
    sheets("RATING").Copy before:=sheets(ThisWorkbook.sheets.Count)
    Application.DisplayAlerts = True
    Set sh = sheets(ThisWorkbook.sheets.Count - 1)
    With sh
'        .Shapes.Range(Array("Image 11")).Delete
'        .Rows("1:14").EntireRow.Hidden = True
'        totC = .Cells(10, .Columns.Count).End(xlToLeft).Column
'        Set Plage = .Range("B14:" & .Cells(18, totC).Address)
'        Call TakePic(objWord, objDoc, Plage, , "SysDyn")

        colD = .Rows("21:22").Find(What:="Drivability Lowest Events", LookAt:=xlWhole).Column
        colC = .Range("colP1").Column
        .Columns(colLettre(colC) & ":" & colLettre(colD)).EntireColumn.Hidden = True
         colD = .Rows("21:22").Find(What:="Responsiveness Lowest Events", LookAt:=xlWhole).Column + 1
         Set plage = .Range("B20:" & .Cells(100, colD).Address)
        Call TakePic(objword, objDoc, plage, , "SysDynGR")
    End With
    Application.DisplayAlerts = False
    sh.Delete
    Application.DisplayAlerts = True
    


'       ObjDoc.InlineShapes(2).Width = 500
End Function



Function MKDynamic(m As Boolean)

             If m = True Then
                  ThisWorkbook.sheets("RATING").Columns("E:L").EntireColumn.Hidden = True
             Else
                 ThisWorkbook.sheets("RATING").Columns("E:L").EntireColumn.Hidden = False
             End If
            ThisWorkbook.sheets("RATING").Shapes("Graphique 9").Left = 10
            ThisWorkbook.sheets("RATING").Shapes("Graphique 10").Left = 150
            ThisWorkbook.sheets("RATING").Shapes("Graphique 11").Left = 270
            ThisWorkbook.sheets("RATING").Shapes("Graphique 12").Left = 390


End Function




Function Remplissage(Fields As Variant, objword As Object, objDoc As Object)
        Dim i As Integer
        Dim j As Integer
        Dim tabF() As String
        Dim tabB() As String
        
        For i = 1 To UBound(Fields)
            tabF = Split(Fields(i), "#")
            If InStr(1, tabF(1), ";") <> 0 Then
                 tabB = Split(tabF(1), ";")
                 For j = 0 To UBound(tabB)
                     objDoc.Bookmarks(tabB(j)).Select
                     objword.Selection.typeText text:=tabF(0)
                 Next j
            Else
                   objDoc.Bookmarks(tabF(1)).Select
                   If tabF(1) = "Signet7" Then
                        objword.Selection.typeText text:=CStr(Date)
                        objDoc.Bookmarks("Places").Select
                        objword.Selection.typeText text:=tabF(0) & " " & CStr(Date)
                   Else
                        objword.Selection.typeText text:=tabF(0)
                   End If
            End If
        Next i
        
         
End Function

Function RedLowPoint(objword As Object, objDoc As Object) As String
        Dim vG As String
        Dim TaBvS() As String
        Dim TaBS() As String
        Dim TabSh() As String
        Dim TabvH() As String
        Dim TabE() As String
        Dim i As Long
        Dim j As Long
        Dim n As Long
        Dim p As Integer
        RedLowPoint = ""
        objDoc.Bookmarks("CommentLow").Select
        objword.Selection.TypeParagraph
        objword.Selection.MoveUp Unit:=5, Count:=1
           vG = getLower
               If vG <> "" Then
                   If InStr(1, vG, "||") = 0 Then
                        ReDim TabSh(0)
                        TabSh(0) = vG
                   Else
                        TabSh = Split(vG, "||")
                   End If
                   n = 0
                   For j = 0 To UBound(TabSh)
                        TaBvS = Split(TabSh(j), "#")
                        If InStr(1, TaBvS(1), ";") = 0 Then
                             ReDim TaBS(0)
                             TaBS(0) = TaBvS(1)
                        Else
                             TaBS = Split(TaBvS(1), ";")
                        End If
                        
                        'Ajout SDV
                        objword.Selection.Font.Bold = True
                        objword.Selection.typeText text:=TaBvS(0)
                        objword.Selection.TypeParagraph
                        objword.Selection.Font.Bold = False
                        'Creation Tableau
                        TabvH = Split(TaBS(0), " > ")
                        ReDim TabE(UBound(TabvH))
                        For p = 0 To UBound(TabvH)
                                TabE(p) = Split(TabvH(p), " : ")(0)
                        Next p
                        objDoc.tables.Add Range:=objword.Selection.Range, numRows:=UBound(TaBS) + 2, NumColumns:= _
                                                                UBound(TabE) + 1, DefaultTableBehavior:=1, AutoFitBehavior:=0
                        'Remplissage Entete
                        For p = 0 To UBound(TabE)
                                objword.Selection.Font.Bold = False
                                objword.Selection.typeText text:=TabE(p)
                                objword.Selection.Shading.BackgroundPatternColor = RGB(192, 192, 192)
                                objword.Selection.MoveRight Unit:=12
                        Next p
                        'Remplissage corps
                        For i = 0 To UBound(TaBS)
                                 TabvH = Split(TaBS(i), " > ")
                                 For p = 0 To UBound(TabvH)
                                        objword.Selection.typeText text:=replace(TabvH(p), TabE(p) & " : ", "")
                                        If p = 1 Then
                                             If CStr(replace(TabvH(p), TabE(p) & " : ", "")) = "1" Then
                                                objword.Selection.Shading.BackgroundPatternColor = RGB(192, 0, 0)
                                             ElseIf CStr(replace(TabvH(p), TabE(p) & " : ", "")) = "2" Then
                                                objword.Selection.Shading.BackgroundPatternColor = RGB(118, 147, 60)
                                             ElseIf CStr(replace(TabvH(p), TabE(p) & " : ", "")) = "3" Then
                                                objword.Selection.Shading.BackgroundPatternColor = RGB(21, 73, 125)
                                             End If
                                        End If
                                        If i < UBound(TaBS) Or p < UBound(TabvH) Then
                                            objword.Selection.MoveRight Unit:=12
                                        Else
'                                            objWord.Selection.MoveDown unit:=5, Count:=1
                                            
                                            objDoc.Bookmarks("Depart_" & n).Select
                                            objword.Selection.TypeParagraph
                                            objword.Selection.TypeParagraph
                                            n = n + 1
                                            objDoc.Bookmarks.Add Range:=objword.Selection, Name:="Depart_" & n
                                            objword.Selection.MoveUp Unit:=5, Count:=1
                                        End If
                                Next p
                        Next i
                   Next j
               End If
       
        
End Function

Function RedLowPointDyn(objword As Object, objDoc As Object) As String
        Dim vG As String
        Dim TaBvS() As String
        Dim TaBS() As String
        Dim TabSh() As String
        Dim TabvH() As String
        Dim TabE() As String
        Dim i As Long
        Dim j As Long
        Dim n As Long
        Dim p As Integer
        RedLowPointDyn = ""
        objDoc.Bookmarks("StartTableOne").Select
        objword.Selection.TypeParagraph
        objword.Selection.MoveUp Unit:=5, Count:=1
           vG = getLowerDyn
               If vG <> "" Then
                   If InStr(1, vG, "||") = 0 Then
                        ReDim TabSh(0)
                        TabSh(0) = vG
                   Else
                        TabSh = Split(vG, "||")
                   End If
                   n = 0
                   For j = 0 To UBound(TabSh)
                        TaBvS = Split(TabSh(j), "#")
                        If InStr(1, TaBvS(1), ";") = 0 Then
                             ReDim TaBS(0)
                             TaBS(0) = TaBvS(1)
                        Else
                             TaBS = Split(TaBvS(1), ";")
                        End If
                        
                        'Ajout SDV
                        objword.Selection.Font.Bold = True
                        objword.Selection.typeText text:=TaBvS(0)
                        objword.Selection.TypeParagraph
                        objword.Selection.Font.Bold = False
                        'Creation Tableau
                        TabvH = Split(TaBS(0), " > ")
                        ReDim TabE(UBound(TabvH))
                        For p = 0 To UBound(TabvH)
                                TabE(p) = Split(TabvH(p), " : ")(0)
                        Next p
                        objDoc.tables.Add Range:=objword.Selection.Range, numRows:=UBound(TaBS) + 2, NumColumns:= _
                                                                UBound(TabE) + 1, DefaultTableBehavior:=1, AutoFitBehavior:=0
                        'Remplissage Entete
                        For p = 0 To UBound(TabE)
                                objword.Selection.Font.Bold = False
                                objword.Selection.typeText text:=TabE(p)
                                objword.Selection.Shading.BackgroundPatternColor = RGB(192, 192, 192)
                                objword.Selection.MoveRight Unit:=12
                        Next p
                        'Remplissage corps
                        For i = 0 To UBound(TaBS)
                                 TabvH = Split(TaBS(i), " > ")
                                 For p = 0 To UBound(TabvH)
                                        objword.Selection.typeText text:=replace(TabvH(p), TabE(p) & " : ", "")
                                        If p = 1 Then
                                             If CStr(replace(TabvH(p), TabE(p) & " : ", "")) = "1" Then
                                                objword.Selection.Shading.BackgroundPatternColor = RGB(192, 0, 0)
                                             ElseIf CStr(replace(TabvH(p), TabE(p) & " : ", "")) = "2" Then
                                                objword.Selection.Shading.BackgroundPatternColor = RGB(118, 147, 60)
                                             ElseIf CStr(replace(TabvH(p), TabE(p) & " : ", "")) = "3" Then
                                                objword.Selection.Shading.BackgroundPatternColor = RGB(21, 73, 125)
                                             End If
                                        End If
                                        If i < UBound(TaBS) Or p < UBound(TabvH) Then
                                            objword.Selection.MoveRight Unit:=12
                                        Else
'                                            objWord.Selection.MoveDown unit:=5, Count:=1
                                            
                                            objDoc.Bookmarks("Dep_" & n).Select
                                            objword.Selection.TypeParagraph
                                            objword.Selection.TypeParagraph
                                            n = n + 1
                                            objDoc.Bookmarks.Add Range:=objword.Selection, Name:="Dep_" & n
                                            objword.Selection.MoveUp Unit:=5, Count:=1
                                        End If
                                Next p
                        Next i
                   Next j
               End If
       
        
End Function

Function getRedLP(onglet As String) As String
         Dim v
         Dim i As Integer
         getRedLP = ""
            v = ThisWorkbook.sheets(onglet).UsedRange.Value
'            If LCase(onglet) = "tip in at deceleration" Then
'                    MsgBox 2
'            End If
            For i = 3 To UBound(v, 1)
            
                If InStr(1, v(i, 15), "RED") <> 0 And (CStr(v(i, 14)) = "1" Or CStr(v(i, 14)) = "2") Then
                        If getRedLP = "" Then
                            getRedLP = onglet & "#" & "P(" & v(i, 14) & ")"
                        Else
                           getRedLP = getRedLP & ";" & "P(" & v(i, 14) & ")"
                        End If
                End If
            Next i
    
    Erase v

End Function


Function getColumnPoint(onglet As String)
  Dim v
  Dim i As Long
  Dim j As Long
  Dim p As Long
  
  getColumnPoint = ""
   v = ThisWorkbook.sheets("structure").UsedRange.Value
        For i = 2 To UBound(v, 1)
            If UCase(v(i, 2)) = UCase(onglet) Then
                j = i + 1
                p = j
                Do While (Len(v(p, 2)) = 0 And j <= UBound(v, 1) And Application.CountA(ThisWorkbook.sheets("structure").Rows(j).EntireRow) > 0)
                    If ThisWorkbook.sheets("structure").Cells(j, 4).Interior.color = RGB(255, 255, 0) Then
                        If getColumnPoint = "" Then getColumnPoint = v(j, 4) Else getColumnPoint = getColumnPoint & ";" & v(j, 4)
                    End If
                    j = j + 1
                    If j <= UBound(v, 1) Then p = j
                Loop
                Exit Function
            End If
        Next i
    Erase v
    
End Function




Function getRangeTable(SDV As String, tabl As String) As Range
    Dim x As Integer, y As Integer, i As Integer, tot As Integer
    Dim r As Range
    Dim col As Long
    
    With ThisWorkbook.Worksheets(SDV)
            Set r = .ListObjects(tabl).Range.Cells(1, 1)
            x = .ListObjects(tabl).Range.Columns.Count
            y = Round((8 - x) / 2, 0)
            tot = (y * 2)
            
            If y - (r.Column - 1) > 0 Then
                i = tot - (r.Column - 1)
            Else
                i = y
            End If
            Set r = .ListObjects(tabl).Range.Cells
            If SDV = "Auto stop" And ThisWorkbook.Worksheets("PARAMETRES GRAPH").Range("D731").Value = "Activè" Then
                col = 34
            Else
                col = ((r.row + r.Rows.Count) + (38 - r.Rows.Count)) - 1
            End If
            If i = tot Then
                Set getRangeTable = .Range(.Cells(r.row, r.Column), .Cells(col, r.Columns.Count + i))
            Else
                y = ((r.Column + r.Columns.Count) - 1) + i
                Set getRangeTable = .Range(.Cells(r.row, r.Column - (tot - i)), .Cells(col, y))
            End If
    End With
   
End Function

Function resetNull(objDoc, objword, x As String)
    Dim p As Integer
    
    objDoc.Bookmarks("S" & x).Select
    objword.Selection.MoveDown 5, -2
    objword.Selection.HomeKey 5
    For p = 1 To 3
         objword.Selection.MoveDown 5, 1, 1
    Next p
    objword.Selection.TypeBackspace
End Function

Function colLettre(NB As Integer)
        Dim tp As String
        While NB > 0
            tp = Chr(((NB - 1) Mod 26) + 65) & tp
            NB = (NB - 1) \ 26
        Wend
        colLettre = tp
End Function












