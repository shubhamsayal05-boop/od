Attribute VB_Name = "Report_PPT"
Option Explicit
Private chemin As String
Private getLower As String
Private getLowerDyn As String
Private SDVListe() As String
Private sommaireTexte  As String
Private numSlide As Integer
Private BlnAddSlideDriv As Boolean
Private BlnAddSlideDynam As Boolean
Public numRows As Integer
Public DivDriv As Boolean
Public drivHeight As Double
Public addDynam As Boolean


Sub CopyVersionPPT(objppt As Object, objPresentation As Object, slideIndex As Integer)
    
    Dim plage As Range
    Dim objslide As Object
    Dim success As Boolean
    Dim ca, cb As Integer
    Dim shp As Object
    
    
 
    
    Set plage = F_Docversions.Range("B6:D11")

    'DoEvents
    'sleep 500
    success = False
    Do While Not success
        On Error Resume Next
 
        
        plage.CopyPicture Appearance:=xlScreen, Format:=xlPicture

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop
    
    
    'DoEvents
    'sleep 500
    
          
    
    Range("B2").Select
 
    Set objslide = objPresentation.Slides(8)
  
    success = False

cb = objslide.Shapes.Count

Do While Not success
    objslide.Shapes.PasteSpecial DataType:=2
    ca = objslide.Shapes.Count
    success = ca > cb
Loop


Set shp = objslide.Shapes(objslide.Shapes.Count)

'DoEvents
With shp
    .Left = 10
    .Top = 370
    .Width = 700
    .Height = 130
End With
   
    
    
End Sub

Function CopyTablePPT(objppt As Object, Optional objPresentation As Object, Optional r As Range, Optional s As shape, Optional x As String)
    
    On Error Resume Next
    Dim i As Integer
    
    For i = 1 To 10
        ERR.Clear
         If Not r Is Nothing Then Call COPYp(r) Else Call COPYp(, s)
        If x <> "" Then
            If Left(x, 1) = "A" Or UCase(Left(x, 3)) = "SYS" Then
                'objPresentation.Bookmarks(x).Select
            Else
                'objPresentation.Bookmarks("S" & x).Select
            End If
        End If
         With objPresentation.Slides(8).Shapes
            .PasteSpecial ''PasteExcelTable ''''''' ' ppPasteBitmap
            
            With .Item(.Count)
                    .Width = 700
                    .Height = 130
                    .Left = 10
                    .Top = 370
                    .LockAspectRatio = msoFalse
                    .ZOrder msoSendToBack
            End With
        End With

        If ERR.Number = 0 Then Exit Function Else Application.Wait Now + TimeValue("0:00:02")
    Next i
    
    
End Function
Function CopyRating_PPT(objppt As Object, objPresentation As Object, ByRef slideIndex As Integer, ByRef visibleCount As Integer, ByRef foundCol As Integer, ByRef targetRow As Integer)
 Dim plage As Range
 Dim plage1 As Range
 Dim plage2 As Range
 Dim plage3 As Range
 Dim plageTest1 As Range
 Dim plageTest2 As Range
 Dim objslide As Object
 Dim slideHeight As Single
 Dim plageHeight As Single
 Dim shp As Variant
 Dim h As Double
 Dim p3Height As Double
 
 
 
 
 Dim objImageBox As PowerPoint.shape
 Dim chemin, NomImage As String
 Dim MyChart As Chart
 Dim ws As Worksheet
 Dim haut, large As Single
 Dim success As Boolean
 Dim allHidden1, allHidden2 As Boolean
 Dim cell As Range
 Dim ca, cb As Integer
 Dim lastCol As String
 Dim blnPlage As Boolean
 Dim rng1, rng2, rng3  As String
 Dim lastRow As Integer
 
 
 
 
  Set objslide = objPresentation.Slides(slideIndex)
  Call CopieSlide(objPresentation, slideIndex)
  
      With objslide.Shapes(1).TextFrame.TextRange
    .text = "Odriv scorecard – Normal mode"  ''"Odriv scorecard – predominant mode"
    '.Font.color.RGB = RGB(0, 0, 255)
    .Font.Bold = True
    .Font.Size = 15
    .Font.Name = "Encode Sans Expanded Light"
    
    End With
        
        
  lastRow = ThisWorkbook.Worksheets("RATING").Cells(ThisWorkbook.Worksheets("RATING").Rows.Count, "D").End(xlUp).row
  lastCol = ColNumToLetter(foundCol + 1)
  
  If visibleCount > 39 Then
    rng1 = "B20:" & lastCol & CStr(targetRow)
    rng2 = "B20:" & lastCol & "22"
    rng3 = "B" & CStr(targetRow + 1) & ":" & lastCol & CStr(lastRow)
    Set plage1 = ThisWorkbook.sheets("RATING").Range(rng1)
    Set plage2 = ThisWorkbook.sheets("RATING").Range(rng2)
    Set plage3 = ThisWorkbook.sheets("RATING").Range(rng3)
    
  Else
    rng1 = "B20:" & lastCol & CStr(lastRow)
    Set plage1 = ThisWorkbook.sheets("RATING").Range(rng1)
  End If
  
  
  
  
   
   
    
    
    
    
  If visibleCount <= 39 Then
  
  
         
         
         Set ws = ThisWorkbook.sheets("RATING")
    ws.Activate
    NomImage = ActiveSheet.Name
    DoEvents
    'sleep 500
    success = False
    Do While Not success
        On Error Resume Next
 
        
        plage1.CopyPicture Appearance:=xlScreen, Format:=xlPicture

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop
    
    DoEvents
    'sleep 500
    
    success = False

cb = objslide.Shapes.Count

Do While Not success
    objslide.Shapes.PasteSpecial DataType:=2
    ca = objslide.Shapes.Count
    success = ca > cb
Loop


Set shp = objslide.Shapes(objslide.Shapes.Count)

DoEvents
With shp
    .lockaspectration = msoTrue
    .Left = 80
    .Top = 65
    .Width = 720
End With
    
    

    
    Set shp = objslide.Shapes(objslide.Shapes.Count)
    h = objslide.Shapes(objslide.Shapes.Count).Height
    shp.LockAspectRatio = msoFalse
    
    If h >= 450 Then

     
     objslide.Shapes(objslide.Shapes.Count).Height = 450
     
    End If
    shp.Width = 780
  numSlide = 5
    
  
    
  Else
         
      
         Set ws = ThisWorkbook.sheets("RATING")
    ws.Activate
    NomImage = ActiveSheet.Name
    DoEvents
    'sleep 500
    success = False
    Do While Not success
        On Error Resume Next
 
        
        plage1.CopyPicture Appearance:=xlScreen, Format:=xlPicture

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop
    
    
    DoEvents
    'sleep 500
    success = False
    Do While Not success
        On Error Resume Next
 
        
        ActiveSheet.Paste: Selection.Name = NomImage

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop
    
    
    haut = ActiveSheet.Shapes(NomImage).Height
    large = ActiveSheet.Shapes(NomImage).Width
                
                
    chemin = ThisWorkbook.Path & "\Graph.png"
    With ActiveSheet

       Set MyChart = .ChartObjects.Add(0, 0, large, haut).Chart
       With MyChart
             .Parent.Activate
             .ChartArea.Format.line.Visible = msoFalse
             DoEvents
             .Paste
             .Export Filename:=chemin, filtername:="PNG"
             .Parent.Delete
       End With
    End With
    Set MyChart = Nothing
    ActiveSheet.Shapes(NomImage).Delete
    Range("B2").Select
    Set objslide = objPresentation.Slides(slideIndex)
    'Set objImageBox = objslide.Shapes.AddPicture(chemin, msoCTrue, msoCTrue, 10, 90, objslide.Master.Width - 50, -1)
    Set objImageBox = objslide.Shapes.AddPicture(chemin, msoCTrue, msoCTrue, 80, 65, 720, -1)
    
                
    Kill (chemin)
 
    Set shp = objslide.Shapes(objslide.Shapes.Count)
    h = objslide.Shapes(objslide.Shapes.Count).Height
    shp.LockAspectRatio = msoFalse
    If h >= 450 Then

     
     objslide.Shapes(objslide.Shapes.Count).Height = 450
    End If
    shp.Width = 780
    
  
        Call CopieSlide(objPresentation, slideIndex + 1)
         
         Set objslide = objPresentation.Slides(slideIndex + 1)
         With objslide.Shapes(1).TextFrame.TextRange
            .text = "Odriv scorecard – Normal mode" ''''''"Odriv scorecard – predominant mode"
            '.Font.color.RGB = RGB(0, 0, 255)
            .Font.Bold = True
            .Font.Size = 15
            .Font.Name = "Encode Sans Expanded Light"
        End With
         
         
         
         
         Set ws = ThisWorkbook.sheets("RATING")
    ws.Activate
    NomImage = ActiveSheet.Name
    DoEvents
    'sleep 500
    success = False
    Do While Not success
        On Error Resume Next
 
        
        plage2.CopyPicture Appearance:=xlScreen, Format:=xlPicture

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop
    
    DoEvents
    'sleep 500
    success = False
    Do While Not success
        On Error Resume Next
 
        
        ActiveSheet.Paste: Selection.Name = NomImage

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop
    
    
    haut = ActiveSheet.Shapes(NomImage).Height
    large = ActiveSheet.Shapes(NomImage).Width
                
                
    chemin = ThisWorkbook.Path & "\Graph.png"
    With ActiveSheet

       Set MyChart = .ChartObjects.Add(0, 0, large, haut).Chart
       With MyChart
             .Parent.Activate
             .ChartArea.Format.line.Visible = msoFalse
             DoEvents
             .Paste
             .Export Filename:=chemin, filtername:="PNG"
             .Parent.Delete
       End With
    End With
    Set MyChart = Nothing
    ActiveSheet.Shapes(NomImage).Delete
    Range("B2").Select
    Set objslide = objPresentation.Slides(slideIndex + 1)
    Set objImageBox = objslide.Shapes.AddPicture(chemin, msoCTrue, msoCTrue, 80, 65, 780, 60)
    Set shp = objslide.Shapes(objslide.Shapes.Count)
    p3Height = 65 + shp.Height
                
    Kill (chemin)

         
         Set ws = ThisWorkbook.sheets("RATING")
    ws.Activate
    NomImage = ActiveSheet.Name
    DoEvents
    'sleep 500
    success = False
    Do While Not success
        On Error Resume Next
 
        
        plage3.CopyPicture Appearance:=xlScreen, Format:=xlPicture

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop
    
    
    DoEvents
    'sleep 500
    success = False
    Do While Not success
        On Error Resume Next
 
        
        ActiveSheet.Paste: Selection.Name = NomImage

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop
    
    
    haut = ActiveSheet.Shapes(NomImage).Height
    large = ActiveSheet.Shapes(NomImage).Width
                
                
    chemin = ThisWorkbook.Path & "\Graph.png"
    With ActiveSheet

       Set MyChart = .ChartObjects.Add(0, 0, large, haut).Chart
       With MyChart
             .Parent.Activate
             .ChartArea.Format.line.Visible = msoFalse
             DoEvents
             .Paste
             .Export Filename:=chemin, filtername:="PNG"
             .Parent.Delete
       End With
    End With
    Set MyChart = Nothing
    ActiveSheet.Shapes(NomImage).Delete
    Range("B2").Select
    Set objslide = objPresentation.Slides(slideIndex + 1)
    Set objImageBox = objslide.Shapes.AddPicture(chemin, msoCTrue, msoCTrue, 80, p3Height, 720, -1)
    'Set objImageBox = objslide.Shapes.AddPicture(chemin, msoCTrue, msoCTrue, 10, 370, objslide.Master.Width - 50, -1)
    
                
    Kill (chemin)
    
    
    
    
    Set shp = objslide.Shapes(objslide.Shapes.Count)
    h = objslide.Shapes(objslide.Shapes.Count).Height
    
    shp.LockAspectRatio = msoFalse
    If h >= 390 Then

     
     objslide.Shapes(objslide.Shapes.Count).Height = 390
    End If
    shp.Width = 780
   
         
         numSlide = 6
    End If
    
End Function


Function REPORTC_PPT(fieldR As Variant)
    Dim objppt As Object
    Dim objPresentation As Object
    Dim FieldSrp As Variant
    Dim c As Object
    Dim nbpages As Integer
    Dim slideIndex As Integer
    Dim Values() As Integer
    Dim visibleCount As Integer
    Dim foundCol As Integer
    Dim targetRow As Integer
    
    
    
    
    On Error GoTo Ers
    Application.EnableEvents = False
    FieldSrp = fieldR
    Unload Preremplissage
    ProgressLoad
    ProgressTitle ("Creating Presentation")
    Call getSDV
    Call CreatePPT
    Set objppt = CreateObject("PowerPoint.Application")
    objppt.Visible = True
    Set objPresentation = objppt.Presentations.Open(chemin)
    Application.DisplayStatusBar = False
    getLower = ""
    getLowerDyn = ""
    ProgressTitle ("Copying Rating")
    slideIndex = 3
   
    
    Call CopyVersionPPT(objppt, objPresentation, 8)

    Call RemplissageTable(objppt, objPresentation, FieldSrp)

    Call CopyRating1_PPT(objppt, objPresentation, slideIndex, FieldSrp)

    Call Find39thVisibleRow(Values)
    
    visibleCount = Values(0)
    targetRow = Values(1)
    foundCol = FindColumnWithValue()
    
   
    
    Call CopyRating_PPT(objppt, objPresentation, slideIndex, visibleCount, foundCol, targetRow)
    

    Call CreateSingleTextboxLayout(objppt, objPresentation, FieldSrp)
    
    ProgressTitle ("MAJ Titres")
    
    
    
    slideIndex = 12
    Call InsertPic_PPT_Format4(objppt, objPresentation, slideIndex)
    
    
    Call Remplissage_Cartouche(objppt, objPresentation, FieldSrp)
    
    
    Call SupprimerDiapositivesVides(objPresentation)
    Call AddSlideGlossary(objPresentation, objppt)
    Call InsererSommaireDynamiqueDeuxNiveaux(objPresentation)
    
    Call RemplirPiedPage(objppt, objPresentation, FieldSrp)
    AppActivate Application.Caption
    
   
    MsgBox "Done", vbInformation, "ODRIV"
    
    ProgressExit
    objPresentation.Save
Ers:
    If ERR.Number <> 0 Then
        If Not objPresentation Is Nothing Then objPresentation.Close
        Application.EnableEvents = True
        Application.DisplayStatusBar = True
        MsgBox ERR.description, vbCritical, "ODRIV"
        ProgressExit
        Unload PleaseWait
    End If
End Function

Function Home_Button(objPresentation As Object)

    
    Dim sld As PowerPoint.slide
    Dim shp As PowerPoint.shape
    Dim foundImage As PowerPoint.shape
    Dim foundSlideNum As Long
    Dim slideIndex As Long
    Dim stopTitle As String
    Dim insertDone As Boolean
    Dim pasteSucceeded As Boolean
    
 
    
    For Each sld In objPresentation.Slides
        For Each shp In sld.Shapes
            If shp.Name = "Immagine 4" Then
                Set foundImage = shp
                foundSlideNum = sld.slideIndex
                Exit For
            End If
        Next shp
        If Not foundImage Is Nothing Then Exit For
    Next sld
 
    If foundImage Is Nothing Then
        MsgBox "Shape named 'Immagine 4' not found."
        Exit Function
    End If
 
    
    slideIndex = foundSlideNum + 2
    stopTitle = "Back UP slides for Top area of improvement"
    insertDone = False
 
    Do While slideIndex <= objPresentation.Slides.Count And Not insertDone
        Set sld = objPresentation.Slides(slideIndex)
 
        On Error Resume Next
        If sld.Shapes.hasTitle Then
            If Trim(sld.Shapes.Title.TextFrame.TextRange.text) = stopTitle Then
                insertDone = True
                Exit Do
            End If
        End If
        On Error GoTo 0
 
        
        foundImage.Copy
        
        pasteSucceeded = False
        Do While Not pasteSucceeded
        On Error Resume Next
        sld.Shapes.Paste
        If ERR.Number = 0 Then
            pasteSucceeded = True
        Else
            ERR.Clear
            DoEvents
        End If
        On Error GoTo 0
        Loop
        With sld.Shapes(sld.Shapes.Count)
            .Left = foundImage.Left
            .Top = 0.4
        End With
 
        slideIndex = slideIndex + 1
    Loop

End Function

Sub SupprimerDiapositivesVides(objPresentation As Object)
    Dim slide As Object
    Dim i As Integer
    Dim shape As Object
    Dim isEmpty As Boolean
    Dim text As String
    
     
    
    ' Parcours des diapositives en sens inverse (afin de ne pas affecter l'indice lors de la suppression)
    For i = objPresentation.Slides.Count To 1 Step -1
''        If i = 11 Then Stop
        Set slide = objPresentation.Slides(i)
        ' Supposer que la diapositive est vide
        isEmpty = True
        ' Vérifier chaque forme sur la diapositive
        For Each shape In slide.Shapes
            
            
            If shape.Name <> "Immagine 4" Then
            
            
            If shape.HasTextFrame Then
                If shape.TextFrame.HasText Then
                    text = shape.TextFrame.TextRange.text
                    If text <> "" And text <> "YYYY/MM/DD" And text <> "Communication Department" Then
                        isEmpty = False
                        Exit For
                    End If
                End If
            End If
            
            ' Si la forme a du texte ou si c'est une forme non vide, on considère la diapositive comme non vide
            If shape.Type <> msoPlaceholder And shape.HasTextFrame Then
                If shape.TextFrame.HasText Then
                    isEmpty = False
                    Exit For ' Quitte dès qu'on trouve un texte
                End If
            End If
            ' Si la forme est un objet graphique (image, tableau, etc.), la diapositive n'est pas vide
            If shape.Type = msoPicture Or shape.Type = msoTable Or shape.Type = msoAutoShape Then
                isEmpty = False
                Exit For
            End If
            
            End If
            
        Next shape
        
        
        ' Si la diapositive est vide, la supprimer
        If isEmpty Then
            slide.Delete
        End If
    Next i
End Sub

Function DeleteEmptySlides(objppt As Object, objPres As Object, slideIndex As Integer)
    Dim objslide As Object ' Individual Slide
    Dim shape As Object
    Dim hasContent As Boolean

   
    ' Loop through the slides in reverse to prevent indexing issues
      ' Loop through the slides in reverse to prevent indexing issues
    For slideIndex = objPres.Slides.Count To 1 Step -1
        Set objslide = objPres.Slides(slideIndex)
        
        hasContent = False ' Reset content flag for each slide
        
        ' Check each shape on the slide
        For Each shape In objslide.Shapes
            ' Check if the shape contains text
            If shape.HasTextFrame Then
                If shape.TextFrame.HasText Then
                    hasContent = True
                    Exit For
                End If
            ' Check if the shape is a table
            ElseIf shape.Type = msoTable Then
                hasContent = True
                Exit For
            ' Check if the shape is a chart
            ElseIf shape.Type = msoChart Then
                hasContent = True
                Exit For
            ' Check if the shape is SmartArt
            ElseIf shape.Type = msoSmartArt Then
                hasContent = True
                Exit For
            ' Check if the shape is a picture
            ElseIf shape.Type = msoPicture Then
                hasContent = True
                Exit For
            End If
        Next shape
        
        ' Delete slide if no content was found
        If Not hasContent Then
            objslide.Delete
        End If
    Next slideIndex
End Function

Sub getSDV()
Dim v As String
Dim c As Range
Dim i As Long
Dim j As Long

i = getLastRowRating
v = ""

    With F_Rating
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
   
End Sub

Function FindColumnWithValue() As Integer

   

    Dim r As Long

    Dim c As Long

    Dim lastCol As Long

    Dim foundCol As Integer
 
    

    foundCol = 0
 
   

    lastCol = WorksheetFunction.Max(F_Rating.Cells(21, F_Rating.Columns.Count).End(xlToLeft).Column, F_Rating.Cells(22, F_Rating.Columns.Count).End(xlToLeft).Column)
 
    

    For r = 21 To 22

        For c = 1 To lastCol
 
            Dim cell As Range

            Set cell = F_Rating.Cells(r, c)
 
            

            If cell.MergeCells Then

                If cell.Address = cell.MergeArea.Cells(1, 1).Address Then

                    If Trim(cell.Value) = "Responsiveness Lowest Events" Then

                        foundCol = c

                        Exit For

                    End If

                End If

            Else

                If Trim(cell.Value) = "Responsiveness Lowest Events" Then

                    foundCol = c

                    Exit For

                End If

            End If
 
        Next c
 
        If foundCol <> 0 Then Exit For

    Next r
 
    FindColumnWithValue = foundCol

End Function

 

 

Sub Find39thVisibleRow(Values() As Integer)
   
    Dim lastRow As Long
    Dim visibleCount As Integer
    Dim currentRow As Long
    Dim targetRow As Integer
    
    
    lastRow = F_Rating.Cells(F_Rating.Rows.Count, "D").End(xlUp).row
    visibleCount = 0
    targetRow = 0
    
    For currentRow = 20 To lastRow
        If Not F_Rating.Rows(currentRow).Hidden Then
            visibleCount = visibleCount + 1
            If visibleCount = 39 Then
                targetRow = currentRow
                
            End If
        End If
    Next currentRow
    
    
    
    
    ReDim Values(0 To 1)
    Values(0) = visibleCount
    Values(1) = targetRow
    
End Sub

Sub CreatePPT()
    
    Dim WDObj As Object
    Dim pptApp As Object
    Dim pptDoc As Object
    
    Application.ScreenUpdating = False
    Set WDObj = F_DNT.OLEObjects("Object 31") '''sheets("DNT").OLEObjects("REPORT")
    WDObj.Activate
    'WDObj.Object.Application.Visible = True
    Set pptApp = WDObj.Object.Application
    Set pptDoc = pptApp.ActivePresentation
    WDObj.Object.Application.Visible = True
    chemin = ThisWorkbook.Path & "/Report_" & ThisWorkbook.Worksheets("HOME").Range("Project") & "_" & replace(Time, ":", "_") & ".pptx"
    pptDoc.SaveAs (chemin)
    pptDoc.Close
    F_Home.Select
    Application.ScreenUpdating = True
    Load PleaseWait
    
    PleaseWait.Show vbModeless
    DoEvents
    
End Sub
Function CopieSlide(objPres As Object, slideIndex As Integer) As Object

Dim objslide As Object
Dim newSlide As Object
Dim pasted As Boolean
    ViderPressePapiers
   ' Copy the blank slide
     
     pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(slideIndex).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
    
    ViderPressePapiers
    Set newSlide = objslide
    
    
End Function

Function CopyHome_PPT(objppt As Object, objPresentation As Object, ByRef slideIndex As Integer)
    Dim plage As Range
    Dim objslide As Object
    Dim newSlide As Object
    
    Set plage = ThisWorkbook.sheets("HOME").Range("B6:L24")
       ' Get the slide to be copied
    Call CopieSlide(objPresentation, slideIndex)
    'Set objSlide = objPresentation.Slides.Add(slideIndex, ppLayoutBlank)
    
    Call TakePic_PPT(objppt, objPresentation, plage, , "SysHome", slideIndex)
    
    
    slideIndex = slideIndex + 1
End Function

Function CopyVersion_PPT(objppt As Object, objPresentation As Object, ByRef slideIndex As Integer)
    Dim plage As Range
    Dim objslide As Object
    Set plage = ThisWorkbook.sheets("DocVersions").Range("B6:D11")
    'Set objSlide = objPresentation.Slides.Add(slideIndex, ppLayoutBlank)
    Call CopieSlide(objPresentation, slideIndex)
    Set objslide = objPresentation.Slides(slideIndex)
    
    Call CopyTable_PPT(objppt, objPresentation, plage, , "SysLink", slideIndex)
    slideIndex = slideIndex + 1
End Function

Function ColNumToLetter(colNum As Integer) As String
    Dim colLetter As String
    colLetter = ""
 
    Do While colNum > 0
        colNum = colNum - 1
        colLetter = Chr((colNum Mod 26) + 65) & colLetter
        colNum = colNum \ 26
    Loop
 
    ColNumToLetter = colLetter
End Function


'Function CopyRating_PPT(objPPT As Object, objPresentation As Object, ByRef slideIndex As Integer)
'    Dim plage As Range
'    Dim plage1 As Range
'    Dim plage2 As Range
'    Dim objSlide As Object
'    Dim slideHeight As Single
'    Dim plageHeight As Single
'    Dim textBoxHeight As Single
'    Dim availableHeight As Single
'
'    ' Set the height of the text box at the bottom of the slide
'    textBoxHeight = 50 ' Adjust as needed
'
'    ' Set the current slide and duplicate it
'    Set objSlide = objPresentation.Slides(slideIndex)
'    Call CopieSlide(objPresentation, slideIndex)
'
'    ' Set slide height and available height
'    slideHeight = objSlide.Parent.PageSetup.slideHeight
'    availableHeight = slideHeight - textBoxHeight - 20 ' Leave a 20-pixel buffer
'
'    ' Define the range and calculate its height
'    Set plage = ThisWorkbook.sheets("RATING").Range("B20:AP74")
'    plageHeight = plage.Height
'
'    ' Compare content height to available slide height
'    If plageHeight > availableHeight Then
'        ' Case 1: Content exceeds slide height, split into two adjusted ranges
'        Set plage1 = ThisWorkbook.sheets("RATING").Range("B20:AP50") ' Adjust range to fit above text box
'        Set plage2 = ThisWorkbook.sheets("RATING").Range("B51:AP74") ' Remainder for second slide
'
'        ' Insert first range on current slide
'        Call TakePic_RatingPPT(objPPT, objPresentation, plage1, , , slideIndex)
'
'        ' Add a new slide for the second range
'        Call CopieSlide(objPresentation, slideIndex + 1)
'        Set objSlide = objPresentation.Slides(slideIndex + 1)
'
'        ' Update title for the second slide
'        With objSlide.Shapes(1).TextFrame.TextRange
'            .text = "Odriv scorecard – Normal mode"
'            .Font.Bold = True
'            .Font.Size = 15
'            .Font.Name = "Encode Sans Expanded Light"
'        End With
'
'        ' Insert second range on the new slide
'        Call TakePic_PPT(objPPT, objPresentation, plage2, , , slideIndex + 1)
'    Else
'        ' Case 2: Content fits within available slide height
'        If plageHeight > availableHeight Then
'            ' Split into two ranges to avoid conflict with the text box
'            Set plage1 = ThisWorkbook.sheets("RATING").Range("B20:AP50")
'            Set plage2 = ThisWorkbook.sheets("RATING").Range("B51:AP74")
'
'            ' Insert first range
'            Call TakePic_PPT(objPPT, objPresentation, plage1, , , slideIndex)
'
'            ' Add a new slide for the second range
'            Call CopieSlide(objPresentation, slideIndex + 1)
'            Set objSlide = objPresentation.Slides(slideIndex + 1)
'
'            ' Update title for the second slide
'            With objSlide.Shapes(1).TextFrame.TextRange
'                .text = "Odriv scorecard – Normal mode"
'                .Font.Bold = True
'                .Font.Size = 15
'                .Font.Name = "Encode Sans Expanded Light"
'            End With
'
'            ' Insert second range
'            Call TakePic_PPT(objPPT, objPresentation, plage2, , , slideIndex + 1)
'        Else
'            ' Single range fits without conflict
'            Call TakePic_PPT(objPPT, objPresentation, plage, , , slideIndex)
'        End If
'    End If
'End Function



Function TakePic_PPTd(objppt As Object, Optional objPresentation As Object, Optional r As Range, Optional s As shape, Optional x As String, Optional slideIndex As Integer)
    On Error Resume Next
    Dim i As Integer
    Dim slideWidth As Single, slideHeight As Single
    Dim leftOffset As Single, topOffset As Single
    Dim textBoxHeight As Single
    Dim originalWidth As Single, originalHeight As Single
    Dim aspectRatio As Single

    ' Set the offsets and text box height
    leftOffset = 30
    topOffset = 150
    textBoxHeight = 50 ' Adjust to match the text box height on the slide

    ' Ensure the slide exists
    If slideIndex > objPresentation.Slides.Count Then
        objPresentation.Slides.Add slideIndex, ppLayoutBlank
    End If

    ' Get slide dimensions
    slideWidth = objPresentation.PageSetup.slideWidth
    slideHeight = objPresentation.PageSetup.slideHeight

    ' Retry loop for pasting content
    For i = 1 To 10
        ERR.Clear

        ' Copy the content
        If Not r Is Nothing Then
            Call COPYp(r)
        Else
            Call COPYp(, s)
        End If

        ' Paste as bitmap
        With objPresentation.Slides(slideIndex).Shapes
            .PasteSpecial ppPasteBitmap

            With .Item(.Count)
                ' Calculate aspect ratio
                originalWidth = .Width
                originalHeight = .Height
                aspectRatio = IIf(originalHeight <> 0, originalWidth / originalHeight, 1)
                
                ' Adjust width and height to fit above the text box
                .LockAspectRatio = msoFalse
                .Width = 900
                .Left = leftOffset
                .Top = topOffset

                ' Ensure the shape does not overlap with the text box
                If (.Top + .Height) > (slideHeight - textBoxHeight) Then
                    .Height = slideHeight - textBoxHeight - topOffset
                    '.Width = 900
                End If

                .ZOrder msoSendToBack
            End With
        End With

        ' Exit if no errors
        If ERR.Number = 0 Then Exit Function

        ' Retry delay
        Application.Wait Now + TimeValue("0:00:02")
    Next i
End Function

Function TakePic_PPT(objppt As Object, Optional objPresentation As Object, Optional r As Range, Optional s As shape, Optional x As String, Optional slideIndex As Integer)
    
    Dim i As Integer
    Dim slideWidth As Single, slideHeight As Single
    Dim leftOffset As Single, topOffset As Single
    Dim textBoxHeight As Single
    Dim originalWidth As Single, originalHeight As Single
    Dim aspectRatio As Single
    Dim success As Boolean
    Dim ca, cb As Integer

    ' Set the offsets and text box height
    leftOffset = 30
    topOffset = 80
    textBoxHeight = 50 ' Adjust to match the text box height on the slide

    On Error Resume Next
    ' Ensure the slide exists
    If slideIndex > objPresentation.Slides.Count Then
        objPresentation.Slides.Add slideIndex, ppLayoutBlank
    End If

    ' Get slide dimensions
    slideWidth = objPresentation.PageSetup.slideWidth
    slideHeight = objPresentation.PageSetup.slideHeight

    ' Retry loop for pasting content
    
        

        ' Copy the content
        If Not r Is Nothing Then
            Call COPYp(r)
        Else
            Call COPYp(, s)
        End If

        
        
        
        success = False                                                                 ' Ajouté par Dhia
        cb = objPresentation.Slides(slideIndex).Shapes.Count

        Do While Not success
            objPresentation.Slides(slideIndex).Shapes.PasteSpecial DataType:=2          ' S'assurer que le paste se fait
            ca = objPresentation.Slides(slideIndex).Shapes.Count
            success = ca > cb
        Loop
        
        
        
        With objPresentation.Slides(slideIndex).Shapes
            

            With .Item(.Count)
                ' Calculate aspect ratio
                originalWidth = .Width
                originalHeight = .Height
                aspectRatio = IIf(originalHeight <> 0, originalWidth / originalHeight, 1)
                
                ' Adjust width and height to fit above the text box
                .Width = slideWidth - (2 * leftOffset)
                .Height = .Width / aspectRatio
                .Left = leftOffset
                .Top = topOffset

                ' Ensure the shape does not overlap with the text box
                If (.Top + .Height) > (slideHeight - textBoxHeight) Then
                    .Height = slideHeight - textBoxHeight - topOffset
                    .Width = .Height * aspectRatio
                End If

                .ZOrder msoSendToBack
            End With
        End With

        

        
End Function



Sub CopyRating1_PPT(objppt As Object, objPresentation As Object, ByRef slideIndex As Integer, Fields As Variant)
    Dim plage As Range
    Dim colD As Integer, colC As Integer, totC As Integer
    Dim objslide As Object
    Dim sourceCell As Range
   
    Set sourceCell = F_Rating.Range("F4")
    
    If F_Home.Range("Milestone") = 4 Then
        sourceCell.EntireColumn.Hidden = True
    End If
    
    sourceCell.EntireRow.Hidden = True

    With F_Rating
        totC = .Cells(10, .Columns.Count).End(xlToLeft).Column
        .Shapes.Range(Array("Image 11")).Visible = msoFalse
        Set plage = .Range("B1:" & .Cells(18, totC).Address)
        
        Call CopieSlide(objPresentation, slideIndex)
         Application.Wait (20)

        Set objslide = objPresentation.Slides(slideIndex)
      
       
        With objslide.Shapes(1).TextFrame.TextRange
        .text = "Odriv scorecard – predominant mode" ''''''.text = "Dynamisme Assessment PCAL DVEE SCORECARD"
        
        .Font.Bold = True
        .Font.Size = 16
        .Font.Name = "Encode Sans Expanded Light"
        End With
        
        Call TakePic_PPT(objppt, objPresentation, plage, , "SysDr", slideIndex)
        slideIndex = slideIndex + 1
        .Shapes.Range(Array("Image 11")).Visible = msoTrue

    End With
    
    sourceCell.EntireColumn.Hidden = False
    sourceCell.EntireRow.Hidden = False
    

End Sub

Function CopyRating2_PPT(objppt As Object, objPresentation As Object, ByRef slideIndex As Integer)
    Dim plage As Range
    Dim sh As Worksheet
    Dim colD As Integer, colC As Integer, totC As Integer
    Dim objslide As Object
    Application.DisplayAlerts = False
    sheets("RATING").Copy before:=sheets(ThisWorkbook.sheets.Count)
    Application.DisplayAlerts = True
    Set sh = sheets(ThisWorkbook.sheets.Count - 1)
    With sh
        colD = .Rows("21:22").Find(What:="Drivability Lowest Events", LookAt:=xlWhole).Column
        colC = .Range("colP1").Column
        .Columns(colLettre(colC) & ":" & colLettre(colD)).EntireColumn.Hidden = True
        colD = .Rows("21:22").Find(What:="Responsiveness Lowest Events", LookAt:=xlWhole).Column + 1
        Set plage = .Range("B20:" & .Cells(100, colD).Address)
        'Set objSlide = objPresentation.Slides.Add(slideIndex, ppLayoutBlank)
        Call CopieSlide(objPresentation, slideIndex)
        Call newSdvSlide_PPT(objslide, "Odriv scorecard – Normal mode", 1.1)
        'Call CopieSlide(objPresentation, slideIndex)
        Call TakePic_PPT(objppt, objPresentation, plage, , "SysDynGR", slideIndex)
        slideIndex = slideIndex + 1
    End With
    Application.DisplayAlerts = False
    sh.Delete
    Application.DisplayAlerts = True
End Function

'Function InsertPic_PPT2(ByRef objPPT As Object, objPres As Object, slideIndex As Integer)
'    Dim i As Integer
'    Dim j As Long
'    Dim H As Integer
'    Dim T As Integer
'    Dim o As String
'    Dim x As String
'    Dim getParamSdv As String
'    Dim v() As String
'    Dim c As Object
'    Dim nbpages As Integer
'    Dim objSlide As Object
'    Dim nbSlides As Integer
'    Dim slideLayout As PpSlideLayout
'    Dim pptShape As shape
'
'    'objDoc.Bookmarks("StartOne").Select
'
'    slideLayout = ppLayoutTitleOnly '''ppLayoutText
'
'    For j = 0 To UBound(SDVListe)
'        o = SDVListe(j)
'        ThisWorkbook.sheets(o).Cells.EntireColumn.Hidden = False
'        For T = 1 To 2
'            ProgressTitle ("Copie des données : " & o)
'            If T = 1 Or (T = 2 And checkCriteriaDyn(o) = True And checkCorrespondancePriorityDyn(o) = True) Then
'                'If j = 0 And T = 1 Then
'                 'Set objSlide = objPres.Slides.Add(slideIndex, slideLayout)
'                      objPres.Slides(8).Copy
'
'                    ' Paste the blank slide into the target presentation
'                     Set objSlide = objPres.Slides.Paste(slideIndex + 1)
'
'                     slideIndex = slideIndex + 1
'               ' Else
'                   'objword.Selection.InsertNewPage
'                  'Set objSlide = objPres.Slides(slideIndex)
'                  'Set objSlide = objPres.Slides.Add(slideIndex, slideLayout)
'
'                    'slideIndex = slideIndex + 1
'                'End If
'               If T = 1 Then Call newSdvSlide_PPT(objSlide, UCase(o) & " DRIVABILITY", "2." & j + 1 & "." & T) Else Call newSdvSlide_PPT(objSlide, UCase(o) & " DYNAMISM", "2." & j + 1 & "." & T)
'
'
'                For i = 1 To 5
'                    x = j & i & T
'
'                        If i = 1 Then
'                            Call insertPart_PPT(objSlide, i)
'                            Call CopySummary_PPT(objPPT, objSlide, o, T, slideIndex)
'
'                        ElseIf i = 2 Or i = 3 Then
'                                If T = 1 Then
'                                    getParamSdv = UpdateGraph.checkGraphEnable(o)
'                                Else
'                                    getParamSdv = UpdateGraphDyn.checkGraphEnable(o)
'                                End If
'                                If getParamSdv <> "" Then
'                                        If i = 2 Then
'                                                v = Split(getParamSdv, ";")
'                                                For H = 0 To UBound(v)
'                                                      If Split(v(H), ":")(0) = "Graphique_0" Or Split(v(H), ":")(0) = "Graphique_00" Then
'                                                           If ThisWorkbook.Worksheets(o).Shapes(Split(v(H), ":")(0)).Visible = True Then
'                                                                Call insertPart_PPT(objSlide, i)
'                                                                If T = 1 Then
'                                                                    Call CopyGraph0_PPT(objPPT, objSlide, o, T, slideIndex)
'                                                                Else
'                                                                    Call CopyGraph0_PPT(objPPT, objSlide, o, T, slideIndex)
'                                                                End If
'                                                           End If
'                                                           Exit For
'                                                      Else
'                                                            If UpdateGraph.checkObject(CStr(Split(v(H), ":")(0)), o) <> "" Then
'                                                                Call insertPart_PPT(objSlide, i)
'                                                                If T = 1 Then
'
'                                                                    Call CopyLeverAS_PPT2(objPPT, objPres, objSlide, o, UpdateGraph.checkObject(CStr(Split(v(H), ":")(0)), o), slideIndex, True)
'                                                                Else
'
'                                                                     Call CopyLeverAS_PPT2(objPPT, objPres, objSlide, o, UpdateGraphDyn.checkObject(CStr(Split(v(H), ":")(0)), o), slideIndex, True)
'                                                                End If
'
'                                                            End If
'                                                            Exit For
'                                                      End If
'                                                Next H
'                                         Else
'                                                If T = 1 Then
'                                                    getParamSdv = UpdateGraph.checkGraphEnable(o)
'                                                Else
'                                                    getParamSdv = UpdateGraphDyn.checkGraphEnable(o)
'                                                End If
'                                               v = Split(getParamSdv, ";")
'                                               If InStr(1, getParamSdv, "Graphique_1") <> 0 Or InStr(1, getParamSdv, "Graphique_11") Then
'                                                    For H = 0 To UBound(v)
'                                                         If Split(v(1), ":")(0) = "Graphique_1" Or InStr(1, getParamSdv, "Graphique_11") Then
'                                                              If ThisWorkbook.Worksheets(o).Shapes(Split(v(1), ":")(0)).Visible = True Then
'                                                                 Call insertPart_PPT(objSlide, i)
'                                                                 Call CopyGraph1_PPT(objPPT, objSlide, o, T)
'
'                                                              End If
'                                                              Exit For
'                                                          End If
'                                                     Next H
'                                                End If
'                                         End If
'                                 End If
'                        ElseIf i = 4 Then
'                            If T = 1 Then
'                                Call CopyPriorityPoints_PPT(objSlide, "Hight", ThisWorkbook.Worksheets(o), i)
'                            Else
'                                Call CopyPriorityPointsDyn_PPT(objSlide, "Hight", ThisWorkbook.Worksheets(o), i)
'                            End If
'                        ElseIf i = 5 Then
'                            If T = 1 Then
'                                 Call CopyPriorityPoints_PPT(objSlide, "Low", ThisWorkbook.Worksheets(o), i)
'                            Else
'                                 Call CopyPriorityPointsDyn_PPT(objSlide, "Low", ThisWorkbook.Worksheets(o), i)
'                            End If
'                        End If
'
'
'        '             If i <= 3 Then
'        '               objDoc.InlineShapes(LastPicNumber(objDoc)).Width = 520
'        '               objDoc.InlineShapes(LastPicNumber(objDoc)).LockAspectRatio = 0
'        '               objDoc.InlineShapes(LastPicNumber(objDoc)).Height = 283.5
'        '            End If
'        '              x = x + 1
'
'                Next i
'            End If
'      Next T
'    Next j
'
'    objPres.Slides(8).Delete
'
'    'Call SupprimerSlidesVides(objPres)
'
'    'Call VerifierPageAvecEnteteVide(objword, objDoc)
'
'End Function
Function TakePic_RatingPPT(objppt As Object, Optional objPresentation As Object, Optional r As Range, Optional s As shape, Optional x As String, Optional slideIndex As Integer)
    On Error Resume Next
    Dim i As Integer
    Dim slideWidth As Single, slideHeight As Single
    Dim leftOffset As Single, topOffset As Single
    Dim textBoxHeight As Single
    Dim originalWidth As Single, originalHeight As Single
    Dim aspectRatio As Single

    ' Set the offsets and text box height
    leftOffset = 30
    topOffset = 80
    textBoxHeight = 50 ' Adjust to match the text box height on the slide

    ' Ensure the slide exists
    If slideIndex > objPresentation.Slides.Count Then
        objPresentation.Slides.Add slideIndex, ppLayoutBlank
    End If

    ' Get slide dimensions
    slideWidth = objPresentation.PageSetup.slideWidth
    slideHeight = objPresentation.PageSetup.slideHeight

    ' Retry loop for pasting content
    For i = 1 To 10
        ERR.Clear

        ' Copy the content
        If Not r Is Nothing Then
            Call COPYp(r)
        Else
            Call COPYp(, s)
        End If

        ' Paste as bitmap
        With objPresentation.Slides(slideIndex).Shapes
            .PasteSpecial ppPasteBitmap

            With .Item(.Count)
                ' Calculate aspect ratio
                originalWidth = .Width
                originalHeight = .Height
                aspectRatio = IIf(originalHeight <> 0, originalWidth / originalHeight, 1)
                
                ' Adjust width and height to fit above the text box
                .Width = slideWidth - (2 * leftOffset)
                .Height = .Width / aspectRatio
                .Left = leftOffset
                .Top = topOffset

                ' Ensure the shape does not overlap with the text box
                If (.Top + .Height) > (slideHeight - textBoxHeight) Then
                    .Height = slideHeight - textBoxHeight - topOffset
                    .Width = .Height * aspectRatio
                End If

                .ZOrder msoSendToBack
            End With
        End With

        ' Exit if no errors
        If ERR.Number = 0 Then Exit Function

        ' Retry delay
        Application.Wait Now + TimeValue("0:00:02")
    Next i
End Function


'Function TakePic_PPT(objPPT As Object, Optional objPresentation As Object, Optional r As Range, Optional s As shape, Optional x As String, Optional slideIndex As Integer)
'    On Error Resume Next
'    Dim i As Integer
'    Dim slideWidth As Single, slideHeight As Single
'
'    If slideIndex > objPresentation.Slides.Count Then
'        objPresentation.Slides.Add slideIndex, ppLayoutBlank '''ppLayoutText
'    End If
'
'      ' Get the slide's width and height
'    slideWidth = objPresentation.PageSetup.slideWidth
'    slideHeight = objPresentation.PageSetup.slideHeight
'
'    For i = 1 To 10
'        ERR.Clear
'        If Not r Is Nothing Then Call COPYp(r) Else Call COPYp(, s)
'        'Add new Slide
'       ' Call NewSlide(objPresentation, slideIndex)
'
'        With objPresentation.Slides(slideIndex).Shapes
'            .PasteSpecial ppPasteBitmap ''DataType:=12  ''2 ' ppPasteEnhancedMetafile
'             ' Apply size and position if specified
'             If x = "SysDr" Then
'                With .Item(.Count)
''                    .Width = 10
''                    .Height = 160
''                    .Left = 30
''                    .Top = 80
''                    .LockAspectRatio = msoFalse
''                    .ZOrder msoSendToBack
'                    .Width = 0
'                    .Height = slideHeight - (2 * 80)
'                    .Left = 30
'                    .Top = 80
'                    .LockAspectRatio = msoFalse
'                    .ZOrder msoSendToBack
'                End With
'              ElseIf x = "SysDynT" Then
'               With .Item(.Count)
'                    .Width = 350
'                    .Height = 450
'                    .Left = 60
'                    .Top = 70
'                    .LockAspectRatio = msoFalse
'                    .ZOrder msoSendToBack
'                End With
'               ElseIf x = "SysDynGR" Then
'                  With .Item(.Count)
'                    .Width = 350
'                    .Height = 450
'                    .Left = 60
'                    .Top = 70
'                    .LockAspectRatio = msoFalse
'                    .ZOrder msoSendToBack
'                End With
'                Else
'                    With .Item(.Count)
''                        .Width = 190
''                        .Height = 290
''                        .Left = 30
''                        .Top = 100
''                        .LockAspectRatio = msoFalse
''                        .ZOrder msoSendToBack
'                        .Width = slideWidth - (2 * 30)
'                        .Height = slideHeight - (2 * 100)
'                        .Left = 30
'                        .Top = 100
'                        .LockAspectRatio = msoFalse
'                        .ZOrder msoSendToBack
'                    End With
'            End If
'        End With
'       'If ERR.Number <> 0 Then Stop
'        If ERR.Number = 0 Then Exit Function Else Application.Wait Now + TimeValue("0:00:02")
'    Next i
'    Exit Function
'
'
'End Function
Sub InsererSommaire_PPT(objPresentation As Object)
    Dim presentation As Object
    Dim sommaireDiapo As Object
    Dim diapo As Object
    Dim i As Integer
    Dim sommaireTexte As String
    Dim Titre As String
    Dim pptShape As Object
    Dim maxLignesParDiapo As Integer
    Dim ligneCount As Integer
    Dim colonneCount As Integer
    Dim lignes() As String
    Dim startLigne As Integer
    Dim ligneTexte As String
    Dim j As Integer
    Dim colonne1Texte As String
    Dim colonne2Texte As String
    Dim slideIndex As Integer
    
    ' Référence à la présentation active
    Set presentation = objPresentation

    ' Nombre maximum de lignes par diapositive
    maxLignesParDiapo = 20 ' Ajuster en fonction du nombre de lignes souhaité par diapositive

    ' Ajouter une diapositive pour le sommaire au début de la présentation
    'slideIndex = 1
    Do
       ' If slideIndex > presentation.Slides.Count Then
           ' presentation.Slides.Add slideIndex, ppLayoutText
       ' End If
        Set sommaireDiapo = presentation.Slides(8)
        
        ' Préparer le texte du sommaire
        sommaireTexte = ""
        For i = 1 To presentation.Slides.Count
            Set diapo = presentation.Slides(i)
            On Error Resume Next
            Titre = diapo.Shapes.Placeholders(1).TextFrame.TextRange.text
            On Error GoTo 0
            If Titre <> "" Then
                sommaireTexte = sommaireTexte & vbCrLf & i & ". " & Titre
            End If
        Next i

        ' Diviser le texte en lignes
        lignes = Split(sommaireTexte, vbCrLf)
        
        ' Initialiser les textes pour les colonnes
        colonne1Texte = ""
        colonne2Texte = ""
        
        ' Remplir les colonnes
        ligneCount = UBound(lignes) + 1
        For i = 0 To ligneCount - 1
            ligneTexte = lignes(i)
            If i < maxLignesParDiapo Then
                colonne1Texte = colonne1Texte & ligneTexte & vbCrLf
            Else
                colonne2Texte = colonne2Texte & ligneTexte & vbCrLf
            End If
        Next i
        
        ' Ajouter le texte aux colonnes
        Set pptShape = sommaireDiapo.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 20, 300, 400)
        pptShape.TextFrame.TextRange.text = colonne1Texte
        With pptShape.TextFrame.TextRange
            .Font.Name = "Arial"
            .Font.Size = 14
            '.Font.color = RGB(0, 0, 255) ' Bleu
        End With

        Set pptShape = sommaireDiapo.Shapes.AddTextbox(msoTextOrientationHorizontal, 320, 20, 300, 400)
        pptShape.TextFrame.TextRange.text = colonne2Texte
        With pptShape.TextFrame.TextRange
            .Font.Name = "Arial"
            .Font.Size = 14
            '.Font.color = RGB(0, 0, 255) ' Bleu
        End With
        
        ' Vérifier si le texte dépasse la diapositive
        If ligneCount > maxLignesParDiapo Then
            ' Passer à la diapositive suivante
            'slideIndex = slideIndex + 1
        Else
            ' Fin de la boucle
            Exit Do
        End If
        
    Loop

    MsgBox "Le sommaire a été ajouté avec succès.", vbInformation
End Sub
Sub InsererSommaire1(objPresentation As Object)
    Dim presentation As Object
    Dim sommaireDiapo As Object
    Dim diapo As Object
    Dim i As Integer
    Dim sommaireTexte As String
    Dim Titre As String
    Dim pptShape As Object
    Dim maxLignesParDiapo As Integer
    Dim ligneCount As Integer
    Dim lignes() As String
    Dim slideIndex As Integer
    Dim colonne1Texte As String
    Dim colonne2Texte As String
    Dim currentLineIndex As Integer
    Dim lineStartIndex As Integer
    Dim maxLinesPerColumn As Integer
    
    ' Référence à la présentation active
    Set presentation = objPresentation

    ' Nombre maximum de lignes par diapositive et par colonne
    maxLignesParDiapo = 20
    maxLinesPerColumn = maxLignesParDiapo / 2

    ' Créer le texte du sommaire dynamiquement
    sommaireTexte = "Sommaire" & vbCrLf & vbCrLf
    
    ' Ajouter une diapositive pour le sommaire au début de la présentation
    slideIndex = 1
    Set sommaireDiapo = presentation.Slides.Add(8, ppLayoutText)
    sommaireDiapo.Shapes.Title.TextFrame.TextRange.text = "Sommaire"
    
    ' Ajouter les titres des diapositives au sommaire
    For i = 1 To presentation.Slides.Count
        Set diapo = presentation.Slides(i)
        On Error Resume Next
        Titre = diapo.Shapes.Placeholders(1).TextFrame.TextRange.text
        On Error GoTo 0
        If Titre <> "" Then
            sommaireTexte = sommaireTexte & i & ". " & Titre & vbCrLf
        End If
    Next i
    
    ' Diviser le texte en lignes
    lignes = Split(sommaireTexte, vbCrLf)
    currentLineIndex = 0
    
    Do
        ' Réinitialiser les textes pour les colonnes
        colonne1Texte = ""
        colonne2Texte = ""
        
        ' Remplir les colonnes
        For i = 0 To maxLinesPerColumn - 1
            lineStartIndex = currentLineIndex + i
            If lineStartIndex <= UBound(lignes) Then
                colonne1Texte = colonne1Texte & lignes(lineStartIndex) & vbCrLf
            End If
        Next i
        
        For i = 0 To maxLinesPerColumn - 1
            lineStartIndex = currentLineIndex + maxLinesPerColumn + i
            If lineStartIndex <= UBound(lignes) Then
                colonne2Texte = colonne2Texte & lignes(lineStartIndex) & vbCrLf
            End If
        Next i
        
        ' Ajouter le texte aux colonnes de la diapositive
        With sommaireDiapo.Shapes.Placeholders(2).TextFrame.TextRange
            .text = colonne1Texte
            .Font.Name = "Arial"
            .Font.Size = 14
            .Font.color = RGB(0, 0, 255) ' Bleu
        End With
        
        Set pptShape = sommaireDiapo.Shapes.AddTextbox(msoTextOrientationHorizontal, 320, 20, 300, 400)
        pptShape.TextFrame.TextRange.text = colonne2Texte
        With pptShape.TextFrame.TextRange
            .Font.Name = "Arial"
            .Font.Size = 14
            .Font.color = RGB(0, 0, 255) ' Bleu
        End With
        
        ' Mettre à jour l'index de ligne et ajouter une nouvelle diapositive si nécessaire
        currentLineIndex = currentLineIndex + (2 * maxLinesPerColumn)
        If currentLineIndex <= UBound(lignes) Then
            slideIndex = slideIndex + 1
            Set sommaireDiapo = presentation.Slides.Add(slideIndex, ppLayoutText)
            sommaireDiapo.Shapes.Title.TextFrame.TextRange.text = "Sommaire - Page " & slideIndex
        End If
        
    Loop While currentLineIndex <= UBound(lignes)

    MsgBox "Le sommaire a été ajouté avec succès.", vbInformation
End Sub
Sub InsererSommaire4(objPresentation As Object)
    Dim presentation As Object
    Dim sommaireDiapo As Object
    Dim diapo As Object
    Dim i As Integer
    Dim sommaireTexte As String
    Dim Titre As String
    Dim pptShape As Object
    Dim slideIndex As Integer
    Dim mainSectionIndex As Integer
    Dim subSectionIndex As Integer

    ' Référence à la présentation active
    Set presentation = objPresentation

    ' Initialiser le texte du sommaire
    sommaireTexte = ""
    
    ' Initialiser les indices de section
    mainSectionIndex = 0
    subSectionIndex = 0

    ' Parcourir les diapositives pour générer le sommaire
    For i = 1 To presentation.Slides.Count
        Set diapo = presentation.Slides(i)
        On Error Resume Next
        Titre = diapo.Shapes.Placeholders(1).TextFrame.TextRange.text
        On Error GoTo 0

        If Titre <> "" Then
            ' Détecter les sections principales et sous-sections
            If InStr(1, Titre, "results", vbTextCompare) > 0 Then
                ' Section principale
                mainSectionIndex = mainSectionIndex + 1
                subSectionIndex = 0
                sommaireTexte = sommaireTexte & mainSectionIndex & ". " & Titre & vbCrLf
            ElseIf mainSectionIndex > 0 Then
                ' Sous-section
                subSectionIndex = subSectionIndex + 1
                sommaireTexte = sommaireTexte & vbTab & mainSectionIndex & "." & subSectionIndex & " " & Titre & vbCrLf
            Else
                ' Si aucune section principale n'a été trouvée
                sommaireTexte = sommaireTexte & Titre & vbCrLf
            End If
        End If
    Next i

    
    ' Ajouter une diapositive pour le sommaire
    slideIndex = 1
    Set sommaireDiapo = presentation.Slides.Add(slideIndex, ppLayoutText)
    sommaireDiapo.Shapes.Title.TextFrame.TextRange.text = "Sommaire"
    
    ' Ajouter le texte du sommaire à la diapositive
    sommaireDiapo.Shapes.Placeholders(2).TextFrame.TextRange.text = sommaireTexte

    MsgBox "Le sommaire a été ajouté avec succès.", vbInformation
End Sub


Sub InsererSommaire3(objPresentation As Object)
    Dim presentation As Object
    Dim sommaireDiapo As Object
    Dim diapo As Object
    Dim i As Integer
    Dim sommaireTexte As String
    Dim Titre As String
    Dim pptShape As Object
    Dim maxLignesParDiapo As Integer
    Dim ligneCount As Integer
    Dim lignes() As String
    Dim slideIndex As Integer
    Dim currentLineIndex As Integer
    Dim maxLinesPerColumn As Integer
    Dim mainSectionIndex As Integer
    Dim subSectionIndex As Integer

    ' Référence à la présentation active
    Set presentation = objPresentation

    ' Nombre maximum de lignes par diapositive et par colonne
    maxLignesParDiapo = 20
    maxLinesPerColumn = maxLignesParDiapo / 2

    ' Initialiser les indices de section
    mainSectionIndex = 0
    subSectionIndex = 0
    
    ' Créer le texte du sommaire dynamiquement
    sommaireTexte = ""

    ' Ajouter les titres des diapositives au sommaire
    For i = 1 To presentation.Slides.Count
        Set diapo = presentation.Slides(i)
        On Error Resume Next
        Titre = diapo.Shapes.Placeholders(1).TextFrame.TextRange.text
        On Error GoTo 0
        
        If Titre <> "" Then
            ' Détecter les sections principales et sous-sections
            If InStr(1, Titre, "results", vbTextCompare) > 0 Then
                ' C'est une section principale
                mainSectionIndex = mainSectionIndex + 1
                subSectionIndex = 0
                sommaireTexte = sommaireTexte & mainSectionIndex & ". " & Titre & vbCrLf
            ElseIf mainSectionIndex > 0 Then
                ' C'est une sous-section
                subSectionIndex = subSectionIndex + 1
                sommaireTexte = sommaireTexte & vbTab & mainSectionIndex & "." & subSectionIndex & " " & Titre & vbCrLf
            Else
                ' Si aucune section principale n'a été trouvée
                sommaireTexte = sommaireTexte & Titre & vbCrLf
            End If
        End If
    Next i
    
    ' Diviser le texte en lignes
    lignes = Split(sommaireTexte, vbCrLf)
    currentLineIndex = 0
    
    ' Ajouter une diapositive pour le sommaire
    slideIndex = 1
    Set sommaireDiapo = presentation.Slides.Add(slideIndex, ppLayoutText)
    sommaireDiapo.Shapes.Title.TextFrame.TextRange.text = "Sommaire"
    
    Do
        ' Réinitialiser les textes pour les colonnes
        Dim colonne1Texte As String
        Dim colonne2Texte As String
        colonne1Texte = ""
        colonne2Texte = ""
        
        ' Remplir les colonnes
        For i = 0 To maxLinesPerColumn - 1
            If currentLineIndex + i <= UBound(lignes) Then
                colonne1Texte = colonne1Texte & lignes(currentLineIndex + i) & vbCrLf
            End If
        Next i
        
        For i = 0 To maxLinesPerColumn - 1
            If currentLineIndex + maxLinesPerColumn + i <= UBound(lignes) Then
                colonne2Texte = colonne2Texte & lignes(currentLineIndex + maxLinesPerColumn + i) & vbCrLf
            End If
        Next i
        
        ' Ajouter le texte aux colonnes de la diapositive
        With sommaireDiapo.Shapes.Placeholders(2).TextFrame.TextRange
            .text = colonne1Texte
            .Font.Name = "Arial"
            .Font.Size = 14
            .Font.color = RGB(0, 0, 255) ' Bleu
        End With
        
        Set pptShape = sommaireDiapo.Shapes.AddTextbox(msoTextOrientationHorizontal, 320, 20, 300, 400)
        pptShape.TextFrame.TextRange.text = colonne2Texte
        With pptShape.TextFrame.TextRange
            .Font.Name = "Arial"
            .Font.Size = 14
            .Font.color = RGB(0, 0, 255) ' Bleu
        End With
        
        ' Mettre à jour l'index de ligne et ajouter une nouvelle diapositive si nécessaire
        currentLineIndex = currentLineIndex + (2 * maxLinesPerColumn)
        If currentLineIndex <= UBound(lignes) Then
            slideIndex = slideIndex + 1
            Set sommaireDiapo = presentation.Slides.Add(slideIndex, ppLayoutText)
            sommaireDiapo.Shapes.Title.TextFrame.TextRange.text = "Sommaire - Page " & slideIndex
        End If
        
    Loop While currentLineIndex <= UBound(lignes)

    MsgBox "Le sommaire a été ajouté avec succès.", vbInformation
End Sub
Sub AddSlideGlossary(objPresentation As Object, objppt As Object)
 
    Dim oleObj As OLEObject
    Dim pptApp As Object
    Dim pptDoc As Object
    Dim pasted As Boolean
    
    
 
    
  
 
    Set oleObj = F_DNT.OLEObjects("Object 32")
    oleObj.Activate
    Set pptApp = oleObj.Object.Application
    Set pptDoc = pptApp.ActivePresentation
    
    ViderPressePapiers
            pasted = False
            Do While Not pasted
                On Error Resume Next
                pptDoc.Slides(1).Copy
                'DoEvents
                'sleep 500
                objPresentation.Slides.Paste (3)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
            Loop
    
    
    
    
 
    pptDoc.Close
 
End Sub

 


 

Sub InsererSommaireDynamiqueDeuxNiveaux(objPresentation As Object)
    Dim presentation As Object
    Dim sommaireDiapo1 As Object
    Dim sommaireDiapo2 As Object
    Dim sommaireDiapo3 As Object
    Dim diapo As Object
    Dim sommaireTexte As String
    Dim sommaireTexte1 As String
    Dim sommaireTexte2 As String
    Dim sommaireTexte3 As String
    Dim Titre As Variant
    Dim pptShape As Object
    Dim slideIndex1 As Integer
    Dim slideIndex2 As Integer
    Dim slideIndex3 As Integer
    Dim maxLignesParDiapo As Integer
    Dim lignes() As String
    Dim ligneCount As Integer
    Dim i As Integer
    Dim titleSet As Object ' Pour stocker les titres ajoutés
    Dim currentLine As Integer ' Pour la plage de lien hypertexte
    Dim nextTitleIndex As Integer ' Pour gérer la numérotation des titres
    Dim separatorLine As Object
    Dim slide As Object
    Dim shape As Object
    Dim text As String
    Dim line As String
    
    
    Dim newNumber As Long
    Dim j As Long
    Dim dotCount As Long
    Dim lastChar As String
    Dim numberStart As Long
    Dim numberText As String
    
    
    
    
    
    ' Gestion des erreurs
    On Error GoTo ErrorHandler

    ' Référence à la présentation active
    Set presentation = objPresentation

    ' Nombre maximum de lignes par diapositive
    maxLignesParDiapo = 20 ' Ajuster selon les besoins

    ' Dictionnaire pour suivre les titres ajoutés
    Set titleSet = CreateObject("Scripting.Dictionary")

    ' Entrées fixes du sommaire avec hyperliens
    sommaireTexte = "1. Test informations" & vbCrLf & "2. oDriv results" & vbCrLf
    
    
    '''''''''''''''''''chercher l'indice de slide qui contient "SUMMARY"
    For Each slide In presentation.Slides
        ' Parcourir toutes les formes (shapes) de la diapositive
        For Each shape In slide.Shapes
            ' Vérifier si la forme contient du texte
            If shape.HasTextFrame Then
                If shape.TextFrame.HasText Then
                    text = shape.TextFrame.TextRange.text
                    ' Tester si "BACK UP" est dans le texte de la forme
                    If InStr(1, text, "SUMMARY", vbTextCompare) > 0 Then
                        slideIndex1 = slide.slideIndex
                        Exit For ' Sortir dès que "BACK UP" est trouvé
                        
                    End If
                End If
            End If
        Next shape
        If slideIndex1 <> 0 Then Exit For
    Next slide
    
    ' Ajouter les hyperliens pour les entrées fixes
    titleSet.Add "1. Test informations", slideIndex1 + 1
    titleSet.Add "2. oDriv results", slideIndex1 + 3
    ' Initialiser l'index pour le prochain titre
    nextTitleIndex = 3

    ' Ajout des titres des diapositives au sommaire de manière dynamique
    ' Ajout des titres des diapositives au sommaire de manière dynamique
For i = 10 To presentation.Slides.Count
    Set diapo = presentation.Slides(i)
      
    ' Vérification si la diapositive est un cartouche
    If Not IsCartouche(diapo) Then
        If diapo.Shapes.Placeholders.Count > 0 Then
            On Error Resume Next
            Titre = diapo.Shapes.Placeholders(1).TextFrame.TextRange.text
             Debug.Print Titre ' Prints the text of each placeholder
            On Error GoTo 0
            
            ' Ajouter le titre si non déjà présent
            '''''If Titre = "2.24 AUTO STOP" Then Stop
            If Len(Titre) > 0 And Not titleSet.Exists(Titre) Then
                ' Calculer la longueur nécessaire pour les points
                Dim lineLength As Integer
                Dim totalLength As Integer
                totalLength = 60 ' Ajustez selon la largeur de votre zone de texte
                If Not Titre Like "*.*" Then
                    lineLength = totalLength - Len(Titre) ''- Len(" (Numero: " & CStr(i) & ")") - 1 ' -1 pour l'espace avant le numéro
                Else
                    lineLength = totalLength - Len(Titre) - 3 '' Len(vbTab)
                End If
                
                If Not Titre Like "*.*" Then
                    ' Ajouter le titre principal
                    If lineLength > 0 Then
                        sommaireTexte = sommaireTexte & nextTitleIndex & ". " & Titre & String(lineLength, ".") & i & vbCrLf
'                        sommaireTexte = sommaireTexte & nextTitleIndex & ". " & Titre & vbCrLf
                       ' Add this after updating each summary text
                    Debug.Print "Building Summary - sommaireTexte: " & vbCrLf & sommaireTexte

                        
                    Else
                        ' Si le titre est trop long, afficher seulement le titre et le numéro
                        sommaireTexte = sommaireTexte & nextTitleIndex & Titre & String(lineLength, ".") & i & vbCrLf
'                         sommaireTexte = sommaireTexte & nextTitleIndex & Titre & vbCrLf
                         Debug.Print "Building Summary - sommaireTexte: " & vbCrLf & sommaireTexte
                    End If
                    nextTitleIndex = nextTitleIndex + 1 ' Incrémenter l'index pour le prochain titre
                Else
                    ' Ajouter le sous-titre
                    sommaireTexte = sommaireTexte & vbTab & Titre & String(lineLength, ".") & i & vbCrLf
'                   sommaireTexte = sommaireTexte & vbTab & Titre & vbCrLf
                     Debug.Print "Building Summary - sommaireTexte: " & vbCrLf & sommaireTexte
                End If
                titleSet.Add Titre, i ' Stocker l'index de la diapositive avec le titre
            End If
        End If
    End If
Next i


    ' Séparer le texte du sommaire en lignes
    lignes = Split(sommaireTexte, vbCrLf)
    ligneCount = UBound(lignes) + 1

    ' Diviser le sommaire en trois parties
    sommaireTexte1 = ""
    sommaireTexte2 = ""
    sommaireTexte3 = ""

    For i = 0 To ligneCount - 1
        If i < maxLignesParDiapo Then
            sommaireTexte1 = sommaireTexte1 & lignes(i) & vbCrLf
        ElseIf i < maxLignesParDiapo * 2 Then
            sommaireTexte2 = sommaireTexte2 & lignes(i) & vbCrLf
        Else
            sommaireTexte3 = sommaireTexte3 & lignes(i) & vbCrLf
        End If
    Next i
    If sommaireTexte3 <> "" Then
    
    For i = 2 To UBound(lignes)
        line = lignes(i)
        dotCount = 0
        Dim hasFiveDots As Boolean: hasFiveDots = False
        ' Check for 5 or more consecutive dots anywhere in the line
        For j = 1 To Len(line)
            If Mid(line, j, 1) = "." Then
                dotCount = dotCount + 1
                If dotCount >= 5 Then
                    hasFiveDots = True
                    Exit For
                End If
            Else
                dotCount = 0 ' reset if non-dot
            End If
        Next j
        ' If line qualifies, increment the ending number
        If hasFiveDots Then
            ' Work backwards to find the number at the end
            For j = Len(line) To 1 Step -1
                lastChar = Mid(line, j, 1)
                If Not IsNumeric(lastChar) Then
                    numberStart = j + 1
                    Exit For
                End If
            Next j
            numberText = Mid(line, numberStart)
            If IsNumeric(numberText) Then
                newNumber = CLng(numberText) + 1
                lignes(i) = Left(line, numberStart - 1) & newNumber
            End If
        End If
    Next i
    End If
    
    
    
    
    

    
    sommaireTexte1 = ""
    sommaireTexte2 = ""
    sommaireTexte3 = ""

    For i = 0 To ligneCount - 1
        If i < maxLignesParDiapo Then
            sommaireTexte1 = sommaireTexte1 & lignes(i) & vbCrLf
        ElseIf i < maxLignesParDiapo * 2 Then
            sommaireTexte2 = sommaireTexte2 & lignes(i) & vbCrLf
        Else
            sommaireTexte3 = sommaireTexte3 & lignes(i) & vbCrLf
        End If
    Next i
    
'    '''''''''''''''''''chercher l'indice de slide qui contient "SUMMARY"
'    For Each slide In presentation.Slides
'        ' Parcourir toutes les formes (shapes) de la diapositive
'        For Each shape In slide.Shapes
'            ' Vérifier si la forme contient du texte
'            If shape.HasTextFrame Then
'                If shape.TextFrame.HasText Then
'                    text = shape.TextFrame.TextRange.text
'                    ' Tester si "BACK UP" est dans le texte de la forme
'                    If InStr(1, text, "SUMMARY", vbTextCompare) > 0 Then
'                        slideIndex1 = slide.slideIndex
'                        Exit For ' Sortir dès que "BACK UP" est trouvé
'
'                    End If
'                End If
'            End If
'        Next shape
'        If slideIndex1 <> 0 Then Exit For
'    Next slide
    ' Affecter les index des diapositives de sommaire
    'slideIndex1 = 9
    'slideIndex2 = slideIndex1 + 1
    slideIndex2 = slideIndex1
    slideIndex3 = slideIndex2 + 1

    ' Ajouter le premier sommaire
    Set sommaireDiapo1 = presentation.Slides(slideIndex1)
    If sommaireTexte3 <> "" Then
        sommaireDiapo1.Copy
        presentation.Slides.Paste (slideIndex1 + 1)
    End If
    
    If sommaireTexte3 <> "" Then
    Set titleSet = CreateObject("Scripting.Dictionary")
    titleSet.Add "1. Test informations", slideIndex1 + 2
    titleSet.Add "2. oDriv results", slideIndex1 + 3
    For i = 10 To presentation.Slides.Count
    Set diapo = presentation.Slides(i)
      
    ' Vérification si la diapositive est un cartouche
    If Not IsCartouche(diapo) Then
        If diapo.Shapes.Placeholders.Count > 0 Then
            On Error Resume Next
            Titre = diapo.Shapes.Placeholders(1).TextFrame.TextRange.text
             
            On Error GoTo 0
            
            
            If Len(Titre) > 0 And Not titleSet.Exists(Titre) Then
                
                
                totalLength = 60 ' Ajustez selon la largeur de votre zone de texte
                If Not Titre Like "*.*" Then
                    lineLength = totalLength - Len(Titre) ''- Len(" (Numero: " & CStr(i) & ")") - 1 ' -1 pour l'espace avant le numéro
                Else
                    lineLength = totalLength - Len(Titre) - 3 '' Len(vbTab)
                End If
                
                If Not Titre Like "*.*" Then
                    ' Ajouter le titre principal
                    If lineLength > 0 Then
                        'sommaireTexte = sommaireTexte & nextTitleIndex & ". " & Titre & String(lineLength, ".") & i & vbCrLf
'                        sommaireTexte = sommaireTexte & nextTitleIndex & ". " & Titre & vbCrLf
                       ' Add this after updating each summary text
                    Debug.Print "Building Summary - sommaireTexte: " & vbCrLf & sommaireTexte

                        
                    Else
                        ' Si le titre est trop long, afficher seulement le titre et le numéro
                        'sommaireTexte = sommaireTexte & nextTitleIndex & Titre & String(lineLength, ".") & i & vbCrLf
'                         sommaireTexte = sommaireTexte & nextTitleIndex & Titre & vbCrLf
                         Debug.Print "Building Summary - sommaireTexte: " & vbCrLf & sommaireTexte
                    End If
                    nextTitleIndex = nextTitleIndex + 1 ' Incrémenter l'index pour le prochain titre
                Else
                    ' Ajouter le sous-titre
                    'sommaireTexte = sommaireTexte & vbTab & Titre & String(lineLength, ".") & i & vbCrLf
'                   sommaireTexte = sommaireTexte & vbTab & Titre & vbCrLf
                     Debug.Print "Building Summary - sommaireTexte: " & vbCrLf & sommaireTexte
                End If
                titleSet.Add Titre, i  ' Stocker l'index de la diapositive avec le titre
            End If
        End If
    End If
Next i
End If
    
    'Set pptShape = sommaireDiapo1.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
    Set pptShape = sommaireDiapo1.Shapes.AddTextbox(msoTextOrientationHorizontal, 6, 100, 550, 400)
    'pptShape.TextFrame.TextRange.text = sommaireTexte1
    
 
    pptShape.TextFrame.TextRange.text = sommaireTexte1
    
   
 
    'pptShape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignRight
    
    
    pptShape.TextFrame.TextRange.Font.Name = "Encode Sans (Corps)"
    ''''MsgBox sommaireTexte1
    pptShape.TextFrame.TextRange.Font.Size = 12
    pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0)
    ApplyHyperlinks titleSet, pptShape

    ' Ajouter le deuxième sommaire si nécessaire
    If sommaireTexte2 <> "" Then
        
        ' Ajouter la ligne séparatrice verticale
        Set separatorLine = sommaireDiapo1.Shapes.AddLine(500, 100, 500, 500) ' Coordonnées pour la ligne (x1, y1, x2, y2)
        separatorLine.line.Weight = 2 ' Épaisseur de la ligne
        separatorLine.line.ForeColor.RGB = RGB(0, 0, 0) ' Couleur de la ligne (ici, noir)
        
        Set sommaireDiapo2 = presentation.Slides(slideIndex2)
        'Set pptShape = sommaireDiapo2.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        Set pptShape = sommaireDiapo2.Shapes.AddTextbox(msoTextOrientationHorizontal, 500, 100, 600, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte2
        pptShape.TextFrame.TextRange.Font.Name = "Encode Sans (Corps)"
        pptShape.TextFrame.TextRange.Font.Size = 12
        pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0)
        ApplyHyperlinks titleSet, pptShape
    End If

    ' Ajouter le troisième sommaire si nécessaire
    If sommaireTexte3 <> "" Then
        Set sommaireDiapo3 = presentation.Slides(slideIndex3)
        'Set pptShape = sommaireDiapo3.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        Set pptShape = sommaireDiapo3.Shapes.AddTextbox(msoTextOrientationHorizontal, 6, 100, 550, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte3
        pptShape.TextFrame.TextRange.Font.Name = "Encode Sans (Corps)"
        pptShape.TextFrame.TextRange.Font.Size = 12
        pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0)
        ApplyHyperlinks titleSet, pptShape
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Une erreur s'est produite : " & ERR.description, vbExclamation
End Sub




Sub CreerTableDesMatieres()
 
    ' Déclaration des objets PowerPoint
    Dim pptApp As Object
    Dim pptPres As Object
    Dim pptSlide As Object
    Dim pptTextbox As Object
    Dim slideIndex As Integer
    Dim titleText As String
    Dim tableOfContents As String
    Dim i As Integer
    ' Initialiser PowerPoint
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    If pptApp Is Nothing Then
        Set pptApp = CreateObject("PowerPoint.Application")
    End If
    On Error GoTo 0
    ' Accéder à la présentation active
    Set pptPres = pptApp.ActivePresentation
    ' Créer une diapositive pour la Table des Matières
    Set pptSlide = pptPres.Slides.Add(1, ppLayoutText)
    pptSlide.Shapes(1).TextFrame.TextRange.text = "Table des Matières"
    pptSlide.Shapes(2).TextFrame.TextRange.text = ""
    ' Initialiser le texte de la table des matières
    tableOfContents = ""
    ' Parcourir toutes les diapositives et extraire les titres
    For i = 2 To pptPres.Slides.Count ' Commence à la diapositive 2 (évite la diapositive de Table des Matières)
        ' Titre de la diapositive (s'il existe)
        On Error Resume Next
        titleText = pptPres.Slides(i).Shapes(1).TextFrame.TextRange.text
        On Error GoTo 0
        ' Ajouter le titre à la table des matières, si un titre est trouvé
        If Len(titleText) > 0 Then
            tableOfContents = tableOfContents & titleText & vbTab & vbTab & "Page " & i & vbCrLf
        End If
    Next i
    ' Insérer la table des matières dans la deuxième zone de texte de la première diapositive
    pptSlide.Shapes(2).TextFrame.TextRange.text = tableOfContents
    ' Aligner les numéros de page à droite
    Dim contentRange As Object
    Set contentRange = pptSlide.Shapes(2).TextFrame.TextRange
    Dim contentLines() As String
    contentLines = Split(tableOfContents, vbCrLf)
    Dim line As Integer
    For line = 0 To UBound(contentLines)
        ' Trouver l'espace avant le numéro de page
        Dim lineText As String
        lineText = contentLines(line)
        Dim spacePos As Integer
        spacePos = InStrRev(lineText, vbTab)
        If spacePos > 0 Then
            ' Extraire la partie avant le numéro de page et le numéro
            Dim titlePart As String
            Dim pagePart As String
            titlePart = Left(lineText, spacePos - 1)
            pagePart = Mid(lineText, spacePos + 1)
            ' Aligner le numéro de page à droite
            contentRange.Paragraphs(line + 1).ParagraphFormat.Alignment = ppAlignLeft
            contentRange.Paragraphs(line + 1).text = titlePart & vbTab & pagePart
        End If
    Next line
    ' Mise en forme (optionnel)
    pptSlide.Shapes(2).TextFrame.TextRange.Font.Name = "Arial"
    pptSlide.Shapes(2).TextFrame.TextRange.Font.Size = 16
    pptSlide.Shapes(2).TextFrame.TextRange.ParagraphFormat.SpaceAfter = 5
    ' Afficher la présentation PowerPoint
    pptApp.Visible = True
End Sub

Sub CopierTexteAvecMiseEnForme()
    ' Déclare une variable pour la présentation et la diapositive
    Dim pptShape As shape
    Dim texteAvecMiseEnForme As String
    ' Exemple de texte avec mise en forme. Remarque : vous pouvez aussi le récupérer d'une autre source (Word, etc.)
    texteAvecMiseEnForme = "Ceci est un texte formaté." & vbCrLf & _
                          "Il peut contenir des" & vbTab & "styles de texte." & vbCrLf & _
                          "Exemple : gras, italique, etc."
 
    ' Ajoute une zone de texte à la diapositive active
    Set pptShape = ActivePresentation.Slides(1).Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 500, 300)
 
    ' Utiliser FormattedText pour copier le texte avec la mise en forme
    pptShape.TextFrame.TextRange.FormattedText = texteAvecMiseEnForme
 
    ' Vous pouvez aussi appliquer un format général si nécessaire
    pptShape.TextFrame.TextRange.Font.Name = "Arial"
    pptShape.TextFrame.TextRange.Font.Size = 14
    pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0) ' Couleur du texte (Noir)
 
End Sub
Function InsererSommaireDynamiqueDeuxNiveauxV14(objPresentation As Object)
    Dim presentation As Object
    Dim sommaireDiapo1 As Object
    Dim sommaireDiapo2 As Object
    Dim sommaireDiapo3 As Object
    Dim diapo As Object
    Dim sommaireTexte As String
    Dim sommaireTexte1 As String
    Dim sommaireTexte2 As String
    Dim sommaireTexte3 As String
    Dim Titre As Variant
    Dim pptShape As Object
    Dim slideIndex1 As Integer
    Dim slideIndex2 As Integer
    Dim slideIndex3 As Integer
    Dim maxLignesParDiapo As Integer
    Dim lignes() As String
    Dim ligneCount As Integer
    Dim i As Integer
    Dim titleSet As Object ' Pour stocker les titres ajoutés
    Dim currentLine As Integer ' Pour la plage de lien hypertexte
    Dim nextTitleIndex As Integer ' Pour gérer la numérotation des titres

    ' Gestion des erreurs
    On Error GoTo ErrorHandler

    ' Référence à la présentation active
    Set presentation = objPresentation

    ' Nombre maximum de lignes par diapositive
    maxLignesParDiapo = 20 ' Ajuster selon les besoins

    ' Dictionnaire pour suivre les titres ajoutés
    Set titleSet = CreateObject("Scripting.Dictionary")

    ' Entrées fixes du sommaire avec hyperliens
    sommaireTexte = "1. Test information" & vbCrLf & "2. oDriv results" & vbCrLf
    
    ' Ajouter les hyperliens pour les entrées fixes
    titleSet.Add "1. Test information", 11 ' Supposons que l'index de la diapo soit 1
    titleSet.Add "2. oDriv results", 12 ' Supposons que l'index de la diapo soit 2

    ' Initialiser l'index pour le prochain titre
    nextTitleIndex = 3

    ' Ajout des titres des diapositives au sommaire de manière dynamique
     ' Ajout des titres des diapositives au sommaire de manière dynamique
    For i = 10 To presentation.Slides.Count
        Set diapo = presentation.Slides(i)
        
        ' Vérification si la diapositive est un cartouche
        If Not IsCartouche(diapo) Then
            If diapo.Shapes.Placeholders.Count > 0 Then
                On Error Resume Next
                Titre = diapo.Shapes.Placeholders(1).TextFrame.TextRange.text
                On Error GoTo 0
                ' Ajouter le titre si non déjà présent
                If Len(Titre) > 0 And Not titleSet.Exists(Titre) Then
                    ' Ajouter le numéro de diapositive au titre
                    If Not Titre Like "*.*" Then
                        ' Ajouter le titre principal
                        sommaireTexte = sommaireTexte & nextTitleIndex & ". " & Titre & " (Numero:..  " & i & ")" & vbCrLf
                        nextTitleIndex = nextTitleIndex + 1 ' Incrémenter l'index pour le prochain titre
                    Else
                        ' Ajouter le sous-titre
                        sommaireTexte = sommaireTexte & vbTab & " " & Titre & " (Numero:  " & i & ")" & vbCrLf
                    End If
                    titleSet.Add Titre, i ' Stocker l'index de la diapositive avec le titre
                End If
            End If
        End If
    Next i

    ' Séparer le texte du sommaire en lignes
    lignes = Split(sommaireTexte, vbCrLf)
    ligneCount = UBound(lignes) + 1

    ' Diviser le sommaire en trois parties
    sommaireTexte1 = ""
    sommaireTexte2 = ""
    sommaireTexte3 = ""

    For i = 0 To ligneCount - 1
        If i < maxLignesParDiapo Then
            sommaireTexte1 = sommaireTexte1 & lignes(i) & vbCrLf
        ElseIf i < maxLignesParDiapo * 2 Then
            sommaireTexte2 = sommaireTexte2 & lignes(i) & vbCrLf
        Else
            sommaireTexte3 = sommaireTexte3 & lignes(i) & vbCrLf
        End If
    Next i

    ' Affecter les index des diapositives de sommaire
    slideIndex1 = 9
    slideIndex2 = slideIndex1 + 1
    slideIndex3 = slideIndex2 + 1

    ' Ajouter le premier sommaire
    Set sommaireDiapo1 = presentation.Slides(slideIndex1)
    Set pptShape = sommaireDiapo1.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
    pptShape.TextFrame.TextRange.text = sommaireTexte1
    pptShape.TextFrame.TextRange.Font.Name = "Encode Sans (Corps)"
    pptShape.TextFrame.TextRange.Font.Size = 15
    pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0)
    ApplyHyperlinks titleSet, pptShape

    ' Ajouter le deuxième sommaire si nécessaire
    If sommaireTexte2 <> "" Then
        Set sommaireDiapo2 = presentation.Slides(slideIndex2)
        Set pptShape = sommaireDiapo2.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte2
        pptShape.TextFrame.TextRange.Font.Name = "Encode Sans (Corps)"
        pptShape.TextFrame.TextRange.Font.Size = 15
        pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0)
        ApplyHyperlinks titleSet, pptShape
    End If

    ' Ajouter le troisième sommaire si nécessaire
    If sommaireTexte3 <> "" Then
        Set sommaireDiapo3 = presentation.Slides(slideIndex3)
        Set pptShape = sommaireDiapo3.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte3
        pptShape.TextFrame.TextRange.Font.Name = "Encode Sans (Corps)"
        pptShape.TextFrame.TextRange.Font.Size = 15
        pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0)
        ApplyHyperlinks titleSet, pptShape
    End If

    Exit Function

ErrorHandler:
    MsgBox "Une erreur s'est produite : " & ERR.description, vbExclamation
End Function



Function InsererSommaireDynamiqueDeuxNiveauxV12(objPresentation As Object)
    Dim presentation As Object
    Dim sommaireDiapo1 As Object
    Dim sommaireDiapo2 As Object
    Dim sommaireDiapo3 As Object
    Dim diapo As Object
    Dim sommaireTexte As String
    Dim sommaireTexte1 As String
    Dim sommaireTexte2 As String
    Dim sommaireTexte3 As String
    Dim Titre As Variant
    Dim pptShape As Object
    Dim slideIndex1 As Integer
    Dim slideIndex2 As Integer
    Dim slideIndex3 As Integer
    Dim maxLignesParDiapo As Integer
    Dim lignes() As String
    Dim ligneCount As Integer
    Dim i As Integer
    Dim titleSet As Object ' Pour stocker les titres ajoutés
    Dim currentLine As Integer ' Pour la plage de lien hypertexte
    Dim nextTitleIndex As Integer ' Pour gérer la numérotation des titres

    ' Gestion des erreurs
    On Error GoTo ErrorHandler

    ' Référence à la présentation active
    Set presentation = objPresentation

    ' Nombre maximum de lignes par diapositive
    maxLignesParDiapo = 20 ' Ajuster selon les besoins

    ' Dictionnaire pour suivre les titres ajoutés
    Set titleSet = CreateObject("Scripting.Dictionary")

    ' Entrées fixes du sommaire
    sommaireTexte = "1. Test information" & vbCrLf & "2. oDriv results" & vbCrLf

    ' Initialiser l'index pour le prochain titre
    nextTitleIndex = 3

    ' Ajout des titres des diapositives au sommaire de manière dynamique
    For i = 10 To presentation.Slides.Count
        Set diapo = presentation.Slides(i)
        
        ' Vérification si la diapositive est un cartouche
        If Not IsCartouche(diapo) Then
            If diapo.Shapes.Placeholders.Count > 0 Then
                On Error Resume Next
                Titre = diapo.Shapes.Placeholders(1).TextFrame.TextRange.text
                On Error GoTo 0
                ' Ajouter le titre si non déjà présent
                If Len(Titre) > 0 And Not titleSet.Exists(Titre) Then
                    ' Vérifier si le titre a un numéro
                    If Not Titre Like "*.*" Then
                        sommaireTexte = sommaireTexte & nextTitleIndex & ". " & Titre & vbCrLf ' Ajout au niveau 1
                        nextTitleIndex = nextTitleIndex + 1 ' Incrémenter l'index pour le prochain titre
                    Else
                        sommaireTexte = sommaireTexte & vbTab & " " & Titre & vbCrLf ' Ajout au niveau 2
                    End If
                    titleSet.Add Titre, i ' Stocker l'index de la diapositive avec le titre
                End If
            End If
        End If
    Next i

    ' Séparer le texte du sommaire en lignes
    lignes = Split(sommaireTexte, vbCrLf)
    ligneCount = UBound(lignes) + 1

    ' Diviser le sommaire en trois parties
    sommaireTexte1 = ""
    sommaireTexte2 = ""
    sommaireTexte3 = ""

    For i = 0 To ligneCount - 1
        If i < maxLignesParDiapo Then
            sommaireTexte1 = sommaireTexte1 & lignes(i) & vbCrLf
        ElseIf i < maxLignesParDiapo * 2 Then
            sommaireTexte2 = sommaireTexte2 & lignes(i) & vbCrLf
        Else
            sommaireTexte3 = sommaireTexte3 & lignes(i) & vbCrLf
        End If
    Next i

    ' Affecter les index des diapositives de sommaire
    slideIndex1 = 9
    slideIndex2 = slideIndex1 + 1
    slideIndex3 = slideIndex2 + 1

    ' Ajouter le premier sommaire
    Set sommaireDiapo1 = presentation.Slides(slideIndex1)
    Set pptShape = sommaireDiapo1.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
    pptShape.TextFrame.TextRange.text = sommaireTexte1
    pptShape.TextFrame.TextRange.Font.Name = "Encode Sans (Corps)"
    pptShape.TextFrame.TextRange.Font.Size = 15
    pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0)
    ApplyHyperlinks titleSet, pptShape

    ' Ajouter le deuxième sommaire si nécessaire
    If sommaireTexte2 <> "" Then
        Set sommaireDiapo2 = presentation.Slides(slideIndex2)
        Set pptShape = sommaireDiapo2.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte2
        pptShape.TextFrame.TextRange.Font.Name = "Encode Sans (Corps)"
        pptShape.TextFrame.TextRange.Font.Size = 15
        pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0)
        ApplyHyperlinks titleSet, pptShape
    End If

    ' Ajouter le troisième sommaire si nécessaire
    If sommaireTexte3 <> "" Then
        Set sommaireDiapo3 = presentation.Slides(slideIndex3)
        Set pptShape = sommaireDiapo3.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte3
        pptShape.TextFrame.TextRange.Font.Name = "Encode Sans (Corps)"
        pptShape.TextFrame.TextRange.Font.Size = 15
        pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0)
        ApplyHyperlinks titleSet, pptShape
    End If

    Exit Function


ErrorHandler:
    MsgBox "Une erreur s'est produite : " & ERR.description, vbExclamation

End Function

' Fonction pour vérifier si la diapositive est un cartouche
Function IsCartouche(diapo As Object) As Boolean
    Dim Titre As String
    On Error Resume Next
    Titre = diapo.Shapes.Placeholders(1).TextFrame.TextRange.text
    On Error GoTo 0
    
    ' Condition pour identifier les diapositives "cartouche"
    If InStr(1, Titre, "Back UP", vbTextCompare) > 0 Or InStr(1, Titre, "Back UP slides for Top area of improvement", vbTextCompare) > 0 Then
        IsCartouche = True
    Else
        IsCartouche = False
    End If
End Function

Sub ApplyHyperlinks(titleSet As Object, pptShape As Object)
    Dim Titre As Variant
    Dim startPos As Integer, endPos As Integer

    For Each Titre In titleSet.keys
        startPos = InStr(pptShape.TextFrame.TextRange.text, Titre)
        endPos = startPos + Len(Titre) - 1

        If startPos > 0 Then
            ' Appliquer un lien hypertexte pour chaque titre
            With pptShape.TextFrame.TextRange.Characters(startPos, Len(Titre))
                .Font.Underline = msoTrue
                .Font.color.RGB = RGB(0, 0, 0)
                .ActionSettings(ppMouseClick).Hyperlink.SubAddress = CStr(titleSet(Titre))
            End With
        End If
    Next Titre
End Sub

Function InsererSommaireDynamiqueDeuxNiveauxV11(objPresentation As Object)
    Dim presentation As Object
    Dim sommaireDiapo1 As Object
    Dim sommaireDiapo2 As Object
    Dim sommaireDiapo3 As Object
    Dim diapo As Object
    Dim sommaireTexte As String
    Dim sommaireTexte1 As String
    Dim sommaireTexte2 As String
    Dim sommaireTexte3 As String
    Dim Titre As Variant
    Dim pptShape As Object
    Dim slideIndex1 As Integer
    Dim slideIndex2 As Integer
    Dim slideIndex3 As Integer
    Dim maxLignesParDiapo As Integer
    Dim lignes() As String
    Dim ligneCount As Integer
    Dim i As Integer
    Dim titleSet As Object ' Pour stocker les titres ajoutés
    Dim currentLine As Integer ' Pour la plage de lien hypertexte

    ' Gestion des erreurs
    On Error GoTo ErrorHandler

    ' Référence à la présentation active
    Set presentation = objPresentation

    ' Nombre maximum de lignes par diapositive
    maxLignesParDiapo = 20 ' Ajuster selon les besoins

    ' Dictionnaire pour suivre les titres ajoutés
    Set titleSet = CreateObject("Scripting.Dictionary")

    ' Entrées fixes du sommaire
    sommaireTexte = "1. Test information" & vbCrLf & "2. oDriv results" & vbCrLf

    ' Ajout des titres des diapositives au sommaire de manière dynamique
    For i = 10 To presentation.Slides.Count
        Set diapo = presentation.Slides(i)
        
        ' Vérification si la diapositive est un cartouche
        If Not IsCartouche(diapo) Then
            If diapo.Shapes.Placeholders.Count > 0 Then
                On Error Resume Next
                Titre = diapo.Shapes.Placeholders(1).TextFrame.TextRange.text
                On Error GoTo 0
                ' Ajouter le titre si non déjà présent
                If Len(Titre) > 0 And Not titleSet.Exists(Titre) Then
                    sommaireTexte = sommaireTexte & vbTab & " " & Titre & vbCrLf
                    titleSet.Add Titre, i ' Stocker l'index de la diapositive avec le titre
                End If
            End If
        End If
    Next i

    ' Séparer le texte du sommaire en lignes
    lignes = Split(sommaireTexte, vbCrLf)
    ligneCount = UBound(lignes) + 1

    ' Diviser le sommaire en trois parties
    sommaireTexte1 = ""
    sommaireTexte2 = ""
    sommaireTexte3 = ""

    For i = 0 To ligneCount - 1
        If i < maxLignesParDiapo Then
            sommaireTexte1 = sommaireTexte1 & lignes(i) & vbCrLf
        ElseIf i < maxLignesParDiapo * 2 Then
            sommaireTexte2 = sommaireTexte2 & lignes(i) & vbCrLf
        Else
            sommaireTexte3 = sommaireTexte3 & lignes(i) & vbCrLf
        End If
    Next i

    ' Affecter les index des diapositives de sommaire
    slideIndex1 = 9
    slideIndex2 = slideIndex1 + 1
    slideIndex3 = slideIndex2 + 1

    ' Ajouter le premier sommaire
    Set sommaireDiapo1 = presentation.Slides(slideIndex1)
    Set pptShape = sommaireDiapo1.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
    pptShape.TextFrame.TextRange.text = sommaireTexte1
    pptShape.TextFrame.TextRange.Font.Name = "Arial"
    pptShape.TextFrame.TextRange.Font.Size = 12
    pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0)
    ApplyHyperlinks titleSet, pptShape

    ' Ajouter le deuxième sommaire si nécessaire
    If sommaireTexte2 <> "" Then
        Set sommaireDiapo2 = presentation.Slides(slideIndex2)
        Set pptShape = sommaireDiapo2.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte2
        pptShape.TextFrame.TextRange.Font.Name = "Arial"
        pptShape.TextFrame.TextRange.Font.Size = 12
        pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0)
        ApplyHyperlinks titleSet, pptShape
    End If

    ' Ajouter le troisième sommaire si nécessaire
    If sommaireTexte3 <> "" Then
        Set sommaireDiapo3 = presentation.Slides(slideIndex3)
        Set pptShape = sommaireDiapo3.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte3
        pptShape.TextFrame.TextRange.Font.Name = "Arial"
        pptShape.TextFrame.TextRange.Font.Size = 12
        pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0)
        ApplyHyperlinks titleSet, pptShape
    End If

    Exit Function

ErrorHandler:
    MsgBox "Une erreur s'est produite : " & ERR.description, vbExclamation

End Function

' Fonction pour vérifier si la diapositive est un cartouche
Function IsCartouche2(diapo As Object) As Boolean
    Dim Titre As String
    On Error Resume Next
    Titre = diapo.Shapes.Placeholders(1).TextFrame.TextRange.text
    On Error GoTo 0
    
    ' Condition pour identifier les diapositives "cartouche"
    If InStr(1, Titre, "Back UP", vbTextCompare) > 0 Or InStr(1, Titre, "Back UP slides for Top area of improvement", vbTextCompare) > 0 Then
        IsCartouche = True
    Else
        IsCartouche = False
    End If
End Function

Sub ApplyHyperlinks2(titleSet As Object, pptShape As Object)
    Dim Titre As Variant
    Dim startPos As Integer, endPos As Integer

    For Each Titre In titleSet.keys
        startPos = InStr(pptShape.TextFrame.TextRange.text, Titre)
        endPos = startPos + Len(Titre) - 1

        If startPos > 0 Then
            ' Appliquer un lien hypertexte pour chaque titre
            With pptShape.TextFrame.TextRange.Characters(startPos, Len(Titre))
                .Font.Underline = msoTrue
                .Font.color.RGB = RGB(0, 0, 0)
                .ActionSettings(ppMouseClick).Hyperlink.SubAddress = CStr(titleSet(Titre))
            End With
        End If
    Next Titre
End Sub

Function InsererSommaireDynamiqueDeuxNiveauxV9(objPresentation As Object)
    Dim presentation As Object
    Dim sommaireDiapo1 As Object
    Dim sommaireDiapo2 As Object
    Dim sommaireDiapo3 As Object
    Dim diapo As Object
    Dim sommaireTexte As String
    Dim sommaireTexte1 As String
    Dim sommaireTexte2 As String
    Dim sommaireTexte3 As String
    Dim Titre As Variant
    Dim pptShape As Object
    Dim slideIndex1 As Integer
    Dim slideIndex2 As Integer
    Dim slideIndex3 As Integer
    Dim maxLignesParDiapo As Integer
    Dim lignes() As String
    Dim ligneCount As Integer
    Dim i As Integer
    Dim titleSet As Object ' To track added titles
    Dim currentLine As Integer ' For hyperlink range
    Dim TitreObj As Object

    ' Error handling setup
    On Error GoTo ErrorHandler

    ' Reference to the active presentation
    Set presentation = objPresentation

    ' Maximum number of lines per slide
    maxLignesParDiapo = 20 ' Adjust as needed

    ' Create a dictionary to track added titles
    Set titleSet = CreateObject("Scripting.Dictionary")

    ' Start with fixed entries
    sommaireTexte = "1. Test information" & vbCrLf & "2. oDriv results" & vbCrLf

    ' Add slide titles to the summary dynamically
    For i = 10 To presentation.Slides.Count
        Set diapo = presentation.Slides(i)
        If diapo.Shapes.Placeholders.Count > 0 Then
            On Error Resume Next
            Titre = diapo.Shapes.Placeholders(1).TextFrame.TextRange.text
            On Error GoTo 0
            ' Check if the title is not already added
            If Len(Titre) > 0 And Not titleSet.Exists(Titre) Then
                sommaireTexte = sommaireTexte & vbTab & " " & Titre & vbCrLf
                titleSet.Add Titre, i ' Store slide index with title
            End If
        End If
    Next i

    ' Split the summary text into lines
    lignes = Split(sommaireTexte, vbCrLf)
    ligneCount = UBound(lignes) + 1

    ' Separate the summary into three texts
    sommaireTexte1 = ""
    sommaireTexte2 = ""
    sommaireTexte3 = ""

    For i = 0 To ligneCount - 1
        If i < maxLignesParDiapo Then
            sommaireTexte1 = sommaireTexte1 & lignes(i) & vbCrLf
        ElseIf i < maxLignesParDiapo * 2 Then
            sommaireTexte2 = sommaireTexte2 & lignes(i) & vbCrLf
        Else
            sommaireTexte3 = sommaireTexte3 & lignes(i) & vbCrLf
        End If
    Next i

    ' Assign slide indices flexibly
    slideIndex1 = 9 ' Adjust starting slide index dynamically if needed
    slideIndex2 = slideIndex1 + 1
    slideIndex3 = slideIndex2 + 1

    ' Add the first slide for the summary
    Set sommaireDiapo1 = presentation.Slides(slideIndex1)
    Set pptShape = sommaireDiapo1.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
    pptShape.TextFrame.TextRange.text = sommaireTexte1
    pptShape.TextFrame.TextRange.Font.Name = "Arial"
    pptShape.TextFrame.TextRange.Font.Size = 12
    pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0)

    ' Apply hyperlinks to each title in sommaireDiapo1
    ApplyHyperlinks titleSet, pptShape

    ' Add the second slide for the summary, if necessary
    If sommaireTexte2 <> "" Then
        Set sommaireDiapo2 = presentation.Slides(slideIndex2)
        Set pptShape = sommaireDiapo2.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte2
        pptShape.TextFrame.TextRange.Font.Name = "Arial"
        pptShape.TextFrame.TextRange.Font.Size = 12
        pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0)
        
        ' Apply hyperlinks to each title in sommaireDiapo2
        ApplyHyperlinks titleSet, pptShape
    End If

    ' Add the third slide for the summary, if necessary
    If sommaireTexte3 <> "" Then
        Set sommaireDiapo3 = presentation.Slides(slideIndex3)
        Set pptShape = sommaireDiapo3.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte3
        pptShape.TextFrame.TextRange.Font.Name = "Arial"
        pptShape.TextFrame.TextRange.Font.Size = 12
        pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0)
        
        ' Apply hyperlinks to each title in sommaireDiapo3
        ApplyHyperlinks titleSet, pptShape
    End If

    Exit Function

ErrorHandler:
    MsgBox "An error occurred: " & ERR.description, vbExclamation

End Function

Sub ApplyHyperlinks1(titleSet As Object, pptShape As Object)
    Dim Titre As Variant
    Dim startPos As Integer, endPos As Integer

    For Each Titre In titleSet.keys
        startPos = InStr(pptShape.TextFrame.TextRange.text, Titre)
        endPos = startPos + Len(Titre) - 1

        If startPos > 0 Then
            ' Apply hyperlink to each title
            With pptShape.TextFrame.TextRange.Characters(startPos, Len(Titre))
                .Font.Underline = msoTrue
                .Font.color.RGB = RGB(0, 0, 0)
                .ActionSettings(ppMouseClick).Hyperlink.SubAddress = CStr(titleSet(Titre))
            End With
        End If
    Next Titre
End Sub



Function InsererSommaireDynamiqueDeuxNiveauxV6(objPresentation As Object)
    Dim presentation As Object
    Dim sommaireDiapo1 As Object
    Dim sommaireDiapo2 As Object
    Dim sommaireDiapo3 As Object
    Dim diapo As Object
    Dim sommaireTexte As String
    Dim sommaireTexte1 As String
    Dim sommaireTexte2 As String
    Dim sommaireTexte3 As String
    Dim Titre As Variant
    Dim pptShape As Object
    Dim slideIndex1 As Integer
    Dim slideIndex2 As Integer
    Dim slideIndex3 As Integer
    Dim maxLignesParDiapo As Integer
    Dim lignes() As String
    Dim ligneCount As Integer
    Dim i As Integer
    Dim titleSet As Object

    ' Référence de la présentation active
    Set presentation = objPresentation

    ' Nombre maximum de lignes par diapo
    maxLignesParDiapo = 20 ' Ajustez si nécessaire

    ' Création d'un dictionnaire pour suivre les titres ajoutés
    Set titleSet = CreateObject("Scripting.Dictionary")

    ' Ajouter les entrées fixes
    sommaireTexte = "1. Test information" & vbCrLf & "2. oDriv results" & vbCrLf

    ' Ajout dynamique des titres des diapositives au sommaire
    For i = 10 To presentation.Slides.Count
        Set diapo = presentation.Slides(i)
        If diapo.Shapes.Placeholders.Count > 0 Then
            On Error Resume Next
            Titre = diapo.Shapes.Placeholders(1).TextFrame.TextRange.text
            On Error GoTo 0
            ' Ajouter le titre s'il n'a pas encore été ajouté
            If Len(Titre) > 0 And Not titleSet.Exists(Titre) Then
                sommaireTexte = sommaireTexte & vbTab & " " & Titre & vbCrLf
                titleSet.Add Titre, i ' Stocker l'index de la diapo avec le titre
            End If
        End If
    Next i

    ' Séparer le texte du sommaire en lignes
    lignes = Split(sommaireTexte, vbCrLf)
    ligneCount = UBound(lignes) + 1

    ' Séparer le sommaire en trois parties
    sommaireTexte1 = ""
    sommaireTexte2 = ""
    sommaireTexte3 = ""

    For i = 0 To ligneCount - 1
        If i < maxLignesParDiapo Then
            sommaireTexte1 = sommaireTexte1 & lignes(i) & vbCrLf
        ElseIf i < maxLignesParDiapo * 2 Then
            sommaireTexte2 = sommaireTexte2 & lignes(i) & vbCrLf
        Else
            sommaireTexte3 = sommaireTexte3 & lignes(i) & vbCrLf
        End If
    Next i

    ' Définir les indices des diapositives
    slideIndex1 = 9
    slideIndex2 = slideIndex1 + 1
    slideIndex3 = slideIndex2 + 1

    ' Ajouter la première diapositive pour le sommaire
    Set sommaireDiapo1 = presentation.Slides(slideIndex1)
    Set pptShape = sommaireDiapo1.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
    pptShape.TextFrame.TextRange.text = sommaireTexte1
    pptShape.TextFrame.TextRange.Font.Name = "Arial"
    pptShape.TextFrame.TextRange.Font.Size = 12
    pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0)

    ' Ajouter les liens hypertexte à chaque titre dans le sommaire de la première diapositive
    Call AjouterHyperliens(pptShape, titleSet)

    ' Ajouter la deuxième diapositive pour le sommaire, si nécessaire
    If sommaireTexte2 <> "" Then
        Set sommaireDiapo2 = presentation.Slides(slideIndex2)
        Set pptShape = sommaireDiapo2.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte2
        pptShape.TextFrame.TextRange.Font.Name = "Arial"
        pptShape.TextFrame.TextRange.Font.Size = 12
        pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0)

        ' Ajouter les liens hypertexte à chaque titre dans le sommaire de la deuxième diapositive
        Call AjouterHyperliens(pptShape, titleSet)
    End If

    ' Ajouter la troisième diapositive pour le sommaire, si nécessaire
    If sommaireTexte3 <> "" Then
        Set sommaireDiapo3 = presentation.Slides(slideIndex3)
        Set pptShape = sommaireDiapo3.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte3
        pptShape.TextFrame.TextRange.Font.Name = "Arial"
        pptShape.TextFrame.TextRange.Font.Size = 12
        pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0)

        ' Ajouter les liens hypertexte à chaque titre dans le sommaire de la troisième diapositive
        Call AjouterHyperliens(pptShape, titleSet)
    End If
End Function

Sub AjouterHyperliens(pptShape As Object, titleSet As Object)
    Dim Titre As Variant
    Dim startPos As Integer
    Dim slideNum As Integer

    ' Ajouter les liens hypertexte à chaque titre dans le sommaire
    For Each Titre In titleSet.keys
        slideNum = titleSet(Titre) ' Numéro de diapositive associé
        startPos = InStr(1, pptShape.TextFrame.TextRange.text, Titre)

        If startPos > 0 Then
            With pptShape.TextFrame.TextRange.Characters(startPos, Len(Titre))
                .Font.Underline = msoTrue
                .Font.color.RGB = RGB(0, 0, 0)
                .ActionSettings(ppMouseClick).Hyperlink.Address = "" ' Clear external link
                .ActionSettings(ppMouseClick).Hyperlink.SubAddress = slideNum & ",0" ' Définir l'adresse de l'hyperlien
                .ActionSettings(ppMouseClick).Hyperlink.TextToDisplay = Titre ' Texte à afficher pour l'hyperlien
            End With
        End If
    Next Titre
End Sub


Function InsererSommaireDynamiqueDeuxNiveauxV7(objPresentation As Object)
    Dim presentation As Object
    Dim sommaireDiapo1 As Object
    Dim sommaireDiapo2 As Object
    Dim sommaireDiapo3 As Object
    Dim diapo As Object
    Dim sommaireTexte As String
    Dim sommaireTexte1 As String
    Dim sommaireTexte2 As String
    Dim sommaireTexte3 As String
    Dim Titre As Variant ' Déclaré comme Variant
    Dim pptShape As Object
    Dim slideIndex1 As Integer
    Dim slideIndex2 As Integer
    Dim slideIndex3 As Integer
    Dim maxLignesParDiapo As Integer
    Dim lignes() As String
    Dim ligneCount As Integer
    Dim i As Integer
    Dim titleSet As Object
       Dim startPos As Integer, endPos As Integer, slideNum As Integer
    ' Référence de la présentation active
    Set presentation = objPresentation

    ' Nombre maximum de lignes par diapo
    maxLignesParDiapo = 20 ' Ajustez si nécessaire

    ' Création d'un dictionnaire pour suivre les titres ajoutés
    Set titleSet = CreateObject("Scripting.Dictionary")

    ' Ajouter les entrées fixes
    sommaireTexte = "1. Test information" & vbCrLf & "2. oDriv results" & vbCrLf

    ' Ajout dynamique des titres des diapositives au sommaire
    For i = 10 To presentation.Slides.Count
        Set diapo = presentation.Slides(i)
        If diapo.Shapes.Placeholders.Count > 0 Then
            On Error Resume Next
            Titre = diapo.Shapes.Placeholders(1).TextFrame.TextRange.text
            On Error GoTo 0
            ' Ajouter le titre s'il n'a pas encore été ajouté
            If Len(Titre) > 0 And Not titleSet.Exists(Titre) Then
                sommaireTexte = sommaireTexte & vbTab & " " & Titre & vbCrLf
                titleSet.Add Titre, i ' Stocker l'index de la diapo avec le titre
            End If
        End If
    Next i

    ' Séparer le texte du sommaire en lignes
    lignes = Split(sommaireTexte, vbCrLf)
    ligneCount = UBound(lignes) + 1

    ' Séparer le sommaire en trois parties
    sommaireTexte1 = ""
    sommaireTexte2 = ""
    sommaireTexte3 = ""

    For i = 0 To ligneCount - 1
        If i < maxLignesParDiapo Then
            sommaireTexte1 = sommaireTexte1 & lignes(i) & vbCrLf
        ElseIf i < maxLignesParDiapo * 2 Then
            sommaireTexte2 = sommaireTexte2 & lignes(i) & vbCrLf
        Else
            sommaireTexte3 = sommaireTexte3 & lignes(i) & vbCrLf
        End If
    Next i

    ' Définir les indices des diapositives
    slideIndex1 = 9
    slideIndex2 = slideIndex1 + 1
    slideIndex3 = slideIndex2 + 1

    ' Ajouter la première diapositive pour le sommaire
    Set sommaireDiapo1 = presentation.Slides(slideIndex1)
    Set pptShape = sommaireDiapo1.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
    pptShape.TextFrame.TextRange.text = sommaireTexte1
    pptShape.TextFrame.TextRange.Font.Name = "Encode Sans (Corps)"
    pptShape.TextFrame.TextRange.Font.Size = 15
    pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0)

    ' Ajouter les liens hypertexte à chaque titre dans le sommaire de la première diapositive
    For Each Titre In titleSet.keys
     
        slideNum = titleSet(Titre) ' Numéro de diapositive associé
        startPos = InStr(1, pptShape.TextFrame.TextRange.text, Titre)

        If startPos > 0 Then
            endPos = startPos + Len(Titre) - 1
            With pptShape.TextFrame.TextRange.Characters(startPos, Len(Titre))
                .Font.Underline = msoTrue
                .Font.color.RGB = RGB(0, 0, 0)
                .ActionSettings(ppMouseClick).Hyperlink.Address = "" ' Clear external link
                .ActionSettings(ppMouseClick).Hyperlink.SubAddress = slideNum & ",0"
            End With
        End If
    Next Titre

    ' Ajouter la deuxième diapositive pour le sommaire, si nécessaire
    If sommaireTexte2 <> "" Then
        Set sommaireDiapo2 = presentation.Slides(slideIndex2)
        Set pptShape = sommaireDiapo2.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte2
        pptShape.TextFrame.TextRange.Font.Name = "Arial"
        pptShape.TextFrame.TextRange.Font.Size = 15
        pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0)

        ' Ajouter les liens hypertexte à chaque titre dans le sommaire de la deuxième diapositive
        For Each Titre In titleSet.keys
            'Dim startPos As Integer, endPos As Integer, slideNum As Integer
            slideNum = titleSet(Titre) ' Numéro de diapositive associé
            startPos = InStr(1, pptShape.TextFrame.TextRange.text, Titre)

            If startPos > 0 Then
                endPos = startPos + Len(Titre) - 1
                With pptShape.TextFrame.TextRange.Characters(startPos, Len(Titre))
                    .Font.Underline = msoTrue
                    .Font.color.RGB = RGB(0, 0, 0)
                     .Font.Name = "Encode Sans (Corps)"
                     .Font.Size = 15
                    .ActionSettings(ppMouseClick).Hyperlink.Address = "" ' Clear external link
                    .ActionSettings(ppMouseClick).Hyperlink.SubAddress = slideNum & ",0"
                End With
            End If
        Next Titre
    End If

    ' Ajouter la troisième diapositive pour le sommaire, si nécessaire
    If sommaireTexte3 <> "" Then
        Set sommaireDiapo3 = presentation.Slides(slideIndex3)
        Set pptShape = sommaireDiapo3.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte3
        pptShape.TextFrame.TextRange.Font.Name = "Encode Sans (Corps)"
        pptShape.TextFrame.TextRange.Font.Size = 15
        pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0)

        ' Ajouter les liens hypertexte à chaque titre dans le sommaire de la troisième diapositive
        For Each Titre In titleSet.keys
            'Dim startPos As Integer, endPos As Integer, slideNum As Integer
            slideNum = titleSet(Titre) ' Numéro de diapositive associé
            startPos = InStr(1, pptShape.TextFrame.TextRange.text, Titre)

            If startPos > 0 Then
                endPos = startPos + Len(Titre) - 1
                With pptShape.TextFrame.TextRange.Characters(startPos, Len(Titre))
                    .Font.Underline = msoTrue
                    .Font.color.RGB = RGB(0, 0, 0)
                    .ActionSettings(ppMouseClick).Hyperlink.Address = "" ' Clear external link
                    .ActionSettings(ppMouseClick).Hyperlink.SubAddress = slideNum & ",0"
                End With
            End If
        Next Titre
    End If

End Function

Function InsererSommaireDynamiqueDeuxNiveauxV8(objPresentation As Object)
    Dim presentation As Object
    Dim sommaireDiapo1 As Object
    Dim sommaireDiapo2 As Object
    Dim sommaireDiapo3 As Object
    Dim diapo As Object
    Dim sommaireTexte As String
    Dim sommaireTexte1 As String
    Dim sommaireTexte2 As String
    Dim sommaireTexte3 As String
    Dim Titre As Variant
    Dim pptShape As Object
    Dim slideIndex1 As Integer
    Dim slideIndex2 As Integer
    Dim slideIndex3 As Integer
    Dim maxLignesParDiapo As Integer
    Dim lignes() As String
    Dim ligneCount As Integer
    Dim i As Integer
    Dim titleSet As Object ' To track added titles
    Dim currentLine As Integer ' For hyperlink range
    Dim TitreObj As Object
    
    ' Error handling setup
    On Error GoTo ErrorHandler

    ' Reference to the active presentation
    Set presentation = objPresentation

    ' Maximum number of lines per slide
    maxLignesParDiapo = 20 ' Adjust as needed

    ' Create a dictionary to track added titles
    Set titleSet = CreateObject("Scripting.Dictionary")

    ' Start with fixed entries
    sommaireTexte = "1. Test information" & vbCrLf & "2. oDriv results" & vbCrLf

    ' Add slide titles to the summary dynamically
    For i = 10 To presentation.Slides.Count
        Set diapo = presentation.Slides(i)
        If diapo.Shapes.Placeholders.Count > 0 Then
            On Error Resume Next
            Titre = diapo.Shapes.Placeholders(1).TextFrame.TextRange.text
            On Error GoTo 0
            ' Check if the title is not already added
            If Len(Titre) > 0 And Not titleSet.Exists(Titre) Then
                sommaireTexte = sommaireTexte & vbTab & " " & Titre & vbCrLf
                titleSet.Add Titre, i ' Store slide index with title
            End If
        End If
    Next i

    ' Split the summary text into lines
    lignes = Split(sommaireTexte, vbCrLf)
    ligneCount = UBound(lignes) + 1

    ' Separate the summary into three texts
    sommaireTexte1 = ""
    sommaireTexte2 = ""
    sommaireTexte3 = ""

    For i = 0 To ligneCount - 1
        If i < maxLignesParDiapo Then
            sommaireTexte1 = sommaireTexte1 & lignes(i) & vbCrLf
        ElseIf i < maxLignesParDiapo * 2 Then
            sommaireTexte2 = sommaireTexte2 & lignes(i) & vbCrLf
        Else
            sommaireTexte3 = sommaireTexte3 & lignes(i) & vbCrLf
        End If
    Next i

    ' Assign slide indices flexibly
    slideIndex1 = 9 ' Adjust starting slide index dynamically if needed
    slideIndex2 = slideIndex1 + 1
    slideIndex3 = slideIndex2 + 1

    ' Add the first slide for the summary
    Set sommaireDiapo1 = presentation.Slides(slideIndex1)
    Set pptShape = sommaireDiapo1.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
    pptShape.TextFrame.TextRange.text = sommaireTexte1
    pptShape.TextFrame.TextRange.Font.Name = "Arial"
    pptShape.TextFrame.TextRange.Font.Size = 12
    pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0) ' Set color to black

    ' Apply hyperlinks to each title in the summary
    currentLine = 1
    For Each Titre In titleSet.keys
        Dim startPos As Integer, endPos As Integer
        startPos = InStr(pptShape.TextFrame.TextRange.text, Titre) ' Find start position of title
        endPos = startPos + Len(Titre) - 1 ' Calculate end position of title

        If startPos > 0 Then
            ' Apply hyperlink to each title
            With pptShape.TextFrame.TextRange.Characters(startPos, Len(Titre))
                .Font.Underline = msoTrue
                .Font.color.RGB = RGB(0, 0, 0) ' Set color to black
                .ActionSettings(ppMouseClick).Hyperlink.SubAddress = CStr(titleSet(Titre)) ' Link to specific slide index
            End With
        End If
        currentLine = currentLine + 1
    Next Titre

    ' Add the second slide for the summary, if necessary
    If sommaireTexte2 <> "" Then
        Set sommaireDiapo2 = presentation.Slides(slideIndex2)
        Set pptShape = sommaireDiapo2.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte2
        pptShape.TextFrame.TextRange.Font.Name = "Arial"
        pptShape.TextFrame.TextRange.Font.Size = 12
        pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0) ' Set color to black
    End If

    ' Add the third slide for the summary, if necessary
    If sommaireTexte3 <> "" Then
        Set sommaireDiapo3 = presentation.Slides(slideIndex3)
        Set pptShape = sommaireDiapo3.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte3
        pptShape.TextFrame.TextRange.Font.Name = "Arial"
        pptShape.TextFrame.TextRange.Font.Size = 12
        pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0) ' Set color to black
    End If

    Exit Function

ErrorHandler:
    MsgBox "An error occurred: " & ERR.description, vbExclamation

End Function



Function InsererSommaireDynamiqueDeuxNiveauxV5(objPresentation As Object)
    Dim presentation As Object
    Dim sommaireDiapo1 As Object
    Dim sommaireDiapo2 As Object
    Dim sommaireDiapo3 As Object
    Dim diapo As Object
    Dim sommaireTexte As String
    Dim sommaireTexte1 As String
    Dim sommaireTexte2 As String
    Dim sommaireTexte3 As String
    Dim Titre As String
    Dim pptShape As Object
    Dim slideIndex1 As Integer
    Dim slideIndex2 As Integer
    Dim slideIndex3 As Integer
    Dim maxLignesParDiapo As Integer
    Dim lignes() As String
    Dim ligneCount As Integer
    Dim i As Integer
    Dim titleSet As Object ' To track added titles

    ' Error handling setup
    On Error GoTo ErrorHandler

    ' Reference to the active presentation
    Set presentation = objPresentation

    ' Maximum number of lines per slide
    maxLignesParDiapo = 20 ' Adjust as needed

    ' Create a dictionary to track added titles
    Set titleSet = CreateObject("Scripting.Dictionary")

    ' Start with fixed entries
    sommaireTexte = "1. Test information" & vbCrLf & "2. oDriv results" & vbCrLf

    ' Add slide titles to the summary dynamically
    For i = 10 To presentation.Slides.Count
        Set diapo = presentation.Slides(i)
        If diapo.Shapes.Placeholders.Count > 0 Then
            On Error Resume Next
            Titre = diapo.Shapes.Placeholders(1).TextFrame.TextRange.text
            On Error GoTo 0
            ' Check if the title is not already added
            If Len(Titre) > 0 And Not titleSet.Exists(Titre) Then
                sommaireTexte = sommaireTexte & vbTab & " " & Titre & vbCrLf
                titleSet.Add Titre, Nothing ' Add title to the set
            End If
        End If
    Next i

    ' Split the summary text into lines
    lignes = Split(sommaireTexte, vbCrLf)
    ligneCount = UBound(lignes) + 1

    ' Separate the summary into three texts
    sommaireTexte1 = ""
    sommaireTexte2 = ""
    sommaireTexte3 = ""

    For i = 0 To ligneCount - 1
        If i < maxLignesParDiapo Then
            sommaireTexte1 = sommaireTexte1 & lignes(i) & vbCrLf
        ElseIf i < maxLignesParDiapo * 2 Then
            sommaireTexte2 = sommaireTexte2 & lignes(i) & vbCrLf
        Else
            sommaireTexte3 = sommaireTexte3 & lignes(i) & vbCrLf
        End If
    Next i

    ' Assign slide indices flexibly
    slideIndex1 = 9 ' Adjust starting slide index dynamically if needed
    slideIndex2 = slideIndex1 + 1
    slideIndex3 = slideIndex2 + 1

    ' Add the first slide for the summary
    Set sommaireDiapo1 = presentation.Slides(slideIndex1)
    Set pptShape = sommaireDiapo1.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
    pptShape.TextFrame.TextRange.text = sommaireTexte1
    pptShape.TextFrame.TextRange.Font.Name = "Arial"
    pptShape.TextFrame.TextRange.Font.Size = 12
    pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0) ' Set color to black

    ' Underline and add internal links for "1. Test information" and "2. oDriv results"
    With pptShape.TextFrame.TextRange.Characters(1, 19)
        .Font.Underline = msoTrue
        .Font.color.RGB = RGB(0, 0, 0) ' Set color to black
        .ActionSettings(ppMouseClick).Hyperlink.SubAddress = "10" ' Navigate to Slide 2
    End With
    With pptShape.TextFrame.TextRange.Characters(21, 14)
        .Font.Underline = msoTrue
        .Font.color.RGB = RGB(0, 0, 0) ' Set color to black
        .ActionSettings(ppMouseClick).Hyperlink.SubAddress = "11" ' Navigate to Slide 3
    End With

    ' Add the second slide for the summary, if necessary
    If sommaireTexte2 <> "" Then
        Set sommaireDiapo2 = presentation.Slides(slideIndex2)
        Set pptShape = sommaireDiapo2.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte2
        pptShape.TextFrame.TextRange.Font.Name = "Arial"
        pptShape.TextFrame.TextRange.Font.Size = 12
        pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0) ' Set color to black
    End If

    ' Add the third slide for the summary, if necessary
    If sommaireTexte3 <> "" Then
        Set sommaireDiapo3 = presentation.Slides(slideIndex3)
        Set pptShape = sommaireDiapo3.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte3
        pptShape.TextFrame.TextRange.Font.Name = "Arial"
        pptShape.TextFrame.TextRange.Font.Size = 12
        pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0) ' Set color to black
    End If

    Exit Function

ErrorHandler:
    MsgBox "An error occurred: " & ERR.description, vbExclamation

End Function


Function InsererSommaireDynamiqueDeuxNiveauxV4(objPresentation As Object)
    Dim presentation As Object
    Dim sommaireDiapo1 As Object
    Dim sommaireDiapo2 As Object
    Dim sommaireDiapo3 As Object
    Dim diapo As Object
    Dim sommaireTexte As String
    Dim sommaireTexte1 As String
    Dim sommaireTexte2 As String
    Dim sommaireTexte3 As String
    Dim Titre As String
    Dim pptShape As Object
    Dim slideIndex1 As Integer
    Dim slideIndex2 As Integer
    Dim slideIndex3 As Integer
    Dim maxLignesParDiapo As Integer
    Dim lignes() As String
    Dim ligneCount As Integer
    Dim i As Integer
    Dim titleSet As Object ' To track added titles

    ' Error handling setup
    On Error GoTo ErrorHandler

    ' Reference to the active presentation
    Set presentation = objPresentation

    ' Maximum number of lines per slide
    maxLignesParDiapo = 20 ' Adjust as needed

    ' Create a dictionary to track added titles
    Set titleSet = CreateObject("Scripting.Dictionary")

    ' Start with fixed entries
    sommaireTexte = "1. Test information" & vbCrLf & "2. oDriv results" & vbCrLf

    ' Add slide titles to the summary dynamically
    For i = 10 To presentation.Slides.Count
        Set diapo = presentation.Slides(i)
        If diapo.Shapes.Placeholders.Count > 0 Then
            On Error Resume Next
            Titre = diapo.Shapes.Placeholders(1).TextFrame.TextRange.text
            On Error GoTo 0
            ' Check if the title is not already added
            If Len(Titre) > 0 And Not titleSet.Exists(Titre) Then
                sommaireTexte = sommaireTexte & vbTab & " " & Titre & vbCrLf
                titleSet.Add Titre, Nothing ' Add title to the set
            End If
        End If
    Next i

    ' Split the summary text into lines
    lignes = Split(sommaireTexte, vbCrLf)
    ligneCount = UBound(lignes) + 1

    ' Separate the summary into three texts
    sommaireTexte1 = ""
    sommaireTexte2 = ""
    sommaireTexte3 = ""

    For i = 0 To ligneCount - 1
        If i < maxLignesParDiapo Then
            sommaireTexte1 = sommaireTexte1 & lignes(i) & vbCrLf
        ElseIf i < maxLignesParDiapo * 2 Then
            sommaireTexte2 = sommaireTexte2 & lignes(i) & vbCrLf
        Else
            sommaireTexte3 = sommaireTexte3 & lignes(i) & vbCrLf
        End If
    Next i

    ' Assign slide indices flexibly
    slideIndex1 = 9 ' Adjust starting slide index dynamically if needed
    slideIndex2 = slideIndex1 + 1
    slideIndex3 = slideIndex2 + 1

    ' Add the first slide for the summary
    Set sommaireDiapo1 = presentation.Slides(slideIndex1)
    Set pptShape = sommaireDiapo1.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
    pptShape.TextFrame.TextRange.text = sommaireTexte1
    pptShape.TextFrame.TextRange.Font.Name = "Arial"
    pptShape.TextFrame.TextRange.Font.Size = 12
    pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0) ' Set color to black

    ' Underline and add hyperlinks for "1. Test information" and "2. oDriv results"
    With pptShape.TextFrame.TextRange.Characters(1, 19)
        .Font.Underline = msoTrue
        .Font.color.RGB = RGB(0, 0, 0) ' Set color to black
        .ActionSettings(ppMouseClick).Hyperlink.Address = "https://example.com/test-info" ' Replace with the actual URL
    End With
    With pptShape.TextFrame.TextRange.Characters(21, 14)
        .Font.Underline = msoTrue
        .Font.color.RGB = RGB(0, 0, 0) ' Set color to black
        .ActionSettings(ppMouseClick).Hyperlink.Address = "https://example.com/odriv-results" ' Replace with the actual URL
    End With

    ' Add the second slide for the summary, if necessary
    If sommaireTexte2 <> "" Then
        Set sommaireDiapo2 = presentation.Slides(slideIndex2)
        Set pptShape = sommaireDiapo2.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte2
        pptShape.TextFrame.TextRange.Font.Name = "Arial"
        pptShape.TextFrame.TextRange.Font.Size = 12
        pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0) ' Set color to black
    End If

    ' Add the third slide for the summary, if necessary
    If sommaireTexte3 <> "" Then
        Set sommaireDiapo3 = presentation.Slides(slideIndex3)
        Set pptShape = sommaireDiapo3.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte3
        pptShape.TextFrame.TextRange.Font.Name = "Arial"
        pptShape.TextFrame.TextRange.Font.Size = 12
        pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 0) ' Set color to black
    End If

    Exit Function

ErrorHandler:
    MsgBox "An error occurred: " & ERR.description, vbExclamation

End Function

Function InsererSommaireDynamiqueDeuxNiveauxV3(objPresentation As Object)
    Dim presentation As Object
    Dim sommaireDiapo1 As Object
    Dim sommaireDiapo2 As Object
    Dim sommaireDiapo3 As Object
    Dim diapo As Object
    Dim sommaireTexte As String
    Dim sommaireTexte1 As String
    Dim sommaireTexte2 As String
    Dim sommaireTexte3 As String
    Dim Titre As String
    Dim pptShape As Object
    Dim slideIndex1 As Integer
    Dim slideIndex2 As Integer
    Dim slideIndex3 As Integer
    Dim maxLignesParDiapo As Integer
    Dim lignes() As String
    Dim ligneCount As Integer
    Dim i As Integer
    Dim sectionNum As Integer
    Dim subSectionNum As Integer
    Dim titleSet As Object ' To track added titles

    ' Error handling setup
    On Error GoTo ErrorHandler

    ' Reference to the active presentation
    Set presentation = objPresentation

    ' Maximum number of lines per slide
    maxLignesParDiapo = 20 ' Adjust as needed

    ' Initialize section and subsection numbers
    sectionNum = 1
    subSectionNum = 1

    ' Create a dictionary to track added titles
    Set titleSet = CreateObject("Scripting.Dictionary")

    ' Start with fixed entries
    sommaireTexte = "1. Test information" & vbCrLf & "2. oDriv results" & vbCrLf

    ' Add slide titles to the summary dynamically
    For i = 9 To presentation.Slides.Count
        Set diapo = presentation.Slides(i)
        If diapo.Shapes.Placeholders.Count > 0 Then
            On Error Resume Next
            Titre = diapo.Shapes.Placeholders(1).TextFrame.TextRange.text
            On Error GoTo 0
            ' Check if the title is not already added
            If Len(Titre) > 0 And Not titleSet.Exists(Titre) Then
                sommaireTexte = sommaireTexte & vbTab & " " & Titre & vbCrLf
                titleSet.Add Titre, Nothing ' Add title to the set
                subSectionNum = subSectionNum + 1
            End If
        End If
    Next i

    ' Split the summary text into lines
    lignes = Split(sommaireTexte, vbCrLf)
    ligneCount = UBound(lignes) + 1

    ' Separate the summary into three texts
    sommaireTexte1 = ""
    sommaireTexte2 = ""
    sommaireTexte3 = ""

    For i = 0 To ligneCount - 1
        If i < maxLignesParDiapo Then
            sommaireTexte1 = sommaireTexte1 & lignes(i) & vbCrLf
        ElseIf i < maxLignesParDiapo * 2 Then
            sommaireTexte2 = sommaireTexte2 & lignes(i) & vbCrLf
        Else
            sommaireTexte3 = sommaireTexte3 & lignes(i) & vbCrLf
        End If
    Next i

    ' Assign slide indices flexibly
    slideIndex1 = 9 ' Adjust starting slide index dynamically if needed
    slideIndex2 = slideIndex1 + 1
    slideIndex3 = slideIndex2 + 1

    ' Add the first slide for the summary
    Set sommaireDiapo1 = presentation.Slides(slideIndex1)
    Set pptShape = sommaireDiapo1.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
    pptShape.TextFrame.TextRange.text = sommaireTexte1
    pptShape.TextFrame.TextRange.Font.Name = "Arial"
    pptShape.TextFrame.TextRange.Font.Size = 12
    pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 255) ' Set color to blue
    
     ' Underline "1. Test information" and "2. oDriv results"
    pptShape.TextFrame.TextRange.Characters(1, 19).Font.Underline = msoTrue
    pptShape.TextFrame.TextRange.Characters(21, 14).Font.Underline = msoTrue

    ' Add the second slide for the summary, if necessary
    If sommaireTexte2 <> "" Then
        Set sommaireDiapo2 = presentation.Slides(slideIndex2)
        Set pptShape = sommaireDiapo2.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte2
        pptShape.TextFrame.TextRange.Font.Name = "Arial"
        pptShape.TextFrame.TextRange.Font.Size = 12
        pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 255) ' Set color to blue
    End If

    ' Add the third slide for the summary, if necessary
    If sommaireTexte3 <> "" Then
        Set sommaireDiapo3 = presentation.Slides.Add(slideIndex3, ppLayoutText)
        Set pptShape = sommaireDiapo3.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte3
        pptShape.TextFrame.TextRange.Font.Name = "Arial"
        pptShape.TextFrame.TextRange.Font.Size = 12
        pptShape.TextFrame.TextRange.Font.color.RGB = RGB(0, 0, 255) ' Set color to blue
    End If

    Exit Function

ErrorHandler:
    MsgBox "An error occurred: " & ERR.description, vbExclamation

End Function

Function InsererSommaireDynamiqueDeuxNiveauxV2(objPresentation As Object)
    Dim presentation As Object
    Dim sommaireDiapo1 As Object
    Dim sommaireDiapo2 As Object
    Dim sommaireDiapo3 As Object
    Dim diapo As Object
    Dim sommaireTexte As String
    Dim sommaireTexte1 As String
    Dim sommaireTexte2 As String
    Dim sommaireTexte3 As String
    Dim Titre As String
    Dim pptShape As Object
    Dim slideIndex1 As Integer
    Dim slideIndex2 As Integer
    Dim slideIndex3 As Integer
    Dim maxLignesParDiapo As Integer
    Dim lignes() As String
    Dim ligneCount As Integer
    Dim i As Integer
    Dim sectionNum As Integer
    Dim subSectionNum As Integer
    Dim titleSet As Object ' To track added titles

    ' Reference to the active presentation
    Set presentation = objPresentation

    ' Maximum number of lines per slide
    maxLignesParDiapo = 20 ' Adjust as needed

    ' Initialize section and subsection numbers
    sectionNum = 1
    subSectionNum = 1

    ' Create a set to track added titles
    Set titleSet = CreateObject("Scripting.Dictionary")

    ' Add slide titles to the summary dynamically
    For i = 8 To presentation.Slides.Count
        Set diapo = presentation.Slides(i)
        If diapo.Shapes.Placeholders.Count > 0 Then ' Check if the placeholder exists
            On Error Resume Next
            Titre = diapo.Shapes.Placeholders(1).TextFrame.TextRange.text
            On Error GoTo 0
            ' Check if the title is not already added
            If Len(Titre) > 0 And Not titleSet.Exists(Titre) Then
                sommaireTexte = sommaireTexte & vbTab & " " & Titre & vbCrLf
                titleSet.Add Titre, Nothing ' Add title to the set
                subSectionNum = subSectionNum + 1
            End If
        End If
    Next i

    ' Split the summary text into lines
    lignes = Split(sommaireTexte, vbCrLf)
    ligneCount = UBound(lignes) + 1

    ' Separate the summary into three texts
    sommaireTexte1 = ""
    sommaireTexte2 = ""
    sommaireTexte3 = ""

    For i = 0 To ligneCount - 1
        If i < maxLignesParDiapo Then
            sommaireTexte1 = sommaireTexte1 & lignes(i) & vbCrLf
        ElseIf i < maxLignesParDiapo * 2 Then
            sommaireTexte2 = sommaireTexte2 & lignes(i) & vbCrLf
        Else
            sommaireTexte3 = sommaireTexte3 & lignes(i) & vbCrLf
        End If
    Next i

    ' Add the first slide for the summary
    slideIndex1 = 9 ' Dynamic slide index
   ' Set sommaireDiapo1 = presentation.Slides.Add(slideIndex1, ppLayoutText)
   Set sommaireDiapo1 = presentation.Slides(slideIndex1)
    Set pptShape = sommaireDiapo1.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
    pptShape.TextFrame.TextRange.text = sommaireTexte1
    pptShape.TextFrame.TextRange.Font.Name = "Arial"
    pptShape.TextFrame.TextRange.Font.Size = 12

    ' Add the second slide for the summary, if necessary
    If sommaireTexte2 <> "" Then
        slideIndex2 = slideIndex1 + 1
       ' Set sommaireDiapo2 = presentation.Slides.Add(slideIndex2, ppLayoutText)
        Set sommaireDiapo2 = presentation.Slides(slideIndex2)
        Set pptShape = sommaireDiapo2.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte2
        pptShape.TextFrame.TextRange.Font.Name = "Arial"
        pptShape.TextFrame.TextRange.Font.Size = 12
    End If

    ' Add the third slide for the summary, if necessary
    If sommaireTexte3 <> "" Then
        slideIndex3 = slideIndex2 + 1
        Set sommaireDiapo3 = presentation.Slides.Add(slideIndex3, ppLayoutText)
        Set pptShape = sommaireDiapo3.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte3
        pptShape.TextFrame.TextRange.Font.Name = "Arial"
        pptShape.TextFrame.TextRange.Font.Size = 12
    End If
End Function

Function InsererSommaireDynamiqueDeuxNiveauxV1(objPresentation As Object)
    Dim presentation As Object
    Dim sommaireDiapo1 As Object
    Dim sommaireDiapo2 As Object
    Dim sommaireDiapo3 As Object
    Dim diapo As Object
    Dim sommaireTexte As String
    Dim sommaireTexte1 As String
    Dim sommaireTexte2 As String
    Dim sommaireTexte3 As String
    Dim Titre As String
    Dim pptShape As Object
    Dim slideIndex1 As Integer
    Dim slideIndex2 As Integer
    Dim slideIndex3 As Integer
    Dim maxLignesParDiapo As Integer
    Dim lignes() As String
    Dim ligneCount As Integer
    Dim i As Integer
    Dim sectionNum As Integer
    Dim subSectionNum As Integer
    Dim titleSet As Object ' To track added titles

    ' Référence à la présentation active
    Set presentation = objPresentation

    ' Nombre maximum de lignes par diapositive
    maxLignesParDiapo = 20 ' Ajustez selon vos besoins

    ' Initialisation des numéros de sections et sous-sections
    sectionNum = 1
    subSectionNum = 1

    ' Créer le texte du sommaire dynamiquement avec deux niveaux de numérotation
    sommaireTexte = sectionNum & ". SYNTHESIS" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & sectionNum & ".1 TEST CONDITIONS" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & sectionNum & ".2 REFERENCES" & vbCrLf
    sectionNum = sectionNum + 1
    sommaireTexte = sommaireTexte & sectionNum & ". OBJECTIVE RESULTS" & vbCrLf

    ' Créer un ensemble pour suivre les titres ajoutés
    Set titleSet = CreateObject("Scripting.Dictionary")

    ' Ajouter les titres des diapositives au sommaire
    For i = 1 To presentation.Slides.Count
        Set diapo = presentation.Slides(i)
        On Error Resume Next
        Titre = diapo.Shapes.Placeholders(1).TextFrame.TextRange.text
        On Error GoTo 0

        ' Ajouter les sections dynamiquement si le titre n'est pas déjà présent
        If Len(Titre) > 0 And Not titleSet.Exists(Titre) Then
            sommaireTexte = sommaireTexte & vbTab & sectionNum & "." & subSectionNum & " " & Titre & vbCrLf
            titleSet.Add Titre, Nothing ' Ajouter le titre à l'ensemble
            subSectionNum = subSectionNum + 1
        End If
    Next i

    ' Ajouter les annexes
    sectionNum = sectionNum + 1
    sommaireTexte = sommaireTexte & sectionNum & ". ANNEXES" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & sectionNum & ".1 OTHER REMARKS" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & sectionNum & ".2 PROBLEMS" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & sectionNum & ".3 OIL DOCINFO LINK" & vbCrLf

    ' Diviser le texte en lignes
    lignes = Split(sommaireTexte, vbCrLf)
    ligneCount = UBound(lignes) + 1

    ' Séparer le sommaire en trois textes
    sommaireTexte1 = ""
    sommaireTexte2 = ""
    sommaireTexte3 = ""

    For i = 0 To ligneCount - 1
        If i < maxLignesParDiapo Then
            sommaireTexte1 = sommaireTexte1 & lignes(i) & vbCrLf
        ElseIf i < maxLignesParDiapo * 2 Then
            sommaireTexte2 = sommaireTexte2 & lignes(i) & vbCrLf
        Else
            sommaireTexte3 = sommaireTexte3 & lignes(i) & vbCrLf
        End If
    Next i

    ' Ajouter la première diapositive pour le sommaire
    slideIndex1 = 9 ' Dynamic slide index
    Call CopieSlide(presentation, slideIndex1)
    Set sommaireDiapo1 = presentation.Slides(slideIndex1)
    Set pptShape = sommaireDiapo1.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
    pptShape.TextFrame.TextRange.text = sommaireTexte1
    pptShape.TextFrame.TextRange.Font.Name = "Arial"
    pptShape.TextFrame.TextRange.Font.Size = 14

    ' Ajouter la deuxième diapositive pour le sommaire, si nécessaire
    If sommaireTexte2 <> "" Then
        slideIndex2 = slideIndex1 + 1
        Call CopieSlide(presentation, slideIndex2)
        Set sommaireDiapo2 = presentation.Slides(slideIndex2)
        Set pptShape = sommaireDiapo2.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte2
        pptShape.TextFrame.TextRange.Font.Name = "Arial"
        pptShape.TextFrame.TextRange.Font.Size = 12
    End If

    ' Ajouter la troisième diapositive pour le sommaire, si nécessaire
    If sommaireTexte3 <> "" Then
        slideIndex3 = slideIndex2 + 1
        Call CopieSlide(presentation, slideIndex3)
        Set sommaireDiapo3 = presentation.Slides(slideIndex3)
        Set pptShape = sommaireDiapo3.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte3
        pptShape.TextFrame.TextRange.Font.Name = "Arial"
        pptShape.TextFrame.TextRange.Font.Size = 12
    End If

    ' MsgBox "Le sommaire à deux niveaux a été ajouté sur trois diapositives avec succès.", vbInformation
End Function

Function InsererSommaireDynamiqueDeuxNiveaux03(objPresentation As Object)
    Dim presentation As Object
    Dim sommaireDiapo1 As Object
    Dim sommaireDiapo2 As Object
    Dim diapo As Object
    Dim sommaireTexte As String
    Dim sommaireTexte1 As String
    Dim sommaireTexte2 As String
    Dim Titre As String
    Dim pptShape As Object
    Dim slideIndex1 As Integer
    Dim slideIndex2 As Integer
    Dim maxLignesParDiapo As Integer
    Dim lignes() As String
    Dim ligneCount As Integer
    Dim i As Integer
    Dim sectionNum As Integer
    Dim subSectionNum As Integer
    Dim titleSet As Object ' To track added titles

    ' Référence à la présentation active
    Set presentation = objPresentation

    ' Nombre maximum de lignes par diapositive
    maxLignesParDiapo = 20 ' Ajustez selon vos besoins

    ' Initialisation des numéros de sections et sous-sections
    sectionNum = 1
    subSectionNum = 1

    ' Créer le texte du sommaire dynamiquement avec deux niveaux de numérotation
    sommaireTexte = sectionNum & ". SYNTHESIS" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & sectionNum & ".1 TEST CONDITIONS" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & sectionNum & ".2 REFERENCES" & vbCrLf
    sectionNum = sectionNum + 1
    sommaireTexte = sommaireTexte & sectionNum & ". OBJECTIVE RESULTS" & vbCrLf

    ' Créer un ensemble pour suivre les titres ajoutés
    Set titleSet = CreateObject("Scripting.Dictionary")

    ' Ajouter les titres des diapositives au sommaire
    For i = 1 To presentation.Slides.Count
        Set diapo = presentation.Slides(i)
        On Error Resume Next
        Titre = diapo.Shapes.Placeholders(1).TextFrame.TextRange.text
        On Error GoTo 0

        ' Ajouter les sections dynamiquement si le titre n'est pas déjà présent
        If Len(Titre) > 0 And Not titleSet.Exists(Titre) Then
            sommaireTexte = sommaireTexte & vbTab & sectionNum & "." & subSectionNum & " " & Titre & vbCrLf
            titleSet.Add Titre, Nothing ' Ajouter le titre à l'ensemble
            subSectionNum = subSectionNum + 1
        End If
    Next i

    ' Ajouter les annexes
    sectionNum = sectionNum + 1
    sommaireTexte = sommaireTexte & sectionNum & ". ANNEXES" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & sectionNum & ".1 OTHER REMARKS" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & sectionNum & ".2 PROBLEMS" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & sectionNum & ".3 OIL DOCINFO LINK" & vbCrLf

    ' Diviser le texte en lignes
    lignes = Split(sommaireTexte, vbCrLf)
    ligneCount = UBound(lignes) + 1

    ' Séparer le sommaire en deux textes
    sommaireTexte1 = ""
    sommaireTexte2 = ""

    For i = 0 To ligneCount - 1
        If i < maxLignesParDiapo Then
            sommaireTexte1 = sommaireTexte1 & lignes(i) & vbCrLf
        Else
            sommaireTexte2 = sommaireTexte2 & lignes(i) & vbCrLf
        End If
    Next i

    ' Ajouter la première diapositive pour le sommaire
    slideIndex1 = 8 ' Dynamic slide index
    Call CopieSlide(presentation, slideIndex1)
    Set sommaireDiapo1 = presentation.Slides(slideIndex1)
    Set pptShape = sommaireDiapo1.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
    pptShape.TextFrame.TextRange.text = sommaireTexte1
    pptShape.TextFrame.TextRange.Font.Name = "Arial"
    pptShape.TextFrame.TextRange.Font.Size = 14

    ' Ajouter la deuxième diapositive pour le sommaire, si nécessaire
    If sommaireTexte2 <> "" Then
      
        Set sommaireDiapo2 = presentation.Slides(slideIndex1)
        Set pptShape = sommaireDiapo2.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte2
        pptShape.TextFrame.TextRange.Font.Name = "Arial"
        pptShape.TextFrame.TextRange.Font.Size = 12
    End If

    ' MsgBox "Le sommaire à deux niveaux a été ajouté sur deux diapositives avec succès.", vbInformation
End Function


Function InsererSommaireDynamiqueDeuxNiveaux02(objPresentation As Object)
    Dim presentation As Object
    Dim sommaireDiapo1 As Object
    Dim sommaireDiapo2 As Object
    Dim diapo As Object
    Dim sommaireTexte As String
    Dim sommaireTexte1 As String
    Dim sommaireTexte2 As String
    Dim Titre As String
    Dim pptShape As Object
    Dim slideIndex1 As Integer
    Dim slideIndex2 As Integer
    Dim maxLignesParDiapo As Integer
    Dim lignes() As String
    Dim ligneCount As Integer
    Dim i As Integer
    Dim sectionNum As Integer
    Dim subSectionNum As Integer

    ' Référence à la présentation active
    Set presentation = objPresentation

    ' Nombre maximum de lignes par diapositive
    maxLignesParDiapo = 20 ' Ajustez selon vos besoins

    ' Initialisation des numéros de sections et sous-sections
    sectionNum = 1
    subSectionNum = 1

    ' Créer le texte du sommaire dynamiquement avec deux niveaux de numérotation
    sommaireTexte = sectionNum & ". SYNTHESIS" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & sectionNum & ".1 TEST CONDITIONS" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & sectionNum & ".2 REFERENCES" & vbCrLf
    sectionNum = sectionNum + 1
    sommaireTexte = sommaireTexte & sectionNum & ". OBJECTIVE RESULTS" & vbCrLf

    ' Ajouter les titres des diapositives au sommaire
    For i = 1 To presentation.Slides.Count
        Set diapo = presentation.Slides(i)
        On Error Resume Next
        Titre = diapo.Shapes.Placeholders(1).TextFrame.TextRange.text
        On Error GoTo 0

        ' Ajouter les sections dynamiquement
        If Len(Titre) > 0 Then
            sommaireTexte = sommaireTexte & vbTab & sectionNum & "." & subSectionNum & " " & Titre & vbCrLf
            subSectionNum = subSectionNum + 1
        End If
    Next i

    ' Ajouter les annexes
    sectionNum = sectionNum + 1
    sommaireTexte = sommaireTexte & sectionNum & ". ANNEXES" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & sectionNum & ".1 OTHER REMARKS" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & sectionNum & ".2 PROBLEMS" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & sectionNum & ".3 OIL DOCINFO LINK" & vbCrLf

    ' Diviser le texte en lignes
    lignes = Split(sommaireTexte, vbCrLf)
    ligneCount = UBound(lignes) + 1

    ' Séparer le sommaire en deux textes
    sommaireTexte1 = ""
    sommaireTexte2 = ""

    For i = 0 To ligneCount - 1
        If i < maxLignesParDiapo Then
            sommaireTexte1 = sommaireTexte1 & lignes(i) & vbCrLf
        Else
            sommaireTexte2 = sommaireTexte2 & lignes(i) & vbCrLf
        End If
    Next i

    ' Ajouter la première diapositive pour le sommaire
    slideIndex1 = 8 ' Dynamic slide index
    Call CopieSlide(presentation, slideIndex1)
    Set sommaireDiapo1 = presentation.Slides(slideIndex1)
    Set pptShape = sommaireDiapo1.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
    pptShape.TextFrame.TextRange.text = sommaireTexte1
    pptShape.TextFrame.TextRange.Font.Name = "Arial"
    pptShape.TextFrame.TextRange.Font.Size = 14

    ' Ajouter la deuxième diapositive pour le sommaire, si nécessaire
    If sommaireTexte2 <> "" Then
      
        Set sommaireDiapo2 = presentation.Slides(slideIndex1)
        Set pptShape = sommaireDiapo2.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte2
        pptShape.TextFrame.TextRange.Font.Name = "Arial"
        pptShape.TextFrame.TextRange.Font.Size = 12
    End If

    ' MsgBox "Le sommaire à deux niveaux a été ajouté sur deux diapositives avec succès.", vbInformation
End Function

Sub InsererSommaireDynamiqueDeuxNiveaux01(objPresentation As Object)
    Dim presentation As Object
    Dim sommaireDiapo1 As Object
    Dim sommaireDiapo2 As Object
    Dim diapo As Object
    Dim sommaireTexte As String
    Dim sommaireTexte1 As String
    Dim sommaireTexte2 As String
    Dim Titre As String
    Dim pptShape As Object
    Dim slideIndex1 As Integer
    Dim slideIndex2 As Integer
    Dim maxLignesParDiapo As Integer
    Dim lignes() As String
    Dim ligneCount As Integer
    Dim i As Integer
    Dim index As Integer
    Dim sectionNum As Integer
    Dim subSectionNum As Integer

    ' Référence à la présentation active
    Set presentation = objPresentation

    ' Nombre maximum de lignes par diapositive
    maxLignesParDiapo = 20 ' Ajustez selon vos besoins

    ' Initialisation des numéros de sections et sous-sections
    sectionNum = 1
    subSectionNum = 1

    ' Créer le texte du sommaire dynamiquement avec deux niveaux de numérotation
    sommaireTexte = sectionNum & ". SYNTHESIS" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & sectionNum & ".1 TEST CONDITIONS" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & sectionNum & ".2 REFERENCES" & vbCrLf
    sectionNum = sectionNum + 1
    sommaireTexte = sommaireTexte & sectionNum & ". OBJECTIVE RESULTS" & vbCrLf

    ' Ajouter les titres des diapositives au sommaire
    For i = 1 To presentation.Slides.Count
        Set diapo = presentation.Slides(i)
        On Error Resume Next
        Titre = diapo.Shapes.Placeholders(1).TextFrame.TextRange.text
        On Error GoTo 0

        ' Ajouter les sections dynamiquement
        If InStr(1, Titre, vbTextCompare) > 0 Or InStr(1, Titre, vbTextCompare) > 0 Then
            sommaireTexte = sommaireTexte & vbTab & sectionNum & "." & subSectionNum & " " & Titre & vbCrLf
            subSectionNum = subSectionNum + 1
       End If
    Next i

    ' Ajouter les annexes
    sectionNum = sectionNum + 1
    sommaireTexte = sommaireTexte & sectionNum & ". ANNEXES" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & sectionNum & ".1 OTHER REMARKS" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & sectionNum & ".2 PROBLEMS" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & sectionNum & ".3 OIL DOCINFO LINK" & vbCrLf

    ' Diviser le texte en lignes
    lignes = Split(sommaireTexte, vbCrLf)
    ligneCount = UBound(lignes) + 1

    ' Séparer le sommaire en deux textes
    sommaireTexte1 = ""
    sommaireTexte2 = ""

    For i = 0 To ligneCount - 1
        If i < maxLignesParDiapo Then
            sommaireTexte1 = sommaireTexte1 & lignes(i) & vbCrLf
        Else
            sommaireTexte2 = sommaireTexte2 & lignes(i) & vbCrLf
        End If
    Next i

    ' Ajouter la première diapositive pour le sommaire
    slideIndex1 = 8 ' Dynamic slide index
    Call CopieSlide(presentation, slideIndex1)
    Set sommaireDiapo1 = presentation.Slides(slideIndex1)
    Set pptShape = sommaireDiapo1.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
    pptShape.TextFrame.TextRange.text = sommaireTexte1
    pptShape.TextFrame.TextRange.Font.Name = "Arial"
    pptShape.TextFrame.TextRange.Font.Size = 14

    ' Ajouter la deuxième diapositive pour le sommaire, si nécessaire
    If sommaireTexte2 <> "" Then
        slideIndex2 = presentation.Slides.Count + 1 ' Dynamic slide index
        Call CopieSlide(presentation, slideIndex2)
        Set sommaireDiapo2 = presentation.Slides(slideIndex2)
        Set pptShape = sommaireDiapo2.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 70, 700, 400)
        pptShape.TextFrame.TextRange.text = sommaireTexte2
        pptShape.TextFrame.TextRange.Font.Name = "Arial"
        pptShape.TextFrame.TextRange.Font.Size = 12
    End If

    'MsgBox "Le sommaire à deux niveaux a été ajouté sur deux diapositives avec succès.", vbInformation
End Sub

Sub InsererSommaireDynamiqueDeuxDiapos(objPresentation As Object)
    Dim presentation As Object
    Dim sommaireDiapo1 As Object
    Dim sommaireDiapo2 As Object
    Dim diapo As Object
    Dim sommaireTexte As String
    Dim sommaireTexte1 As String
    Dim sommaireTexte2 As String
    Dim Titre As String
    Dim pptShape As Object
    Dim slideIndex1 As Integer
    Dim slideIndex2 As Integer
    Dim maxLignesParDiapo As Integer
    Dim lignes() As String
    Dim ligneCount As Integer
    Dim i As Integer
    Dim index As Integer
    
    ' Référence à la présentation active
    Set presentation = objPresentation

    ' Nombre maximum de lignes par diapositive
    maxLignesParDiapo = 20 ' Ajustez selon vos besoins

    ' Créer le texte du sommaire dynamiquement
    sommaireTexte = "1. SYNTHESIS" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & "1.1 TEST CONDITIONS" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & "1.2 REFERENCES" & vbCrLf
    sommaireTexte = sommaireTexte & "2. OBJECTIVE RESULTS" & vbCrLf

    ' Ajouter les titres des diapositives au sommaire
    For i = 1 To presentation.Slides.Count
        Set diapo = presentation.Slides(i)
        On Error Resume Next
        Titre = diapo.Shapes.Placeholders(1).TextFrame.TextRange.text
        On Error GoTo 0
        
        ' Ajouter les sections dynamiquement
        If InStr(1, Titre, "DRIVABILITY", vbTextCompare) > 0 Or InStr(1, Titre, "DYNAMISM", vbTextCompare) > 0 Then
            sommaireTexte = sommaireTexte & vbTab & Titre & vbCrLf
        End If
    Next i
    
    ' Ajouter les annexes
    sommaireTexte = sommaireTexte & "3. ANNEXES" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & "3.1 OTHER REMARKS" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & "3.2 PROBLEMS" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & "3.3 OIL DOCINFO LINK" & vbCrLf
    
    ' Diviser le texte en lignes
    lignes = Split(sommaireTexte, vbCrLf)
    ligneCount = UBound(lignes) + 1
    
    ' Séparer le sommaire en deux textes
    sommaireTexte1 = ""
    sommaireTexte2 = ""
    
    For i = 0 To ligneCount - 1
        If i < maxLignesParDiapo Then
            sommaireTexte1 = sommaireTexte1 & lignes(i) & vbCrLf
        Else
            sommaireTexte2 = sommaireTexte2 & lignes(i) & vbCrLf
        End If
    Next i
    
    ' Ajouter la première diapositive pour le sommaire
    slideIndex1 = 8
    Call CopieSlide(presentation, slideIndex1)
    Set sommaireDiapo1 = presentation.Slides(slideIndex1)
     Set pptShape = sommaireDiapo1.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 20 + 50, 700, 70)
      'Set sommaireDiapo1 = presentation.Slides.Add(slideIndex1, ppLayoutText)
    'sommaireDiapo1.Shapes.title.TextFrame.TextRange.Text = "Summary"
   pptShape.TextFrame.TextRange.text = sommaireTexte1
    pptShape.TextFrame.TextRange.Font.Name = "Arial"
   pptShape.TextFrame.TextRange.Font.Size = 14
       
    ' Ajouter la deuxième diapositive pour le sommaire, si nécessaire
    If sommaireTexte2 <> "" Then
        slideIndex2 = 9
         Set sommaireDiapo2 = presentation.Slides(slideIndex2)
         Set pptShape = sommaireDiapo2.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 20 + 50, 700, 70)
       
       ' Set sommaireDiapo2 = presentation.Slides.Add(slideIndex2, ppLayoutText)
       ' sommaireDiapo2.Shapes.title.TextFrame.TextRange.Text = "Summary (Part 2)"
        'Set sommaireDiapo2 = presentation.Slides(9)
        pptShape.TextFrame.TextRange.text = sommaireTexte2
        pptShape.TextFrame.TextRange.Font.Name = "Arial"
        pptShape.TextFrame.TextRange.Font.Size = 12
        
    End If

    MsgBox "Le sommaire a été ajouté sur deux diapositives avec succès.", vbInformation
End Sub


Sub InsererSommaireDynamique(objPresentation As Object)
    Dim presentation As Object
    Dim sommaireDiapo As Object
    Dim diapo As Object
    Dim sommaireTexte As String
    Dim Titre As String
    Dim pptShape As Object
    Dim slideIndex As Integer
    Dim mainSectionIndex As Integer
    Dim subSectionIndex1 As Integer
    Dim subSectionIndex2 As Integer
    Dim i As Integer

    ' Référence à la présentation active
    Set presentation = objPresentation

    ' Initialiser les indices de section
    mainSectionIndex = 1
    subSectionIndex1 = 1
    subSectionIndex2 = 1
    
    ' Créer le texte du sommaire dynamiquement
    sommaireTexte = "1. SYNTHESIS" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & "1.1 TEST CONDITIONS" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & "1.2 REFERENCES" & vbCrLf
    sommaireTexte = sommaireTexte & "2. OBJECTIVE RESULTS" & vbCrLf

    ' Ajouter les titres des diapositives au sommaire
    For i = 1 To presentation.Slides.Count
        Set diapo = presentation.Slides(i)
        On Error Resume Next
        Titre = diapo.Shapes.Placeholders(1).TextFrame.TextRange.text
        On Error GoTo 0
        
        ' Si le titre contient certains mots-clés, générer des sections et sous-sections
        If InStr(1, Titre, "DRIVABILITY", vbTextCompare) > 0 Then
            sommaireTexte = sommaireTexte & vbTab & "2." & subSectionIndex1 & "." & subSectionIndex2 & " " & Titre & vbCrLf
            subSectionIndex2 = subSectionIndex2 + 1
        ElseIf InStr(1, Titre, "DYNAMISM", vbTextCompare) > 0 Then
            sommaireTexte = sommaireTexte & vbTab & "2." & subSectionIndex1 & "." & subSectionIndex2 & " " & Titre & vbCrLf
            subSectionIndex2 = subSectionIndex2 + 1
        Else
            ' Réinitialiser le sous-index pour une nouvelle section
            subSectionIndex2 = 1
            subSectionIndex1 = subSectionIndex1 + 1
        End If
    Next i
    
    ' Ajouter les annexes
    sommaireTexte = sommaireTexte & "3. ANNEXES" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & "3.1 OTHER REMARKS" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & "3.2 PROBLEMS" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & "3.3 OIL DOCINFO LINK" & vbCrLf
    
    ' Ajouter une diapositive pour le sommaire
    slideIndex = 1
    Set sommaireDiapo = presentation.Slides.Add(slideIndex, ppLayoutText)
    sommaireDiapo.Shapes.Title.TextFrame.TextRange.text = "Summary"
    
    ' Ajouter le texte du sommaire à la diapositive
    sommaireDiapo.Shapes.Placeholders(2).TextFrame.TextRange.text = sommaireTexte

    MsgBox "Le sommaire a été ajouté avec succès.", vbInformation
End Sub


Sub InsererSommaire2(objPresentation As Object)
    Dim presentation As Object
    Dim sommaireDiapo As Object
    Dim diapo As Object
    Dim i As Integer
    Dim sommaireTexte As String
    Dim Titre As String
    Dim pptShape As Object
    Dim maxLignesParDiapo As Integer
    Dim ligneCount As Integer
    Dim lignes() As String
    Dim slideIndex As Integer
    Dim colonne1Texte As String
    Dim colonne2Texte As String
    Dim currentLineIndex As Integer
    Dim lineStartIndex As Integer
    Dim maxLinesPerColumn As Integer
    
    ' Référence à la présentation active
    Set presentation = objPresentation

    ' Nombre maximum de lignes par diapositive et par colonne
    maxLignesParDiapo = 20
    maxLinesPerColumn = maxLignesParDiapo / 2

    ' Créer le texte du sommaire dynamiquement
    sommaireTexte = "1. Test information" & vbCrLf
    sommaireTexte = sommaireTexte & "2. oDriv results" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & "2.1 Drive away creep eng on" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & "2.2 Drive away creep eng off" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & "2.3 Drive away standing start eng on" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & "2.4 Drive away standing start eng off" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & "2.6 Tip in at deceleration" & vbCrLf
    sommaireTexte = sommaireTexte & vbTab & "2.5 Power-on upshift" & vbCrLf
    sommaireTexte = sommaireTexte & "3. AVL Drive Report Generator results" & vbCrLf

    ' Diviser le texte en lignes
    lignes = Split(sommaireTexte, vbCrLf)
    currentLineIndex = 0
    
    ' Ajouter une diapositive pour le sommaire
    slideIndex = 1
    Set sommaireDiapo = presentation.Slides.Add(slideIndex, ppLayoutText)
    sommaireDiapo.Shapes.Title.TextFrame.TextRange.text = "Sommaire"
    
    Do
        ' Réinitialiser les textes pour les colonnes
        colonne1Texte = ""
        colonne2Texte = ""
        
        ' Remplir les colonnes
        For i = 0 To maxLinesPerColumn - 1
            lineStartIndex = currentLineIndex + i
            If lineStartIndex <= UBound(lignes) Then
                colonne1Texte = colonne1Texte & lignes(lineStartIndex) & vbCrLf
            End If
        Next i
        
        For i = 0 To maxLinesPerColumn - 1
            lineStartIndex = currentLineIndex + maxLinesPerColumn + i
            If lineStartIndex <= UBound(lignes) Then
                colonne2Texte = colonne2Texte & lignes(lineStartIndex) & vbCrLf
            End If
        Next i
        
        ' Ajouter le texte aux colonnes de la diapositive
        With sommaireDiapo.Shapes.Placeholders(2).TextFrame.TextRange
            .text = colonne1Texte
            .Font.Name = "Arial"
            .Font.Size = 14
            .Font.color = RGB(0, 0, 255) ' Bleu
        End With
        
        Set pptShape = sommaireDiapo.Shapes.AddTextbox(msoTextOrientationHorizontal, 320, 20, 300, 400)
        pptShape.TextFrame.TextRange.text = colonne2Texte
        With pptShape.TextFrame.TextRange
            .Font.Name = "Arial"
            .Font.Size = 14
            .Font.color = RGB(0, 0, 255) ' Bleu
        End With
        
        ' Mettre à jour l'index de ligne et ajouter une nouvelle diapositive si nécessaire
        currentLineIndex = currentLineIndex + (2 * maxLinesPerColumn)
        If currentLineIndex <= UBound(lignes) Then
            slideIndex = slideIndex + 1
            Set sommaireDiapo = presentation.Slides.Add(slideIndex, ppLayoutText)
            sommaireDiapo.Shapes.Title.TextFrame.TextRange.text = "Sommaire - Page " & slideIndex
        End If
        
    Loop While currentLineIndex <= UBound(lignes)

    MsgBox "Le sommaire a été ajouté avec succès.", vbInformation
End Sub


Sub InsererSommaire(objPresentation)
    Dim presentation As Object
    Dim sommaireDiapo As Object
    Dim diapo As Object
    Dim i As Integer
'    Dim sommaireTitre As String
'    Dim sommaireTexte  As String
    Dim Titre As String
    Dim pptShape As Object
    Dim objslide As Object
    Dim pptShape1 As Object
    
    ' Référence à la présentation active
    Set presentation = objPresentation

    ' Ajouter une diapositive pour le sommaire au début de la présentation
    Set sommaireDiapo = presentation.Slides(8)
'    objPresentation.Slides(10).Copy
    ' Paste the blank slide into the target presentation
'    Set objSlide = objPresentation.Slides.Paste(11)
'    sommaireDiapo.Shapes.title.TextFrame.TextRange.Text = "Sommaire"

    ' Parcourir les diapositives de la présentation existante
'    sommaireTexte = ""
'    For i = 11 To presentation.Slides.Count
'        Set diapo = presentation.Slides(i)
'        titre = diapo.Shapes.title.TextFrame.TextRange.Text
'        sommaireTexte = sommaireTexte & vbCrLf & i - 1 & ". " & titre
'    Next i

    ' Ajouter le texte du sommaire à la diapositive de sommaire
'    sommaireDiapo.Shapes.Placeholders(2).TextFrame.TextRange.Text = sommaireTexte
    Set pptShape = sommaireDiapo.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 20 + 50, 700, 70)
'    Set pptShape1 = objSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 20 + 50, 700, 70)
    ''''.AddTextbox(msoTextOrientationHorizontal, positionLeft, positionTop, largeur, hauteur)
    
    pptShape.TextFrame.TextRange.text = sommaireTexte
    
'     Dim lignes() As String
'    Dim ligneStart As Integer
'    Dim texteExtraite As String
'    Dim i As Integer
    ' Diviser le texte en lignes
'    lignes = Split(sommaireTexte, vbCrLf)
'    ligneStart = 15
'    ' Construire le texte extrait à partir de la ligne spécifiée
'    If ligneStart <= UBound(lignes) + 1 Then
'        For i = ligneStart - 1 To UBound(lignes)
'            texteExtraite = texteExtraite & lignes(i) & vbCrLf
'        Next i
'        ' Enlever le dernier saut de ligne ajouté à la fin
'        texteExtraite = Left(texteExtraite, Len(texteExtraite) - Len(vbCrLf))
'    Else
'        texteExtraite = "La ligne de départ est au-delà du nombre de lignes du texte."
'    End If
'    pptShape1.TextFrame.TextRange.Text = texteExtraite
    ' Modifier le style du texte
    With pptShape.TextFrame.TextRange
        ' Changer la police
        .Font.Name = "Arial"
        ' Changer la taille de la police
        .Font.Size = 14
        ' Changer la couleur du texte
        .Font.color = RGB(0, 0, 255) ' Bleu
        ' Changer le gras
        .Font.Bold = msoFalse
        ' Changer l'italique
        .Font.Italic = msoFalse
        ' Changer le soulignement
'        .Font.Underline = msoUnderlineSingleLine
    End With
    
'    With pptShape1.TextFrame.TextRange
'        ' Changer la police
'        .Font.Name = "Arial"
'        ' Changer la taille de la police
'        .Font.Size = 14
'        ' Changer la couleur du texte
'        .Font.color = RGB(0, 0, 255) ' Bleu
'        ' Changer le gras
'        .Font.Bold = msoFalse
'        ' Changer l'italique
'        .Font.Italic = msoFalse
'        ' Changer le soulignement
''        .Font.Underline = msoUnderlineSingleLine
'    End With
'    sommaireDiapo.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 20 + 20, 700, 50).TextFrame.TextRange.Text = sommaireTexte
'    ' Ajuster la mise en forme du texte (facultatif)
'    With sommaireDiapo.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 20 + 20, 700, 50)
''        .ParagraphFormat.Bullet.Type = ppBulletUnnumbered
'         .Font.Size = 10
'    End With

''    MsgBox "Le sommaire a été ajouté avec succès.", vbInformation
End Sub

Function newSdvSlide_PPT(objslide As Object, nameSdv As String, Numb As String)


 Dim i As Long
 Dim objshape As Object
 On Error Resume Next
 
    ' Reference the title shape of the slide
'Set objShape = objSlide.Shapes.Text

' Set the text, font, boldness, and size
With objslide.Shapes(1).TextFrame.TextRange
        .text = Numb & " " & nameSdv
        '.text = nameSdv
        '.Font.color.RGB = RGB(0, 0, 255)
        .Font.Bold = msoTrue
        .Font.Size = 18
End With

sommaireTexte = sommaireTexte & vbCrLf & objslide.Shapes(1).TextFrame.TextRange.text

 On Error GoTo 0



'   With objSlide.Shapes.title.TextFrame.TextRange
'        .Text = numb & " " & NameSdv
'        .Font.Bold = msoTrue
'        .Font.Size = 20
'  End With
'     With objSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 20 + 20, 700, 50).TextFrame.TextRange
'        .Text = numb & " " & NameSdv
'        .Font.Bold = msoTrue
'        .Font.Size = 20
'    End With
End Function

Function RemplirTestInformation(objPres As Object, nameSdv As String, slideIndex As Integer)

    Dim objslide As Object
   Set objslide = objPres.Slides(slideIndex)

    ' Set the text, font, boldness, and size
    With objslide.Shapes(1).TextFrame.TextRange
            .text = nameSdv
            '.text = nameSdv
            '.Font.color.RGB = RGB(0, 0, 255)
            .Font.Bold = msoTrue
            .Font.Size = 18
    End With
    
    sommaireTexte = sommaireTexte & vbCrLf & objslide.Shapes(1).TextFrame.TextRange.text
    
    

 On Error GoTo 0
End Function

Function insertPart_PPT(objPresentation As Object, objslide As Object, i As Integer)
    Dim TaBS(7) As String
    Dim slideHeight As Single
    Dim slideWidth As Single
    
 
    ' Define titles
    TaBS(1) = "DRIVABILITY"
    TaBS(2) = "RESPONSIVENESS"
    TaBS(3) = "RESPONSIVENESS"
    TaBS(4) = "DRIVABILITY"
    TaBS(5) = "DRIVABILITY"
    
    TaBS(6) = "RESPONSIVENESS"
    TaBS(7) = "RESPONSIVENESS"
    ' Get slide dimensions from the slide itself, not the master
    slideHeight = objslide.Parent.PageSetup.slideHeight
    slideWidth = objslide.Parent.PageSetup.slideWidth

   If i = 1 Then
      'DoEvents
      With objslide.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 31 + 31 * i, 700, 50).TextFrame.TextRange
        .text = TaBS(i)
        .Font.color.RGB = RGB(0, 0, 127)
        .Font.Bold = msoTrue
        .Font.Size = 14
      End With
         
    ElseIf i = 2 Then
        'DoEvents
        With objslide.Shapes.AddTextbox(msoTextOrientationHorizontal, 500, 21 + 21 * i, 700, 50).TextFrame.TextRange
        .text = TaBS(i)
        .Font.color.RGB = RGB(0, 0, 127)
        .Font.Bold = msoTrue
        .Font.Size = 14
      End With
    
    ElseIf i = 4 Or i = 7 Then
      'DoEvents
      With objslide.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 31 + 31 * 1, 700, 50).TextFrame.TextRange
        .text = TaBS(i)
        .Font.color.RGB = RGB(0, 0, 127)
        .Font.Bold = msoTrue
        .Font.Size = 14
      End With
      
    ElseIf i = 6 Then
             'DoEvents
             With objslide.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 270, 700, 50).TextFrame.TextRange
                .text = TaBS(i)
                .Font.color.RGB = RGB(0, 0, 127)
                .Font.Bold = msoTrue
                .Font.Size = 14
            End With
            
 
    End If


End Function

Function insertPart_PPT3(objslide As Object, i As Integer)
 Dim objshape As Object
 On Error Resume Next

    Dim TaBS(6) As String
    TaBS(1) = "Synthesis"
    'TaBS(2) = "Points visualisation 1"
    'TaBS(3) = "Points visualisation 2"
    'TaBS(4) = "Highest Criticality to improve:"
   ' TaBS(5) = "Lowest Criticality to improve:"



    ' Reference the title shape of the slide
Set objshape = objslide.Shapes.Title

' Set the text, font, boldness, and size
With objshape.TextFrame.TextRange
        .text = TaBS(i)
        .Font.color.RGB = RGB(0, 0, 255)
        .Font.Bold = msoTrue
        .Font.Size = 18
End With

 On Error GoTo 0
    'With objSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 19 + 19 * i, 700, 50).TextFrame.TextRange
        '.Text = TaBS(i)
        '.Font.color.RGB = RGB(0, 0, 255)
        '.Font.Bold = msoTrue
        '.Font.Size = 18
   ' End With
End Function
Function CopyTable_PPT(objppt As Object, Optional objPresentation As Object, Optional r As Range, Optional s As shape, Optional x As String, Optional slideIndex As Integer)
    On Error Resume Next
    Dim i As Integer

    ''slideIndex = 2 ' Adjust as necessary

    For i = 1 To 10
        ERR.Clear
        If Not r Is Nothing Then Call COPYp(r) Else Call COPYp(, s)
        With objPresentation.Slides(slideIndex).Shapes
            .PasteSpecial DataType:=ppPasteOLEObject ''''12             ''''2 ' ppPasteEnhancedMetafile
               With .Item(.Count)
                .Width = 90
                .Height = 180
                .Left = 30
                .Top = 90
                .LockAspectRatio = msoFalse
                .ZOrder msoSendToBack
            End With
            
        End With
        If ERR.Number <> 0 Then Stop
        If ERR.Number = 0 Then Exit Function Else Application.Wait Now + TimeValue("0:00:02")
    Next i

End Function

Function TakePicSelection_PPT4(objppt As Object, objslide As Object, Optional r As Range, Optional s As Object, Optional x As Integer, Optional NameGH As String, Optional leverAs As Boolean)
    
    On Error Resume Next
    Dim shp As Object
    
    ' Check if the provided object is a Range
    If Not r Is Nothing Then
        ' Copy the range and paste it into the slide as a bitmap
        r.Copy
        Set shp = objslide.Shapes.PasteSpecial(DataType:=ppPasteBitmap)(1)
        
'        If NameGH = "PPTENTETE" Then
'            shp.Left = 6
'            shp.Top = 90
'            shp.LockAspectRatio = msoFalse
'            shp.Height = objSlide.Master.Height - 150
'            shp.Width = objSlide.Master.Width - 50
'            shp.LockAspectRatio = msoTrue
'        End If
        
        If leverAs = True Then
            ' Apply styles for Range
            MsgBox "1"
            shp.Left = 180
            shp.Top = 90
            shp.Width = 150
            shp.Height = 300
        Else
        If x = 1 Then
            
            shp.Left = 100
            shp.Top = 100
            shp.Width = 150
            shp.Height = 200
         ElseIf NameGH <> "PPTENTETE" Then
            
            shp.Left = 500
            shp.Top = 100
            shp.Width = 150
            shp.Height = 200
            
        End If
        End If
        
    ElseIf Not s Is Nothing Then
        ' Copy the shape
        s.Copy
        Set shp = objslide.Shapes.PasteSpecial(DataType:=ppPasteBitmap)(1)
        
        ' Check the name of the shape and apply styles accordingly
        If s.Name = "Groupage" And NameGH = "Graphique_0" Or NameGH = "Graphique_00" Then
            If x = 1 Then
                ' Apply styles for "graphe1"
                shp.Left = 100
                shp.Top = 300
                shp.Width = 100
                shp.Height = 200
            ElseIf x = 2 Then
                shp.Left = 500
                shp.Top = 300
                shp.Width = 100
                shp.Height = 200
            End If
            ElseIf NameGH = "Graphique_1" Or NameGH = "Graphique_11" Then
            If x = 1 Then
                ' Apply different styles for "graphe0"
                shp.Left = 100
                shp.Top = 160
                shp.Width = 100
                shp.Height = 200
            ElseIf x = 2 Then
                shp.Left = 500
                shp.Top = 160
                shp.Width = 100
                shp.Height = 200
                
            End If
'        Else
'            ' Apply default styles for other shapes
'            shp.Left = 150
'            shp.Top = 150
'            shp.Width = 350
'            shp.Height = 320
        End If
     End If
  
    On Error GoTo 0
End Function

Function TakePicSelection_PPT3(objppt As Object, objslide As Object, Optional r As Range, Optional s As Object, Optional x As Integer, Optional NameGH As String, Optional leverAs As Boolean)
    
    On Error Resume Next
    Dim shp As Object
    
    ' Check if the provided object is a Range
    If Not r Is Nothing Then
        ' Copy the range and paste it into the slide as a bitmap
        r.Copy
        Set shp = objslide.Shapes.PasteSpecial(DataType:=ppPasteBitmap)(1)
        
        If leverAs = True Then
            ' Apply styles for Range
            shp.Left = 180
            shp.Top = 90
            shp.Height = 300 '300
            shp.Width = 200 '150
        Else
        If x = 1 Then
            shp.Left = 100
            shp.Top = 100
            shp.Height = 250 '200
            shp.Width = 300 '100
         Else
            shp.Left = 500
            shp.Top = 100
            shp.Height = 250 '200
            shp.Width = 300 '150
        End If
        End If
        
    ElseIf Not s Is Nothing Then
        ' Copy the shape
        s.Copy
        Set shp = objslide.Shapes.PasteSpecial(DataType:=ppPasteBitmap)(1)
        
        ' Check the name of the shape and apply styles accordingly
        If s.Name = "Groupage" And NameGH = "Graphique_0" Or NameGH = "Graphique_00" Then
            If x = 1 Then
                ' Apply styles for "graphe1"
                shp.Left = 100
                shp.Top = 300
                shp.Height = 250 '200
                shp.Width = 300 '100
            ElseIf x = 2 Then
                shp.Left = 500
                shp.Top = 300
                shp.Height = 250 '200
                shp.Width = 300 '100
            End If
            ElseIf NameGH = "Graphique_1" Or NameGH = "Graphique_11" Then
            If x = 1 Then
                ' Apply different styles for "graphe0"
                shp.Left = 100
                shp.Top = 160
                shp.Height = 250 '200
                shp.Width = 300 '100
            ElseIf x = 2 Then
                shp.Left = 500
                shp.Top = 160
                shp.Height = 250 '200
                shp.Width = 300 '100
                
            End If
'        Else
'            ' Apply default styles for other shapes
'            shp.Left = 150
'            shp.Top = 150
'            shp.Width = 350
'            shp.Height = 320
        End If
     End If
  
    On Error GoTo 0
End Function
Function TakePicSelection_PPT(objppt As Object, objslide As Object, Optional r As Range, Optional s As Object, Optional NameGH As String, Optional leverAs As Boolean)
    
    On Error Resume Next
    Dim shp As Object
    
    ' Check if the provided object is a Range
    If Not r Is Nothing Then
        ' Copy the range and paste it into the slide as a bitmap
        r.Copy
        Set shp = objslide.Shapes.PasteSpecial(DataType:=ppPasteBitmap)(1)
        
        If leverAs = True Then
            ' Apply styles for Range
            shp.Left = 180
            shp.Top = 90
            shp.Width = 150
            shp.Height = 300
        Else
            shp.Left = 130
            shp.Top = 80
            shp.Width = 150
            shp.Height = 200
        End If
        
        
    ElseIf Not s Is Nothing Then
        ' Copy the shape
        s.Copy
        Set shp = objslide.Shapes.PasteSpecial(DataType:=ppPasteBitmap)(1)
        
        ' Check the name of the shape and apply styles accordingly
        If s.Name = "Groupage" And NameGH = "Graphique_0" Or NameGH = "Graphique_00" Then
            ' Apply styles for "graphe1"
            shp.Left = 100
            shp.Top = 310
            shp.Width = 100
            shp.Height = 150
            
        ElseIf NameGH = "Graphique_1" Or NameGH = "Graphique_11" Then
            ' Apply different styles for "graphe0"
            shp.Left = 350
            shp.Top = 310
            shp.Width = 100
            shp.Height = 150
            
        Else
            ' Apply default styles for other shapes
            shp.Left = 150
            shp.Top = 150
            shp.Width = 350
            shp.Height = 320
        End If
    End If
    
    On Error GoTo 0
End Function
Function TakePicSelection_PPT2(objppt As Object, objPres As Object, Optional objslide As Object, Optional r As Range, Optional s As Object, Optional slideIndex As Integer)
    
    On Error Resume Next
    'Dim i As Integer
    Dim shp As Object
    
      
    'For i = 1 To 10
        ERR.Clear
        
        ' If range is provided, copy it; otherwise, copy the shape
        If Not r Is Nothing Then
            r.Copy
        Else
            s.Copy
        End If
          Set shp = objslide.Shapes.PasteSpecial(DataType:=ppPasteBitmap)(1)
          
        ' Paste the copied content into the PowerPoint slide as a bitmap
        With shp
            .LockAspectRatio = msoTrue
            .PasteSpecial DataType:=ppPasteBitmap
            .Left = 200 ' Adjust position
            .Top = 140  ' Adjust position
            .Width = 150
            .Height = 300
          ' Apply styles to existing shapes on the slide
      
        End With
         
        Set objslide = objPres.Slides.Add(slideIndex, ppLayoutBlank)
                    
        slideIndex = slideIndex + 1
        
        
        ' Check if the paste was successful
        If ERR.Number = 0 Then
            ' Optionally, add space after the pasted content
            ' Insert a new line or paragraph depending on your needs
            With objslide.Shapes
                .AddTextbox(msoTextOrientationHorizontal, 10, objslide.Shapes.Count * 50 + 50, 500, 50).TextFrame.TextRange.text = vbCrLf
            End With
            Exit Function
        Else
            ' Wait for 2 seconds before retrying
            Application.Wait Now + TimeValue("0:00:02")
        End If
    'Next i
End Function

Function TakePicSelection_PPT31(objppt As Object, objslide As Object, Optional r As Range, Optional s As Object)
    
    On Error Resume Next
    Dim i As Integer
    
    For i = 1 To 10
        ERR.Clear
        
        ' If range is provided, copy it; otherwise, copy the shape
        If Not r Is Nothing Then
            r.Copy
        ElseIf Not s Is Nothing Then
            s.Copy
        End If
        
        ' Paste the copied content into the PowerPoint slide as a bitmap
        With objslide.Shapes
            .PasteSpecial DataType:=ppPasteBitmap
        End With
        
        ' Check if the paste was successful
        If ERR.Number = 0 Then
            ' Optionally, add space after the pasted content
            ' Insert a new line or paragraph depending on your needs
            With objslide.Shapes
                .AddTextbox(msoTextOrientationHorizontal, 10, .Count * 50 + 50, 500, 50).TextFrame.TextRange.text = vbCrLf
            End With
            Exit Function
        Else
            ' Wait for 2 seconds before retrying
            Application.Wait Now + TimeValue("0:00:02")
        End If
    Next i
End Function
'Function TakePicSelection_PPT(objPPT As Object, Optional objSlide As Object, Optional r As Range, Optional s As shape, Optional slideIndex As Integer)
'    On Error Resume Next
'    Dim i As Integer
'
'    For i = 1 To 10
'        ERR.Clear
'
'        ' Copy the range or shape
'        If Not r Is Nothing Then
'            Call COPYp(r)
'        Else
'            Call COPYp(, s)
'        End If
'
'        ' Paste the picture into the PowerPoint slide
'        If Not objSlide Is Nothing Then
'           With objSlide.Shapes.PasteSpecial(DataType:=ppPasteBitmap)
'            .LockAspectRatio = msoTrue
'            .Left = 90 ' Adjust position
'            .Top = 170  ' Adjust position
'            .Width = 200
'            .Height = 350
'
'            End With
'
'
'        End If
'
'        Set objSlide = objSlide.Add(slideIndex, ppLayoutText)
'        slideIndex = slideIndex + 1
'
'        ' Check if the operation was successful
'        If ERR.Number = 0 Then
'            Exit Function
'        Else
'            Application.Wait Now + TimeValue("0:00:02")
'        End If
'    Next i
'End Function
 Function CopySummary_PPT(objppt As Object, objslide As Object, SDV As String, x As Integer, slideIndex As Integer)
    
    
    Dim ws As Worksheet
    Dim success As Boolean
    Dim shp As Object
    Dim cb, ca As Integer
    
    
    Set ws = ThisWorkbook.Worksheets(SDV)
    
    If x = 1 Then
        
    

'DoEvents
'sleep 500

success = False
 
   
    Do While Not success
        On Error Resume Next
 
        ws.Range("B4:K22").CopyPicture Appearance:=xlScreen, Format:=xlPicture
        

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop

'DoEvents
'sleep 500


success = False

cb = objslide.Shapes.Count

Do While Not success
    objslide.Shapes.PasteSpecial DataType:=2
    ca = objslide.Shapes.Count
    success = ca > cb
Loop


Set shp = objslide.Shapes(objslide.Shapes.Count)

DoEvents
With shp
    .Left = 31
    .Top = 86
    .Width = 388
End With
   
   

    ElseIf x = 2 Then
        
        
            
           
           
           
           
          
        'DoEvents
        'sleep 500
success = False
Do While Not success
        On Error Resume Next
 
        
        ws.Range("BI4:BR22").CopyPicture Appearance:=xlScreen, Format:=xlPicture

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop

        
'DoEvents
'sleep 500


success = False

cb = objslide.Shapes.Count

Do While Not success
    objslide.Shapes.PasteSpecial DataType:=2
    ca = objslide.Shapes.Count
    success = ca > cb
Loop


Set shp = objslide.Shapes(objslide.Shapes.Count)


'DoEvents
With shp
    .Left = 500
    .Top = 86
    .Width = 388
End With

     
    End If
    
End Function
Function CopyGraph0_PPT(objppt As Object, objslide As Object, SDV As String, x As Integer, slideIndex As Integer, sh As String)
    Dim chartObj As ChartObject
    Dim success As Boolean
    Dim cb, ca As Integer
    Dim shp As Object
    Dim ws As Worksheet
    Dim obj As shape
    
    
    
    Set ws = ThisWorkbook.Worksheets(sh)
    ws.Columns.Hidden = False
    
    If x = 1 Then
        
        
        
        
        Set chartObj = ws.ChartObjects("Graphique_0")
        If chartObj.Visible = True Then
                
'DoEvents
'sleep 500

success = False
 
   
    Do While Not success
        On Error Resume Next
 
        chartObj.CopyPicture Appearance:=xlScreen, Format:=xlPicture
        

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop

'DoEvents
'sleep 500


success = False

cb = objslide.Shapes.Count

Do While Not success
    objslide.Shapes.PasteSpecial DataType:=2
    ca = objslide.Shapes.Count
    success = ca > cb
Loop


Set shp = objslide.Shapes(objslide.Shapes.Count)

'DoEvents
With shp
    .LockAspectRatio = msoFalse
    .Left = 31
    .Top = 315
    .Width = 388
    .Height = 202
End With
            
       Set obj = ws.Shapes("Keys")
        
            
       success = False
 
   
    Do While Not success
        On Error Resume Next
 
        obj.CopyPicture Appearance:=xlScreen, Format:=xlPicture
        

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop

'DoEvents
'sleep 500


success = False

cb = objslide.Shapes.Count

Do While Not success
    objslide.Shapes.PasteSpecial DataType:=2
    ca = objslide.Shapes.Count
    success = ca > cb
Loop


Set shp = objslide.Shapes(objslide.Shapes.Count)

'DoEvents
With shp
    .LockAspectRatio = msoFalse
    .Left = 375
    .Top = 400
    .Width = 40
    .Height = 110
End With
            
            
       
        End If
        

                
    ElseIf x = 2 Then
        
      
        
        Set chartObj = ws.ChartObjects("Graphique_00")
        If chartObj.Visible = True Then
        
        
        'DoEvents
        'sleep 500

        success = False
 
   
        Do While Not success
        On Error Resume Next
 
        chartObj.Chart.CopyPicture Appearance:=xlScreen, Format:=xlPicture
        

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop

'DoEvents
'sleep 500


success = False

cb = objslide.Shapes.Count

Do While Not success
    objslide.Shapes.PasteSpecial DataType:=2
    ca = objslide.Shapes.Count
    success = ca > cb
Loop


Set shp = objslide.Shapes(objslide.Shapes.Count)

'DoEvents
With shp
    .LockAspectRatio = msoFalse
    .Left = 500
    .Top = 315
    .Width = 388
    .Height = 202
End With

 Set obj = ws.Shapes("Keyss")
        
            
       success = False
 
   
    Do While Not success
        On Error Resume Next
 
        obj.CopyPicture Appearance:=xlScreen, Format:=xlPicture
        

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop

'DoEvents
'sleep 500


success = False

cb = objslide.Shapes.Count

Do While Not success
    objslide.Shapes.PasteSpecial DataType:=2
    ca = objslide.Shapes.Count
    success = ca > cb
Loop


Set shp = objslide.Shapes(objslide.Shapes.Count)

'DoEvents
With shp
    .LockAspectRatio = msoFalse
    .Left = 844
    .Top = 400
    .Width = 40
    .Height = 110
End With
            
            
 

    
      End If
    End If
           
End Function
Function CopyGraph1_PPT(objppt As Object, objPres As Object, objslide As Object, SDV As String, x As Integer, slideIndex As Integer)
              
 
    Dim chartObj As ChartObject
    Dim shp As Object
    Dim ca, cb As Integer
    Dim success As Boolean
    Dim ws As Worksheet
    
    
    Set ws = ThisWorkbook.Worksheets(SDV)
    ws.Columns.Hidden = False
    
    If x = 1 Then
        
        
        
        
        Set chartObj = ws.ChartObjects("Graphique_1")
        If chartObj.Visible = True Then
     
      
'DoEvents
'sleep 500

success = False
 
   
    Do While Not success
        On Error Resume Next
 
        chartObj.CopyPicture Appearance:=xlScreen, Format:=xlPicture
        

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop

'DoEvents
'sleep 500


success = False

cb = objslide.Shapes.Count

Do While Not success
    objslide.Shapes.PasteSpecial DataType:=2
    ca = objslide.Shapes.Count
    success = ca > cb
Loop


Set shp = objslide.Shapes(objslide.Shapes.Count)

'DoEvents
With shp
    .LockAspectRatio = msoFalse
    .Left = 31
    .Top = 135
    .Width = 320
    .Height = 220
End With
                 
        End If
        
                
    ElseIf x = 2 Then
        
          
        Set chartObj = ws.ChartObjects("Graphique_11")
        
        If chartObj.Visible = True Then
        'DoEvents
        'sleep 500

        success = False
 
   
        Do While Not success
        On Error Resume Next
 
        chartObj.Chart.CopyPicture Appearance:=xlScreen, Format:=xlPicture
        

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop

'DoEvents
'sleep 500


success = False

cb = objslide.Shapes.Count

Do While Not success
    objslide.Shapes.PasteSpecial DataType:=2
    ca = objslide.Shapes.Count
    success = ca > cb
Loop


Set shp = objslide.Shapes(objslide.Shapes.Count)

'DoEvents
With shp
    .LockAspectRatio = msoFalse
    .Left = 500
    .Top = 135
    .Width = 320
    .Height = 220
End With
               


               
   End If
               
               
               
    End If

    
End Function

Function CopyLeverAS_PPT(objslide As Object, SDV As String, tables As String)
    Dim ws As Worksheet
    Dim rng As Range
    
    Set ws = ThisWorkbook.Worksheets(SDV)
    
    ' Define the range you want to copy
    ' Adjust this range according to your data
    Set rng = ws.Range(tables)
    
    ' Copy the range
    rng.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    
    ' Paste the picture into the slide
    'With objSlide.Shapes.PasteSpecial ''(ppPasteBitmap)
        '.LockAspectRatio = msoTrue
        '.top = 100
        '.Left = 50
    'End With
    
      With objslide.Shapes.PasteSpecial(DataType:=ppPasteEnhancedMetafile)
                .LockAspectRatio = msoTrue
                .Left = 100 ' Adjust position
                .Top = 100  ' Adjust position
                .Width = 200
                .Height = 400
                
            End With
End Function
Function CopyPriorityPoints_PPT(objppt As Object, objPres As Object, objslide As Object, Filt As String, Shts As Worksheet, Pos As Integer, slideIndex As Integer)
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
    Dim shapeWidth As Long
    Dim pasted As Boolean
    
    ' Setting initial width for the shape to be inserted in PowerPoint
    shapeWidth = 500

    With Shts
        ' Find the maximum row in columns 13 to 15
        For x = 13 To 15
            If .Cells(.Rows.Count, x).End(xlUp).row > TotalRow Then TotalRow = .Cells(.Rows.Count, x).End(xlUp).row
        Next x
        
        colonne = getLastColumnDrivability(Shts.Name) - 1
        
        ' Hide specific columns and check priority filter
        Call HideC3(Shts.Name, "driv")
        If FilterPriority_PPT(Shts, Filt) = True And colonne <> 0 Then
            ' Find the last row for the selected columns
            For x = 13 To 15
                If .Cells(.Rows.Count, x).End(xlUp).row > lastRow Then lastRow = .Cells(.Rows.Count, x).End(xlUp).row
            Next x
            Set plage = .Range(.Cells(3, 13), .Cells(lastRow, colonne))
            
            ' Criticality check and construction of priority points string
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
             ViderPressePapiers
             
            pasted = False
            Do While Not pasted
                On Error Resume Next
                objPres.Slides(9).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
            Loop
             
             ViderPressePapiers
            ' Insert a part into the PowerPoint slide at the specified position
            'Call insertPart_PPT(objPPT, objSlide, pos)
            
            ' Take a screenshot of the selection and insert it into the PowerPoint slide
            Call TakePicSelection_PPT3(objppt, objslide, plage)
            
            ' Resize the last inserted shape (picture) in PowerPoint
            objslide.Shapes(objslide.Shapes.Count).LockAspectRatio = msoFalse
            objslide.Shapes(objslide.Shapes.Count).Width = shapeWidth
            
            ' Unfilter the sheet after processing
            Call UnFilterPriority(Shts, TotalRow)
        End If
    End With
End Function

Sub CopyPriorityPoints_PPT3(objppt As Object, objPres As Object, objslide As Object, Filt As String, Shts As Worksheet, Pos As Integer, slideIndex As Integer, o As String, jj As Long, BlnAddSlide As Boolean)
    Dim plage As Range
    Dim colonne As Integer
    Dim lastRow As Long
    Dim TotalRow As Long
    Dim currow As Long
    Dim lastRowP1 As Long
    Dim x As Long
    Dim i As Integer
    Dim j As Integer
    Dim cD
    Dim trouveCol As Boolean
    Dim cG As String
    Dim tabPS() As String
    Dim colPS As String
    Dim plageentete As Range
    Dim plage1 As Range
    Dim BlnPlage2 As Boolean
    Dim objImageBox As PowerPoint.shape
    Dim chemin, NomImage As String
    Dim MyChart As Chart
    Dim ws As Worksheet
    Dim haut, large As Single
    Dim success As Boolean
    Dim pasted As Boolean
    Dim ca, cb As Integer
    Dim shp As Object
    Dim grandPlage As Range
    Dim plageEnt As Range
    Dim scaledHeight As Double
    Dim h As Double
    Dim MaxHeight, entHeight As Double
   
    
    
    Set ws = ThisWorkbook.Worksheets(o)
    With Shts
         For x = 13 To 15
            If .Cells(.Rows.Count, x).End(xlUp).row > TotalRow Then TotalRow = .Cells(.Rows.Count, x).End(xlUp).row
        Next x
        
        
       colonne = getLastColumnDrivability(Shts.Name) - 1
     
        Call HideC3(Shts.Name, "driv")
        If FilterPriority_PPT(Shts, Filt) = True And colonne <> 0 Then
            addDynam = True
        Else
            addDynam = False
        End If
        
        If FilterPriority_PPT(Shts, Filt) = True And colonne <> 0 Then
            For x = 13 To 15
                If .Cells(.Rows.Count, x).End(xlUp).row > lastRow Then lastRow = .Cells(.Rows.Count, x).End(xlUp).row
            Next x
            Set grandPlage = ws.Range(.Cells(7, 13), .Cells(lastRow, colonne))
            Set plageEnt = ws.Range(.Cells(3, 13), .Cells(6, colonne))
            scaledHeight = grandPlage.Height * 940 / grandPlage.Width
            drivHeight = 0
            entHeight = 940 / plageEnt.Width * plageEnt.Height
            MaxHeight = 515 - plageEnt.Top - entHeight
            
            
            
            If scaledHeight <= MaxHeight Then
                BlnPlage2 = False
                DivDriv = False
            Else
                BlnPlage2 = True
                DivDriv = True
            End If
            
            
            ViderPressePapiers
            pasted = False
            Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                'DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
            Loop
            
            
            ViderPressePapiers
            Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
            slideIndex = slideIndex + 1
        
        
            Call insertPart_PPT(objPres, objslide, 4)
           
            
            
            
            

            
            If Not BlnPlage2 Then
            
            
                
                Set ws = Shts
               
                'DoEvents
                'sleep 500
                success = False
                Do While Not success
        On Error Resume Next
 
        
        plageEnt.CopyPicture Appearance:=xlScreen, Format:=xlPicture

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop
                
                
                DoEvents
                'sleep 500
                
                success = False

cb = objslide.Shapes.Count

Do While Not success
    objslide.Shapes.PasteSpecial DataType:=2
    ca = objslide.Shapes.Count
    success = ca > cb
Loop


Set shp = objslide.Shapes(objslide.Shapes.Count)

DoEvents
With shp
    .Left = 10
    .Top = 90
    .Width = objslide.Master.Width - 20
End With
drivHeight = drivHeight + shp.Height
h = 90 + shp.Height


                'DoEvents
                'sleep 500
                success = False
                Do While Not success
        On Error Resume Next
 
        
        grandPlage.CopyPicture Appearance:=xlScreen, Format:=xlPicture

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop
                
                
                'DoEvents
                
                
                success = False

cb = objslide.Shapes.Count

Do While Not success
    objslide.Shapes.PasteSpecial DataType:=2
    ca = objslide.Shapes.Count
    success = ca > cb
Loop


Set shp = objslide.Shapes(objslide.Shapes.Count)

DoEvents
With shp
    .Left = 10
    .Top = h - 1
    .Width = objslide.Master.Width - 20
End With
drivHeight = drivHeight + shp.Height

          
            Else
                
                Call divgraphique(plageEnt, Filt, o, objPres, slideIndex, jj, objppt, objslide, grandPlage)
                BlnPlage2 = False
            End If

            
            BlnAddSlideDriv = True
        Else
            BlnAddSlideDriv = False
            
        End If
    End With
    
End Sub

Function InsererEntete(objppt As Object, objslide As Object, entete As Range, o As String)


Dim objImageBox As PowerPoint.shape
Dim chemin, NomImage As String
Dim MyChart As Chart
Dim ws As Worksheet
Dim haut, large As Single
Dim success As Boolean
Dim shp As Object
Dim ca, cb As Integer

    
    
    
    Set ws = ThisWorkbook.Worksheets(o)
    ws.Activate
    NomImage = ActiveSheet.Name
    DoEvents
    'sleep 500
    success = False
    Do While Not success
        On Error Resume Next
 
        
        entete.CopyPicture Appearance:=xlScreen, Format:=xlPicture

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop
    
    DoEvents
    'sleep 500
    
    success = False

cb = objslide.Shapes.Count

Do While Not success
    objslide.Shapes.PasteSpecial DataType:=2
    ca = objslide.Shapes.Count
    success = ca > cb
Loop


Set shp = objslide.Shapes(objslide.Shapes.Count)

DoEvents
With shp
    .Left = 10
    .Top = 90
    .Width = objslide.Master.Width - 20
End With
    
    
    Call verifDim(objslide)
                
    





End Function

Function InsererTable(objppt As Object, objslide As Object, plage As Range, o As String, h As Double)


Dim objImageBox As PowerPoint.shape
Dim chemin, NomImage As String
Dim MyChart As Chart
Dim ws As Worksheet
Dim haut, large As Single
Dim success As Boolean
Dim ca, cb As Integer
Dim shp As Object
Dim GetSubRangeByHeight As Range
Dim newHeight, newWidth As Long
Dim MaxWidth, MaxHeight, originalWidth, originalHeight, scaledWeight As Long
Dim tempRange As Range
Dim finalRange As Range
Dim scaledTempHeight, scaledHeight, tempHeight As Long
Dim i As Integer

    Dim startRow As Long, endRow As Long

    Dim totalHeight As Double

    Dim colStart As Long, colEnd As Long
 
    
    Set ws = ThisWorkbook.Worksheets(o)
    ws.Activate
    NomImage = ActiveSheet.Name
    DoEvents
    'sleep 500
    success = False
    Do While Not success
        On Error Resume Next
 
        
        plage.CopyPicture Appearance:=xlScreen, Format:=xlPicture

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop
    
    
    DoEvents
    'sleep 500
    
    success = False

cb = objslide.Shapes.Count

Do While Not success
    objslide.Shapes.PasteSpecial DataType:=2
    ca = objslide.Shapes.Count
    success = ca > cb
Loop




Set shp = objslide.Shapes(objslide.Shapes.Count)



With shp
    .LockAspectRatio = msoTrue
    .Width = objslide.Master.Width - 20
    .Left = 10
    .Top = h
End With




Set plage = ws.Range("BT6:DA143")
MaxWidth = objslide.Master.Width - 20 ' fixed for all slides
MaxHeight = 315 ' max space allowed in slide
 
' Step 1: check full range
originalWidth = plage.Width
originalHeight = plage.Height
scaledHeight = (MaxWidth / originalWidth) * originalHeight
 
If scaledHeight > MaxHeight Then
    ' Export full range

    ' Step 2: loop through rows until subrange height fits
    For i = 1 To plage.Rows.Count
        Set tempRange = ws.Range(plage.Rows(1), plage.Rows(i))
        tempHeight = tempRange.Height
        scaledTempHeight = (MaxWidth / originalWidth) * tempHeight
 
        If scaledTempHeight > MaxHeight Then
            Set finalRange = ws.Range(plage.Rows(1), plage.Rows(i - 1))
            Exit For
        End If
    Next i
End If



Set ws = ThisWorkbook.Worksheets(o)
    ws.Activate
    NomImage = ActiveSheet.Name
    DoEvents
    'sleep 500
    success = False
    Do While Not success
        On Error Resume Next
 
        
        finalRange.CopyPicture Appearance:=xlScreen, Format:=xlPicture

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop
    
    
    DoEvents
    'sleep 500
    
    success = False

cb = objslide.Shapes.Count

Do While Not success
    objslide.Shapes.PasteSpecial DataType:=2
    ca = objslide.Shapes.Count
    success = ca > cb
Loop




Set shp = objslide.Shapes(objslide.Shapes.Count)



With shp
    .LockAspectRatio = msoTrue
    .Width = objslide.Master.Width - 20
    .Left = 10
    .Top = h
End With

   

    
MsgBox "dhia"

    
                
    



End Function

Sub ConstPlages(comptPlage As Integer, Shts As Worksheet, colPS As String, tabPS() As String, trouveCol As Boolean, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    
    Dim cD
    Dim j As Integer
    
    With Shts
    
    If comptPlage = 2 Then
        For Each cD In plage1.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage1.Rows.Count + plage1.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage2.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage2.Rows.Count + plage2.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
    
    ElseIf comptPlage = 3 Then
        
        For Each cD In plage1.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage1.Rows.Count + plage1.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage2.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage2.Rows.Count + plage2.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage3.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage3.Rows.Count + plage3.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
    ElseIf comptPlage = 4 Then
        
        For Each cD In plage1.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage1.Rows.Count + plage1.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage2.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage2.Rows.Count + plage2.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage3.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage3.Rows.Count + plage3.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage4.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage4.Rows.Count + plage4.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
    
    ElseIf comptPlage = 5 Then
        
        For Each cD In plage1.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage1.Rows.Count + plage1.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage2.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage2.Rows.Count + plage2.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage3.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage3.Rows.Count + plage3.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage4.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage4.Rows.Count + plage4.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage5.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage5.Rows.Count + plage5.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
    ElseIf comptPlage = 6 Then
        
        For Each cD In plage1.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage1.Rows.Count + plage1.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage2.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage2.Rows.Count + plage2.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage3.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage3.Rows.Count + plage3.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage4.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage4.Rows.Count + plage4.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage5.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage5.Rows.Count + plage5.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage6.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage6.Rows.Count + plage6.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
    
    ElseIf comptPlage = 7 Then
        
        For Each cD In plage1.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage1.Rows.Count + plage1.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage2.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage2.Rows.Count + plage2.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage3.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage3.Rows.Count + plage3.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage4.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage4.Rows.Count + plage4.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage5.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage5.Rows.Count + plage5.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage6.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage6.Rows.Count + plage6.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage7.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage7.Rows.Count + plage7.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
    
    ElseIf comptPlage = 8 Then
        
        For Each cD In plage1.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage1.Rows.Count + plage1.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage2.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage2.Rows.Count + plage2.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage3.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage3.Rows.Count + plage3.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage4.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage4.Rows.Count + plage4.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage5.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage5.Rows.Count + plage5.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage6.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage6.Rows.Count + plage6.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage7.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage7.Rows.Count + plage7.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage8.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage8.Rows.Count + plage8.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
    
    ElseIf comptPlage = 9 Then
        
        For Each cD In plage1.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage1.Rows.Count + plage1.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage2.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage2.Rows.Count + plage2.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage3.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage3.Rows.Count + plage3.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage4.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage4.Rows.Count + plage4.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage5.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage5.Rows.Count + plage5.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage6.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage6.Rows.Count + plage6.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage7.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage7.Rows.Count + plage7.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage8.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage8.Rows.Count + plage8.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage9.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage9.Rows.Count + plage9.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
    ElseIf comptPlage = 10 Then
        
        For Each cD In plage1.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage1.Rows.Count + plage1.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage2.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage2.Rows.Count + plage2.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage3.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage3.Rows.Count + plage3.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage4.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage4.Rows.Count + plage4.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage5.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage5.Rows.Count + plage5.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage6.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage6.Rows.Count + plage6.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage7.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage7.Rows.Count + plage7.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage8.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage8.Rows.Count + plage8.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage9.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage9.Rows.Count + plage9.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage10.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage10.Rows.Count + plage10.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
    ElseIf comptPlage = 11 Then
        
        For Each cD In plage1.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage1.Rows.Count + plage1.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage2.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage2.Rows.Count + plage2.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage3.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage3.Rows.Count + plage3.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage4.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage4.Rows.Count + plage4.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage5.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage5.Rows.Count + plage5.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage6.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage6.Rows.Count + plage6.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage7.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage7.Rows.Count + plage7.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage8.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage8.Rows.Count + plage8.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage9.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage9.Rows.Count + plage9.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage10.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage10.Rows.Count + plage10.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage11.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage11.Rows.Count + plage11.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
    ElseIf comptPlage = 12 Then
        
        For Each cD In plage1.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage1.Rows.Count + plage1.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage2.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage2.Rows.Count + plage2.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage3.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage3.Rows.Count + plage3.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage4.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage4.Rows.Count + plage4.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage5.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage5.Rows.Count + plage5.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage6.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage6.Rows.Count + plage6.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage7.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage7.Rows.Count + plage7.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage8.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage8.Rows.Count + plage8.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage9.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage9.Rows.Count + plage9.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage10.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage10.Rows.Count + plage10.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage11.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage11.Rows.Count + plage11.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
    For Each cD In plage12.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage12.Rows.Count + plage12.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
    
    ElseIf comptPlage > 12 Then
        
        For Each cD In plage1.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage1.Rows.Count + plage1.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage2.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage2.Rows.Count + plage2.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage3.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage3.Rows.Count + plage3.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage4.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage4.Rows.Count + plage4.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage5.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage5.Rows.Count + plage5.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage6.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage6.Rows.Count + plage6.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage7.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage7.Rows.Count + plage7.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage8.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage8.Rows.Count + plage8.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage9.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage9.Rows.Count + plage9.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage10.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage10.Rows.Count + plage10.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage11.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage11.Rows.Count + plage11.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage12.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage12.Rows.Count + plage12.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage13.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 13) & " > Priority : " & Shts.Cells(cD.row, 14)
                    For j = 0 To UBound(tabPS)
                           If Not .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, .Rows(6).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage13.Rows.Count + plage13.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
    
    End If
    
    End With
                                
                                
End Sub

Sub ConstPlagesDyn(comptPlage As Integer, Shts As Worksheet, colPS As String, tabPS() As String, trouveCol As Boolean, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    
    Dim cD
    Dim j As Integer
    
    With Shts
    
    
    If comptPlage = 2 Then
        For Each cD In plage1.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage1.Rows.Count + plage1.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage2.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage2.Rows.Count + plage2.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
    
    ElseIf comptPlage = 3 Then
        
        For Each cD In plage1.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage1.Rows.Count + plage1.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage2.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage2.Rows.Count + plage2.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage3.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage3.Rows.Count + plage3.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
    ElseIf comptPlage = 4 Then
        
        For Each cD In plage1.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage1.Rows.Count + plage1.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage2.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage2.Rows.Count + plage2.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage3.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage3.Rows.Count + plage3.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage4.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage4.Rows.Count + plage4.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
    
    ElseIf comptPlage = 5 Then
        
        For Each cD In plage1.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage1.Rows.Count + plage1.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage2.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage2.Rows.Count + plage2.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage3.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage3.Rows.Count + plage3.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage4.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage4.Rows.Count + plage4.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage5.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage5.Rows.Count + plage5.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
    ElseIf comptPlage = 6 Then
        
        For Each cD In plage1.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage1.Rows.Count + plage1.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage2.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage2.Rows.Count + plage2.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage3.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage3.Rows.Count + plage3.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage4.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage4.Rows.Count + plage4.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage5.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage5.Rows.Count + plage5.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage6.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage6.Rows.Count + plage6.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
    
    ElseIf comptPlage = 7 Then
        
        For Each cD In plage1.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage1.Rows.Count + plage1.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage2.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage2.Rows.Count + plage2.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage3.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage3.Rows.Count + plage3.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage4.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage4.Rows.Count + plage4.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage5.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage5.Rows.Count + plage5.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage6.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage6.Rows.Count + plage6.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage7.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage7.Rows.Count + plage7.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
    
    ElseIf comptPlage = 8 Then
        
        For Each cD In plage1.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage1.Rows.Count + plage1.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage2.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage2.Rows.Count + plage2.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage3.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage3.Rows.Count + plage3.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage4.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage4.Rows.Count + plage4.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage5.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage5.Rows.Count + plage5.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage6.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage6.Rows.Count + plage6.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage7.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage7.Rows.Count + plage7.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage8.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage8.Rows.Count + plage8.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
    
    ElseIf comptPlage = 9 Then
        
        For Each cD In plage1.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage1.Rows.Count + plage1.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage2.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage2.Rows.Count + plage2.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage3.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage3.Rows.Count + plage3.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage4.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage4.Rows.Count + plage4.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage5.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage5.Rows.Count + plage5.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage6.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage6.Rows.Count + plage6.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage7.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage7.Rows.Count + plage7.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage8.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage8.Rows.Count + plage8.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage9.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage9.Rows.Count + plage9.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
    ElseIf comptPlage = 10 Then
        
        For Each cD In plage1.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage1.Rows.Count + plage1.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage2.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage2.Rows.Count + plage2.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage3.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage3.Rows.Count + plage3.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage4.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage4.Rows.Count + plage4.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage5.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage5.Rows.Count + plage5.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage6.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage6.Rows.Count + plage6.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage7.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage7.Rows.Count + plage7.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage8.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage8.Rows.Count + plage8.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage9.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage9.Rows.Count + plage9.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage10.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage10.Rows.Count + plage10.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
    ElseIf comptPlage = 11 Then
        
        For Each cD In plage1.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage1.Rows.Count + plage1.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage2.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage2.Rows.Count + plage2.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage3.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage3.Rows.Count + plage3.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage4.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage4.Rows.Count + plage4.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage5.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage5.Rows.Count + plage5.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage6.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage6.Rows.Count + plage6.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage7.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage7.Rows.Count + plage7.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage8.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage8.Rows.Count + plage8.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage9.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage9.Rows.Count + plage9.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage10.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage10.Rows.Count + plage10.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage11.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage11.Rows.Count + plage11.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
    ElseIf comptPlage = 12 Then
        
        For Each cD In plage1.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage1.Rows.Count + plage1.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage2.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage2.Rows.Count + plage2.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage3.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage3.Rows.Count + plage3.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage4.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage4.Rows.Count + plage4.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage5.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage5.Rows.Count + plage5.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage6.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage6.Rows.Count + plage6.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage7.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage7.Rows.Count + plage7.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage8.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage8.Rows.Count + plage8.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage9.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage9.Rows.Count + plage9.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage10.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage10.Rows.Count + plage10.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage11.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage11.Rows.Count + plage11.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
    For Each cD In plage12.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage12.Rows.Count + plage12.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
    
    ElseIf comptPlage > 12 Then
        
        For Each cD In plage1.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage1.Rows.Count + plage1.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage2.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage2.Rows.Count + plage2.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage3.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage3.Rows.Count + plage3.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage4.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage4.Rows.Count + plage4.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage5.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage5.Rows.Count + plage5.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage6.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage6.Rows.Count + plage6.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage7.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage7.Rows.Count + plage7.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage8.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage8.Rows.Count + plage8.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage9.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage9.Rows.Count + plage9.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage10.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage10.Rows.Count + plage10.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage11.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage11.Rows.Count + plage11.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage12.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage12.Rows.Count + plage12.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD
        
        For Each cD In plage13.Rows
           
            If cD.row > 6 And Shts.Rows(cD.row).Hidden = False Then
                    colPS = colPS & "Criticality : " & Shts.Cells(cD.row, 72) & " > Priority : " & Shts.Cells(cD.row, 72)
                    For j = 0 To UBound(tabPS)
                           If Not Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole) Is Nothing Then
                                 colPS = colPS & " > " & tabPS(j) & " : " & Shts.Cells(cD.row, Shts.Range("BT6:" & Shts.Cells(6, colonne).Address).Find(What:=tabPS(j), LookAt:=xlWhole).Column)
                                 trouveCol = True
                           End If
                    Next j
                    If cD.row < (plage13.Rows.Count + plage13.row) - 1 Then
                          colPS = colPS & ";"
                    End If
              End If
        Next cD

    
                               
    End If
    
    End With
                                
                                
End Sub

Sub divgraphiqueDyn2(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, comptPlage As Integer, objppt As Object, objslide As Object, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    Dim h As Double
    Dim objshp As Variant
    Dim pasted As Boolean
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage2, o, h)
        
        Call verifDim(objslide)
End Sub
Sub divgraphiqueDyn3(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, comptPlage As Integer, objppt As Object, objslide As Object, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    Dim h As Double
    Dim objshp As Variant
    Dim pasted As Boolean
    
    
    
    ViderPressePapiers
    
    pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
        
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage2, o, h)
        
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage3, o, h)
        
        Call verifDim(objslide)
End Sub
Sub divgraphiqueDyn4(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, comptPlage As Integer, objppt As Object, objslide As Object, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    Dim h As Double
    Dim objshp As Variant
    Dim pasted As Boolean
    
    ViderPressePapiers
    
    pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
        
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage2, o, h)
        
        Call verifDim(objslide)
        ViderPressePapiers
        
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage3, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage4, o, h)
        Call verifDim(objslide)
End Sub
Sub divgraphiqueDyn5(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, comptPlage As Integer, objppt As Object, objslide As Object, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    Dim h As Double
    Dim objshp As Variant
    Dim pasted As Boolean
    
    ViderPressePapiers
    
    pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
        
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage2, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage3, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage4, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage5, o, h)
        Call verifDim(objslide)
End Sub
Sub divgraphiqueDyn6(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, comptPlage As Integer, objppt As Object, objslide As Object, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    Dim h As Double
    Dim objshp As Variant
    Dim pasted As Boolean
    
    ViderPressePapiers
    
    
    pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
        
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage2, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage3, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage4, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage5, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage6, o, h)
        Call verifDim(objslide)
End Sub
Sub divgraphiqueDyn7(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, comptPlage As Integer, objppt As Object, objslide As Object, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    Dim h As Double
    Dim objshp As Variant
    Dim pasted As Boolean
    
    ViderPressePapiers
    
    pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage2, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage3, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage4, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage5, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage6, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage7, o, h)
        Call verifDim(objslide)
End Sub
Sub divgraphiqueDyn8(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, comptPlage As Integer, objppt As Object, objslide As Object, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    Dim h As Double
    Dim objshp As Variant
    Dim pasted As Boolean
    
    ViderPressePapiers
    
    pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage2, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage3, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage4, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage5, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage6, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage7, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage8, o, h)
        Call verifDim(objslide)
End Sub
Sub divgraphiqueDyn9(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, comptPlage As Integer, objppt As Object, objslide As Object, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    Dim h As Double
    Dim objshp As Variant
    Dim pasted As Boolean
    
    ViderPressePapiers
    
    pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage2, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage3, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage4, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage5, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage6, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage7, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage8, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage9, o, h)
        Call verifDim(objslide)
End Sub
Sub divgraphiqueDyn10(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, comptPlage As Integer, objppt As Object, objslide As Object, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    Dim h As Double
    Dim objshp As Variant
    Dim pasted As Boolean
    
    ViderPressePapiers
    
    pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage2, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage3, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage4, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage5, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage6, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage7, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage8, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage9, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage10, o, h)
        Call verifDim(objslide)
End Sub
Sub divgraphiqueDyn11(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, comptPlage As Integer, objppt As Object, objslide As Object, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    Dim h As Double
    Dim objshp As Variant
    Dim pasted As Boolean
    
    ViderPressePapiers
    
    pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage2, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
       pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage3, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage4, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage5, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage6, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage7, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage8, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage9, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage10, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage11, o, h)
        Call verifDim(objslide)
End Sub
Sub divgraphiqueDyn12(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, comptPlage As Integer, objppt As Object, objslide As Object, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    Dim h As Double
    Dim objshp As Variant
    Dim pasted As Boolean
    
    
    ViderPressePapiers
    
    pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage2, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage3, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage4, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage5, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage6, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage7, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage8, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage9, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage10, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage11, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage12, o, h)
        Call verifDim(objslide)
End Sub
Sub divgraphiqueDyn13(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, comptPlage As Integer, objppt As Object, objslide As Object, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    Dim h As Double
    Dim objshp As Variant
    Dim pasted As Boolean
    
    ViderPressePapiers
    
    pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage2, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage3, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage4, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage5, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage6, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage7, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage8, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage9, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage10, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage11, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage12, o, h)
        Call verifDim(objslide)
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage13, o, h)
        Call verifDim(objslide)
End Sub
Sub divgraphiqueDyn(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, objppt As Object, objslide As Object, grandPlage As Range)
    
    Dim objshp As Variant
    Dim objImageBox As PowerPoint.shape
    
    Dim ws As Worksheet
    Dim haut, large As Single
    Dim success As Boolean
    Dim ca, cb As Integer
    Dim shp As Object
    
 
    
    
    Dim exportRange As Range
    Dim tempRange As Range
 
    Dim originalWidth As Double, originalHeight As Double
    Dim scaledHeight As Double
    Dim MaxWidth As Double, MaxHeight As Double
 
    Dim rowStart As Long, rowEnd As Long, lastRow As Long
    Dim h As Double
    Dim pasted As Boolean
    Dim pages As Integer
    Dim entHeight As Double
    
    
  
    pages = 0
    Set ws = ThisWorkbook.Worksheets(o)
    MaxWidth = objslide.Master.Width - 20
    entHeight = 940 / ent.Width * ent.Height
    MaxHeight = 515 - ent.Top - entHeight
    
 
        originalWidth = grandPlage.Width
        originalHeight = grandPlage.Height
        lastRow = grandPlage.Rows.Count
 
        rowStart = 1
 
        Do While rowStart < lastRow
    For rowEnd = rowStart To lastRow
        Set tempRange = ws.Range(grandPlage.Rows(rowStart), grandPlage.Rows(rowEnd))
        scaledHeight = (MaxWidth / originalWidth) * tempRange.Height
        
        If scaledHeight > MaxHeight Then
            If rowEnd = rowStart Then
                
                Set exportRange = tempRange
                rowEnd = rowEnd + 1
                Exit For
            Else
                Set exportRange = ws.Range(grandPlage.Rows(rowStart), grandPlage.Rows(rowEnd - 1))
                Exit For
            End If
        End If
 
        If rowEnd = lastRow Then
            Set exportRange = tempRange
            Exit For
        End If
    Next rowEnd
    
    
    
            ViderPressePapiers
            If pages <> 0 Then
            pasted = False
            Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                'DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 7)
        End If
        
        
        Set ws = ThisWorkbook.Worksheets(o)

    'DoEvents
    'sleep 500
    success = False
    Do While Not success
        On Error Resume Next
 
        
        ent.CopyPicture Appearance:=xlScreen, Format:=xlPicture

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop
    
    'DoEvents
    'sleep 500
   
    
    
    success = False

cb = objslide.Shapes.Count

Do While Not success
    objslide.Shapes.PasteSpecial DataType:=2
    ca = objslide.Shapes.Count
    success = ca > cb
Loop


Set shp = objslide.Shapes(objslide.Shapes.Count)

'DoEvents
With shp
    
    .Left = 10
    .Top = 90
    .Width = objslide.Master.Width - 20
End With

h = shp.Height + 90
   Set ws = ThisWorkbook.Worksheets(o)
   
    
    'DoEvents
    'sleep 500
    success = False
    Do While Not success
        On Error Resume Next
 
        
        exportRange.CopyPicture Appearance:=xlScreen, Format:=xlPicture

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop
    
    'DoEvents
    'sleep 500
    
    success = False

cb = objslide.Shapes.Count

Do While Not success
    objslide.Shapes.PasteSpecial DataType:=2
    ca = objslide.Shapes.Count
    success = ca > cb
Loop


Set shp = objslide.Shapes(objslide.Shapes.Count)

DoEvents
With shp
    .LockAspectRatio = msoTrue
    .Left = 10
    .Top = h - 1
    .Width = objslide.Master.Width - 20
End With


            
 
            rowStart = rowEnd
            pages = pages + 1
        Loop
    

               
End Sub



Function verifDim1(obslide As Object)

Dim shp As Variant
Dim hight As Double


Set shp = obslide.Shapes(obslide.Shapes.Count)
hight = obslide.Shapes(obslide.Shapes.Count).Height
'MsgBox hight
If hight >= 425 Then

     shp.LockAspectRatio = msoFalse
     obslide.Shapes(obslide.Shapes.Count).Height = 410
End If
    
    
End Function

Function verifDim(obslide As Object)

Dim shp As Variant
Dim hight As Double


Set shp = obslide.Shapes(obslide.Shapes.Count)
hight = obslide.Shapes(obslide.Shapes.Count).Height
'MsgBox hight
If hight >= 332 Then

     shp.LockAspectRatio = msoFalse
     obslide.Shapes(obslide.Shapes.Count).Height = 312
End If

    
    
End Function

Sub divgraphique3(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, comptPlage As Integer, objppt As Object, objslide As Object, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    Dim h As Double
    Dim objshp As Variant
    Dim pasted As Boolean
    
    ViderPressePapiers
    
    pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage2, o, h)
        
        
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage3, o, h)
        
End Sub

Sub divgraphique2(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, comptPlage As Integer, objppt As Object, objslide As Object, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    Dim h As Double
    Dim objshp As Variant
    Dim pasted As Boolean
    
    ViderPressePapiers
    
    pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        Call insertPart_PPT(objPres, objslide, 4)
        slideIndex = slideIndex + 1
        
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage2, o, h)
End Sub

Sub divgraphique4(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, comptPlage As Integer, objppt As Object, objslide As Object, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    Dim h As Double
    Dim objshp As Variant
    Dim pasted As Boolean
    
    ViderPressePapiers
    
    pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage2, o, h)
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage3, o, h)
        
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage4, o, h)

End Sub

Sub divgraphique5(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, comptPlage As Integer, objppt As Object, objslide As Object, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    Dim h As Double
    Dim objshp As Variant
    Dim pasted As Boolean
    
    ViderPressePapiers
    
    pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage2, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage3, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage4, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage5, o, h)
End Sub

Sub divgraphique6(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, comptPlage As Integer, objppt As Object, objslide As Object, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    Dim h As Double
    Dim objshp As Variant
    Dim pasted As Boolean
    
    ViderPressePapiers
    
    pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage2, o, h)
        
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage3, o, h)
        
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage4, o, h)
        
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage5, o, h)
        
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, ent, o, h)
End Sub

Sub divgraphique7(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, comptPlage As Integer, objppt As Object, objslide As Object, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    Dim h As Double
    Dim objshp As Variant
    Dim pasted As Boolean
    
    ViderPressePapiers
    
    pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage2, o, h)
        
        
        ViderPressePapiers
        
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage3, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage4, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage5, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage6, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage7, o, h)

End Sub

Sub divgraphique8(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, comptPlage As Integer, objppt As Object, objslide As Object, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    Dim h As Double
    Dim objshp As Variant
    Dim pasted As Boolean
    
        ViderPressePapiers
    
    pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage2, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage3, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage4, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage5, o, h)
        
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage6, o, h)
        
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage7, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage8, o, h)

End Sub

Sub divgraphique9(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, comptPlage As Integer, objppt As Object, objslide As Object, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    Dim h As Double
    Dim objshp As Variant
    Dim pasted As Boolean
    
    
        ViderPressePapiers
    
    pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage2, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage3, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage4, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage5, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage6, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage7, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage8, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage9, o, h)

End Sub

Sub divgraphique10(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, comptPlage As Integer, objppt As Object, objslide As Object, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    Dim h As Double
    Dim objshp As Variant
    Dim pasted As Boolean
    
    
        ViderPressePapiers
    
    pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage2, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage3, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage4, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage5, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage6, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage7, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage8, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage9, o, h)
        
        
        ViderPressePapiers
    
    pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage10, o, h)
End Sub

Sub divgraphique11(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, comptPlage As Integer, objppt As Object, objslide As Object, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    Dim h As Double
    Dim objshp As Variant
    Dim pasted As Boolean
    
    
        ViderPressePapiers
    
    pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage2, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage3, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage4, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage5, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage6, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage7, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage8, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage9, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage10, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage11, o, h)

End Sub

Sub divgraphique12(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, comptPlage As Integer, objppt As Object, objslide As Object, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    Dim h As Double
    Dim objshp As Variant
    Dim pasted As Boolean
    
    
        ViderPressePapiers
    
    pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        Call insertPart_PPT(objPres, objslide, 4)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage2, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        Call insertPart_PPT(objPres, objslide, 4)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage3, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        Call insertPart_PPT(objPres, objslide, 4)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage4, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        Call insertPart_PPT(objPres, objslide, 4)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage5, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        Call insertPart_PPT(objPres, objslide, 4)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage6, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        Call insertPart_PPT(objPres, objslide, 4)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage7, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        Call insertPart_PPT(objPres, objslide, 4)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage8, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        Call insertPart_PPT(objPres, objslide, 4)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage9, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        Call insertPart_PPT(objPres, objslide, 4)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage10, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        Call insertPart_PPT(objPres, objslide, 4)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage11, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        Call insertPart_PPT(objPres, objslide, 4)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage12, o, h)
        

End Sub

Sub divgraphique13(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, comptPlage As Integer, objppt As Object, objslide As Object, plage1 As Range, plage2 As Range, plage3 As Range, plage4 As Range, plage5 As Range, plage6 As Range, plage7 As Range, plage8 As Range, plage9 As Range, plage10 As Range, plage11 As Range, plage12 As Range, plage13 As Range)
    Dim h As Double
    Dim objshp As Variant
    Dim pasted As Boolean
    
    
        ViderPressePapiers
    
    pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        Call insertPart_PPT(objPres, objslide, 4)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage2, o, h)
        
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        Call insertPart_PPT(objPres, objslide, 4)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage3, o, h)
        
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        Call insertPart_PPT(objPres, objslide, 4)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage4, o, h)
        
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        Call insertPart_PPT(objPres, objslide, 4)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage5, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        Call insertPart_PPT(objPres, objslide, 4)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage6, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        Call insertPart_PPT(objPres, objslide, 4)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage7, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        Call insertPart_PPT(objPres, objslide, 4)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage8, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        Call insertPart_PPT(objPres, objslide, 4)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage9, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        Call insertPart_PPT(objPres, objslide, 4)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage10, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        Call insertPart_PPT(objPres, objslide, 4)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage11, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        Call insertPart_PPT(objPres, objslide, 4)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage12, o, h)
        
        
        ViderPressePapiers
        
        pasted = False
        Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        Call insertPart_PPT(objPres, objslide, 4)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        Call InsererEntete(objppt, objslide, ent, o)
        h = objslide.Shapes(objslide.Shapes.Count).Height + 90
        Call InsererTable(objppt, objslide, plage13, o, h)

End Sub

Sub divgraphique(ent As Range, Filt As String, o As String, objPres As Object, slideIndex As Integer, jj As Long, objppt As Object, objslide As Object, grandPlage As Range)
    
    
    
    Dim objshp As Variant
    Dim objImageBox As PowerPoint.shape
    
    Dim ws As Worksheet
    Dim haut, large As Single
    Dim success As Boolean
    Dim ca, cb As Integer
    Dim shp As Object
    
 
    
    
    Dim exportRange As Range
    Dim tempRange As Range
 
    Dim originalWidth As Double, originalHeight As Double
    Dim scaledHeight As Double
    Dim MaxWidth As Double, MaxHeight As Double
 
    Dim rowStart As Long, rowEnd As Long, lastRow As Long
    Dim h As Double
    Dim pasted As Boolean
    Dim pages As Integer
    Dim Espace As Double
    Dim entHeight As Double
    

   
    pages = 0
    Set ws = ThisWorkbook.Worksheets(o)
    MaxWidth = objslide.Master.Width - 20
    entHeight = 940 / ent.Width * ent.Height
    MaxHeight = 515 - ent.Top - entHeight
 
        originalWidth = grandPlage.Width
        originalHeight = grandPlage.Height
        lastRow = grandPlage.Rows.Count
 
        rowStart = 1
 
        Do While rowStart < lastRow
    For rowEnd = rowStart To lastRow
        Set tempRange = ws.Range(grandPlage.Rows(rowStart), grandPlage.Rows(rowEnd))
        scaledHeight = (MaxWidth / originalWidth) * tempRange.Height
 
        If scaledHeight > MaxHeight Then
            If rowEnd = rowStart Then
                
                Set exportRange = tempRange
                rowEnd = rowEnd + 1
                Exit For
            Else
                Set exportRange = ws.Range(grandPlage.Rows(rowStart), grandPlage.Rows(rowEnd - 1))
                Exit For
            End If
        End If
 
        If rowEnd = lastRow Then
            Set exportRange = tempRange
            Exit For
        End If
    Next rowEnd
 
 
            ViderPressePapiers
            If pages <> 0 Then
            pasted = False
            Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                'DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
        Loop
        
        ViderPressePapiers
        Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
        slideIndex = slideIndex + 1
        
        Call insertPart_PPT(objPres, objslide, 4)
        End If
        
        
        Set ws = ThisWorkbook.Worksheets(o)
    'DoEvents
    'sleep 500
    success = False
    Do While Not success
        On Error Resume Next
 
        
        ent.CopyPicture Appearance:=xlScreen, Format:=xlPicture

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop
    
    'DoEvents
    'sleep 500
  
    
    
    success = False

cb = objslide.Shapes.Count

Do While Not success
    objslide.Shapes.PasteSpecial DataType:=2
    ca = objslide.Shapes.Count
    success = ca > cb
Loop


Set shp = objslide.Shapes(objslide.Shapes.Count)

'DoEvents
With shp
    
    .Left = 10
    .Top = 90
    .Width = objslide.Master.Width - 20
End With
h = shp.Height + 90
   Set ws = ThisWorkbook.Worksheets(o)
    
    
    'DoEvents
    'sleep 500
    success = False
    Do While Not success
        On Error Resume Next
 
        
        exportRange.CopyPicture Appearance:=xlScreen, Format:=xlPicture

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop
    
    'DoEvents
    'sleep 500
    
    success = False

cb = objslide.Shapes.Count

Do While Not success
    objslide.Shapes.PasteSpecial DataType:=2
    ca = objslide.Shapes.Count
    success = ca > cb
Loop


Set shp = objslide.Shapes(objslide.Shapes.Count)

'DoEvents
With shp
    .LockAspectRatio = msoTrue
    .Left = 10
    .Top = h - 1
    .Width = objslide.Master.Width - 20
End With

            
 
            rowStart = rowEnd
            pages = pages + 1
        Loop
    
    
                
End Sub

Function CopyPriorityPointsDyn_PPT3(objppt As Object, objPres As Object, objslide As Object, Filt As String, Shts As Worksheet, Pos As Integer, slideIndex As Integer, o As String, jj As Long, BlnAddSlide As Boolean)
    Dim plageEnt As Range
    Dim plage As Range
    Dim colonne As Integer
    Dim lastRow As Long
    Dim TotalRow As Long
    Dim currow As Long
    Dim lastRowP1 As Long
    Dim x As Long
    Dim i As Integer
    Dim j As Integer
    Dim cD
    Dim trouveCol As Boolean
    Dim cG As String
    Dim tabPS() As String
    Dim colPS As String
    Dim plage1 As Range
    Dim BlnPlage2 As Boolean
    Dim objImageBox As PowerPoint.shape
    Dim chemin, NomImage As String
    Dim MyChart As Chart
    Dim ws As Worksheet
    Dim haut, large As Single
    Dim success As Boolean
    Dim pasted As Boolean
    Dim ca, cb As Integer
    Dim shp As Object
    Dim grandPlage As Range
    Dim scaledHeight As Double
    Dim scaledEntHeight As Double
    Dim h As Double
    Dim dynHeight As Double
    Dim entHeight, MaxHeight As Double
    
    
    
    
    Set ws = ThisWorkbook.Worksheets(o)
    With Shts
         For x = 72 To 74
            If .Cells(.Rows.Count, x).End(xlUp).row > TotalRow Then TotalRow = .Cells(.Rows.Count, x).End(xlUp).row
        Next x
      
       colonne = getLastColumnDinamyc(Shts.Name) - 1
       
        Call HideC3(Shts.Name, "dyn")
        If FilterPriorityDyn_PPT(Shts, Filt) = True And colonne <> 0 Then
            For x = 72 To 74
                If .Cells(.Rows.Count, x).End(xlUp).row > lastRow Then lastRow = .Cells(.Rows.Count, x).End(xlUp).row
            Next x
            Set grandPlage = ws.Range(.Cells(7, 72), .Cells(lastRow, colonne))
            Set plageEnt = ws.Range(.Cells(3, 72), .Cells(6, colonne))
            scaledHeight = grandPlage.Height * 940 / grandPlage.Width
            scaledEntHeight = plageEnt.Height * 940 / plageEnt.Width
            dynHeight = scaledHeight + scaledEntHeight
            entHeight = 940 / plageEnt.Width * plageEnt.Height
            MaxHeight = 515 - plageEnt.Top - entHeight
            
            
            If scaledHeight <= MaxHeight Then
                BlnPlage2 = False
            Else
                BlnPlage2 = True
            End If
            
            
            If DivDriv = False And dynHeight <= 215 And drivHeight <= 180 And addDynam = True Then
                
            Call insertPart_PPT(objPres, objslide, 6)
            Set ws = Shts
                'DoEvents
                'sleep 500
                success = False
                Do While Not success
        On Error Resume Next
 
        
        plageEnt.CopyPicture Appearance:=xlScreen, Format:=xlPicture

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop
                
                
                'DoEvents
                'sleep 500
                
                success = False

cb = objslide.Shapes.Count

Do While Not success
    objslide.Shapes.PasteSpecial DataType:=2
    ca = objslide.Shapes.Count
    success = ca > cb
Loop


Set shp = objslide.Shapes(objslide.Shapes.Count)

DoEvents
With shp
    .Left = 10
    .Top = 300
    .Width = objslide.Master.Width - 20
End With

h = 300 + shp.Height




Set ws = Shts
                
                'DoEvents
                'sleep 500
                success = False
                Do While Not success
        On Error Resume Next
 
        
        grandPlage.CopyPicture Appearance:=xlScreen, Format:=xlPicture

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop
                
                
                'DoEvents
                'sleep 500
                
                success = False

cb = objslide.Shapes.Count

Do While Not success
    objslide.Shapes.PasteSpecial DataType:=2
    ca = objslide.Shapes.Count
    success = ca > cb
Loop


Set shp = objslide.Shapes(objslide.Shapes.Count)

DoEvents
With shp
    .Left = 10
    .Top = h - 1
    .Width = objslide.Master.Width - 20
End With

                
                
                
                
                
                
            Else
            ViderPressePapiers
            pasted = False
            Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                'DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
            Loop
            
            
            ViderPressePapiers
            Call newSdvSlide_PPT(objslide, UCase(o), "2." & jj + 1)
            slideIndex = slideIndex + 1
        
        
            Call insertPart_PPT(objPres, objslide, 7)
            
            If Not BlnPlage2 Then
                
                Set ws = Shts
                'DoEvents
                'sleep 500
                success = False
                Do While Not success
        On Error Resume Next
 
        
        plageEnt.CopyPicture Appearance:=xlScreen, Format:=xlPicture

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop
                
                
                'DoEvents
                'sleep 500
                
                success = False

cb = objslide.Shapes.Count

Do While Not success
    objslide.Shapes.PasteSpecial DataType:=2
    ca = objslide.Shapes.Count
    success = ca > cb
Loop


Set shp = objslide.Shapes(objslide.Shapes.Count)

DoEvents
With shp
    .Left = 10
    .Top = 90
    .Width = objslide.Master.Width - 20
End With

h = 90 + shp.Height




Set ws = Shts
                
                'DoEvents
                'sleep 500
                success = False
                Do While Not success
        On Error Resume Next
 
        
        grandPlage.CopyPicture Appearance:=xlScreen, Format:=xlPicture

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop
                
                
                'DoEvents
                'sleep 500
                
                success = False

cb = objslide.Shapes.Count

Do While Not success
    objslide.Shapes.PasteSpecial DataType:=2
    ca = objslide.Shapes.Count
    success = ca > cb
Loop


Set shp = objslide.Shapes(objslide.Shapes.Count)

DoEvents
With shp
    .Left = 10
    .Top = h - 1
    .Width = objslide.Master.Width - 20
End With


            
           
            
           
            Else
                Call divgraphiqueDyn(plageEnt, Filt, o, objPres, slideIndex, jj, objppt, objslide, grandPlage)
                BlnPlage2 = False
            End If
            End If
            Call UnFilterPriority(Shts, TotalRow)
       
        End If
    End With
End Function
Sub CopyPriorityPointsDyn_PPT(objslide As Object, Filt As String, Shts As Worksheet, Pos As Integer)
    Dim plage As Range
    Dim colonne As Integer
    Dim lastRow As Long
    Dim TotalRow As Long
    Dim x As Long
    Dim cD As Range
    Dim trouveCol As Boolean
    Dim cG As String
    Dim tabPS() As String
    Dim colPS As String
    
    On Error Resume Next
    
    With Shts
        For x = 72 To 74
            If .Cells(.Rows.Count, x).End(xlUp).row > TotalRow Then TotalRow = .Cells(.Rows.Count, x).End(xlUp).row
        Next x
        
        colonne = getLastColumnDinamyc(Shts.Name) - 1
        
        Call HideC3(Shts.Name, "dyn")
        If FilterPriorityDyn_PPT(Shts, Filt) = True And colonne <> 0 Then
            For x = 72 To 74
                If .Cells(.Rows.Count, x).End(xlUp).row > lastRow Then lastRow = .Cells(.Rows.Count, x).End(xlUp).row
            Next x
            Set plage = .Range(.Cells(3, 72), .Cells(lastRow, colonne))
            
            ' Copy the range
            plage.CopyPicture Appearance:=xlScreen, Format:=xlPicture
            
            ' Paste the picture into the slide
            'With objSlide.Shapes.PasteSpecial '''(ppPasteBitmap)
                '.LockAspectRatio = msoTrue
                '.top = 100
                '.Left = 50
            'End With
            
            ' With objSlide.Shapes ''(ppPasteBitmap)
        '.PasteSpecial ppPasteEnhancedMetafile
        '.PasteSpecial.LockAspectRatio = msoTrue
        '.PasteSpecial.top = 100   ' Adjust the top position as needed
        '.PasteSpecial.Left = 50   ' Adjust the left position as needed
        '.PasteSpecial.width = 250 ' Adjust the width as needed
        '.PasteSpecial.height = 300 ' Adjust the height as needed
    'End With

            
            Call UnFilterPriority(Shts, TotalRow)
        End If
    End With
    On Error GoTo 0
End Sub
Function Remplissage_PPT2(Fields As Variant, objppt As Object, objPresentation As Object)
    Dim i As Integer
    Dim j As Integer
    Dim tabF() As String
    Dim tabB() As String
    Dim slide As Object
    Dim shape As Object
    Dim foundShape As Boolean

    For i = 1 To UBound(Fields)
        tabF = Split(Fields(i), "#")
        If InStr(1, tabF(1), ";") <> 0 Then
            tabB = Split(tabF(1), ";")
            For j = 0 To UBound(tabB)
                foundShape = False
                For Each slide In objPresentation.Slides
                    For Each shape In slide.Shapes
                        If shape.HasTextFrame Then
                            If shape.TextFrame.TextRange.text = tabB(j) Then
                                shape.TextFrame.TextRange.text = tabF(0)
                                foundShape = True
                                Exit For
                            End If
                        End If
                    Next shape
                    If foundShape Then Exit For
                Next slide
            Next j
        Else
            foundShape = False
            For Each slide In objPresentation.Slides
                For Each shape In slide.Shapes
                    If shape.HasTextFrame Then
                        If shape.TextFrame.TextRange.text = tabF(1) Then
                            If tabF(1) = "Signet7" Then
                                shape.TextFrame.TextRange.text = CStr(Date)
                               ' For Each shape In slide.Shapes
                                    If shape.HasTextFrame Then
                                        If shape.TextFrame.TextRange.text = "Places" Then
                                            shape.TextFrame.TextRange.text = tabF(0) & " " & CStr(Date)
                                            Exit For
                                        End If
                                    End If
                              '  Next shape
                            Else
                                shape.TextFrame.TextRange.text = tabF(0)
                            End If
                            foundShape = True
                            Exit For
                        End If
                    End If
                Next shape
                If foundShape Then Exit For
            Next slide
        End If
    Next i
    'slideIndex = slideIndex + 1 ' Increment slideIndex after processing
End Function
Sub RemplirPiedPage(ByRef objppt As Object, objPresentation As Object, Fields As Variant)
    Dim slide As Object
    Dim piedShape As Object
    Dim piedTexte As String
    Dim tabF() As String
    Dim i As Integer
    Dim hasTitle As Boolean
    Dim shape   As Object
    Dim text As String
    Dim Indice_BackUp As String
    Dim Var_Indice_BackUp As Variant
    Dim k As Integer
    Dim Bln_Fond_White As Boolean
    
    
 hasTitle = False ' Initialize the flag



'''''''''''''''''''''Chercher l'indice de Slide qui contient Back Up
  ' Parcourir toutes les diapositives
    Indice_BackUp = "1"
    For Each slide In objPresentation.Slides
        ' Parcourir toutes les formes (shapes) de la diapositive
        For Each shape In slide.Shapes
            ' Vérifier si la forme contient du texte
            If shape.HasTextFrame Then
                If shape.TextFrame.HasText Then
                    text = shape.TextFrame.TextRange.text
                    ' Tester si "BACK UP" est dans le texte de la forme
                    If InStr(1, text, "Back UP", vbTextCompare) > 0 Then
                        'Trouve = True
                        'Exit For ' Sortir dès que "BACK UP" est trouvé
                        Indice_BackUp = Indice_BackUp & "," & slide.slideIndex
                    End If
                End If
            End If
        Next shape
        ' Si le texte a été trouvé dans une diapositive, afficher un message et sortir
'        If Trouve Then
'            MsgBox "La diapositive " & slide.slideIndex & " contient 'BACK UP'.", vbInformation
'            Exit Sub
'        End If
    Next slide
    
    Var_Indice_BackUp = Split(Indice_BackUp, ",")
    
    ' Boucle à travers chaque diapositive
For Each slide In objPresentation.Slides
      ' Loop through the shapes in the slide to check for the title placeholder
    For i = 1 To slide.Shapes.Count
        If slide.Shapes(i).Type = msoPlaceholder Then
            If slide.Shapes(i).PlaceholderFormat.Type = ppPlaceholderTitle Then
                Set piedShape = slide.Shapes(i)
                hasTitle = True
                Exit For
            End If
        End If
    Next i
   If slide.slideIndex <> 1 Then
   
        For i = 1 To UBound(Fields)
          tabF = Split(Fields(i), "#")
        Debug.Print "The value of tabF is: " & Join(tabF, ", ")
          
          If UCase(tabF(1)) = UCase("DepTV") Then
           
                If hasTitle Then
                              Set piedShape = slide.Shapes.Title
                              
                              ' Set the text, font, boldness, and size
                              With piedShape.TextFrame.TextRange
                                      .text = tabF(0)
                                      ' Test if tabF(0) equals 1 to set the font color to white
                                       Bln_Fond_White = False
                                       For k = 0 To UBound(Var_Indice_BackUp)
                                            If slide.slideIndex = CDbl(Var_Indice_BackUp(k)) Then
                                                .Font.color.RGB = RGB(255, 255, 255) ' White
                                                Bln_Fond_White = True
                                                Exit For
                                            End If
                                        Next k
                                        
                                        If Not Bln_Fond_White Then
                                            .Font.color.RGB = RGB(0, 0, 128) ' Default color (Navy Blue)
                                        End If
                                     .Font.Italic = msoTrue
                                      .ParagraphFormat.Alignment = ppAlignLeft
                                      .Font.Size = 11
                              End With
                Else
                                   Set piedShape = slide.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, slide.Master.Height - 20, slide.Master.Width, 30)
                                          With piedShape.TextFrame.TextRange
                                              .text = tabF(0)
                                              .Font.Size = 11
                                              .Font.Italic = msoTrue
                                              .ParagraphFormat.Alignment = ppAlignLeft
'                                               If slide.slideIndex = 1 Then
'                                                    .Font.color.RGB = RGB(255, 255, 255) ' White
'                                                Else
'                                                    .Font.color.RGB = RGB(0, 0, 128) ' Default color (Navy Blue)
'                                                End If
                                                
                                                Bln_Fond_White = False
                                                For k = 0 To UBound(Var_Indice_BackUp)
                                                     If slide.slideIndex = CDbl(Var_Indice_BackUp(k)) Then
                                                         .Font.color.RGB = RGB(255, 255, 255) ' White
                                                         Bln_Fond_White = True
                                                         Exit For
                                                     End If
                                                 Next k
                                                 
                                                 If Not Bln_Fond_White Then
                                                     .Font.color.RGB = RGB(0, 0, 128) ' Default color (Navy Blue)
                                                 End If
                                                 
                                          End With
                  End If
        End If
         
         If UCase(tabF(1)) = UCase("Domains") Then
           
            'Set piedShape = slide.Shapes.title
                If hasTitle Then
                 Set piedShape = slide.Shapes.Title
                    ' Set the text, font, boldness, and size
                    With piedShape.TextFrame.TextRange
                            .text = tabF(0)
'                             If slide.slideIndex = 1 Then
'                                                    .Font.color.RGB = RGB(255, 255, 255) ' White
'                                                Else
'                                                    .Font.color.RGB = RGB(0, 0, 128) ' Default color (Navy Blue)
'                                                End If
                            
                            Bln_Fond_White = False
                            For k = 0 To UBound(Var_Indice_BackUp)
                                 If slide.slideIndex = CDbl(Var_Indice_BackUp(k)) Then
                                     .Font.color.RGB = RGB(255, 255, 255) ' White
                                     Bln_Fond_White = True
                                     Exit For
                                 End If
                             Next k
                             
                             If Not Bln_Fond_White Then
                                 .Font.color.RGB = RGB(0, 0, 128) ' Default color (Navy Blue)
                             End If
                                                 
                            .Font.Italic = msoTrue
                            .ParagraphFormat.Alignment = ppAlignCenter
                            .Font.Size = 11
                    End With
                Else
                             Set piedShape = slide.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, slide.Master.Height - 20, slide.Master.Width, 30)
                                    With piedShape.TextFrame.TextRange
                                       If UCase(tabF(0)) = "DRIVABILITY-DYNAMISM" Then
                                        .text = "Drivability Assessment"
                                        Else
                                        .text = tabF(0)
                                        End If
                                        .Font.Size = 11
                                        .Font.Italic = msoTrue
                                        .ParagraphFormat.Alignment = ppAlignCenter
'                                        If slide.slideIndex = 1 Then
'                                            .Font.color.RGB = RGB(255, 255, 255) ' White
'                                        Else
'                                            .Font.color.RGB = RGB(0, 0, 128) ' Default color (Navy Blue)
'                                        End If
                                        
                                        Bln_Fond_White = False
                                        For k = 0 To UBound(Var_Indice_BackUp)
                                             If slide.slideIndex = CDbl(Var_Indice_BackUp(k)) Then
                                                 .Font.color.RGB = RGB(255, 255, 255) ' White
                                                 Bln_Fond_White = True
                                                 Exit For
                                             End If
                                         Next k
                                         
                                         If Not Bln_Fond_White Then
                                             .Font.color.RGB = RGB(0, 0, 128) ' Default color (Navy Blue)
                                         End If
                             
                                    End With
                      End If
        End If
          If UCase(tabF(1)) = UCase("Signet7") Then
            ' Ajouter le pied de page
          

' If title shape exists, apply text and formatting
                If hasTitle Then
                 Set piedShape = slide.Shapes.Title
                    With piedShape.TextFrame.TextRange
                        .text = tabF(0) ' Assign the text from your array
'                         If slide.slideIndex = 1 Then
'                                            .Font.color.RGB = RGB(255, 255, 255) ' White
'                                        Else
'                                            .Font.color.RGB = RGB(0, 0, 128) ' Default color (Navy Blue)
'                                        End If ' Set font color to blue
                        
                        Bln_Fond_White = False
                        For k = 0 To UBound(Var_Indice_BackUp)
                             If slide.slideIndex = CDbl(Var_Indice_BackUp(k)) Then
                                 .Font.color.RGB = RGB(255, 255, 255) ' White
                                 Bln_Fond_White = True
                                 Exit For
                             End If
                         Next k
                         
                         If Not Bln_Fond_White Then
                             .Font.color.RGB = RGB(0, 0, 128) ' Default color (Navy Blue)
                         End If
                         
                        .Font.Italic = msoTrue ' Set italic style
                        .ParagraphFormat.Alignment = ppAlignCenter ' Align center
                        .Font.Size = 11 ' Set font size to 12
                    End With
                Else
                     Set piedShape = slide.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, slide.Master.Height - 20, slide.Master.Width, 30)
                            
                            With piedShape.TextFrame.TextRange
                                '.text = tabF(0)
                                .text = "C2-Sensible"
                                .Font.Size = 11
                                .Font.Italic = msoTrue
                                .ParagraphFormat.Alignment = ppAlignRight
'                                 If slide.slideIndex = 1 Then
'                                            .Font.color.RGB = RGB(255, 255, 255) ' White
'                                 Else
'                                            .Font.color.RGB = RGB(0, 0, 128) ' Default color (Navy Blue)
'                                 End If
                                
                                Bln_Fond_White = False
                                For k = 0 To UBound(Var_Indice_BackUp)
                                     If slide.slideIndex = CDbl(Var_Indice_BackUp(k)) Then
                                         .Font.color.RGB = RGB(255, 255, 255) ' White
                                         Bln_Fond_White = True
                                         Exit For
                                     End If
                                 Next k
                                 
                                 If Not Bln_Fond_White Then
                                     .Font.color.RGB = RGB(0, 0, 128) ' Default color (Navy Blue)
                                 End If
                                 
                            End With
                
                End If
         End If
         
        Next i
        End If
    Next slide
End Sub


'Function RemplirPiedPage(ByRef objPPT As Object, objPresentation As Object, Fields As Variant)
'    Dim slide As Object
'    Dim piedShape As Object
'    Dim piedTexte As String
'    Dim tabF() As String
'    Dim i As Integer
'    Dim hasTitle As Boolean
'
' hasTitle = False ' Initialize the flag
'
'
'
'    ' Boucle à travers chaque diapositive
'For Each slide In objPresentation.Slides
'      ' Loop through the shapes in the slide to check for the title placeholder
'    For i = 1 To slide.Shapes.Count
'        If slide.Shapes(i).Type = msoPlaceholder Then
'            If slide.Shapes(i).PlaceholderFormat.Type = ppPlaceholderTitle Then
'                Set piedShape = slide.Shapes(i)
'                hasTitle = True
'                Exit For
'            End If
'        End If
'    Next i
'
'        For i = 1 To UBound(Fields)
'          tabF = Split(Fields(i), "#")
'          If UCase(tabF(1)) = UCase("DepTV") Then
'
'                If hasTitle Then
'                              Set piedShape = slide.Shapes.Title
'
'                              ' Set the text, font, boldness, and size
'                              With piedShape.TextFrame.TextRange
'                                      .text = tabF(0)
'                                      ' Test if tabF(0) equals 1 to set the font color to white
'                                       If slide.slideIndex = 1 Then
'                                            .Font.color.RGB = RGB(255, 255, 255) ' White
'                                        Else
'                                            .Font.color.RGB = RGB(0, 0, 128) ' Default color (Navy Blue)
'                                        End If
'                                     .Font.Italic = msoTrue
'                                      .ParagraphFormat.Alignment = ppAlignLeft
'                                      .Font.Size = 11
'                              End With
'                Else
'                                   Set piedShape = slide.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, slide.Master.Height - 30, slide.Master.Width, 30)
'                                          With piedShape.TextFrame.TextRange
'                                              .text = tabF(0)
'                                              .Font.Size = 11
'                                              .Font.Italic = msoTrue
'                                              .ParagraphFormat.Alignment = ppAlignLeft
'                                               If slide.slideIndex = 1 Then
'                                                    .Font.color.RGB = RGB(255, 255, 255) ' White
'                                                Else
'                                                    .Font.color.RGB = RGB(0, 0, 128) ' Default color (Navy Blue)
'                                                End If
'                                          End With
'                  End If
'        End If
'
'         If UCase(tabF(1)) = UCase("Domains") Then
'
'            'Set piedShape = slide.Shapes.title
'                If hasTitle Then
'                 Set piedShape = slide.Shapes.Title
'                    ' Set the text, font, boldness, and size
'                    With piedShape.TextFrame.TextRange
'                            .text = tabF(0)
'                             If slide.slideIndex = 1 Then
'                                                    .Font.color.RGB = RGB(255, 255, 255) ' White
'                                                Else
'                                                    .Font.color.RGB = RGB(0, 0, 128) ' Default color (Navy Blue)
'                                                End If
'                            .Font.Italic = msoTrue
'                            .ParagraphFormat.Alignment = ppAlignCenter
'                            .Font.Size = 11
'                    End With
'                Else
'                             Set piedShape = slide.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, slide.Master.Height - 30, slide.Master.Width, 30)
'                                    With piedShape.TextFrame.TextRange
'                                        .text = tabF(0)
'                                        .Font.Size = 11
'                                        .Font.Italic = msoTrue
'                                        .ParagraphFormat.Alignment = ppAlignCenter
'                                        If slide.slideIndex = 1 Then
'                                            .Font.color.RGB = RGB(255, 255, 255) ' White
'                                        Else
'                                            .Font.color.RGB = RGB(0, 0, 128) ' Default color (Navy Blue)
'                                        End If
'                                    End With
'                      End If
'        End If
'          If UCase(tabF(1)) = UCase("Signet7") Then
'            ' Ajouter le pied de page
'
'
'' If title shape exists, apply text and formatting
'                If hasTitle Then
'                 Set piedShape = slide.Shapes.Title
'                    With piedShape.TextFrame.TextRange
'                        .text = tabF(0) ' Assign the text from your array
'                         If slide.slideIndex = 1 Then
'                                            .Font.color.RGB = RGB(255, 255, 255) ' White
'                                        Else
'                                            .Font.color.RGB = RGB(0, 0, 128) ' Default color (Navy Blue)
'                                        End If ' Set font color to blue
'                        .Font.Italic = msoTrue ' Set italic style
'                        .ParagraphFormat.Alignment = ppAlignCenter ' Align center
'                        .Font.Size = 11 ' Set font size to 12
'                    End With
'                Else
'                     Set piedShape = slide.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, slide.Master.Height - 30, slide.Master.Width, 30)
'                            With piedShape.TextFrame.TextRange
'                                .text = tabF(0)
'                                .Font.Size = 11
'                                .Font.Italic = msoTrue
'                                .ParagraphFormat.Alignment = ppAlignRight
'                                 If slide.slideIndex = 1 Then
'                                            .Font.color.RGB = RGB(255, 255, 255) ' White
'                                 Else
'                                            .Font.color.RGB = RGB(0, 0, 128) ' Default color (Navy Blue)
'                                 End If
'                            End With
'
'                End If
'         End If
'        Next i
'
'    Next slide
'End Function
Function RemplirEnteteEtPiedPage(ByRef objppt As Object, objPresentation As Object, Fields As Variant, nameSdv As String, num As String)
    Dim slide As Object
    Dim enteteShape As shape
    Dim piedShape As shape
    Dim enteteTexte As String
    Dim piedTexte As String
    Dim tabF() As String
     
    ' Boucle à travers chaque diapositive
    For Each slide In objPresentation.Slides
        ' Ajouter l'en-tête
        Set enteteShape = slide.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, slide.Master.Width, 30)
        With enteteShape.TextFrame.TextRange
            .text = nameSdv
            .Font.Size = 14
            .Font.Bold = msoTrue
            .ParagraphFormat.Alignment = ppAlignLeft
        End With
        
        For i = 1 To UBound(Fields)
          tabF = Split(Fields(i), "#")
          If UCase(tabF(1)) = UCase("DepTV") Then
            ' Ajouter le pied de page
            Set piedShape = slide.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, slide.Master.Height - 20, slide.Master.Width, 30)
            With piedShape.TextFrame.TextRange
                .text = tabF(0)
                .Font.Size = 12
                .Font.Italic = msoTrue
                .ParagraphFormat.Alignment = ppAlignLeft
            End With
         End If
         
         If UCase(tabF(1)) = UCase("Domains") Then
            ' Ajouter le pied de page
            Set piedShape = slide.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, slide.Master.Height - 20, slide.Master.Width, 30)
            With piedShape.TextFrame.TextRange
                .text = tabF(0)
                .Font.Size = 12
                .Font.Italic = msoTrue
                .ParagraphFormat.Alignment = ppAlignCenter
            End With
         End If
        Next i
        
    Next slide
End Function
Function InspectShapes(objppt As Object, objPres As Object)
    Dim objslide As Object
    Dim shp As Object
    Dim i As Integer
    Dim shapeInfo As String
    
    ' Set the slide number (e.g., 1st slide)
    Set objslide = objPres.Slides(2)
    
    ' Loop through all shapes on the slide
    For i = 1 To objslide.Shapes.Count
        Set shp = objslide.Shapes(i)
        
        ' Collect shape details
        shapeInfo = "Shape " & i & " Type: "
        
        ' Check if it's a text shape
        If shp.HasTextFrame Then
            If shp.TextFrame.HasText Then
                shapeInfo = shapeInfo & "Text - " & shp.TextFrame.TextRange.text & vbCrLf
            Else
                shapeInfo = shapeInfo & "Empty Text Frame" & vbCrLf
            End If
        End If
        
        ' Check if it's a picture
        If shp.Type = msoPicture Then
            shapeInfo = shapeInfo & "Picture" & vbCrLf
        End If
        
        ' Check if it's a table
        If shp.HasTable Then
            shapeInfo = shapeInfo & "Table with " & shp.table.Rows.Count & " rows and " & shp.table.Columns.Count & " columns" & vbCrLf
        End If
        
        ' Check if it's a group
        If shp.Type = msoGroup Then
            shapeInfo = shapeInfo & "Group of shapes" & vbCrLf
        End If
        
        ' Display the shape info in a message box or debug window
        Debug.Print shapeInfo
    Next i
End Function
Sub GetTableDetails(objppt As Object, objPres As Object)
    Dim objslide As Object
    Dim shp As Object
    Dim cell As Object
    Dim i As Integer
    Dim j As Integer
    Dim rowsCount As Integer
    Dim colsCount As Integer
    Dim shapeInfo As String
    Dim k As Integer
    
    Dim sourceCell As Range
    Dim subtitleText As String
    Dim cellColor As Long
    Dim textColor As Long
     
    ' Set the slide to inspect (e.g., slide 2)
    Set objslide = objPres.Slides(2)
    
     ' Set the source cell (F4) in Excel
    Set sourceCell = F_Rating.Range("F4")
    
    ' Get the text from the source cell
    subtitleText = sourceCell.Value
    
      ' Get the background color (Interior Color) of the cell
    cellColor = sourceCell.Interior.color
    
     ' Get the text color of the source cell
    textColor = sourceCell.Font.color
    
    
    ' Loop through all shapes on the slide
    For i = 1 To objslide.Shapes.Count
        Set shp = objslide.Shapes(4)
        
        ' Check if the shape contains a table
        If shp.HasTable Then
            ' Get the number of rows and columns
            rowsCount = shp.table.Rows.Count
            colsCount = shp.table.Columns.Count
            
            ' Display the information for each table
            shapeInfo = "Shape " & i & " is a table with " & rowsCount & " rows and " & colsCount & " columns."
            'MsgBox shapeInfo
            
            ' Insert text into each cell of the table
            'For j = 1 To rowsCount
                'For K = 1 To colsCount
                    ' Modify this line to insert your desired text
                    
                    shp.table.cell(1, 2).shape.TextFrame.TextRange.text = subtitleText
                  
                    ' Set the background color of the table cell
                    shp.table.cell(1, 2).shape.Fill.ForeColor.RGB = cellColor
                    
                    ' Set the text color for the cell
                   shp.table.cell(1, 2).shape.TextFrame.TextRange.Font.color = textColor
               ' Next K
           ' Next j
        Else
            ' Optionally, you can output information about non-table shapes
            Debug.Print "Shape " & i & " is not a table."
        End If
    Next i
End Sub

Function GetTableDetails01(objppt As Object, objPres As Object)
    Dim objslide As Object
    Dim shp As Object
    Dim i As Integer
    Dim rowsCount As Integer
    Dim colsCount As Integer
    Dim shapeInfo As String
    
    ' Set the slide to inspect (e.g., slide 1)
    Set objslide = objPres.Slides(2)
    
    ' Loop through all shapes on the slide
    For i = 1 To objslide.Shapes.Count
        Set shp = objslide.Shapes(i)
        
        ' Check if the shape contains a table
        If shp.HasTable Then
            ' Get the number of rows and columns
            rowsCount = shp.table.Rows.Count
            colsCount = shp.table.Columns.Count
            
            ' Display the information for each table
            shapeInfo = "Shape " & i & " is a table with " & rowsCount & " rows and " & colsCount & " columns."
            MsgBox shapeInfo
        Else
            ' Optionally, you can output information about non-table shapes
            Debug.Print "Shape " & i & " is not a table."
        End If
    Next i
End Function
Function InsertTextInGreenRow(objppt As Object, objPres As Object)
    Dim objslide As Object
    Dim shp As Object
    Dim tableIndex As Integer
    Dim targetTable As table
    Dim row As Integer
    Dim col As Integer
    
    ' Set the slide where the table is located (e.g., slide 1)
    Set objslide = objPres.Slides(2)
    
    ' Initialize table counter
    tableIndex = 0
    
    ' Loop through the shapes to find the 2nd table
    For Each shp In objslide.Shapes
        If shp.HasTable Then
            tableIndex = tableIndex + 1
            
            ' Check if it's the second table
            If tableIndex = 1 Then
                Set targetTable = shp.table
                Exit For
            End If
        End If
    Next shp
    
    ' Check if the target table was found
    If Not targetTable Is Nothing Then
        ' Insert text into the green row (assuming it's row 2 and column 2)
        row = 2 ' Assuming the green row is row 2
        col = 1 ' You can adjust the column number accordingly
        
        ' Insert text into the targeted cell
        targetTable.cell(row, col).shape.TextFrame.TextRange.text = "Inserted Green Row Text"
    Else
        MsgBox "Table 2 not found on this slide!"
    End If
End Function


Function InsertTextInTable01(objppt As Object, objPres As Object)
    Dim objslide As Object
    Dim shp As Object
    Dim tableIndex As Integer
    Dim targetTable As table
    Dim row As Integer
    Dim col As Integer
    
    ' Set the slide where the table is located (e.g., slide 1)
    Set objslide = objPres.Slides(2)
    
    ' Loop through the shapes to find the 2nd table (tableIndex = 2)
    tableIndex = 0 ' Initialize table counter
    
    For Each shp In objslide.Shapes
        ' Check if the shape is a table
        If shp.HasTable Then
            tableIndex = tableIndex + 1 ' Increment table index
            
            ' Check if it's the second table
            If tableIndex = 2 Then
                Set targetTable = shp.table ' Get the reference to the second table
                Exit For ' Exit loop after finding the table
            End If
        End If
    Next shp
    
    ' If the second table was found, insert text in column 1
    If Not targetTable Is Nothing Then
        col = 1 ' Specify column 1
        
        ' Loop through all rows in the table and insert text in column 1
        For row = 1 To targetTable.Rows.Count
            targetTable.cell(row, col).shape.TextFrame.TextRange.text = "Inserted Text " & row
        Next row
    Else
        MsgBox "Table 2 not found on this slide!"
    End If
End Function

Sub RemplissageTable(objppt As Object, objPres As Object, Fields As Variant)
   Dim objslide As slide
    Dim i As Integer
    Dim tabF() As String
    Dim projectText As String
    Dim StagesText As String
    Dim DepTVText As String
    
    
    Call GetTableDetails(objppt, objPres)
    
   
    Set objslide = objPres.Slides(2)
    
    For i = 1 To UBound(Fields)
        tabF = Split(Fields(i), "#")
        If InStr(1, tabF(1), ";") <> 0 Then
            projectText = tabF(0)
            UpdateContentText objslide.Shapes(2), projectText, "Projects"
        ElseIf UCase(tabF(1)) = UCase("Stages") Then
            StagesText = tabF(0)
            UpdateContentText objslide.Shapes, StagesText, "Stages"
      
        End If
        
    Next i
    
End Sub
Function UpdateContentText(ByRef shape As Object, ByVal text As String, typeText As String)

If typeText = "Projects" Then
    With shape.TextFrame
        .TextRange.text = text
        .TextRange.Font.color.RGB = RGB(255, 255, 255)
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Size = 16
        .TextRange.Font.Name = "Encode Sans (Corps)"
    End With
ElseIf typeText = "Stages" Then
  With shape.AddTextbox(msoTextOrientationHorizontal, 20, 125, 700, 50).TextFrame.TextRange
        .text = text
        .Font.color.RGB = RGB(255, 255, 255)
        .Font.Bold = msoTrue
        .Font.Size = 12
        .Font.Name = "Encode Sans (Corps)"
  End With

End If

End Function
Sub CreateSingleTextboxLayout(objppt As Object, objPres As Object, Fields As Variant)
    Dim slide As Object
    Dim shapeText As Object
    Dim textContent As String
    Dim projectText As String
    Dim LocDatTestText As String
    Dim AcStatus As String
    Dim vehOption As String
     
    Dim tabF() As String
    Dim i As Integer
    
    ' Référencer la diapositive actuelle (8ème slide)
    '''''''''''''''''Modif 03/02/2025
    
    If numSlide = 5 Then
        Set slide = objPres.Slides(10)
    Else
        Set slide = objPres.Slides(11)
    End If
    
    ' Parcourir les champs passés en paramètres
    For i = 1 To UBound(Fields)
        ' Convertir Fields(i) en chaîne pour éviter les erreurs de type
        tabF = Split(CStr(Fields(i)), "#")
        
        If UCase(tabF(1)) = UCase("LocDatTest") Then
            LocDatTestText = tabF(0)
        ElseIf UCase(tabF(1)) = UCase("AcStatus") Then
            AcStatus = tabF(0)
         ' Vérifier que le deuxième élément contient un point-virgule
        ElseIf UCase(tabF(1)) = UCase("vehOption") Then
            vehOption = tabF(0)
            
            ' Construire le texte avec des tabulations pour aligner les valeurs
            textContent = "Location and date of test: " & LocDatTestText & vbTab & vbTab & "Vehicle weight: 2000 kg" & vbCrLf & _
                          "A/C status: " & AcStatus & vbTab & vbTab & "Tyre dimensions: 235/55 R19" & vbCrLf & _
                          "Vehicle mileage at start of test: " & vehOption & vbTab & vbTab & "V@1000rpm:"
            
            ' Ajouter la zone de texte avec dimensions définies
            Set shapeText = slide.Shapes.AddTextbox(msoTextOrientationHorizontal, 15, 70, 750, 150)
            With shapeText.TextFrame.TextRange
                .text = textContent
                .Font.Size = 14
                .Font.color.RGB = RGB(0, 0, 0)
                .Font.Bold = msoTrue
            End With
            
            ' Ajouter les tabulations pour l'alignement
            With shapeText.TextFrame.Ruler.TabStops
                .Add ppTabStopRight, 650 ' Position à droite
            End With
        End If
    Next i
End Sub

Function Remplissage_Cartouche(ByRef objppt As Object, objPres As Object, Fields As Variant)
    Dim objslide As slide
    Dim i As Integer
    Dim tabF() As String
    Dim projectText As String
    Dim CarNumbText As String
    Dim StandardText As String
    Dim StagesText As String
    Dim GoalsText As String
    Dim DepTVText As String
    Dim SendFromText As String
    Dim TelFromText As String
    Dim MailFromText As String
    Dim LocDatTestText As String
    
    ' Sélectionner la première diapositive
    Set objslide = objPres.Slides(1)
    
    ' Boucle sur les champs fournis
    For i = 1 To UBound(Fields)
         tabF = Split(Fields(i), "#")
        
        ' Vérifier si le champ contient un point-virgule
        If InStr(1, tabF(1), ";") <> 0 Then
            projectText = tabF(0)
            ' Mettre à jour la forme avec le texte du projet
           UpdateText objslide.Shapes(6), projectText & vbCrLf, "", "Zone1"
        
       ElseIf UCase(tabF(1)) = UCase("Stages") Then
            StagesText = tabF(0)
            StagesText = "Cal maturity :" & StagesText
            UpdateText objslide.Shapes(6), projectText & vbCrLf & StagesText, "Zone1"
            
        ElseIf UCase(tabF(1)) = UCase("Goals") Then
            GoalsText = tabF(0)
            GoalsText = "Assessment type: " & GoalsText
             UpdateText objslide.Shapes(6), projectText & vbCrLf & StagesText & vbCrLf & GoalsText, "Zone1"
             
        ElseIf UCase(tabF(1)) = UCase("DepTV") Then
            DepTVText = tabF(0)
              UpdateText objslide.Shapes(3), DepTVText & vbCrLf & SendFromText & vbCrLf & TelFromText & vbCrLf & MailFromText & vbCrLf & LocDatTestText, "Zone3"
        
        ElseIf UCase(tabF(1)) = UCase("SendFrom") Then
            SendFromText = tabF(0)
            UpdateText objslide.Shapes(3), DepTVText & vbCrLf & SendFromText, "Zone4"
        
        ElseIf UCase(tabF(1)) = UCase("TelFrom") Then
            TelFromText = tabF(0)
             UpdateText objslide.Shapes(3), DepTVText & vbCrLf & SendFromText & vbCrLf & TelFromText, "Zone4"
        
        ElseIf UCase(tabF(1)) = UCase("MailFrom") Then
            MailFromText = LCase(tabF(0)) ' Convert email to lowercase
             UpdateText objslide.Shapes(3), DepTVText & vbCrLf & SendFromText & vbCrLf & TelFromText & vbCrLf & MailFromText, "Zone4"
        
        ElseIf UCase(tabF(1)) = UCase("Signet7") Then
            LocDatTestText = tabF(0)
            UpdateText objslide.Shapes(3), DepTVText & vbCrLf & SendFromText & vbCrLf & TelFromText & vbCrLf & MailFromText & vbCrLf & LocDatTestText, "Zone2"
        End If
    Next i
End Function


Function UpdateText(ByRef shape As Object, ByVal text As String, Optional ByVal style As String = "", Optional typeOfText As String, Optional ByRef tabF As Variant)
   With shape.TextFrame
  
        .TextRange.text = text
      
        ' Default font settings
        .TextRange.Font.color.RGB = RGB(255, 255, 255) ' White text
        '.TextRange.Font.Bold = msoTrue ' Bold text
        
        ' Style-specific settings based on provided style parameter
        Select Case style
            Case "Title"
                .TextRange.Font.Size = 18 ' Larger font for title
                .TextRange.ParagraphFormat.Alignment = ppAlignCenter ' Center alignment for title
                .TextRange.Font.Name = "Encode Sans Expanded Light"
                
            Case "Zone1"
                .TextRange.Font.Size = 18 ' Smaller font for subtitles
                .TextRange.Font.color.RGB = RGB(255, 255, 255)
                If typeOfText = UCase("Standars") Then
                .TextRange.Font.Name = "Encode Sans Expanded Light"
                Else
                 .TextRange.Font.Name = "Encode Sans (Corps)"
                End If
                '.MarginLeft = 20  ' Adjust as needed
                '.MarginRight = 10
                '.MarginTop = 10
                '.MarginBottom = 0
                .TextRange.ParagraphFormat.Alignment = ppAlignLeft
                
                
            Case "Zone2"
               
                .TextRange.Font.Size = 18 '
                .TextRange.Font.color.RGB = RGB(255, 255, 255)
                'shape.Line.Visible = msoFalse ' Hide borders if necessary
                ' Adjusting margins (padding)
                '.MarginLeft = 20  ' Adjust as needed
                '.MarginRight = 10
                '.MarginTop = 10
                '.MarginBottom = 5
                .TextRange.Font.Name = "Encode Sans (Corps)"
                .TextRange.ParagraphFormat.Alignment = ppAlignLeft
                
             Case "Zone3"
             
                
                .TextRange.Font.color.RGB = RGB(255, 255, 255)
                .TextRange.Font.Name = "Encode Sans (Corps)"
                .TextRange.ParagraphFormat.Alignment = ppAlignLeft
                ' Special case for email, adjust size
                If InStr(1, text, "@") > 0 Then ' Check if the text contains an email
                    .TextRange.Font.Size = 11 ' Set email font size to 11
                Else
                .TextRange.Font.Size = 18 ' Default for other Zone3 text
                End If
                
            Case "Zone4"
                .TextRange.Font.color.RGB = RGB(255, 255, 255)
                .TextRange.Font.Name = "Encode Sans (Corps)"
                .TextRange.ParagraphFormat.Alignment = ppAlignLeft
                ' Special case for email, adjust size
                
                .TextRange.Font.Size = 11
               
                
            Case Else
                .TextRange.Font.Size = 18 ' Default font size for other text
                .TextRange.Font.Name = "Encode Sans (Corps)"
        End Select

        
    End With

End Function

' Sous-fonction pour la mise à jour des formes


'Function Remplissage_Cartouche(ByRef objPPT As Object, objPres As Object, fields As Variant)
'    Dim objSlide As slide
'    Dim enteteShape As shape
'    Dim piedShape As shape
'    Dim i As Integer
'    Dim j As Integer
'    Dim tabF() As String
'    Dim tabB() As String
'    Dim projectText As String
'    Dim CarNumbText As String
'    Dim StandardText As String
'
'
'
'      Set objSlide = objPres.Slides(1)
'
'        For i = 1 To UBound(fields)
'            tabF = Split(fields(i), "#")
'            If InStr(i, tabF(1), ";") <> 0 Then
'                 tabB = Split(tabF(1), ";")
'                 projectText = tabF(0)
'                 For j = 0 To UBound(tabB)
'                     With objSlide.Shapes(2).TextFrame.TextRange
'                        .Text = tabF(0)
'                        .Font.color.RGB = RGB(255, 255, 255)
'                        .Font.Bold = msoTrue
'                        .Font.Size = 18
'                    End With
'                 Next j
'             End If
'            ElseIf tabF(1) = "TetTab1_2" Then
'            CarNumbText = tabF(0)
'                     With objSlide.Shapes(6).TextFrame.TextRange
'                        .Text = projectText & vbCrLf & CarNumbText
'                        .Font.color.RGB = RGB(255, 255, 255)
'                        .Font.Bold = msoTrue
'                        .Font.Size = 18
'                    End With
'             End If
'             ElseIf tabF(1) = "Standars" Then
'             StandardText = tabF(0)
'                     With objSlide.Shapes(6).TextFrame.TextRange
'                        .Text = projectText & vbCrLf & CarNumbText & vbCrLf & StandardText
'                        .Font.color.RGB = RGB(255, 255, 255)
'                        .Font.Bold = msoTrue
'                        .Font.Size = 18
'                    End With
'             End If
'
'
'       Next i
'
'End Function
Function Remplissage_PPT002(Fields As Variant, objppt As Object, objPres As Object)
    Dim i As Integer
    Dim tabF() As String
    Dim objslide As Object
    Dim shapeTextRange As Object
    Dim shapeName As String

    ' Set the slide (assuming the cartouche is on slide 1)
    Set objslide = objPres.Slides(1)

    ' Loop through each field to fill the cartouche
    For i = 1 To UBound(Fields)
        tabF = Split(Fields(i), "#")
        shapeName = tabF(1)

        ' Optimized to handle only the cartouche shapes
        Select Case shapeName
            Case "Signet7"
                ' Set the date in shape 4
                Set shapeTextRange = objslide.Shapes(4).TextFrame.TextRange
                With shapeTextRange
                    .text = CStr(Date)
                    .Font.Bold = msoTrue
                    .Font.Size = 18
                End With

                ' Set additional text in shape 5
                Set shapeTextRange = objslide.Shapes(5).TextFrame.TextRange
                With shapeTextRange
                    .text = tabF(0) & " " & CStr(Date)
                    .Font.Bold = msoTrue
                    .Font.Size = 18
                End With

            Case Else
                ' General case for other cartouche shapes (assumed to be shape 3)
                Set shapeTextRange = objslide.Shapes(3).TextFrame.TextRange
                With shapeTextRange
                    .text = tabF(0) & vbCrLf & "Additional String" & vbCrLf & "Another String"
                    ' Format the first string
                    With .Characters(1, Len(tabF(0))).Font
                        .Bold = msoTrue
                        .Size = 18
                    End With
                    ' Format the second string
                    With .Characters(Len(tabF(0)) + 2, Len("Additional String")).Font
                        .Bold = msoFalse
                        .Size = 14
                    End With
                    ' Format the third string
                    With .Characters(Len(tabF(0)) + Len("Additional String") + 4, Len("Another String")).Font
                        .Bold = msoTrue
                        .Size = 12
                    End With
                End With
        End Select
    Next i
End Function

Function Remplissage_PPT(Fields As Variant, objppt As Object, objPres As Object)
    Dim i As Integer
    Dim j As Integer
    Dim tabF() As String
    Dim tabB() As String
    Dim objslide As Object
    Dim shapeTextRange As Object
    Dim shapeName As String

    ' Set the slide
    Set objslide = objPres.Slides(1)

    ' Loop through each field
    For i = 1 To UBound(Fields)
        tabF = Split(Fields(i), "#")
        shapeName = tabF(1)

        If InStr(1, shapeName, ";") > 0 Then
            ' If multiple shapes are specified
            tabB = Split(shapeName, ";")
            For j = 0 To UBound(tabB)
                ' Set the text and format for each shape
                Set shapeTextRange = objslide.Shapes(2).TextFrame.TextRange
                With shapeTextRange
                    .text = tabF(0)
                    .Font.Bold = msoTrue
                    .Font.Size = 18
                End With
            Next j
        Else
            ' Handle special case for "Signet7"
            If shapeName = "Signet7" Then
                Set shapeTextRange = objslide.Shapes(4).TextFrame.TextRange
                With shapeTextRange
                    .text = CStr(Date)
                    .Font.Bold = msoTrue
                    .Font.Size = 18
                End With

                shapeName = tabF(0)
                Set shapeTextRange = objslide.Shapes(5).TextFrame.TextRange
                With shapeTextRange
                    .text = tabF(0) & " " & CStr(Date)
                    .Font.Bold = msoTrue
                    .Font.Size = 18
                End With
            Else
                ' Handle general case
                Set shapeTextRange = objslide.Shapes(3).TextFrame.TextRange
                With shapeTextRange
                    .text = tabF(0)
                    .Font.Bold = msoTrue
                    .Font.Size = 18
                End With
            End If
        End If
    Next i
End Function

Function Remplissage_PPT4(Fields As Variant, objppt As Object, objPres As Object)
        Dim i As Integer
        Dim j As Integer
        Dim tabF() As String
        Dim tabB() As String
        Dim objslide As Object
        
         Set objslide = objPres.Slides(1)
         
        For i = 1 To UBound(Fields)
            tabF = Split(Fields(i), "#")
            If InStr(1, tabF(1), ";") <> 0 Then
                 tabB = Split(tabF(1), ";")
                 For j = 0 To UBound(tabB)
                     'objDoc.Bookmarks(tabB(j)).Select
                     'objword.Selection.TypeText Text:=tabF(0)
                    
                     With objslide.Shapes(tabB(j)).TextFrame.TextRange
                        .text = tabF(0)
                        '.Font.color.RGB = RGB(0, 0, 255)
                        .Font.Bold = msoTrue
                        .Font.Size = 18
                    End With
                 Next j
            Else
                   'objDoc.Bookmarks(tabF(1)).Select
                   If tabF(1) = "Signet7" Then
                       ' objword.Selection.TypeText Text:=CStr(Date)
                       ' objDoc.Bookmarks("Places").Select
                       ' objword.Selection.TypeText Text:=tabF(0) & " " & CStr(Date)
                      With objslide.Shapes(tabF(1)).TextFrame.TextRange
                        .text = CStr(Date)
                        '.Font.color.RGB = RGB(0, 0, 255)
                        .Font.Bold = msoTrue
                        .Font.Size = 18
                       End With
                       With objslide.Shapes(tabF(0)).TextFrame.TextRange
                        .text = tabF(0) & " " & CStr(Date)
                        '.Font.color.RGB = RGB(0, 0, 255)
                        .Font.Bold = msoTrue
                        .Font.Size = 18
                       End With
                       
                       
                   Else
                        'objword.Selection.TypeText Text:=tabF(0)
                          With objslide.Shapes(tabF(0)).TextFrame.TextRange
                        .text = tabF(0)
                        '.Font.color.RGB = RGB(0, 0, 255)
                        .Font.Bold = msoTrue
                        .Font.Size = 18
                       End With
                   End If
            End If
        Next i
        
         
End Function

Function Remplissage_PPT3(Fields As Variant, objppt As Object, objslide As Object, slideIndex As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim tabF() As String
    Dim tabB() As String
    Dim shp As Object
    
    For i = 1 To UBound(Fields)
        tabF = Split(Fields(i), "#")
        
        If InStr(1, tabF(1), ";") <> 0 Then
            tabB = Split(tabF(1), ";")
            For j = 0 To UBound(tabB)
             On Error Resume Next
                Set shp = objslide.Shapes(tabB(j))
                 On Error GoTo 0
                If Not shp Is Nothing Then
                    If shp.HasTextFrame Then
                        If shp.TextFrame.HasText Then
                            shp.TextFrame.TextRange.text = tabF(0)
                        End If
                    End If
                End If

            Next j
        Else
            Set shp = objslide.Shapes(tabF(1))
            If Not shp Is Nothing Then
                If tabF(1) = "Signet7" Then
                    shp.TextFrame.TextRange.text = CStr(Date)
                    Set shp = objslide.Shapes("Places")
                    If Not shp Is Nothing Then
                        shp.TextFrame.TextRange.text = tabF(0) & " " & CStr(Date)
                    End If
                Else
                    shp.TextFrame.TextRange.text = tabF(0)
                End If
            End If
        End If
    Next i
End Function
Sub InsertPic_PPT_Format4(ByRef objppt As Object, objPres As Object, slideIndex As Integer)
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
    Dim objslide As Object
    Dim nbSlides As Integer
    Dim slideLayout As PpSlideLayout
    Dim pptShape As shape
    Dim pictureFound As Boolean
    Dim shape As Object
    Dim BlnAddSlide As Boolean
    Dim ShapeDriv As Object
    Dim objName As String
    Dim pasted, pasteSucceeded As Boolean
    Dim shp As PowerPoint.shape
    
    
    If numSlide = 5 Then
        slideIndex = 11
        For Each shp In objPres.Slides(10).Shapes
            If shp.Name = "Immagine 4" Then
                shp.Copy
                pasteSucceeded = False
                Do While Not pasteSucceeded
                    On Error Resume Next
                    objPres.Slides(5).Shapes.Paste
                    If ERR.Number = 0 Then
                        pasteSucceeded = True
                    Else
                        ERR.Clear
                        'DoEvents
                    End If
                    On Error GoTo 0
                Loop
                With objPres.Slides(5).Shapes(objPres.Slides(5).Shapes.Count)
                    .Left = shp.Left
                    .Top = 0.4
                End With
                Exit For
            End If
        Next shp
        
            
            
    ElseIf numSlide = 6 Then
        slideIndex = 12
        For Each shp In objPres.Slides(11).Shapes
            If shp.Name = "Immagine 4" Then
                shp.Copy
                pasteSucceeded = False
                Do While Not pasteSucceeded
                    On Error Resume Next
                    objPres.Slides(6).Shapes.Paste
                    If ERR.Number = 0 Then
                        pasteSucceeded = True
                    Else
                        ERR.Clear
                        'DoEvents
                    End If
                    On Error GoTo 0
                Loop
                With objPres.Slides(6).Shapes(objPres.Slides(6).Shapes.Count)
                    .Left = shp.Left
                    .Top = 0.4
                End With
                Exit For
            End If
        Next shp
    End If
    
    
    
    
    
    slideLayout = ppLayoutTitleOnly
        ViderPressePapiers
    pasted = False
    Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                'DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
    Loop
        
        ViderPressePapiers
                     slideIndex = slideIndex + 1
    
    Dim BlnGraphDriv As Boolean
    Dim BlnGraphDyna As Boolean
    
    For j = 0 To UBound(SDVListe)
        BlnGraphDriv = False
        BlnGraphDyna = False
        o = SDVListe(j)
        ThisWorkbook.sheets(o).Cells.EntireColumn.Hidden = False
        ThisWorkbook.sheets(o).UsedRange.EntireRow.Hidden = False
        'DoEvents
           pictureFound = False
           
    
        ViderPressePapiers
             
            pasted = False
            Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                'DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
            Loop
        
        ViderPressePapiers
                     slideIndex = slideIndex + 1
         
        For t = 1 To 2
            ProgressTitle ("Copie des données : " & o)
            If t = 1 Or (t = 2 And checkCriteriaDyn(o) = True And checkCorrespondancePriorityDyn(o) = True) Then
               
                If t = 1 Then
                    Call newSdvSlide_PPT(objslide, UCase(o), "2." & j + 1)
                    Call insertPart_PPT(objPres, objslide, 1)
                    Call insertPart_PPT(objPres, objslide, 2)
                    
                Else
                    Call newSdvSlide_PPT(objslide, UCase(o), "2." & j + 1)
                End If
                For i = 1 To 5
                    x = j & i & t

                        If i = 1 Then
                           
                            Call CopySummary_PPT(objppt, objslide, o, t, slideIndex)

                        ElseIf i = 2 Then
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
                                                                
                                                                If t = 1 Then
                                                                    BlnGraphDriv = True
                                                                    Call CopyGraph0_PPT(objppt, objslide, o, t, slideIndex, o)
                                                                Else
                                                                    BlnGraphDyna = True
                                                                    Call CopyGraph0_PPT(objppt, objslide, o, t, slideIndex, o)
                                                                End If
                                                           End If
                                                           Exit For
                                                      Else
                                                             If UpdateGraph.checkObject(CStr(Split(v(h), ":")(0)), o) <> "" Then
                                                                objName = UpdateGraph.checkObject(CStr(Split(v(h), ":")(0)), o)
                                                                
                                                                
                                                                
                                                                    Call LevGraphs(objppt, objslide, o, t, slideIndex, objName)

                                                                    

                                                            End If
                                                            Exit For
                                                      End If
                                                Next h
                                        End If
                                 
                                 Else
                                    If i = 2 Then
                                        BlnGraphDriv = True
                                        BlnGraphDyna = True
                                        
                                        Call CopyGraph0_PPT(objppt, objslide, o, t, slideIndex, o)
                                    End If
                                    
                                    
                                 End If
                                 
                      
                        End If
                        

                Next i
            End If
           
        
      Next t
      
              
      Call InsertPic_PPT_Format2(objppt, objPres, objslide, slideIndex, j, BlnGraphDriv, BlnGraphDyna)
      
        
                
                
      For t = 1 To 2
        For i = 4 To 5
          If i = 4 Then
                  
                  If t = 1 Then

                        Call CopyPriorityPoints_PPT3(objppt, objPres, objslide, "Hight", ThisWorkbook.Worksheets(o), i, slideIndex, o, j, BlnAddSlide)
                        
                        ThisWorkbook.sheets(o).UsedRange.EntireRow.Hidden = False
                        DoEvents
                  Else
                        
                        Call CopyPriorityPointsDyn_PPT3(objppt, objPres, objslide, "Hight", ThisWorkbook.Worksheets(o), i, slideIndex, o, j, BlnAddSlide)
              
                  End If
                  
              End If
        Next i
      Next t
      
    BlnAddSlide = False
    BlnAddSlideDriv = False
    BlnAddSlideDynam = False

    Next j

  

End Sub

Function LevGraphs(objppt As Object, objslide As Object, SDV As String, x As Integer, slideIndex As Integer, objName As String)
    
   
    Dim ws As Worksheet
    Dim success As Boolean
    Dim rng As Range
    Dim ca, cb As Integer
    Dim shp As Object
    
    Set ws = ThisWorkbook.Worksheets(SDV)
    
    If x = 1 Then
        
    

'DoEvents
'sleep 500

Set rng = ws.Range(objName)

If Not rng Is Nothing Then
success = False
 
   
    Do While Not success
        On Error Resume Next
 
        Range(objName).CopyPicture Appearance:=xlScreen, Format:=xlPicture
        

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop

'DoEvents
'sleep 500




Range("B2").Select


success = False

cb = objslide.Shapes.Count

Do While Not success
    objslide.Shapes.PasteSpecial DataType:=2
    ca = objslide.Shapes.Count
    success = ca > cb
Loop


Set shp = objslide.Shapes(objslide.Shapes.Count)

'DoEvents
With shp
    .LockAspectRatio = msoFalse
    .Left = 30
    .Top = 310
    .Width = 115
    .Height = 205
End With




End If
        
    ElseIf x = 2 Then
        
        
    
        'DoEvents
        'sleep 500
        
Set rng = ws.Range(objName)
If Not rng Is Nothing Then
success = False
Do While Not success
        On Error Resume Next
 
        
        Range(objName).CopyPicture Appearance:=xlScreen, Format:=xlPicture

        If ERR.Number = 0 Then
            success = True
        Else
            ERR.Clear
        End If
    Loop

        


Range("B2").Select




success = False

cb = objslide.Shapes.Count

Do While Not success
    objslide.Shapes.PasteSpecial DataType:=2
    ca = objslide.Shapes.Count
    success = ca > cb
Loop


Set shp = objslide.Shapes(objslide.Shapes.Count)

'DoEvents
With shp
    .LockAspectRatio = msoFalse
    .Left = 500
    .Top = 310
    .Width = 115
    .Height = 205
End With




End If
       
    End If
    
End Function

Function CopyPriorityPointsDyn_PPT33(objppt As Object, objPres As Object, objslide As Object, Filt As String, Shts As Worksheet, Pos As Integer, slideIndex As Integer)
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
    Dim pasted As Boolean
    
    
                     
                     
    With Shts
         For x = 72 To 74
            If .Cells(.Rows.Count, x).End(xlUp).row > TotalRow Then TotalRow = .Cells(.Rows.Count, x).End(xlUp).row
        Next x
      
       colonne = getLastColumnDinamyc(Shts.Name) - 1
      
        Call HideC3(Shts.Name, "dyn")
        If FilterPriorityDyn_PPT(Shts, Filt) = True And colonne <> 0 Then
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
            'Call insertPart_PPT(objPres, objSlide, pos)
        ViderPressePapiers
            
            
            pasted = False
            Do While Not pasted
                On Error Resume Next
                objPres.Slides(9).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
            Loop
        
        
        ViderPressePapiers
            
            Call TakePicSelection_PPT3(objppt, objslide, plage)
            With objslide.Shapes(objslide.Shapes.Count)
                .LockAspectRatio = msoTrue
                .Width = 500 ' Set the width of the shape
            End With
            Call UnFilterPriority(Shts, TotalRow)
       
        End If
    End With
End Function

Function InsertPic_PPT_Format2(ByRef objppt As Object, objPres As Object, objslide As Object, slideIndex As Integer, j As Long, BlnGraphDriv As Boolean, BlnGraphDyna As Boolean)
    Dim i As Integer
    Dim h As Integer
    Dim t As Integer
    Dim o As String
    Dim x As String
    Dim getParamSdv As String
    Dim v() As String
    Dim c As Object
    Dim nbpages As Integer
  
    Dim nbSlides As Integer
    Dim slideLayout As PpSlideLayout
    Dim pptShape As shape
    Dim pasted As Boolean
   
   
        o = SDVListe(j)
        
      
        ViderPressePapiers
          
          pasted = False
          Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
          Loop
        
        ViderPressePapiers
                     slideIndex = slideIndex + 1
'
            ProgressTitle ("Copie des données : " & o)
    For t = 1 To 2
'            If t = 1 Then Call newSdvSlide_PPT(objSlide, UCase(o), "2." & j + 1) Else Call newSdvSlide_PPT(objSlide, UCase(o), "2." & j + 1)
            
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
                                            ''If UCase(o) = "ACCEL CST LOAD" Then Stop
                                            If t = 1 Then Call newSdvSlide_PPT(objslide, UCase(o), "2." & j + 1) Else Call newSdvSlide_PPT(objslide, UCase(o), "2." & j + 1)
                                            
                                            If t = 1 Then
                                                Call insertPart_PPT(objPres, objslide, 1)
                                                If BlnGraphDriv Then Call CopyGraph1_PPT(objppt, objPres, objslide, o, t, slideIndex)
                                            Else
                                                Call insertPart_PPT(objPres, objslide, 2)
                                                If BlnGraphDyna Then Call CopyGraph1_PPT(objppt, objPres, objslide, o, t, slideIndex)
                                            End If
'                                            Call CopyGraph1_PPT(objPPT, objPres, objSlide, o, t, slideIndex)
                                             'Call CopySummary_PPT(objPPT, objSlide, o, T, slideIndex)
                                            
                                         End If
                                         
                                     End If
                                      Exit For
                                Next h
                           End If
                  
      Next t
End Function
Function InsertPic_PPT2(ByRef objppt As Object, objPres As Object, slideIndex As Integer)
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
    Dim objslide As Object
    Dim nbSlides As Integer
    Dim slideLayout As PpSlideLayout
    Dim pptShape As shape
    Dim pasted As Boolean
   
    slideLayout = ppLayoutTitleOnly '''ppLayoutText
    
    For j = 0 To UBound(SDVListe)
        o = SDVListe(j)
        ThisWorkbook.sheets(o).Cells.EntireColumn.Hidden = False
        For t = 1 To 2
            ProgressTitle ("Copie des données : " & o)
            If t = 1 Or (t = 2 And checkCriteriaDyn(o) = True And checkCorrespondancePriorityDyn(o) = True) Then
                'If j = 0 And T = 1 Then
        ViderPressePapiers
                      
            pasted = False
            Do While Not pasted
                On Error Resume Next
                objPres.Slides(numSlide).Copy
                DoEvents
                'sleep 500
                Set objslide = objPres.Slides.Paste(slideIndex + 1)
                If ERR.Number = 0 Then
                    pasted = True
                Else
                    ERR.Clear
                    'sleep 500
                End If
                On Error GoTo 0
            Loop
        
        ViderPressePapiers
                     slideIndex = slideIndex + 1
                     
                     
               ' Else 'End If
            If t = 1 Then Call newSdvSlide_PPT(objslide, UCase(o) & " DRIVABILITY", "2." & j + 1 & "." & t) Else Call newSdvSlide_PPT(objslide, UCase(o) & " DYNAMISM", "2." & j + 1 & "." & t)
            
                For i = 1 To 5
                    x = j & i & t
                    
                        If i = 1 Then
                            Call insertPart_PPT(objslide, i)
                            Call CopySummary_PPT(objppt, objslide, o, t, slideIndex)
                           
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
                                                                Call insertPart_PPT(objslide, i)
                                                                If t = 1 Then
                                                                    Call CopyGraph0_PPT(objppt, objslide, o, t, slideIndex)
                                                                Else
                                                                    Call CopyGraph0_PPT(objppt, objslide, o, t, slideIndex)
                                                                End If
                                                           End If
                                                           Exit For
                                                      Else
                                                            If UpdateGraph.checkObject(CStr(Split(v(h), ":")(0)), o) <> "" Then
                                                                Call insertPart_PPT(objslide, i)
                                                                If t = 1 Then
                                                                    
                                                                    Call CopyLeverAS_PPT2(objppt, objPres, objslide, o, UpdateGraph.checkObject(CStr(Split(v(h), ":")(0)), o), slideIndex, True)
                                                                Else
                                                
                                                                     Call CopyLeverAS_PPT2(objppt, objPres, objslide, o, UpdateGraphDyn.checkObject(CStr(Split(v(h), ":")(0)), o), slideIndex, True)
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
                                                                 Call insertPart_PPT(objslide, i)
                                                                 Call CopyGraph1_PPT(objppt, objPres, objslide, o, t, slideIndex = 8)
                                                                 
                                                              End If
                                                              Exit For
                                                          End If
                                                     Next h
                                                End If
                                         End If
                                 End If
                        ElseIf i = 4 Then
                            If t = 1 Then
                                Call CopyPriorityPoints_PPT(objslide, "Hight", ThisWorkbook.Worksheets(o), i)
                            Else
                                Call CopyPriorityPointsDyn_PPT(objslide, "Hight", ThisWorkbook.Worksheets(o), i)
                            End If
                        ElseIf i = 5 Then
                            If t = 1 Then
                                 Call CopyPriorityPoints_PPT(objslide, "Low", ThisWorkbook.Worksheets(o), i)
                            Else
                                 Call CopyPriorityPointsDyn_PPT(objslide, "Low", ThisWorkbook.Worksheets(o), i)
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
    
   ' objPres.Slides(8).Delete
    
    'Call SupprimerSlidesVides(objPres)
    
    'Call VerifierPageAvecEnteteVide(objword, objDoc)
    
End Function
'Function InsertPic_PPT2(ByRef objPPT As Object, objPres As Object, slideIndex As Integer)
'    Dim i As Integer
'    Dim j As Long
'    Dim H As Integer
'    Dim T As Integer
'    Dim o As String
'    Dim x As String
'    Dim getParamSdv As String
'    Dim v() As String
'    Dim c As Object
'    Dim nbpages As Integer
'    Dim objSlide As Object
'    Dim nbSlides As Integer
'    Dim slideLayout As PpSlideLayout
'    Dim pptShape As shape
'
'    slideLayout = ppLayoutTitleOnly '''ppLayoutText
'
'    For j = 0 To UBound(SDVListe)
'        o = SDVListe(j)
'        ThisWorkbook.sheets(o).Cells.EntireColumn.Hidden = False
'        For T = 1 To 2
'            ProgressTitle ("Copie des données : " & o)
'            If T = 1 Or (T = 2 And checkCriteriaDyn(o) = True And checkCorrespondancePriorityDyn(o) = True) Then
'                'If j = 0 And T = 1 Then
'                      objPres.Slides(8).Copy
'                     Set objSlide = objPres.Slides.Paste(slideIndex + 1)
'                     slideIndex = slideIndex + 1
'               ' Else 'End If
'            If T = 1 Then Call newSdvSlide_PPT(objSlide, UCase(o) & " DRIVABILITY", "2." & j + 1 & "." & T) Else Call newSdvSlide_PPT(objSlide, UCase(o) & " DYNAMISM", "2." & j + 1 & "." & T)
'
'                For i = 1 To 5
'                    x = j & i & T
'
'                        If i = 1 Then
'                            Call insertPart_PPT(objSlide, i)
'                            Call CopySummary_PPT(objPPT, objSlide, o, T, slideIndex)
'
'                        ElseIf i = 2 Or i = 3 Then
'                                If T = 1 Then
'                                    getParamSdv = UpdateGraph.checkGraphEnable(o)
'                                Else
'                                    getParamSdv = UpdateGraphDyn.checkGraphEnable(o)
'                                End If
'                                If getParamSdv <> "" Then
'                                        If i = 2 Then
'                                                v = Split(getParamSdv, ";")
'                                                For H = 0 To UBound(v)
'                                                      If Split(v(H), ":")(0) = "Graphique_0" Or Split(v(H), ":")(0) = "Graphique_00" Then
'                                                           If ThisWorkbook.Worksheets(o).Shapes(Split(v(H), ":")(0)).Visible = True Then
'                                                                Call insertPart_PPT(objSlide, i)
'                                                                If T = 1 Then
'                                                                    Call CopyGraph0_PPT(objPPT, objSlide, o, T, slideIndex)
'                                                                Else
'                                                                    Call CopyGraph0_PPT(objPPT, objSlide, o, T, slideIndex)
'                                                                End If
'                                                           End If
'                                                           Exit For
'                                                      Else
'                                                            If UpdateGraph.checkObject(CStr(Split(v(H), ":")(0)), o) <> "" Then
'                                                                Call insertPart_PPT(objSlide, i)
'                                                                If T = 1 Then
'
'                                                                    Call CopyLeverAS_PPT2(objPPT, objPres, objSlide, o, UpdateGraph.checkObject(CStr(Split(v(H), ":")(0)), o), slideIndex, True)
'                                                                Else
'
'                                                                     Call CopyLeverAS_PPT2(objPPT, objPres, objSlide, o, UpdateGraphDyn.checkObject(CStr(Split(v(H), ":")(0)), o), slideIndex, True)
'                                                                End If
'
'                                                            End If
'                                                            Exit For
'                                                      End If
'                                                Next H
'                                         Else
'                                                If T = 1 Then
'                                                    getParamSdv = UpdateGraph.checkGraphEnable(o)
'                                                Else
'                                                    getParamSdv = UpdateGraphDyn.checkGraphEnable(o)
'                                                End If
'                                               v = Split(getParamSdv, ";")
'                                               If InStr(1, getParamSdv, "Graphique_1") <> 0 Or InStr(1, getParamSdv, "Graphique_11") Then
'                                                    For H = 0 To UBound(v)
'                                                         If Split(v(1), ":")(0) = "Graphique_1" Or InStr(1, getParamSdv, "Graphique_11") Then
'                                                              If ThisWorkbook.Worksheets(o).Shapes(Split(v(1), ":")(0)).Visible = True Then
'                                                                 Call insertPart_PPT(objSlide, i)
'                                                                 Call CopyGraph1_PPT(objPPT, objPres, objSlide, o, T, slideIndex = 8)
'
'                                                              End If
'                                                              Exit For
'                                                          End If
'                                                     Next H
'                                                End If
'                                         End If
'                                 End If
'                        ElseIf i = 4 Then
'                            If T = 1 Then
'                                Call CopyPriorityPoints_PPT(objSlide, "Hight", ThisWorkbook.Worksheets(o), i)
'                            Else
'                                Call CopyPriorityPointsDyn_PPT(objSlide, "Hight", ThisWorkbook.Worksheets(o), i)
'                            End If
'                        ElseIf i = 5 Then
'                            If T = 1 Then
'                                 Call CopyPriorityPoints_PPT(objSlide, "Low", ThisWorkbook.Worksheets(o), i)
'                            Else
'                                 Call CopyPriorityPointsDyn_PPT(objSlide, "Low", ThisWorkbook.Worksheets(o), i)
'                            End If
'                        End If
'
'
'        '             If i <= 3 Then
'        '               objDoc.InlineShapes(LastPicNumber(objDoc)).Width = 520
'        '               objDoc.InlineShapes(LastPicNumber(objDoc)).LockAspectRatio = 0
'        '               objDoc.InlineShapes(LastPicNumber(objDoc)).Height = 283.5
'        '            End If
'        '              x = x + 1
'
'                Next i
'            End If
'      Next T
'    Next j
'
'   ' objPres.Slides(8).Delete
'
'    'Call SupprimerSlidesVides(objPres)
'
'    'Call VerifierPageAvecEnteteVide(objword, objDoc)
'
'End Function


Function FilterPriority_PPT(Shts As Worksheet, Filt As String) As Boolean
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
    FilterPriority_PPT = False
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
                   
                '''''''''''''''''Modif 23/01/2025
                      For i = 7 To lastRow
'                          If (.Cells(i, colonne)) <> VCrit(1) Then
                          If IsNumeric((.Cells(i, colonne))) Then
                                If (.Cells(i, colonne)) >= 5 Then
                                    If r Is Nothing Then Set r = .Cells(i, colonne) Else Set r = Union(r, .Cells(i, colonne))
                                Else
                                   FilterPriority_PPT = True
                                End If
                            
                          Else
                                FilterPriority_PPT = True
                            End If
                     Next i
                     If Not r Is Nothing And FilterPriority_PPT = True Then r.EntireRow.Hidden = True
                     
              
       
'         Call UnFilterPriority(Shts)
    End With
    
End Function

Function FilterPriorityDyn_PPT(Shts As Worksheet, Filt As String) As Boolean
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
    FilterPriorityDyn_PPT = False
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
                      '''''''''''''''Modif 23/01/2025
''                          If (.Cells(i, colonne)) <> VCrit(1) Then
                            
                            If IsNumeric((.Cells(i, colonne))) Then
                                If (.Cells(i, colonne)) >= 5 Then
                                    If r Is Nothing Then Set r = .Cells(i, colonne) Else Set r = Union(r, .Cells(i, colonne))
                                Else
                                    FilterPriorityDyn_PPT = True
                                End If
                            Else
                                    FilterPriorityDyn_PPT = True
                            End If
                                
                     Next i
                     If Not r Is Nothing And FilterPriorityDyn_PPT = True Then r.EntireRow.Hidden = True
                     
               
       
'         Call UnFilterPriorityDyn(Shts)
    End With
    
End Function








