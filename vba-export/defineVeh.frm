Attribute VB_Name = "defineVeh"
Attribute VB_Base = "0{E9CB3676-F0BC-4200-9B50-47163FF7176B}{3D9D65F4-06C5-439F-AC52-367F566264F3}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False





Option Explicit

Private Sub CommandButton1_Click()
  Dim c As Range
  Dim test As Boolean
  
  
        
        test = ValueExistsInVehicleRange(Me.tEMPS)
        If test = False And Me.tEMPS <> "" Then
        ProgressLoad
        With ThisWorkbook.Worksheets("CONFIGURATIONS")
                Call hideShowTarget(False)
                If ThisWorkbook.sheets("HOME").Range("Fuel").Value = "" Or ThisWorkbook.sheets("HOME").Range("Gears").Value = "" Or ThisWorkbook.sheets("HOME").Range("Software") = "" Or ThisWorkbook.sheets("HOME").Range("Prestation").Value = "" Or ThisWorkbook.sheets("HOME").Range("DriveVersion").Value = "" Or ThisWorkbook.sheets("HOME").Range("Milestone").Value = "" Or ThisWorkbook.sheets("HOME").Range("Area").Value = "" Then
                    addVehRating (Me.tEMPS)
                Else
                    addVehRating2 (Me.tEMPS)
                End If
                Call hideShowTarget(True)
                
                If ThisWorkbook.sheets("HOME").Range("Fuel").Value = "" Or ThisWorkbook.sheets("HOME").Range("Gears").Value = "" Or ThisWorkbook.sheets("HOME").Range("Software") = "" Or ThisWorkbook.sheets("HOME").Range("Prestation").Value = "" Or ThisWorkbook.sheets("HOME").Range("DriveVersion").Value = "" Or ThisWorkbook.sheets("HOME").Range("Milestone").Value = "" Or ThisWorkbook.sheets("HOME").Range("Area").Value = "" Then
                    addVehGraphStatus (Me.tEMPS)
                Else
                    addVehGraphStatus2 (Me.tEMPS)
                End If
                Call UpdateYVal
                addVehTotPoint
                Set c = .Range("VEHICLE")
                Set c = c.Offset(1, 0)
                While c.Value <> ""
                    Set c = c.Offset(1, 0)
                    
                    
          
                    
                    
                Wend
                If .Range("A" & c.row & ":B" & c.row).MergeCells = False Then
                    .Rows(c.row + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                    .Range("A" & c.row & ":B" & c.row).Borders(1).LineStyle = xlContinuous
                    .Range("A" & c.row & ":B" & c.row).Borders(2).LineStyle = xlContinuous
                    .Range("A" & c.row & ":B" & c.row).Borders(3).LineStyle = xlContinuous
                    .Range("A" & c.row & ":B" & c.row).Borders(4).LineStyle = xlContinuous
                    .Range("A" & c.row & ":B" & c.row).Merge
                    .Range("A" & c.row & ":B" & c.row) = Me.tEMPS
                End If
           End With
           
         Unload Me
         MsgBox "Terminé", vbInformation, "ODRIV"
         ElseIf Me.tEMPS = "" Then
            MsgBox "Insérer une véhicule", vbCritical, "ODRIV"
            Exit Sub
         ElseIf test = True Then
            MsgBox "Target vehicule existe dèjà", vbCritical, "ODRIV"
            Exit Sub
        End If
        Unload PleaseWait
 
 
End Sub

Function addVehRating(nameVeh As String)
    Dim r
    Dim lastC
    Dim colD
    Dim j As Integer
    Dim ws As Worksheet
    Dim lastRow As Integer
    Dim startCell As Range
    Dim i As Integer
    Dim fillColor As Long
    Dim shp As shape
    Dim startRow As Integer
    
    
       
       
    j = 0

    Set ws = ThisWorkbook.sheets("CONFIGURATIONS")
    Set startCell = ws.Range("VEHICLE")
 
    
    startRow = startCell.row
 
    
    lastRow = ws.Cells(ws.Rows.Count, startCell.Column).End(xlUp).row
    For i = startRow To lastRow
        If ws.Cells(i, 1) = "" Then
            Exit For
        Else
            j = j + 1
        End If
    Next i
    
    
    
     fillColor = ThisWorkbook.sheets("Palette").Cells(j + 1, 1).Interior.color
    
    
    With sheets("RATING")
'        Application.ScreenUpdating = False
        .Activate
        colD = .Rows("21:22").Find(What:="Drivability Lowest Events", LookAt:=xlWhole).Column
        .Columns(colD).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Range(.Cells(21, colD), .Cells(22, colD)).Merge
        .Cells(21, colD) = nameVeh
        
        colD = .Rows("21:22").Find(What:="Responsiveness Lowest Events", LookAt:=xlWhole).Column
        .Columns(colD).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Range(.Cells(21, colD), .Cells(22, colD)).Merge
        .Cells(21, colD) = nameVeh
        
        lastC = .Cells(10, .Columns.Count).End(xlToLeft).Column
        .Columns(lastC).Copy Destination:=.Cells(1, lastC + 1)
        Application.CutCopyMode = False
        .Cells(10, lastC + 1) = nameVeh
        .Cells(16, lastC + 1) = nameVeh
        
        
        
        
    End With
    Set ws = ThisWorkbook.sheets("RATING")
    For Each shp In ws.Shapes
        ' Check if shape is in the specified column
        If shp.TopLeftCell.Column = lastC + 1 Then
            ' Check if shape name contains "Triangle"
            If InStr(1, shp.Name, "Triangle", vbTextCompare) > 0 Then
                On Error Resume Next
                shp.Fill.ForeColor.RGB = fillColor
                On Error GoTo 0
            End If
        End If
    Next shp
    sheets("CONFIGURATIONS").Activate
'    Application.ScreenUpdating = True
End Function

Function addVehRating2(nameVeh As String)
    Dim r
    Dim lastC
    Dim colD
    Dim j As Integer
    Dim ws, ws2 As Worksheet
    Dim lastRow, lastRow2 As Integer
    Dim startCell As Range
    Dim i As Integer
    Dim fillColor As Long
    Dim shp As shape
    Dim startRow, startRow2 As Integer
    
    
       
       
    

    j = 1
   
    
    Set ws = ThisWorkbook.sheets("CONFIGURATIONS")
    Set ws2 = ThisWorkbook.sheets("Graph_status")
    Set startCell = ws.Range("VEHICLE")

    startRow = startCell.row
    startRow2 = 2
    lastRow = ws.Cells(ws.Rows.Count, startCell.Column).End(xlUp).row
    lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).row
    For i = startRow2 To lastRow2
        If ws2.Cells(i, 1) = "index rouge" Then
            Exit For
        ElseIf ws2.Cells(i, 5).Interior.color <> RGB(255, 255, 255) Then
            j = j + 1
        End If
    Next i
    
    
    
     fillColor = ThisWorkbook.sheets("Palette").Cells(j + 1, 1).Interior.color
    
    
    With sheets("RATING")
'        Application.ScreenUpdating = False
        .Activate
        colD = .Rows("21:22").Find(What:="Drivability Lowest Events", LookAt:=xlWhole).Column
        .Columns(colD).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Range(.Cells(21, colD), .Cells(22, colD)).Merge
        .Cells(21, colD) = nameVeh
        
        colD = .Rows("21:22").Find(What:="Responsiveness Lowest Events", LookAt:=xlWhole).Column
        .Columns(colD).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Range(.Cells(21, colD), .Cells(22, colD)).Merge
        .Cells(21, colD) = nameVeh
        
        lastC = .Cells(10, .Columns.Count).End(xlToLeft).Column
        .Columns(lastC).Copy Destination:=.Cells(1, lastC + 1)
        Application.CutCopyMode = False
        .Cells(10, lastC + 1) = nameVeh
        .Cells(16, lastC + 1) = nameVeh
        
        
        
        
    End With
    Set ws = ThisWorkbook.sheets("RATING")
    For Each shp In ws.Shapes
        ' Check if shape is in the specified column
        If shp.TopLeftCell.Column = lastC + 1 Then
            ' Check if shape name contains "Triangle"
            If InStr(1, shp.Name, "Triangle", vbTextCompare) > 0 Then
                On Error Resume Next
                shp.Fill.ForeColor.RGB = fillColor
                On Error GoTo 0
            End If
        End If
    Next shp
    sheets("CONFIGURATIONS").Activate
'    Application.ScreenUpdating = True
End Function

'Function addVehRating(nameVeh As String)
'    Dim r
'    Dim lastC
'    Dim cold
'
'    With sheets("RATING")
'        cold = .Rows("21:22").Find(What:="Drivability Lowest Events", LookAt:=xlWhole).Column
'        .Columns(cold).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'        .Cells(21, cold) = nameVeh
'
'        cold = .Rows("21:22").Find(What:="Dynamism Lowest Events", LookAt:=xlWhole).Column
'        .Columns(cold).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'        .Cells(22, cold) = nameVeh
'
'        lastC = .Cells(10, .Columns.Count).End(xlToLeft).Column
'        .Columns(lastC).Copy Destination:=.Cells(1, lastC + 1)
'        Application.CutCopyMode = False
'        .Cells(10, lastC + 1) = nameVeh
'        .Cells(16, lastC + 1) = nameVeh
'    End With
'
'End Function

Function addVehGraphStatus2(nameVeh As String)
    Dim r
    Dim lastr
    Dim i As Integer
    Dim num As Integer
    Dim fillColor As Long
    Dim Pos As Integer
    Dim j As Integer
    Dim ws, ws2 As Worksheet
    Dim startCell As Range
    Dim lastRow, lastRow2 As Long
    Dim startRow, startRow2 As Long
    
    j = 1
    num = 0
    
    Set ws = ThisWorkbook.sheets("CONFIGURATIONS")
    Set ws2 = ThisWorkbook.sheets("Graph_status")
    Set startCell = ws.Range("VEHICLE")

    startRow = startCell.row
    startRow2 = 2
    lastRow = ws.Cells(ws.Rows.Count, startCell.Column).End(xlUp).row
    lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).row
    For i = startRow2 To lastRow2
        If ws2.Cells(i, 1) = "index rouge" Then
            Exit For
        ElseIf ws2.Cells(i, 5).Interior.color <> RGB(255, 255, 255) Then
            j = j + 1
        End If
    Next i
    
    
   
    fillColor = ThisWorkbook.sheets("Palette").Cells(j + 1, 1).Interior.color
    

    With sheets("Graph_status")
        lastr = .Cells(.Rows.Count, 1).End(xlUp).row
        For i = 1 To lastr
            If .Cells(i, 1) = "index rouge" Then
                
                num = num + 1
                .Rows(i - 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                .Cells(i - 1, 1) = nameVeh
                .Cells(i - 1, 4) = .Cells(i - 2, 4) + 8.2857
                .Cells(i - 1, 5).Interior.color = fillColor
                If num Mod 2 <> 0 Then
                    
                    .Cells(i - 1, 3).formula = "=(B" & (i - 1) & "/100)^GLOBALPUISS"
                Else
                    .Cells(i - 1, 3).formula = "=((B" & (i - 1) & "/100))*100"
                End If
                i = i + 1
           End If
        Next i
        Application.CutCopyMode = False
    End With
End Function

Function addVehGraphStatus(nameVeh As String)
    Dim r
    Dim lastr
    Dim i As Integer
    Dim num As Integer
    Dim fillColor As Long
    Dim Pos As Integer
    Dim j As Integer
    Dim ws As Worksheet
    Dim startCell As Range
    Dim lastRow As Long
    Dim startRow As Long
    
    j = 0
    num = 0
    
    Set ws = ThisWorkbook.sheets("CONFIGURATIONS")
    Set startCell = ws.Range("VEHICLE")
 
    
    startRow = startCell.row
 
    
    lastRow = ws.Cells(ws.Rows.Count, startCell.Column).End(xlUp).row
    For i = startRow To lastRow
        If ws.Cells(i, 1) = "" Then
            Exit For
        Else
            j = j + 1
        End If
    Next i
    
    
    
    fillColor = ThisWorkbook.sheets("Palette").Cells(j + 1, 1).Interior.color
    
    
    
    
    
    
    
    With sheets("Graph_status")
        lastr = .Cells(.Rows.Count, 1).End(xlUp).row
        For i = 1 To lastr
            If .Cells(i, 1) = "index rouge" Then
                
                num = num + 1
                .Rows(i - 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                .Cells(i - 1, 1) = nameVeh
                .Cells(i - 1, 4) = .Cells(i - 2, 4) + 8.2857
                .Cells(i - 1, 5).Interior.color = fillColor
                If num Mod 2 <> 0 Then
                    
                    .Cells(i - 1, 3).formula = "=(B" & (i - 1) & "/100)^GLOBALPUISS"
                Else
                    .Cells(i - 1, 3).formula = "=((B" & (i - 1) & "/100))*100"
                End If
                i = i + 1
           End If
        Next i
        Application.CutCopyMode = False
    End With
End Function

Function addVehTotPoint()
    Dim r
    Dim lastr, lastC
    Dim i As Integer
    
    With sheets("totalPoint")
        .Columns("S:DK").EntireColumn.Delete
        lastr = sheets("RATING").Cells(.Rows.Count, 4).End(xlUp).row
        lastC = sheets("RATING").Cells(lastr, sheets("RATING").Columns.Count).End(xlToLeft).Column
        sheets("RATING").Range(sheets("RATING").Cells(lastr, 2), sheets("RATING").Cells(lastr, lastC)).Copy Destination:=.Range("S1")
        Application.CutCopyMode = False
       
    End With

End Function


Private Sub UserForm_Activate()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

End Sub


Function ValueExistsInVehicleRange(valueToCheck As String) As Boolean
    Dim ws As Worksheet
    Dim startCell As Range
    Dim currentCell As Range
 
    Set ws = ThisWorkbook.sheets("CONFIGURATIONS")
    Set startCell = ws.Range("VEHICLE")
    Set currentCell = startCell
 
    ' Default result
    ValueExistsInVehicleRange = False
 
    ' Loop until an empty cell is found
    Do While currentCell.Value <> ""
        If UCase(currentCell.Value) = UCase(valueToCheck) Then
            ValueExistsInVehicleRange = True
            Exit Function
        End If
        Set currentCell = currentCell.Offset(1, 0)
    Loop
End Function

Private Sub UpdateYVal()
 
    Dim ws As Worksheet

    Dim lastRow As Long

    Dim i As Long, j As Long

    Dim GI As Collection

    Dim IR As Collection

    Dim Values As Collection

    Dim n As Long

    Dim stepVal As Double

    Dim minVal As Double, maxVal As Double

    minVal = 4

    maxVal = 99
 
    Set ws = ThisWorkbook.sheets("Graph_status")

    Set GI = New Collection

    Set IR = New Collection

    Set Values = New Collection
 
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
 
    

    For i = 2 To lastRow

        If ws.Cells(i, 1).Value = "Global index" Then

            GI.Add i + 1

        ElseIf ws.Cells(i, 1).Value = "index rouge" Then

            IR.Add i - 2

        End If

    Next i
 
   

    For i = 1 To GI.Count

        n = IR(i) - GI(i) + 1

        stepVal = (maxVal - minVal) / (n - 1)

        If i = 1 Then

            

            For j = 0 To n - 1

                ws.Cells(GI(i) + j, 4).Value = minVal + stepVal * j

                Values.Add minVal + stepVal * j

            Next j
 
        Else

            

            For j = 1 To n

                ws.Cells(GI(i) + j - 1, 4).Value = Values(j)

            Next j

        End If

    Next i
 
End Sub

 

