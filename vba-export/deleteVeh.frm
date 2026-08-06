Attribute VB_Name = "deleteVeh"
Attribute VB_Base = "0{E6DAA70C-BBD7-451E-AAD1-C3A634E7C2F1}{0A45080B-992D-41FD-8827-34DCAA1890F7}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False



Sub DeleteFromTargetVehicle(ByRef arr() As String)

    Dim lastRow As Long
    Dim i, j As Long
    Dim targetValue As String
 
 
    lastRow = Feuil22.Cells(Feuil22.Rows.Count, "C").End(xlUp).row
 
    For j = LBound(arr) To UBound(arr)
        targetValue = arr(j)
        For i = lastRow To 1 Step -1
            If Feuil22.Cells(i, "C").Value = targetValue Then
                Feuil22.Rows(i).Delete
            End If
        Next i
    Next j


End Sub

Private Sub CommandButton1_Click()
 
    Dim selectedValues() As String
    Dim countSelected As Long
    Dim Pos() As Integer
    
    
    Dim i As Integer

    countSelected = 0
    For i = 0 To Me.ListVeh.ListCount - 1
        If Me.ListVeh.Selected(i) Then countSelected = countSelected + 1
    Next i
    If countSelected = 0 Then
        MsgBox "Sélectionner une véhicule au moins", vbCritical, "Odriv"
        Exit Sub
    Else
        ProgressLoad
        Call hideShowTarget(False)
        Call GetSelectedValues(selectedValues)
        Call GetPosConfig(selectedValues, Pos)
        Call DeleteSelectedVehicles(selectedValues, Pos)
        Call DeleteFromGraphStatus(selectedValues)
        Call DeleteRatingColumns(selectedValues)
        Call DeleteFromTargetVehicle(selectedValues)
        Call UpdatePalette
        Call UpdateMarqueurs
        Call hideShowTarget(True)
        Unload Me
    End If

    MsgBox "Véhicules supprimées.", vbInformation, "ODRIV"
    Unload PleaseWait
End Sub
 
 
Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim rngVehicle As Range
    Dim startRow As Long
    Dim i As Long
    Dim currentValue As String
    Set ws = ThisWorkbook.Worksheets("CONFIGURATIONS")
    Set rngVehicle = ws.Range("VEHICLE")
    startRow = rngVehicle.row + rngVehicle.Rows.Count
    Me.ListVeh.Clear
    i = startRow
    Do
        currentValue = ws.Cells(i, rngVehicle.Column).Value
        If Trim(currentValue) = "" Then Exit Do
        Me.ListVeh.AddItem currentValue
        i = i + 1
    Loop
End Sub
 
 

Private Sub GetSelectedValues(ByRef arr() As String)
    Dim i As Long
    Dim countSelected As Long
    countSelected = 0
    For i = 0 To Me.ListVeh.ListCount - 1
        If Me.ListVeh.Selected(i) Then countSelected = countSelected + 1
    Next i
    If countSelected = 0 Then
        MsgBox "Sélectionner une véhicule au moins", vbCritical, "Odriv"
        Exit Sub
    End If
    ReDim arr(1 To countSelected)
    Dim idx As Long
    idx = 1
    For i = 0 To Me.ListVeh.ListCount - 1
        If Me.ListVeh.Selected(i) Then
            arr(idx) = Me.ListVeh.list(i)
            idx = idx + 1
        End If
    Next i
End Sub

 
Sub DeleteSelectedVehicles(selectedValues() As String, Pos() As Integer)
    Dim ws As Worksheet
    Dim rngVehicle As Range
    Dim startRow As Integer
    Dim currentRow As Long
    Dim valueToDelete As String
    Dim i As Long
    Dim j As Integer
    Dim startCell As Range
    Dim lastRow As Integer
    Dim r As Integer
    Dim test As Boolean
    Dim k As Integer
    
    

    
    For i = LBound(selectedValues) To UBound(selectedValues)
        
       
        valueToDelete = selectedValues(i)
        
        
        
        test = False
    j = Pos(i)
    
    Set ws = ThisWorkbook.sheets("Palette")
    ws.Activate
    r = 2
    Do While test = False
    
    
        With ws.Cells(r, 1).Interior
            ' Check if fill is white or no fill
            If .ColorIndex = xlNone Or .color = RGB(255, 255, 255) Then
                FirstWhiteFillRowInCol1 = r
                test = True
            Else
            r = r + 1
            End If
        End With
    Loop
    
    
    'ws.Cells(r, 1).Interior.color = ws.Cells(j, 1).Interior.color
    
    'ws.Rows(j).Delete
    For k = i To UBound(selectedValues)
        Pos(k) = Pos(k) - 1
    Next k
        Set ws = ThisWorkbook.Worksheets("CONFIGURATIONS")
        Set rngVehicle = ws.Range("VEHICLE")
        startRow = rngVehicle.row
        currentRow = startRow
        Do While Trim(ws.Cells(currentRow, rngVehicle.Column).Value) <> ""
            If ws.Cells(currentRow, rngVehicle.Column).Value = valueToDelete Then
                ws.Rows(currentRow).Delete
                Exit Do
            Else
                currentRow = currentRow + 1
            End If
        Loop
    Next i
End Sub
 
Sub DeleteFromGraphStatus(selectedValues() As String)
    Dim ws As Worksheet
    Dim i As Long
    Dim currentRow As Long
    Dim lastRow As Long
    Dim valueToDelete As String
    Set ws = ThisWorkbook.Worksheets("Graph_status")
    For i = LBound(selectedValues) To UBound(selectedValues)
        valueToDelete = selectedValues(i)
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        For currentRow = lastRow To 1 Step -1
            If ws.Cells(currentRow, 1).Value = valueToDelete Then
                ws.Rows(currentRow).Delete
            End If
        Next currentRow
    Next i
End Sub
 
 
Sub DeleteRatingColumns(selectedValues() As String)
    Dim ws As Worksheet
    Dim i As Long
    Dim selectedValue As String
    Dim cellFound As Range
    Dim searchRange As Range
    Dim colNum As Long
    Dim firstAddress As String
    Set ws = ThisWorkbook.Worksheets("RATING")
    
    
    ws.Activate
    For i = LBound(selectedValues) To UBound(selectedValues)
        
        For j = 1 To 2
        
        selectedValue = selectedValues(i)
        Set searchRange = ws.Rows("21:22")
        Set cellFound = searchRange.Find(What:=selectedValue, LookAt:=xlWhole)
        If Not cellFound Is Nothing Then
            firstAddress = cellFound.Address
            ws.Columns(cellFound.Column).Delete
      
        End If
        
        Next j
        Set searchRange = ws.Rows(10)
        Set cellFound = searchRange.Find(What:=selectedValue, LookAt:=xlWhole)
        If Not cellFound Is Nothing Then
            ws.Columns(cellFound.Column).Delete
        End If
    Next i

    
End Sub


Private Sub GetPosConfig(ByRef inputArr() As String, ByRef lengthsArr() As Integer)
    Dim i As Long
    Dim arrSize As Long
    Dim j As Integer
    Dim ws As Worksheet
    Dim k As Integer
    Dim startCell As Range
    Dim startRow, lastRow As Integer
    Dim valueToDelete As String
    
    
    
    
    arrSize = UBound(inputArr) - LBound(inputArr) + 1
    ReDim lengthsArr(LBound(inputArr) To UBound(inputArr))
 
    For i = LBound(inputArr) To UBound(inputArr)
        
        j = 0

    Set ws = ThisWorkbook.sheets("CONFIGURATIONS")
    Set startCell = ws.Range("VEHICLE")
    startRow = startCell.row
    
    
 
    valueToDelete = inputArr(i)
    lastRow = ws.Cells(ws.Rows.Count, startCell.Column).End(xlUp).row
    For k = startRow To lastRow
        If ws.Cells(k, 1) = valueToDelete Then
            Exit For
        Else
            j = j + 1
        End If
    Next k
        
        
        
        lengthsArr(i) = j + 1
    Next i
    Exit Sub
 

End Sub

Private Sub UpdateY()
    Dim ws As Worksheet
    Dim lastRow As Integer
    Dim i As Integer, j As Integer
    Dim GI As Collection
    Dim IR As Collection
    Dim Values As Collection
 
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
        If i = 1 Then
            For j = GI(i) To IR(i)
                ws.Cells(j, 4).Value = 8.2857 * (j - 2)
                Values.Add 8.2857 * (j - 2)
            Next j
            
            
        Else
        
        For j = 1 To (IR(i) - GI(i) + 1)
            ws.Cells(GI(i) + j - 1, 4).Value = Values(j)
        Next j
        End If
    Next i
End Sub

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

    maxVal = 100
 
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

Sub UpdatePalette()


    Dim ws As Worksheet
    Dim startCell As Range
    Dim c As Range
    Dim lastRow, numRows, i, j, h As Integer
    
 
    Set ws = ThisWorkbook.Worksheets("CONFIGURATIONS")
    Set startCell = ws.Range("VEHICLE")
 
    
    Set c = startCell.Offset(1, 0)
 
    numRows = 0
    Do While Trim(c.Value) <> ""
        numRows = numRows + 1
        Set c = c.Offset(1, 0)
    Loop
    
    lastRow = ThisWorkbook.Worksheets("Graph_status").Cells(Rows.Count, "A").End(xlUp).row
    
    
    For i = 1 To lastRow
        If ThisWorkbook.Worksheets("Graph_status").Cells(i, 1) = "Global index" Then
            h = 2
            For j = i + 1 To i + numRows
                ThisWorkbook.sheets("Graph_status").Cells(j, 5).Interior.color = ThisWorkbook.sheets("Palette").Cells(h, 1).Interior.color
                h = h + 1
            Next j
        End If
    Next i
    

End Sub

Sub UpdateMarqueurs()
 
    Dim ws As Worksheet

    Dim shp As shape

    Dim colD As Long, lastCol As Long

    Dim h As Long, i As Long

    Dim hasTriangle As Boolean

    Dim colColor As Long
 
    Set ws = ThisWorkbook.Worksheets("RATING")
 
    colD = ws.Rows(10).Find(What:="Tested vehicle", LookAt:=xlWhole).Column

    lastCol = ws.Cells(10, ws.Columns.Count).End(xlToLeft).Column
 
    h = 2
 
    For i = colD + 1 To lastCol
 
        hasTriangle = False

        colColor = ThisWorkbook.sheets("Palette").Cells(h, 1).Interior.color
 
        For Each shp In ws.Shapes

            If shp.TopLeftCell.Column = i Then

                If InStr(1, shp.Name, "Triangle", vbTextCompare) > 0 Then
 

                    shp.Fill.ForeColor.RGB = colColor

                    hasTriangle = True
 
                End If

            End If

        Next shp
 
        If hasTriangle Then h = h + 1
 
    Next i
 
End Sub
