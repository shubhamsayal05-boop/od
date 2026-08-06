Attribute VB_Name = "TargetVehiclePoint"
Function getTargetColonnePage(colname As String, part As String)
    Dim i As Integer
    
    getTargetColonnePage = 0
    With ThisWorkbook.sheets("RATING")
        If part = "DRIV" Then
            i = .Rows("21:22").Find(What:="Driveability Index", LookAt:=xlWhole).Column
            While Len(.Cells(21, i)) > 0 And .Cells(21, i) <> "Responsiveness Index"
                If colname = .Cells(21, i) Then
                    getTargetColonnePage = i
                    Exit Function
                End If
                i = i + 1
            Wend
            
        Else
             i = .Rows("21:22").Find(What:="Responsiveness Index", LookAt:=xlWhole).Column
            While Len(.Cells(21, i)) > 0
                If colname = .Cells(21, i) Then
                    getTargetColonnePage = i
                    Exit Function
                End If
                i = i + 1
            Wend
        End If
    End With
End Function

Function getTargetColonneCible(colname As String)
     getTargetColonneCible = 0
    With ThisWorkbook.sheets("RATING")
         If Not .Rows("10:10").Find(What:=colname, LookAt:=xlWhole) Is Nothing Then
            getTargetColonneCible = .Rows("10:10").Find(What:=colname, LookAt:=xlWhole).Column
         End If
    End With
End Function

Function getTargetRowGraphStatusDrivability(colname As String)
    Dim lastr As Integer
    Dim i As Integer
    
    getTargetRowGraphStatusDrivability = 0
    With ThisWorkbook.sheets("Graph_status")
          lastr = .Cells(.Rows.Count, 1).End(xlUp).row
          
          i = .Columns("A:E").Find(What:="DRIVABILITY", LookAt:=xlWhole).row
          While i <= lastr And .Cells(i, 1) <> "DYNAMIC"
             If .Cells(i, 1) = colname Then
                getTargetRowGraphStatusDrivability = i
                Exit Function
             End If
             i = i + 1
          Wend
        
    End With
End Function

Function getTargetRowGraphStatusDynamic(colname As String)
    Dim lastr As Integer
    Dim i As Integer
    
    getTargetRowGraphStatusDynamic = 0
    With ThisWorkbook.sheets("Graph_status")
          lastr = .Cells(.Rows.Count, 1).End(xlUp).row
          
          i = .Columns("A:E").Find(What:="RESPONSIVENESS", LookAt:=xlWhole).row
          While i <= lastr
             If .Cells(i, 1) = colname Then
                getTargetRowGraphStatusDynamic = i
                Exit Function
             End If
             i = i + 1
          Wend
        
    End With
End Function

Function getTauxRowGraphStatusDrivability(colname As String)
    Dim lastr As Integer
    Dim i As Integer
    
    getTauxRowGraphStatusDrivability = 0
    With ThisWorkbook.sheets("Graph_status")
          lastr = .Cells(.Rows.Count, 1).End(xlUp).row
          
          i = .Columns("A:E").Find(What:="DRIVABILITY", LookAt:=xlWhole).row
          i = i + 2
          While .Cells(i, 1) <> "Global index"
             i = i + 1
          Wend
          
          While i <= lastr And .Cells(i, 1) <> "DYNAMIC"
             If .Cells(i, 1) = colname Then
                getTauxRowGraphStatusDrivability = i
                Exit Function
             End If
             i = i + 1
          Wend
        
    End With
End Function


Function getTauxRowGraphStatusDynamic(colname As String)
    Dim lastr As Integer
    Dim i As Integer
    
    getTauxRowGraphStatusDynamic = 0
    With ThisWorkbook.sheets("Graph_status")
          lastr = .Cells(.Rows.Count, 1).End(xlUp).row
          
          i = .Columns("A:E").Find(What:="RESPONSIVENESS", LookAt:=xlWhole).row
          i = i + 2
          While .Cells(i, 1) <> "Global index"
             i = i + 1
          Wend
          
          While i <= lastr
             If .Cells(i, 1) = colname Then
                getTauxRowGraphStatusDynamic = i
                Exit Function
             End If
              i = i + 1
          Wend
        
    End With
End Function

Function CalcNoteGlobal()
   Dim i As Integer
   Dim tableVeh() As String
   Dim x, y
   If InStr(1, ThisWorkbook.sheets("HOME").Range("C23").Value, ",") <> 0 Then
      tableVeh = Split(ThisWorkbook.sheets("HOME").Range("C23").Value, ",")
   Else
     ReDim tableVeh(0)
     tableVeh(0) = ThisWorkbook.sheets("HOME").Range("C23").Value
   End If
   
   For i = 0 To UBound(tableVeh)
        If getTargetColonnePage(tableVeh(i), "DRIV") <> 0 And getTargetColonnePage(tableVeh(i), "DYN") <> 0 Then
             x = getTargetRowGraphStatusDrivability(tableVeh(i))
             y = getTargetRowGraphStatusDynamic(tableVeh(i))
             With ThisWorkbook.sheets("Graph_status")
                .Cells(x, 2) = GetNoteGlobalTarget("driv", tableVeh(i))
                .Cells(x, 2) = IIf(.Cells(x, 2) = -555, "", Round(.Cells(x, 2), 1))
                .Cells(y, 2) = GetNoteGlobalTarget("dyn", tableVeh(i))
                .Cells(y, 2) = IIf(.Cells(y, 2) = -555, "", Round(.Cells(y, 2), 1))
             End With
        End If
   Next i
   
End Function

Function CalculIndexTarget()
    hideShowTarget (True)
    Call UpdateMarqueursRating
    Call UpdateGraphStatus
    CalcNoteGlobal
    CalcTauxGlobal
    TracePoint
   
End Function

Function CalcTauxGlobal()
   Dim i As Integer
   Dim tableVeh() As String
   Dim x, y
   If InStr(1, ThisWorkbook.sheets("HOME").Range("C23").Value, ",") <> 0 Then
      tableVeh = Split(ThisWorkbook.sheets("HOME").Range("C23").Value, ",")
   Else
     ReDim tableVeh(0)
     tableVeh(0) = ThisWorkbook.sheets("HOME").Range("C23").Value
   End If
   
   For i = 0 To UBound(tableVeh)
        If getTargetColonnePage(tableVeh(i), "DRIV") <> 0 And getTargetColonnePage(tableVeh(i), "DYN") <> 0 Then
             x = getTauxRowGraphStatusDrivability(tableVeh(i))
             y = getTauxRowGraphStatusDynamic(tableVeh(i))
             With ThisWorkbook.sheets("Graph_status")
                .Cells(x, 2) = GetTaux(tableVeh(i))
                .Cells(y, 2) = GetTauxDyn(tableVeh(i))
             End With
        End If
   Next i
   
End Function
Function hideShowTarget(hide As Boolean)
        Dim i As Integer
        Dim colname As String
      
        
        
       With ThisWorkbook.sheets("RATING")
             If hide = True Then
                 
                  colname = ThisWorkbook.sheets("HOME").Range("C23").Value
                 
                   i = .Rows("21:22").Find(What:="Driveability Index", LookAt:=xlWhole).Column
                    i = i + 1
                    While Len(.Cells(21, i)) > 0 And .Cells(21, i) <> "Drivability Lowest Events"
                        If InStr(1, "," & colname & ",", "," & .Cells(21, i) & ",") = 0 Then
                            .Columns(i).EntireColumn.Hidden = True
                        End If
                        i = i + 1
                    Wend
                          
                    i = .Rows("21:22").Find(What:="Responsiveness Index", LookAt:=xlWhole).Column
                    i = i + 1
                    While Len(.Cells(21, i)) > 0 And .Cells(21, i) <> "Responsiveness Lowest Events"
                        If InStr(1, "," & colname & ",", "," & .Cells(21, i) & ",") = 0 Then
                            .Columns(i).EntireColumn.Hidden = True
                        End If
                        i = i + 1
                    Wend
                    
                     
                    i = .Rows("10:10").Find(What:="Tested vehicle", LookAt:=xlWhole).Column + 1
                    While Len(.Cells(10, i)) > 0
                        If InStr(1, "," & colname & ",", "," & .Cells(10, i) & ",") = 0 Then
                            .Columns(i).EntireColumn.Hidden = True
                        End If
                        i = i + 1
                    Wend
            Else
                 .Columns.Hidden = False
            End If
       End With
        
End Function
Function removePoint(nameGraph As String)
    Dim i As Integer
    Dim j As Integer
    Dim graph() As String
    
    Dim KJ As FullSeriesCollection
    graph = Split(nameGraph, ";")
    
    With ThisWorkbook.Worksheets("RATING")
        For i = 0 To UBound(graph)
            If .ChartObjects(graph(i)).Chart.FullSeriesCollection.Count > 4 Then
                For j = .ChartObjects(graph(i)).Chart.FullSeriesCollection.Count To 4 Step -1
                   If j = 4 Then Exit For
                   .ChartObjects(graph(i)).Chart.FullSeriesCollection(j).Delete
                Next j
            End If
        Next i
    End With
End Function
Sub FJFJ()
    TracePoint
End Sub
Function TracePoint()
        Dim colname
        Dim nameGraph As String
        Dim graph() As String
        Dim i As Integer, j As Integer, k As Integer, LigneStatus As Integer
        Dim color As String, pointName As String
        
        nameGraph = "Graphique 1;Graphique 2;Graphique 3;Graphique 4"
        colname = ThisWorkbook.sheets("HOME").Range("C23").Value
        graph = Split(nameGraph, ";")
        Call removePoint(nameGraph)
        
        With ThisWorkbook.Worksheets("RATING")
         i = .Rows("10:10").Find(What:="Tested vehicle", LookAt:=xlWhole).Column + 1
         k = 5
                While Len(.Cells(10, i)) > 0
                    If InStr(1, "," & colname & ",", "," & .Cells(10, i) & ",") <> 0 Then
                       
                        For j = 0 To UBound(graph)
                             
                             .ChartObjects(graph(j)).Chart.SeriesCollection.NewSeries
                             
                            If j = 0 Then LigneStatus = getTargetRowGraphStatusDrivability(.Cells(10, i).Value)
                            If j = 2 Then LigneStatus = getTargetRowGraphStatusDynamic(.Cells(10, i).Value)
                            If j = 1 Then LigneStatus = getTauxRowGraphStatusDrivability(.Cells(10, i).Value)
                            If j = 3 Then LigneStatus = getTauxRowGraphStatusDynamic(.Cells(10, i).Value)
                              
                             
                             .ChartObjects(graph(j)).Chart.FullSeriesCollection(k).ChartType = xlXYScatter
                             .ChartObjects(graph(j)).Chart.FullSeriesCollection(k).Name = "=Graph_status!$A$" & LigneStatus
                             .ChartObjects(graph(j)).Chart.FullSeriesCollection(k).XValues = "=Graph_status!$C$" & LigneStatus
                             .ChartObjects(graph(j)).Chart.FullSeriesCollection(k).Values = "=Graph_status!$D$" & LigneStatus
                             
                             .ChartObjects(graph(j)).Chart.FullSeriesCollection(k).MarkerSize = 24
                             .ChartObjects(graph(j)).Chart.FullSeriesCollection(k).MarkerStyle = 3
                             
                             color = rgbFromInterior(sheets("Graph_status").Cells(LigneStatus, 5).Interior.color)
                             .ChartObjects(graph(j)).Chart.FullSeriesCollection(k).Format.Fill.ForeColor.RGB = RGB(Split(color, ";")(0), Split(color, ";")(1), Split(color, ";")(2))
                             .ChartObjects(graph(j)).Chart.FullSeriesCollection(k).Format.line.ForeColor.RGB = RGB(Split(color, ";")(0), Split(color, ";")(1), Split(color, ";")(2))
                             
                             pointName = getPointColumn(i, "driv")
                             If pointName <> "" Then .Shapes(pointName).Fill.ForeColor.RGB = RGB(Split(color, ";")(0), Split(color, ";")(1), Split(color, ";")(2))
                              
                              pointName = getPointColumn(i, "dyn")
                             If pointName <> "" Then .Shapes(pointName).Fill.ForeColor.RGB = RGB(Split(color, ";")(0), Split(color, ";")(1), Split(color, ";")(2))
                             
                             
                        Next j
                        k = k + 1
                    End If
                    i = i + 1
                Wend
         End With
         
         
End Function

Function rgbFromInterior(colorInterior As Long) As String
    Dim r As Integer, g As Integer, b As Integer
    
    r = colorInterior Mod 256
    g = (colorInterior \ 256) Mod 256
    b = (colorInterior \ 256 \ 256) Mod 256
    
    rgbFromInterior = r & ";" & g & ";" & b
    
End Function

Function getPointColumn(col As Integer, part As String) As String
            Dim shp As shape
            getPointColumn = ""
            For Each shp In ThisWorkbook.Worksheets("RATING").Shapes
                If shp.AutoShapeType = msoShapeIsoscelesTriangle Then
                    If shp.TopLeftCell.Column = col Then
                        If part = "driv" And shp.TopLeftCell.row >= 9 And shp.TopLeftCell.row <= 10 Then
                            getPointColumn = shp.Name
                        ElseIf part = "dyn" And shp.TopLeftCell.row >= 14 And shp.TopLeftCell.row <= 15 Then
                            getPointColumn = shp.Name
                        End If
                    End If
                End If
            Next shp
End Function








