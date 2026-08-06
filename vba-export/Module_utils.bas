Attribute VB_Name = "Module_utils"
Option Explicit
Function lastRowByColumn(onglet As String, col As Integer)
   With ThisWorkbook.sheets(onglet)
        lastRowByColumn = .Cells(.Rows.Count, col).End(xlUp).row
   End With
End Function
Function EventAndScreen(v As Boolean)
   Application.EnableEvents = v
   Application.ScreenUpdating = v
End Function
Sub MajTargets()
   
    Dim v
    Dim i As Long
    Dim j As Integer
    Dim c As Range
    Dim tableVeh() As String
    Dim cur_cell_cal
    
    Application.StatusBar = "Mise à jour des target"
    i = 23
'    While Len(ThisWorkbook.sheets("RATING").Cells(i, 4)) > 0
'        If Not ThisWorkbook.sheets("RATING").Cells(i, 14).MergeCells Then
'            ThisWorkbook.sheets("RATING").Cells(i, 14).ClearContents
'        End If
'        i = i + 1
'    Wend
  
   If InStr(1, ThisWorkbook.sheets("HOME").Range("C23").Value, ",") <> 0 Then
      tableVeh = Split(ThisWorkbook.sheets("HOME").Range("C23").Value, ",")
   Else
     ReDim tableVeh(0)
     tableVeh(0) = ThisWorkbook.sheets("HOME").Range("C23").Value
   End If
   
   For j = 0 To UBound(tableVeh)
            v = ThisWorkbook.sheets("TARGET VEHICLE").UsedRange.Value
            For i = 2 To UBound(v, 1)
            
                If StrComp(ThisWorkbook.sheets("HOME").Range("DriveVersion").Value, v(i, 2), vbTextCompare) = 0 And _
                   StrComp(tableVeh(j), v(i, 3), vbTextCompare) = 0 And _
                   StrComp(ThisWorkbook.sheets("HOME").Range("Mode").Value, v(i, 4), vbTextCompare) = 0 Then
                    If sheetExists(v(i, 1)) Then
                        If checkCriteria(v(i, 1)) = True And checkCorrespondancePriority(v(i, 1)) = True Then
                                Set c = ThisWorkbook.sheets("RATING").Range("D23").CurrentRegion.Find(What:=v(i, 1), LookAt:=xlWhole)
                                If c Is Nothing Then
                                    MsgBox v(i, 1) & " non trouvé feuille RATING !", vbCritical
                                Else
                                    If Application.International(xlDecimalSeparator) = "," Then
                                        cur_cell_cal = replace(v(i, 5), ".", ",")
                                    Else
                                        cur_cell_cal = replace(v(i, 5), ",", ".")
                                    End If
                                    If IsNumeric(cur_cell_cal) Then
                                        ThisWorkbook.sheets("RATING").Cells(c.row, getTargetColonnePage(tableVeh(j), "DRIV")).Value = cur_cell_cal
                                        ThisWorkbook.sheets(v(i, 1)).Range("K5").Value = cur_cell_cal
                                    End If
                                      
                                       
                                End If
                        End If
                        
                         If checkCriteriaDyn(v(i, 1)) = True And checkCorrespondancePriorityDyn(v(i, 1)) = True Then
                                Set c = ThisWorkbook.sheets("RATING").Range("D23").CurrentRegion.Find(What:=v(i, 1), LookAt:=xlWhole)
                                If c Is Nothing Then
                                    MsgBox v(i, 1) & " non trouvé feuille RATING !", vbCritical
                                Else
                                  
                                    If Application.International(xlDecimalSeparator) = "," Then
                                        cur_cell_cal = replace(v(i, 6), ".", ",")
                                    Else
                                        cur_cell_cal = replace(v(i, 6), ",", ".")
                                    End If
                                       
                                            If IsNumeric(cur_cell_cal) Then
                                          
                                                ThisWorkbook.sheets("RATING").Cells(c.row, getTargetColonnePage(tableVeh(j), "DYN")).Value = cur_cell_cal
                                                ThisWorkbook.sheets(v(i, 1)).Range("BR5").Value = cur_cell_cal
                                            End If
                                       
                                End If
                        End If
                    End If
                End If
            Next i
            Erase v
            
     Next j
    Application.StatusBar = False

    
End Sub
   
Function CamFull(onglet As String, typeOff As String)
       Dim TabBoL(2) As Boolean
       Dim tabV(2) As String
       Dim cellule(9) As String
       Dim i As shape
       Dim n As Long, j As Long
       Dim sh As Worksheet
       Dim tGetSHap(2) As String
       Dim p
       
       If typeOff = "driv" Then
           tGetSHap(0) = "Graphique P1"
           tGetSHap(1) = "Graphique P2"
           tGetSHap(2) = "Graphique P3"
           n = 1
       Else
           tGetSHap(0) = "Graphique P11"
           tGetSHap(1) = "Graphique P12"
           tGetSHap(2) = "Graphique P13"
           n = 11
       End If
       
       If Len(onglet) > 0 Then Set sh = ThisWorkbook.sheets(onglet) Else Set sh = ActiveSheet
       TabBoL(0) = False
       TabBoL(1) = False
       TabBoL(2) = False
       For Each i In sh.Shapes
          If i.Name = tGetSHap(0) Then TabBoL(0) = True
          If i.Name = tGetSHap(1) Then TabBoL(1) = True
          If i.Name = tGetSHap(2) Then TabBoL(2) = True
       Next i
       If TabBoL(0) = False Or TabBoL(1) = False Or TabBoL(2) = False Then Exit Function
       
       If typeOff = "driv" Then
            cellule(1) = "K14"
            cellule(2) = "K11"
            cellule(3) = "K17"
            cellule(4) = "K15"
            cellule(5) = "K12"
            cellule(6) = "K18"
            cellule(7) = "K16"
            cellule(8) = "K13"
            cellule(9) = "K19"
      Else
            cellule(1) = "BR14"
            cellule(2) = "BR11"
            cellule(3) = "BR17"
            cellule(4) = "BR15"
            cellule(5) = "BR12"
            cellule(6) = "BR18"
            cellule(7) = "BR16"
            cellule(8) = "BR13"
            cellule(9) = "BR19"
      End If
      
      For n = n To n + 2
        If n = 1 Or n = 11 Then
          j = 0
        ElseIf n = 2 Or n = 12 Then
          j = 3
        ElseIf n = 3 Or n = 13 Then
          j = 6
        End If
        If sh.Range(cellule(j + 1)).Value < 0 Then tabV(0) = 0 Else tabV(0) = sh.Range(cellule(j + 1)).Value
        If sh.Range(cellule(j + 2)).Value < 0 Then tabV(1) = 0 Else tabV(1) = sh.Range(cellule(j + 2)).Value
        If sh.Range(cellule(j + 3)).Value < 0 Then tabV(2) = 0 Else tabV(2) = sh.Range(cellule(j + 3)).Value
        If Len(tabV(0)) = 0 Then tabV(0) = 0
        If Len(tabV(1)) = 0 Then tabV(1) = 0
        If Len(tabV(2)) = 0 Then tabV(2) = 0
        sh.ChartObjects("Graphique P" & n).Chart.FullSeriesCollection(2).Values = "{" & tabV(0) & "," & tabV(1) & "," & tabV(2) & "}"
        For p = 1 To 3
            If p = 1 Then sh.ChartObjects("Graphique P" & n).Chart.SeriesCollection(1).Points(p).Format.Fill.ForeColor.RGB = 65535
            If p = 1 Then sh.ChartObjects("Graphique P" & n).Chart.SeriesCollection(1).Points(p).Format.Fill.Transparency = 0
            If p = 2 Then sh.ChartObjects("Graphique P" & n).Chart.SeriesCollection(1).Points(p).Format.Fill.ForeColor.RGB = 32512
            If p = 2 Then sh.ChartObjects("Graphique P" & n).Chart.SeriesCollection(1).Points(p).Format.Fill.Transparency = 0
            If p = 3 Then sh.ChartObjects("Graphique P" & n).Chart.SeriesCollection(1).Points(p).Format.Fill.ForeColor.RGB = 255
            If p = 3 Then sh.ChartObjects("Graphique P" & n).Chart.SeriesCollection(1).Points(p).Format.Fill.Transparency = 0
            If p = 1 Then sh.ChartObjects("Graphique P" & n).Chart.SeriesCollection(2).Points(p).Format.Fill.ForeColor.RGB = 26367
            If p = 1 Then sh.ChartObjects("Graphique P" & n).Chart.SeriesCollection(2).Points(p).Format.Fill.Transparency = 0.49
            If p = 2 Then sh.ChartObjects("Graphique P" & n).Chart.SeriesCollection(2).Points(p).Format.Fill.ForeColor.RGB = 16777215
            If p = 2 Then sh.ChartObjects("Graphique P" & n).Chart.SeriesCollection(2).Points(p).Format.Fill.Transparency = 1
            If p = 3 Then sh.ChartObjects("Graphique P" & n).Chart.SeriesCollection(2).Points(p).Format.Fill.ForeColor.RGB = 255
            If p = 3 Then sh.ChartObjects("Graphique P" & n).Chart.SeriesCollection(2).Points(p).Format.Fill.Transparency = 0.49
        Next
     Next n
End Function






















