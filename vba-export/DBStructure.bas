Attribute VB_Name = "DBStructure"
Option Explicit

Function getInterval()
    Dim req As Object
    
    Set req = db.Request("Select IntervalCol, TableCol from dataInterval order by N°")
    
    With ThisWorkbook.Worksheets("DBStructure")
        .Range("A2:B1000").ClearContents
        If Not req Is Nothing Then
            .Range("A2").CopyFromRecordset req
            req.Close
            Set req = Nothing
        End If
    End With
End Function


Function getEntete()
    Dim req As Object
    
    Set req = db.Request("Select IdCol, TableCol, DescriptionCol from entete order by N°")
    
    With ThisWorkbook.Worksheets("DBStructure")
        .Range("D2:F2000").ClearContents
        If Not req Is Nothing Then
            .Range("D2").CopyFromRecordset req
            req.Close
            Set req = Nothing
        End If
    End With
End Function

Function getTableByCol(numCol As Integer) As String
    Dim r As Range
    Dim lastRow As Integer
    getTableByCol = ""
    With ThisWorkbook.Worksheets("DBStructure")
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
        For Each r In .Range("A2:A" & lastRow)
            If numCol >= Split(r.Value, "-")(0) And numCol <= Split(r.Value, "-")(1) Then
                getTableByCol = r.Offset(0, 1).Value
                Exit Function
            End If
        Next r
    End With
End Function

Function getTableByDescription(description As String) As String
  Dim r As Range
  getTableByDescription = ""
 
  With ThisWorkbook.Worksheets("DBStructure")
        Set r = .Columns(6).Cells.Find(What:=description, LookAt:=xlWhole)
        If Not r Is Nothing Then
              getTableByDescription = r.Offset(0, -1).Value
        End If
   End With
   

End Function

Function getColumnByDescription(description As String) As String
  Dim r As Range
  
  getColumnByDescription = ""
  With ThisWorkbook.Worksheets("DBStructure")
        Set r = .Columns(6).Find(What:=description, LookAt:=xlWhole)
        If Not r Is Nothing Then
              getColumnByDescription = r.Offset(0, -2).Value
        End If
   End With
End Function

Function getTableByListCol(list As String) As String
   Dim r As Range
   Dim tabl() As String
   Dim listCol As String
   Dim i As Long
   
   tabl = Split(list, ", ")
   getTableByListCol = ""
 
   With ThisWorkbook.Worksheets("DBStructure")
        For i = 0 To UBound(tabl)
             Set r = .Columns(4).Find(What:=replace(tabl(i), " ", ""), LookAt:=xlWhole)
             If Not r Is Nothing Then
                   If InStr(1, ";" & listCol & ";", ";" & r.Offset(0, 1) & ";") = 0 Then
                        If listCol = "" Then listCol = r.Offset(0, 1) Else listCol = listCol & ";" & r.Offset(0, 1)
                   End If
             End If
        Next i
        
        getTableByListCol = listCol
   End With

End Function
Function getListColFromTable(list As String, tableCol As String) As String
   Dim r As Range
   Dim tabl() As String
   Dim listCol As String
   Dim i As Long
   
   tabl = Split(list, ", ")
   getListColFromTable = ""
 
   With ThisWorkbook.Worksheets("DBStructure")
        For i = 0 To UBound(tabl)
             Set r = .Columns(4).Find(What:=replace(tabl(i), " ", ""), LookAt:=xlWhole)
             If Not r Is Nothing Then
                   If InStr(1, ", " & listCol & ", ", ", " & r.Offset(0, 1) & ", ") = 0 And r.Offset(0, 1) = tableCol Then
                        If listCol = "" Then listCol = r.Value Else listCol = listCol & ", " & r.Value
                   End If
             End If
        Next i
        
        getListColFromTable = listCol
   End With

End Function


Function getDescriptionByCol(colN As String) As String
        Dim r As Range
  
        getDescriptionByCol = ""
        With ThisWorkbook.Worksheets("DBStructure")
              Set r = .Columns(4).Find(What:=colN, LookAt:=xlWhole)
              If Not r Is Nothing Then
                    getDescriptionByCol = r.Offset(0, 2).Value
              End If
         End With
End Function







