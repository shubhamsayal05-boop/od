Attribute VB_Name = "SupprimeEvent"
Option Explicit
Function createDeleteEvent()
    On Error Resume Next
    Application.CommandBars("Del_Event").Delete
    On Error GoTo 0
    Application.CommandBars.Add Name:="Del_Event", position:=msoBarPopup, Temporary:=True
    
    With Application.CommandBars("Del_Event").Controls.Add(msoControlButton)
        .Caption = "Delete"
         .FaceId = 2985
         '.FaceId = 3001
        .OnAction = "deleteById"
    End With
     
End Function
Function createDeleteEventDyn()
    On Error Resume Next
    Application.CommandBars("Del_Event").Delete
    On Error GoTo 0
    Application.CommandBars.Add Name:="Del_Event", position:=msoBarPopup, Temporary:=True
    
    With Application.CommandBars("Del_Event").Controls.Add(msoControlButton)
        .Caption = "Delete"
         .FaceId = 2985
         '.FaceId = 3001
        .OnAction = "deleteByIdDyn"
    End With
     
End Function
Function deleteById()
    Dim r As Range
    Dim col As Integer
    Dim Tots As Long
    Dim lignD As Long
    Dim cSup As Range
    Dim cDrivDyn As Range
    Dim rGet As Range
    Dim colIdSup As String
    Dim x As Integer
    Dim getR As Integer
    Dim RqOdb As Object
    Dim idc As String
    Dim rs0 As Variant
    Dim rs1 As Variant
    Dim rs2 As Variant
    Dim rs3 As Variant
    
   
    idc = getDbId(ThisWorkbook.Worksheets("Home").Range("idProjects"))
    Set RqOdb = db.GetOdb(val(idc))
    
    ThisWorkbook.Worksheets("Data").AutoFilterMode = False
    col = getLastColumnDrivability(ActiveSheet.Name)
    x = getLastColumnDinamyc(ActiveSheet.Name)
    If Selection.Cells.Count = 1 Then
           If Len(Cells(Selection.row, col)) > 0 Then
                rs0 = db.GetListValue("Select * From DATAID where N°=" & Cells(Selection.row, col).Value, RqOdb)
                rs1 = db.GetListValue("Select * From DATASUB1 where IDDATA=" & Cells(Selection.row, col).Value, RqOdb)
                rs2 = db.GetListValue("Select * From DATASUB2 where IDDATA=" & Cells(Selection.row, col).Value, RqOdb)
                rs3 = db.GetListValue("Select * From DATASUB3 where IDDATA=" & Cells(Selection.row, col).Value, RqOdb)
                If Not isEmpty(rs0) Then Call db.Execute("Delete * From DATAID where N°=" & Cells(Selection.row, col).Value, RqOdb)
                If Not isEmpty(rs1) Then Call db.Execute("Delete * From DATASUB1 where IDDATA=" & Cells(Selection.row, col).Value, RqOdb)
                If Not isEmpty(rs2) Then Call db.Execute("Delete * From DATASUB2 where IDDATA=" & Cells(Selection.row, col).Value, RqOdb)
                If Not isEmpty(rs3) Then Call db.Execute("Delete * From DATASUB3 where IDDATA=" & Cells(Selection.row, col).Value, RqOdb)
                
                lignD = ThisWorkbook.sheets("Data").Columns(1).Find(What:=Cells(Selection.row, col).Value, LookAt:=xlWhole).row
                ThisWorkbook.sheets("Data").Rows(lignD).EntireRow.Delete
                getR = searchById(x, Cells(Selection.row, col))
                Range("M" & Selection.row & ":" & Cells(Selection.row, col).Address).Delete Shift:=xlUp
                If getR <> 0 Then Range("BT" & getR & ":" & Cells(getR, x).Address).Delete Shift:=xlUp
          End If
    Else
            Application.ScreenUpdating = False
            Tots = TotEventSheet(ActiveSheet.Name)
           
            For Each r In Selection
                    If r.row > 6 And Len(Cells(r.row, col)) > 0 And r.row <= Tots Then
                         If colIdSup = "" Then colIdSup = Cells(r.row, col) Else colIdSup = colIdSup & ", " & Cells(r.row, col)
                         getR = searchById(x, Cells(r.row, col))
                         If cSup Is Nothing Then
                             Set cSup = Range("M" & r.row & ":" & Cells(r.row, col).Address)
                             If getR <> 0 Then
                                  If cDrivDyn Is Nothing Then
                                     Set cDrivDyn = Range("BT" & getR & ":" & Cells(getR, x).Address)
                                  Else
                                     Set cDrivDyn = Union(cDrivDyn, Range("BT" & getR & ":" & Cells(getR, x).Address))
                                  End If
                            End If
                        Else
                            Set cSup = Union(cSup, Range("M" & r.row & ":" & Cells(r.row, col).Address))
                             If getR <> 0 Then
                                  If cDrivDyn Is Nothing Then
                                     Set cDrivDyn = Range("BT" & getR & ":" & Cells(getR, x).Address)
                                  Else
                                     Set cDrivDyn = Union(cDrivDyn, Range("BT" & getR & ":" & Cells(getR, x).Address))
                                  End If
                            End If
                       
                       End If
                        
                         lignD = ThisWorkbook.sheets("Data").Columns(1).Find(What:=ActiveSheet.Cells(r.row, col), LookAt:=xlWhole).row
                         If rGet Is Nothing Then Set rGet = ThisWorkbook.sheets("Data").Cells(lignD, 1) Else _
                            Set rGet = Union(rGet, ThisWorkbook.sheets("Data").Cells(lignD, 1))
                            
                    End If
            Next r
            rs0 = db.GetListValue("Select * From DATAID where [N°] IN (" & colIdSup & ")", RqOdb)
            rs1 = db.GetListValue("Select * From DATASUB1 where IDDATA IN (" & colIdSup & ")", RqOdb)
            rs2 = db.GetListValue("Select * From DATASUB2 where IDDATA IN (" & colIdSup & ")", RqOdb)
            rs3 = db.GetListValue("Select * From DATASUB3 where IDDATA IN (" & colIdSup & ")", RqOdb)
            If Not isEmpty(rs0) Then Call db.Execute("Delete * From DATAID where [N°] IN (" & colIdSup & ")", RqOdb)
            If Not isEmpty(rs1) Then Call db.Execute("Delete * From DATASUB1 where IDDATA IN   (" & colIdSup & ")", RqOdb)
            If Not isEmpty(rs2) Then Call db.Execute("Delete * From DATASUB2 where IDDATA IN   (" & colIdSup & ")", RqOdb)
            If Not isEmpty(rs3) Then Call db.Execute("Delete * From DATASUB3 where IDDATA IN   (" & colIdSup & ")", RqOdb)
            
            db.CloseSudbConn
            ActiveSheet.Cells.EntireColumn.Hidden = False
            If Not cSup Is Nothing Then cSup.Delete Shift:=xlUp
            If Not cDrivDyn Is Nothing Then cDrivDyn.Delete Shift:=xlUp
            If Not rGet Is Nothing Then rGet.EntireRow.Delete
            Call gotoD("driv")
            Application.ScreenUpdating = True
    End If
End Function

Function deleteByIdDyn()
    Dim r As Range
    Dim col As Integer
    Dim Tots As Long
    Dim lignD As Long
    Dim cSup As Range
    Dim cDrivDyn As Range
    Dim rGet As Range
    Dim colIdSup As String
    Dim x As Integer
    Dim getR As Integer
    Dim RqOdb As Object
    Dim idc As String
    Dim rs0 As Variant
    Dim rs1 As Variant
    Dim rs2 As Variant
    Dim rs3 As Variant
   
    idc = getDbId(ThisWorkbook.Worksheets("Home").Range("idProjects"))
    Set RqOdb = db.GetOdb(val(idc))
    ThisWorkbook.Worksheets("Data").AutoFilterMode = False
    col = getLastColumnDinamyc(ActiveSheet.Name)
    x = getLastColumnDrivability(ActiveSheet.Name)
    
    If Selection.Cells.Count = 1 Then
           If Len(Cells(Selection.row, col)) > 0 Then
                rs0 = db.GetListValue("Select * From DATAID where N°=" & Cells(Selection.row, col).Value, RqOdb)
                rs1 = db.GetListValue("Select * From DATASUB1 where IDDATA=" & Cells(Selection.row, col).Value, RqOdb)
                rs2 = db.GetListValue("Select * From DATASUB2 where IDDATA=" & Cells(Selection.row, col).Value, RqOdb)
                rs3 = db.GetListValue("Select * From DATASUB3 where IDDATA=" & Cells(Selection.row, col).Value, RqOdb)
                If Not isEmpty(rs0) Then Call db.Execute("Delete * From DATAID where N°=" & Cells(Selection.row, col).Value, RqOdb)
                If Not isEmpty(rs1) Then Call db.Execute("Delete * From DATASUB1 where IDDATA=" & Cells(Selection.row, col).Value, RqOdb)
                If Not isEmpty(rs2) Then Call db.Execute("Delete * From DATASUB2 where IDDATA=" & Cells(Selection.row, col).Value, RqOdb)
                If Not isEmpty(rs3) Then Call db.Execute("Delete * From DATASUB3 where IDDATA=" & Cells(Selection.row, col).Value, RqOdb)
                
                lignD = ThisWorkbook.sheets("Data").Columns(1).Find(What:=Cells(Selection.row, col).Value, LookAt:=xlWhole).row
                ThisWorkbook.sheets("Data").Rows(lignD).EntireRow.Delete
                Range("BT" & Selection.row & ":" & Cells(Selection.row, col).Address).Delete Shift:=xlUp
                getR = searchById(x, Cells(Selection.row, col))
                Range("M" & getR & ":" & Cells(getR, x).Address).Delete Shift:=xlUp
          End If
    Else
            Application.ScreenUpdating = False
            Tots = TotEventSheet(ActiveSheet.Name)
           
            For Each r In Selection
                    If r.row > 6 And Len(Cells(r.row, col)) > 0 And r.row <= Tots Then
                         If colIdSup = "" Then colIdSup = Cells(r.row, col) Else colIdSup = colIdSup & ", " & Cells(r.row, col)
                         getR = searchById(x, Cells(r.row, col))
                         If cSup Is Nothing Then
                             Set cSup = Range("BT" & r.row & ":" & Cells(r.row, col).Address)
                              If getR <> 0 Then
                                  If cDrivDyn Is Nothing Then
                                     Set cDrivDyn = Range("M" & getR & ":" & Cells(getR, x).Address)
                                  Else
                                     Set cDrivDyn = Union(cDrivDyn, Range("M" & getR & ":" & Cells(getR, x).Address))
                                  End If
                            End If
                        Else
                             Set cSup = Union(cSup, Range("BT" & r.row & ":" & Cells(r.row, col).Address))
                             If getR <> 0 Then
                                  If cDrivDyn Is Nothing Then
                                     Set cDrivDyn = Range("M" & getR & ":" & Cells(getR, x).Address)
                                  Else
                                     Set cDrivDyn = Union(cDrivDyn, Range("M" & getR & ":" & Cells(getR, x).Address))
                                  End If
                            End If
                            
                        End If
                         
                         lignD = ThisWorkbook.sheets("Data").Columns(1).Find(What:=ActiveSheet.Cells(r.row, col), LookAt:=xlWhole).row
                         If rGet Is Nothing Then Set rGet = ThisWorkbook.sheets("Data").Cells(lignD, 1) Else _
                            Set rGet = Union(rGet, ThisWorkbook.sheets("Data").Cells(lignD, 1))
                            
                    End If
            Next r
            
            rs0 = db.GetListValue("Select * From DATAID where [N°] IN (" & colIdSup & ")", RqOdb)
            rs1 = db.GetListValue("Select * From DATASUB1 where IDDATA IN (" & colIdSup & ")", RqOdb)
            rs2 = db.GetListValue("Select * From DATASUB2 where IDDATA IN (" & colIdSup & ")", RqOdb)
            rs3 = db.GetListValue("Select * From DATASUB3 where IDDATA IN (" & colIdSup & ")", RqOdb)
            If Not isEmpty(rs0) Then Call db.Execute("Delete * From DATAID where [N°] IN (" & colIdSup & ")", RqOdb)
            If Not isEmpty(rs1) Then Call db.Execute("Delete * From DATASUB1 where IDDATA IN   (" & colIdSup & ")", RqOdb)
            If Not isEmpty(rs2) Then Call db.Execute("Delete * From DATASUB2 where IDDATA IN   (" & colIdSup & ")", RqOdb)
            If Not isEmpty(rs3) Then Call db.Execute("Delete * From DATASUB3 where IDDATA IN   (" & colIdSup & ")", RqOdb)
           
            ActiveSheet.Cells.EntireColumn.Hidden = False
            If Not cSup Is Nothing Then cSup.Delete Shift:=xlUp
            If Not rGet Is Nothing Then rGet.EntireRow.Delete
            If Not cDrivDyn Is Nothing Then cDrivDyn.Delete Shift:=xlUp
            Call gotoD("dyn")
            
            db.CloseSudbConn
             
            Application.ScreenUpdating = True
    End If
    
    
End Function



Function searchById(col As Integer, id As String)
    Dim j As Integer
    j = 7
    searchById = 0
    While Len(Cells(j, col)) > 0
        If Cells(j, col) = id Then
            searchById = j
            Exit Function
        End If
        j = j + 1
    Wend
End Function




