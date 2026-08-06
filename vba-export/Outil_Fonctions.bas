Attribute VB_Name = "Outil_Fonctions"
Option Explicit
Function SDV2Ncrit(ByVal SDV As String) As Integer
    Dim o As Long, n As Long
    Dim st As String
    Dim NB As Integer
    NB = 0
    SDV2Ncrit = 0
    
     st = getNumberRow(SDV)
     If st = "" Then Exit Function
     o = val(Split(st, ";")(1))
     n = val(Split(st, ";")(0)) + 1
     Do While o >= n
            With ThisWorkbook.Worksheets("structure")
                If .Range("C" & n) = "criteria" Then
                   NB = NB + 1
                End If
            End With
            n = n + 1
     Loop
 
    SDV2Ncrit = NB
    
End Function

Function SDV2Nrow(ByVal SDV As String) As Integer
    'SDV2Ncrit = ThisWorkbook.Sheets("Structure").Cells(4, numSDV(sdv) + 3).Value
    
    Dim v
    Dim i As Long
    Dim NB As String
    NB = 0
    
    v = ThisWorkbook.sheets("structure").UsedRange.Columns(2).Value
    For i = 2 To UBound(v, 1)
        If StrComp(SDV, v(i, 1), vbTextCompare) = 0 Then
            NB = i
        End If
    Next i
    
    SDV2Nrow = NB
    
    Erase v
End Function
Function order(ByVal SDV As String, position As Integer) As String
Dim i As Integer
Dim j As Integer
With ThisWorkbook.Worksheets("RATING")
                i = getLastRowRating
                For j = 23 To i
                  If LCase(.Range("D" & j).Value) = LCase(SDV) Then
                      If position = 0 Then
                        order = .Range("E" & j).Address
                      ElseIf position = 1 Then
                        order = .Range("W" & j).Address
                      End If
                      Exit Function
                  End If
                Next j
                 
End With
End Function


Function tauxPts(ByVal SDV As String, ByVal Milestone As Integer) As Variant
     Dim rowSDV As Integer
    Dim k As Integer
    Dim taux(6) As Integer
    Dim col As Integer
    Dim prio As String
    Dim r As Range
    
    prio = ThisWorkbook.sheets("HOME").Range("Prestation").Value
    
'    If Len(sdv) >= 31 Then
'        sdv = Left(sdv, 30)
'    End If
    With ThisWorkbook.sheets("SETTINGS")
        
        Set r = .UsedRange.Find(What:=SDV, LookIn:=xlValues, LookAt:=xlWhole)
        
        If r Is Nothing Then
            Exit Function
        End If
        
        For k = 1 To 6
             taux(k) = .Cells(r.row + k + 3, Milestone + 1).Value
        Next
        
    End With
    

    tauxPts = taux
    
End Function

Function NbMinPts(ByVal SDV As String, ByVal Milestone As Integer) As Variant
    Dim rowSDV As Integer
    Dim k As Integer
    Dim NbMin(6) As Integer
    Dim r As Range
    Dim col As Integer
    Dim prio As String
    
    prio = ThisWorkbook.sheets("HOME").Range("Prestation").Value
    
'     If Len(sdv) >= 31 Then
'        sdv = Left(sdv, 30)
'    End If

   With ThisWorkbook.sheets("SETTINGS")
        
        Set r = .UsedRange.Find(What:=SDV, LookIn:=xlValues, LookAt:=xlWhole)
        
        If r Is Nothing Then
            Exit Function
        End If
        
        For k = 1 To 6
            NbMin(k) = .Cells(r.row + k + 3, Milestone + 7).Value
        Next
        
    End With

    NbMinPts = NbMin

End Function

Function Weight(ByVal SDV As String) As Variant
    Dim rowSDV As Integer
    Dim r As Range
    Dim col As Integer
    Dim prio As String
    
    prio = ThisWorkbook.sheets("HOME").Range("Prestation").Value
    Weight = 0
    With ThisWorkbook.sheets("SETTINGS")
        Set r = .UsedRange.Find(What:=SDV, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not r Is Nothing Then
             With ThisWorkbook.sheets("SETTINGS")
                      Weight = .Cells(r.row + 11, 3).Value
             End With
        End If
        
    End With

End Function

Function OvMinPts(ByVal SDV As String) As Variant
    Dim rowSDV As Integer
    Dim r As Range
    Dim prio As String
    Dim col As Integer
    prio = ThisWorkbook.sheets("HOME").Range("Prestation").Value
    
    OvMinPts = 0
    With ThisWorkbook.sheets("SETTINGS")
        Set r = .UsedRange.Find(What:=SDV, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not r Is Nothing Then
             With ThisWorkbook.sheets("SETTINGS")
                      OvMinPts = .Cells(r.row + 11, 11).Value
             End With
        End If
        
    End With
    
sortir:
   If ERR.Number <> 0 Then
    OvMinPts = 0
    ERR.Clear
   End If
End Function

Function ExecuteAccessQuery(req As String, dbS As Object)
  
   Dim rst As Object
    Set rst = dbS.QueryDefs("SupReq")
    rst.sql = req
    rst.Execute


End Function

Function getLastColumnDrivability(SDV As String) As Integer
    Dim r As Range
    Dim i As Integer
    
    With ThisWorkbook.sheets(SDV)
            Set r = .Range("M6")
            While Len(r.Value) > 0
                    i = r.Column
                    Set r = r.Offset(0, 1)
            Wend
            getLastColumnDrivability = i
    End With
    
End Function


Function getLastColumnDinamyc(SDV As String) As Integer
    Dim r As Range
    Dim i As Integer
    
    With ThisWorkbook.sheets(SDV)
            Set r = .Range("BT6")
            While Len(r.Value) > 0
                    i = r.Column
                    Set r = r.Offset(0, 1)
            Wend
            getLastColumnDinamyc = i
    End With
    
End Function

Sub goDynamyc()
   If ActiveSheet.Range("A2:BG2").EntireColumn.Hidden = True Then
        Call gotoD("driv")
    Else
        Call gotoD("dyn")
    End If
End Sub
Function gotoD(typeOff As String)
        Dim targ As Range
        Dim wnd As Window
        Dim scaling As Long
        
        If typeOff = "driv" Then
            Set targ = ActiveSheet.Range("A1:AG62")
            ActiveSheet.Range("A2:BG2").EntireColumn.Hidden = False
            ActiveSheet.Range("BH2:FH2").EntireColumn.Hidden = True
            ActiveSheet.Shapes.Range(Array("TITRESNAME")).TextFrame2.TextRange.Characters.text = "DRIVABILITY"
            ActiveSheet.Range("B1").Select
            Call HideC3(ActiveSheet.Name, "driv")
            ActiveSheet.AutoFilterMode = False
        ElseIf typeOff = "dyn" Then
            Set targ = ActiveSheet.Range("BH1:CO62")
            ActiveSheet.Range("A2:BG2").EntireColumn.Hidden = True
            ActiveSheet.Range("BH2:FH2").EntireColumn.Hidden = False
            ActiveSheet.Shapes.Range(Array("TITRESNAME")).TextFrame2.TextRange.Characters.text = "RESPONSIVENESS"
            ActiveSheet.Range("BI1").Select
            Call HideC3(ActiveSheet.Name, "dyn")
            ActiveSheet.AutoFilterMode = False
        End If
        
        Set wnd = ActiveWindow
        wnd.ScrollColumn = 1
      
        
        
End Function
Function initDriv(onglet As String, partPath As String)
        If partPath = "driv" Then
            ThisWorkbook.sheets(onglet).Range("A2:BG2").EntireColumn.Hidden = False
            ThisWorkbook.sheets(onglet).Range("BH2:FH2").EntireColumn.Hidden = True
            ThisWorkbook.sheets(onglet).Shapes.Range(Array("TITRESNAME")).TextFrame2.TextRange.Characters.text = "DRIVABILITY"
            Call HideC3(ThisWorkbook.sheets(onglet).Name, "driv")
        Else
            ThisWorkbook.sheets(onglet).Range("A2:BG2").EntireColumn.Hidden = True
            ThisWorkbook.sheets(onglet).Range("BH2:FH2").EntireColumn.Hidden = False
            ThisWorkbook.sheets(onglet).Shapes.Range(Array("TITRESNAME")).TextFrame2.TextRange.Characters.text = "RESPONSIVENESS"
            Call HideC3(ThisWorkbook.sheets(onglet).Name, "dyn")
        End If
End Function

Function getLastRowRating()
    Dim DL As Integer
    With ThisWorkbook.sheets("RATING")
        DL = .Cells(.Rows.Count, 2).End(xlUp).row
        If .Cells(.Rows.Count, 4).End(xlUp).row > DL Then DL = .Cells(.Rows.Count, 4).End(xlUp).row
        getLastRowRating = DL
    End With
End Function

Function MaskEmptySdv()
    Dim DL As Integer
    Dim i
    Dim tJ
    With ThisWorkbook.sheets("RATING")
        DL = getLastRowRating
        .Rows(2 & ":" & DL).EntireRow.Hidden = False
        For i = 23 To DL
            If Len(.Range("B" & i)) > 0 Then
                tJ = i
            End If
            If .Range("C" & i).Font.Size <> 2 Then
                .Rows(i).EntireRow.Hidden = True
            ElseIf tJ <> 0 Then
                    .Rows(tJ).EntireRow.Hidden = False
                     tJ = 0
            End If
        Next i
    End With
End Function

Sub AutoBackupWorkbook()
    Dim saveInterval As String
    Dim nextBackupTime As Double
    Dim backupPath As String
    Dim rRow As Integer
    
    rRow = ThisWorkbook.sheets("CONFIGURATIONS").Range("auto_saves").row
    If val(ThisWorkbook.sheets("CONFIGURATIONS").Range("B" & rRow + 2)) < 5 Then
        MsgBox "Temps Minimum 5 minutes", vbCritical, "ODRIV"
        Exit Sub
    End If
    If ThisWorkbook.sheets("CONFIGURATIONS").Range("B" & rRow + 1) = 1 Then
        If (Dir(ThisWorkbook.sheets("CONFIGURATIONS").Range("B" & rRow + 3) & "\", vbDirectory) = vbNullString) Then
            MkDir ThisWorkbook.sheets("CONFIGURATIONS").Range("B" & rRow + 3) & "\"
        End If
        saveInterval = ThisWorkbook.sheets("CONFIGURATIONS").Range("B" & rRow + 2)
        backupPath = ThisWorkbook.sheets("CONFIGURATIONS").Range("B" & rRow + 3) & "\" & ThisWorkbook.Name & Format(Now, "yyyymmdd") & ".xlsm"
        ThisWorkbook.SaveCopyAs backupPath
        nextBackupTime = Now + TimeValue("00:" & saveInterval & ":00")
        Application.OnTime nextBackupTime, "AutoBackupWorkbook"
   End If

End Sub

Function getDbId(id As String) As String
    Dim req As Object
    Dim ids As String
    Set req = db.Request("Select db_name From projet" & db.AnneeEnCours & " Where code IN (" & id & ")")
    If Not req Is Nothing Then
        While Not req.EOF
            If InStr(1, "," & ids & ",", "," & Right(req.Fields(0).Value, 1) & ",") = 0 Then
                ids = Right(req.Fields(0).Value, 1)
            End If
            req.MoveNext
        Wend
        req.Close
    End If
    getDbId = ids
End Function
Function dbNameBalancer()
    Dim dbCount(1 To 4) As Integer
    Dim i As Integer, minCount As Integer, dbIndex As Integer
    Dim dbName As String
    
    dbCount(1) = countProject(1)
    dbCount(2) = countProject(2)
    dbCount(3) = countProject(3)
    dbCount(4) = countProject(4)
    
    minCount = 51
    dbIndex = -1
    
    For i = 1 To 4
        If dbCount(i) < minCount And dbCount(i) < 50 Then
            minCount = dbCount(i)
            dbIndex = i
        End If
    Next i
    
    If dbIndex <> -1 Then
        dbName = "_OdrivDB_" & dbIndex
    Else
        dbName = "_OdrivDB_4"
    End If
    dbNameBalancer = dbName
End Function

Function countProject(conn As Integer)
    Dim RS As Object
    Dim nbMax As Integer
    Dim RqOdb As Object
    Dim tot As Integer
    
   tot = 55
    Set RqOdb = db.GetOdb(conn)
   
    Set RS = db.Request("SELECT COUNT (*) As ProjectCount From projet", RqOdb)
    
    If Not RS Is Nothing Then
        tot = RS.Fields("ProjectCount").Value
    End If
    countProject = tot
End Function

Function orderColDyn(onglet As String, Optional orderP As Integer)
    Dim lrow
    Dim lcol
    Dim lig
    Dim rangeTry
    
    With ThisWorkbook.sheets(onglet)
       lcol = getLastColumnDinamyc(onglet)
        For lig = 72 To lcol
           If .Cells(.Rows.Count, lig).End(xlUp).row > lrow Then lrow = .Cells(.Rows.Count, lig).End(xlUp).row
        Next lig
        
        Set rangeTry = .Range(.Cells(7, 73), .Cells(lrow, lcol))
        If lrow > 6 Then
            If .Cells(.Rows.Count, 74).End(xlUp).row > 6 Or _
               .Cells(.Rows.Count, 78).End(xlUp).row > 6 Or _
               .Cells(.Rows.Count, 79).End(xlUp).row > 6 Then
               
               With .Sort
                        .SortFields.Clear
'                        .SortFields.Add key:=rangeTry.Columns(2), order:=xlAscending, CustomOrder:=orderP
                         If orderP = 0 Then
                            .SortFields.Add key:=rangeTry.Columns(2), order:=xlAscending, CustomOrder:="RED,RED +,YELLOW,GREEN"
                        Else
                            .SortFields.Add key:=rangeTry.Columns(2), order:=xlAscending, CustomOrder:="GREEN,YELLOW,RED,RED +"
                        End If
                        .SortFields.Add key:=rangeTry.Columns(4), order:=xlAscending
                        .SortFields.Add key:=rangeTry.Columns(5), order:=xlAscending
                        
                        .SetRange rangeTry
                        .Header = xlNo
                        .MatchCase = False
                        .Orientation = xlTopToBottom
                        ThisWorkbook.sheets(onglet).Activate
                        .Apply

            End With
'            .Range(.Cells(7, 73), .Cells(LRow, LCol)).Sort Key1:=.Range("BV7"), Key2:=.Range("BX7"), Key3:=.Range("BY7")
            
            End If
        End If
        
        End With
      
End Function
        
      

Function orderCol(onglet As String, Optional orderP As Integer)
    Dim lrow
    Dim lcol
    Dim lig
    Dim rangeTry
    
    With ThisWorkbook.sheets(onglet)
    
     lcol = getLastColumnDrivability(onglet)
        For lig = 14 To lcol
           If .Cells(.Rows.Count, lig).End(xlUp).row > lrow Then lrow = .Cells(.Rows.Count, lig).End(xlUp).row
        Next lig
        
        Set rangeTry = .Range(.Cells(7, 14), .Cells(lrow, lcol))
        If lrow > 6 Then
            If .Cells(.Rows.Count, 15).End(xlUp).row > 6 Or _
               .Cells(.Rows.Count, 19).End(xlUp).row > 6 Or _
               .Cells(.Rows.Count, 20).End(xlUp).row > 6 Then
                
               With .Sort
                        .SortFields.Clear
                        If orderP = 0 Then
                            .SortFields.Add key:=rangeTry.Columns(2), order:=xlAscending, CustomOrder:="RED,RED +,YELLOW,GREEN"
                        Else
                            .SortFields.Add key:=rangeTry.Columns(2), order:=xlAscending, CustomOrder:="GREEN,YELLOW,RED,RED +"
                        End If
                        .SortFields.Add key:=rangeTry.Columns(4), order:=xlAscending
                        .SortFields.Add key:=rangeTry.Columns(5), order:=xlAscending
                        
                        .SetRange rangeTry
                        .Header = xlNo
                        .MatchCase = False
                        .Orientation = xlTopToBottom
                        ThisWorkbook.sheets(onglet).Activate
                        .Apply
'                .Range(.Cells(7, 14), .Cells(LRow, LCol)).Sort Key1:=.Range("O7"), Key2:=.Range("Q7"), Key3:=.Range("R7")
            End With
            
            End If
            
        End If
        
    End With
       
End Function


