Attribute VB_Name = "InsertDB"
Option Explicit

Sub CreateColonne()
  Dim r As Range
  Dim colon As Object
  Dim RsRows As Object
  Dim c
  Dim i As Integer
  Dim Vals As String
  
  Set colon = CreateObject("Scripting.Dictionary")
  Set RsRows = CreateObject("ADODB.Recordset")
  RsRows.ActiveConnection = db.GetOdb
  RsRows.Properties("Jet OLEDB:Locking Granularity") = 1
  RsRows.Open "[data]", db.GetOdb, 1, 3, 2
  
  For Each c In RsRows.Fields
         If Not colon.Exists(UCase(c.Name)) Then
                colon.Add key:=UCase(c.Name), Item:=UCase(c.Name)
            End If
  Next c
  If Not RsRows Is Nothing Then RsRows.Close
  Set RsRows = Nothing
  

    Set r = ThisWorkbook.sheets("structure").Cells(2, 3)

     While Not Len(r.Value) = 0
         
         If Len(r.Offset(0, 2).Value) > 0 And InStr(1, replace(r.Offset(0, 2), " ", ""), ",") <> 0 Then
            Vals = replace(r.Offset(0, 2).Value, ".", "")
            For i = 1 To Len(r.Offset(0, 2).Value)
              Vals = replace(Vals, "  ", " ")
            Next i
            
            If Not colon.Exists(UCase(Vals)) Then
                db.Execute ("ALTER TABLE DATA ADD [" & Vals & "] varchar(255) Null")
                colon.Add key:=UCase(Vals), Item:=UCase(Vals)
            End If
         ElseIf Len(r.Offset(0, 1).Value) > 0 And r.Value = "criteria" And Len(r.Offset(0, 2).Value) = 0 Then
             Vals = replace(r.Offset(0, 1).Value, ".", "")
             For i = 1 To Len(r.Offset(0, 1).Value)
               Vals = replace(Vals, "  ", " ")
             Next i
             
             If Not colon.Exists(UCase(Vals)) Then
                db.Execute ("ALTER TABLE DATA ADD [" & Vals & "] varchar(255) Null")
                colon.Add key:=UCase(Vals), Item:=UCase(Vals)
            End If
         End If
        
         Set r = r.Offset(1, 0)
     Wend

End Sub

Function chargeVal(Optional v As String, Optional nomColEv As Variant)
  Dim i As Long
  Dim j As Long
  Dim startId As Long
  Dim table As Object
  Dim tableId As Object
  Dim columnId As String
  Dim RsRows As Object
  Dim c
  Dim req As Object
  Dim ViD As String
  Dim tabId As String, description As String, colDb As String
  Dim k  As Variant
  Dim tablSub() As String
  Dim saveEntete As Object
  Dim RqOdb As Object
  Dim idc As String
  Dim reqDatas As String
  
  
  ViD = db.GetValue("Select ID FROM Projet Where id=" & ThisWorkbook.Worksheets("HOME").Range("UNIQUEP").Value)
  idc = getDbId(ThisWorkbook.Worksheets("Home").Range("idProjects"))
  Set RqOdb = db.GetOdb(val(Right(idc, 1)))
  startId = 0
  Call DBStructure.getEntete
  
  If v = "" Then
           Set table = initTable
           Set tableId = CreateObject("Scripting.Dictionary")
           Set saveEntete = CreateObject("Scripting.Dictionary")
           
           With ThisWorkbook.sheets("DATA")
            
               For j = 3 To .Range("A65000").End(xlUp).row
                    Set RsRows = table("dataId")
                    RsRows.addnew
                    RsRows("code") = ThisWorkbook.Worksheets("HOME").Range("AT32").Value
                    RsRows("UNIQUENAME") = val(ViD)
                   
                    RsRows("Sous situation de vie, Sub Event Name") = .Range(nomColEv & j)
                    RsRows.Update
                    If startId = 0 Then
                        startId = val(db.GetValue("SELECT Max(dataId.N°) FROM dataId;", RqOdb))
                    Else
                        startId = startId + 1
                    End If
                    
                    
                    If j = 3 Then
                            For i = 1 To .Cells(2, .Columns.Count).End(xlToLeft).Column
                                  If .Cells(2, i).Value = "Start time of the Sub Event" Then .Cells(1, i).Value = "Temps debut d'analyse"
                                  description = UCase(replace(.Cells(1, i).Value & ", " & .Cells(2, i).Value, ".", ""))
                                   tabId = DBStructure.getTableByDescription(description)
                                   If table.Exists(tabId) Then
                                         Set RsRows = table(tabId)
                                         If Not tableId.Exists(tabId) Then
                                               tableId.Add key:=tabId, Item:=tabId
                                               RsRows.addnew
                                               RsRows("idData") = startId
                                         End If
                                         colDb = DBStructure.getColumnByDescription(description)
                                         RsRows(colDb) = .Cells(j, i).Value
                                         If Not saveEntete.Exists(colDb & "#" & tabId) Then saveEntete.Add key:=colDb & "#" & tabId, Item:=i
                                         If InStr(1, ", " & columnId & ", ", ", " & colDb & ", ") = 0 Then
                                            If columnId = "" Then columnId = colDb Else columnId = columnId & ", " & colDb
                                         End If
                                      
                                   End If
                            Next i
                    Else
                            For Each k In saveEntete.keys
                                Set RsRows = table(Split(k, "#")(1))
                                If Not tableId.Exists(Split(k, "#")(1)) Then
                                      tableId.Add key:=Split(k, "#")(1), Item:=Split(k, "#")(1)
                                      RsRows.addnew
                                      RsRows("idData") = startId
                                End If
                                RsRows(Split(k, "#")(0)) = .Cells(j, saveEntete(k)).Value
                            Next k
                    End If
                    
                    For Each k In tableId.keys
                            table(k).Update
                            tableId.Remove k
                    Next k
                    ProgressTitle ("Chargement des données dans la base" & vbCrLf & j & "/" & .Range("A65000").End(xlUp).row)
                
               Next j
               
           End With
           
           Set req = db.Request("Select ColonneDb fROM projet where id=" & ViD)
           If Not req Is Nothing Then
             If Len(req.Fields(0).Value) > 0 And Len(columnId) > 0 Then
                columnId = addUpdateColId(columnId, req.Fields(0).Value)
             End If
             req.Close
             Set req = Nothing
           End If
           
            
    
  
         
           If Len(columnId) > 0 Then
                db.Execute ("Update projet set ColonneDb=" & Chr(34) & columnId & Chr(34) & " where id=" & ViD)
                
                idc = getDbId(ViD)
                Set RqOdb = db.GetOdb(val(idc))
                db.Execute "Update projet set ColonneDb=" & Chr(34) & columnId & Chr(34) & " where id=" & ViD, RqOdb
           End If
           Set table = Nothing
           Set RsRows = Nothing
    End If
    
    'If v <> "" Then
    '    Set req = db.Request(selectDatas & " And [Sous situation de vie, Sub Event Name] In(" & v & ")")
    'Else
    '    Set req = db.Request(selectDatas)
    'End If
    reqDatas = selectDatas(v)
    If reqDatas = "" Then
        MsgBox "Aucune Données", vbCritical, "ODRIV"
        Exit Function
    End If
'    tablSub = Split(selectDatas(v), ";")
    ReDim tablSub(0)
    tablSub(0) = reqDatas
    
    With ThisWorkbook.sheets("DATA")
         .UsedRange.Clear
         startId = 0
         
         idc = getDbId(ThisWorkbook.Worksheets("Home").Range("idProjects"))
         Set RqOdb = db.GetOdb(val(idc))
       
         For j = 0 To UBound(tablSub)
             Set req = db.Request(tablSub(j), RqOdb)
            
             If Not req Is Nothing Then
                 
                 If .Cells(1, .Columns.Count).End(xlToLeft).Column = 1 Then
                     .Range("A2").CopyFromRecordset req
                 Else
                     .Cells(2, .Cells(1, .Columns.Count).End(xlToLeft).Column + 1).CopyFromRecordset req
                End If
                
                 For i = 0 To req.Fields.Count - 1
                     startId = startId + 1
                     .Cells(1, startId) = IIf(Left(req.Fields(i).Name, 4) = "col_", DBStructure.getDescriptionByCol(req.Fields(i).Name), req.Fields(i).Name)
                     If IsNumeric(.Cells(2, startId)) = True And Len(.Cells(2, startId)) > 0 Then
                         .Columns(startId).TextToColumns Destination:=.Cells(1, startId), DataType:=xlDelimited, _
                             TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                             Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
                             :=Array(1, 1), TrailingMinusNumbers:=True
                      End If
                 Next i
                 req.Close
             End If
        Next j
   End With
    
    'If Not req Is Nothing Then req.Close
    Set req = Nothing
    
End Function

Function addUpdateColId(id As String, updateId As String)
    Dim i As Integer
    Dim tabAllId() As String
    Dim idGet As String
    
    idGet = id
    tabAllId = Split(updateId, ", ")
    For i = 0 To UBound(tabAllId)
        If InStr(1, ", " & id & ", ", ", " & tabAllId(i) & ", ") = 0 Then
            idGet = idGet & ", " & tabAllId(i)
        End If
    Next i
    
    addUpdateColId = idGet
    
End Function
Function initTable() As Object
        Dim r As Range
        Dim table As Object
        Dim RsRows As Object
        Dim lastRow As Integer
        Dim idc As String
        Dim RqOdb As Object
        
        Set table = CreateObject("Scripting.Dictionary")
        
        idc = getDbId(ThisWorkbook.Worksheets("Home").Range("idProjects"))
        Set RqOdb = db.GetOdb(val(Right(idc, 1)))
        
        With ThisWorkbook.Worksheets("DBStructure")
            lastRow = .Cells(.Rows.Count, 2).End(xlUp).row
            For Each r In .Range("B2:B" & lastRow)
                    If Not table.Exists(r) Then
                        Set RsRows = CreateObject("ADODB.Recordset")
                        RsRows.ActiveConnection = RqOdb
                        RsRows.Properties("Jet OLEDB:Locking Granularity") = 1
                        RsRows.Open r.Value, RqOdb, 2, 3, 2
                        table.Add key:=r.Value, Item:=RsRows
                     End If
            Next r
            
            Set RsRows = CreateObject("ADODB.Recordset")
            RsRows.ActiveConnection = RqOdb
            RsRows.Properties("Jet OLEDB:Locking Granularity") = 1
            RsRows.Open "dataId", RqOdb, 2, 3, 2
            table.Add key:="dataId", Item:=RsRows
             
        End With
        Set initTable = table
        If Not RsRows Is Nothing Then Set RsRows = Nothing
End Function

Function CVAL()
    Dim Totalevents As Double
    Dim evt As Double, k As Double
    Dim c As Range
    Dim NEVENTS As Long
    Dim LeverMeca As Boolean, LeverElec As Boolean
    Dim nom As String
    
    Dim Name_Tend As String, Name_Tstart As String
   

    NEVENTS = ThisWorkbook.Worksheets("DATA").Range("A65000").End(xlUp).row
     ThisWorkbook.Worksheets("VIERGE").Cells.EntireColumn.Hidden = False
      With ThisWorkbook.Worksheets("structure")
             Set c = ThisWorkbook.sheets("structure").Range("B2")
             ThisWorkbook.Worksheets("VIERGE").Visible = True
               Do While Len(c.Offset(0, 1).Value) > 0
                 If Len(c.Value) > 0 Then
                        If Not ThisWorkbook.sheets("DATA").Columns(3).Cells.Find(What:=c.Value, LookAt:=xlWhole) Is Nothing Then
                           Call GenSdVSheets(c.Value)
                        End If
                End If
              Set c = c.Offset(1, 0)
            Loop
        ThisWorkbook.Worksheets("VIERGE").Visible = False
    End With
     
    Call FiltresOff
    With ThisWorkbook.Worksheets("DATA")
            Name_Tend = nomcol("Val_A_Tend", "Selector_Lever_Position")
            Name_Tstart = nomcol("Val_A_Tend", "Selector_Lever_Position")
            If NEVENTS <= 2 Then
                NEVENTS = 2
            End If
            Totalevents = .Range("A65000").End(xlUp).row
            LeverElec = False
            LeverMeca = False
            evt = NEVENTS
            While LeverMeca = False And LeverElec = False And evt <= Totalevents
                evt = evt + 1
                If (.Range(Name_Tend & evt).Value = 4 Or .Range(Name_Tstart & evt).Value = 4) And .Range(Name_Tend & evt).Value <> .Range(Name_Tstart & evt).Value Then
                    LeverMeca = True
                ElseIf (.Range(Name_Tend & evt).Value = 0 Or .Range(Name_Tstart & evt).Value = 0) And .Range(Name_Tend & evt).Value <> .Range(Name_Tstart & evt).Value Then
                    LeverElec = True
                End If
            Wend
            If LeverMeca = True And LeverElec = False Then
                For k = NEVENTS + 1 To Totalevents
                    .Range(Name_Tend & k) = .Range(Name_Tend & k) - 1
                    .Range(Name_Tstart & k) = .Range(Name_Tstart & k) - 1
                Next k
            End If
        End With
        ThisWorkbook.Worksheets("HOME").Range("Project") = ThisWorkbook.sheets("DATA").Range(nomcol("Vehicle Configuration", "Vehicle Configuration Name") & 3)
        ThisWorkbook.Worksheets("DATA").UsedRange.Offset(1, 0).Interior.color = RGB(255, 255, 255)
        ThisWorkbook.Worksheets("DATA").UsedRange.AutoFilter
        ThisWorkbook.Worksheets("DATA").Rows(1).AutoFilter
        nom = ThisWorkbook.sheets("HOME").Range("Fuel").Value & " PREMIUM " & ThisWorkbook.sheets("HOME").Range("Prestation").Value & " (" & ThisWorkbook.sheets("HOME").Range("DriveVersion").Value & ")"
        Call MajAll_WTP(ThisWorkbook.sheets("HOME").Range("Mode").Value, nom)
        ThisWorkbook.Worksheets("DATA").Visible = xlSheetHidden
  
End Function





Function selectDatas(Optional v As Variant)
          Dim req As Object, reqSub As Object
          Dim table() As String, talbleId As Variant
          Dim getTable As String
          Dim concatJoin As String
          Dim i As Integer
          Dim j As Long
          Dim getsubid As String
          Dim RqOdb As Object
          Dim idc As String
          
          selectDatas = ""
           idc = getDbId(ThisWorkbook.Worksheets("Home").Range("idProjects"))
          Set RqOdb = db.GetOdb(val(idc))

          Set req = db.Request("Select ColonneDb from Projet where id=" & ThisWorkbook.Worksheets("HOME").Range("UNIQUEP").Value)
          If Not req Is Nothing Then
                 If IsNull(req.Fields(0).Value) = True Then Exit Function
                 getTable = DBStructure.getTableByListCol(req.Fields(0).Value)
                 If InStr(1, getTable, ";") = 0 Then
                     ReDim table(0)
                     table(0) = getTable
                 Else
                    table = Split(getTable, ";")
                 End If
                 
               
                
                For i = 0 To UBound(table)
                        
                        If i = 0 Then
'''                            If v = "" Then
'''                                Set reqSub = db.Request("Select N° from dataId where uniquename=" & ThisWorkbook.Worksheets("HOME").Range("UNIQUEP").Value & " order by N°", RqOdb)
'''                            Else
'''                                 Set reqSub = db.Request("Select N° from dataId where uniquename=" & ThisWorkbook.Worksheets("HOME").Range("UNIQUEP").Value _
'''                                            & " And [Sous situation de vie, Sub Event Name] In(" & v & ") order by N°", RqOdb)
'''                            End If
'''
'''                            If Not reqSub Is Nothing Then
'''                                    talbleId = reqSub.getrows
''''                                    For j = 0 To UBound(talbleId, 2)
''''                                            If getsubid = "" Then getsubid = CStr(talbleId(0, j)) Else getsubid = getsubid & ", " & CStr(talbleId(0, j))
''''                                    Next j
'''                                   getsubid = "idData >=" & CStr(talbleId(0, j)) & " And idData <=" & CStr(talbleId(0, UBound(talbleId, 2)))
'''
'''                            End If
                         End If
                         
                      
                            concatJoin = concatJoin & ", " & getListColFromTable(req.Fields(0).Value, table(i))
                            talbleId = IIf(talbleId <> "", talbleId & "," & table(i), table(i))
                       
                        
                Next i
                
                
                 concatJoin = "Select dataId.N°, dataId.code, dataId.[Sous situation de vie, Sub Event Name] " & concatJoin & " FROM "
                 If UBound(Split(talbleId, ",")) = 2 Then concatJoin = concatJoin & "((dataId LEFT JOIN dataSub1 ON dataId.N° = dataSub1.idData) LEFT JOIN dataSub2 ON dataId.N° = dataSub2.idData) LEFT JOIN dataSub3 ON dataId.N° = dataSub3.idData"
                 If UBound(Split(talbleId, ",")) = 1 Then concatJoin = concatJoin & "(dataId LEFT JOIN dataSub1 ON dataId.N° = dataSub1.idData) LEFT JOIN dataSub2 ON dataId.N° = dataSub2.idData"
                 If UBound(Split(talbleId, ",")) = 0 Then concatJoin = concatJoin & "dataId LEFT JOIN dataSub1 ON dataId.N° = dataSub1.idData"

                 If v <> "" Then
                    concatJoin = concatJoin & " where " _
                    & "dataId.uniquename=" & ThisWorkbook.Worksheets("HOME").Range("UNIQUEP").Value & " And dataId.[Sous situation de vie, Sub Event Name] In(" & v & ")  ORDER BY dataId.N°"
                Else
                      concatJoin = concatJoin & " where " _
                    & "dataId.uniquename=" & ThisWorkbook.Worksheets("HOME").Range("UNIQUEP").Value & " ORDER BY dataId.N°"
                End If
              
              
                 selectDatas = concatJoin
              
                
                Set req = Nothing
          End If
End Function









































