Attribute VB_Name = "UpdateTAP"
Option Explicit
Private TotCount(4) As Integer
Private TabList(4) As String
Private NotFound As String
Private pwt_conf As String


Function checkChange() As Boolean
    Dim i As Integer
    Dim c As Range
   
    TabList(1) = getList("ENGINE")
    TabList(2) = getList("GEARBOX")
    TabList(3) = getList("NBGEAR")
    TabList(4) = getList("AREA")
    TotCount(1) = 0
    TotCount(2) = 0
    TotCount(3) = 0
    TotCount(4) = 0
    
    checkChange = False
    With ThisWorkbook.Worksheets("CONFIGURATIONS ARRAY")
         NotFound = ";"
         Set c = .Range("B2")
                        
          While Application.CountA(.Range("B" & c.row & ":G" & c.row)) > 0 Or .Range("C" & c.row).Interior.color = 855309
                If UCase(c.Value) = "ENGINE TYPE" Then
                    If checkConfigPresent(c.row + 1, TabList(1), c.Column, 1) = False Then checkChange = True
                End If
                
                If UCase(c.Offset(0, 3).Value) = "GEARBOX TYPE" Then
                    If checkConfigPresent(c.row + 1, TabList(2), c.Offset(0, 3).Column, 2) = False Then checkChange = True
                End If
                
                If UCase(c.Value) = "NUMBER OF GEARS" Then
                    If checkConfigPresent(c.row + 1, TabList(3), c.Column, 3) = False Then checkChange = True
                End If
                
                If UCase(c.Offset(0, 3).Value) = "AREA" Then
                    If checkConfigPresent(c.row + 1, TabList(4), c.Offset(0, 3).Column, 4) = False Then checkChange = True
                End If
               
                Set c = c.Offset(1, 0)
            Wend
            
'            For i = 3 To 9
'                 If getCorrespondance(.Range("B" & i)) <> 0 Then
'                     If getChange(i, TabList(getCorrespondance(.Range("B" & i))), getCorrespondance(.Range("B" & i))) = False Then checkChange = True
'                 End If
'            Next i

        If checkChange = True Then
            Call putValue
            Call putAllConfigValue
            Call putAllPowertrainValue
       End If

    End With

End Function
Function checkConfigPresent(i As Long, config As String, col As Integer, id As Integer) As Boolean
    Dim r As Range
    Dim mat As Integer
    Dim nbCount As Integer
    Dim strGet As String
    With ThisWorkbook.sheets("CONFIGURATIONS ARRAY")
        checkConfigPresent = True
        Set r = .Cells(i, col + 1)
        If id <= 2 Then mat = 0 Else mat = 12
        While r.Interior.color = 855309
            If Len(r.Value) > 0 Then nbCount = nbCount + 1
            If InStr(1, ";" & UCase(config) & ";", ";" & UCase(r.Value) & ";") = 0 Then
                     checkConfigPresent = False
                     NotFound = NotFound & id & "." & r.row - mat & ";"
            End If
            Set r = r.Offset(1, 0)
        Wend
        strGet = config
        strGet = Left(strGet, Len(strGet) - 1)
        strGet = Right(strGet, Len(strGet) - 1)
        If nbCount <> UBound(Split(strGet, ";")) + 1 Then
            checkConfigPresent = False
        End If
        
    End With
End Function
Function remplirListe(i As Long, Liste As String, col As Integer)
    Dim r As Range
    Dim TblV() As String
    
    With ThisWorkbook.sheets("CONFIGURATIONS ARRAY")
        TblV = Split(Liste, ";")
        Set r = .Cells(i, col + 1)
        
        While r.Interior.color = 855309
            r.ClearContents
            Set r = r.Offset(1, 0)
        Wend
        
        Set r = .Cells(i, col + 1)
        For i = 0 To UBound(TblV)
             r.Value = TblV(i)
             Set r = r.Offset(1, 0)
        Next i
        
    End With
        
End Function
Function remplirSetting(i As Long, Liste As String, col As Integer, id As Integer)
    Dim r As Range
    Dim TblV() As String
    Dim TblVal() As String
    Dim getNul As Boolean
    Dim colon As Object
    Dim j As Integer
    
    getNul = True
    Set colon = CreateObject("Scripting.Dictionary")
    If TotCount(id) = 0 Then getNul = False
    With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
        TblV = Split(Liste, ";")
        Set r = .Cells(i, col + 1)
        If checkSame(r, Liste) = True Then Exit Function
        
        If TotCount(id) = 0 Then
                While r.Interior.color = 855309
                        If Not colon.Exists(UCase(r.Value)) Then
                            colon.Add key:=UCase(r.Value), Item:=UCase(r.Offset(0, 1).Value)
                        End If
                        TotCount(id) = TotCount(id) + 1
                        .Range(r, r.Offset(0, 1)).ClearContents
                        Set r = r.Offset(1, 0)
                Wend
        Else
            For j = i To i + TotCount(id)
                 If Not colon.Exists(UCase(r.Value)) Then
                         colon.Add key:=UCase(r.Value), Item:=UCase(r.Offset(0, 1).Value)
                End If
                Set r = r.Offset(1, 0)
            Next j
            .Range(.Cells(i, col + 1), .Cells(j - 1, col + 2)).ClearContents
        End If
        
        Set r = .Cells(i, col + 1)
        ReDim TblVal(UBound(TblV), 1)
        For i = 0 To UBound(TblV)
             TblVal(i, 0) = TblV(i)
             TblVal(i, 1) = IIf(colon.Exists(TblV(i)), colon(TblV(i)), "")
        Next i
         r.Resize(UBound(TblVal, 1) + 1, 2).Value = TblVal
        
    End With
        
End Function
Function checkSame(r As Range, Liste As String) As Boolean
    Dim TblV() As String
    Dim i As Integer
    Dim c As Range
    
    Set c = r
    checkSame = True
    TblV = Split(Liste, ";")
    For i = 0 To UBound(TblV)
         If UCase(c.Value) <> UCase(TblV(i)) Then
                checkSame = False
                Exit Function
         End If
         Set c = c.Offset(1, 0)
    Next i
    If Len(c.Value) > 0 Then checkSame = False
    
        
    
    
End Function
Function putValue() As Boolean
        Dim j As Integer
        Dim i As Integer
        Dim lc As Integer
        Dim TblV() As String
        Dim StrS As String
        Dim l As Integer
        Dim c As Range
        
        With ThisWorkbook.Worksheets("CONFIGURATIONS ARRAY")
              Set c = .Range("B2")
                               
              While Application.CountA(.Range("B" & c.row & ":G" & c.row)) > 0 Or .Range("C" & c.row).Interior.color = 855309
                        If UCase(c.Value) = "ENGINE TYPE" Then
                                StrS = Left(TabList(1), Len(TabList(1)) - 1)
                                StrS = Right(StrS, Len(StrS) - 1)
                                Call remplirListe(c.row + 1, StrS, c.Column)
                        End If
                        
                        If UCase(c.Offset(0, 3).Value) = "GEARBOX TYPE" Then
                                StrS = Left(TabList(2), Len(TabList(2)) - 1)
                                StrS = Right(StrS, Len(StrS) - 1)
                                Call remplirListe(c.row + 1, StrS, c.Offset(0, 3).Column)
                        End If
                        
                        If UCase(c.Value) = "NUMBER OF GEARS" Then
                                StrS = Left(TabList(3), Len(TabList(3)) - 1)
                                StrS = Right(StrS, Len(StrS) - 1)
                                Call remplirListe(c.row + 1, StrS, c.Column)
                        End If
                        
                        If UCase(c.Offset(0, 3).Value) = "AREA" Then
                                StrS = Left(TabList(4), Len(TabList(4)) - 1)
                                StrS = Right(StrS, Len(StrS) - 1)
                                Call remplirListe(c.row + 1, StrS, c.Offset(0, 3).Column)
                        End If
               
                Set c = c.Offset(1, 0)
            Wend
            

                
'               For j = 3 To 9 Step 2
'                           l = l + 1
'                           StrS = Left(TabList(l), Len(TabList(l)) - 1)
'                           StrS = Right(StrS, Len(StrS) - 1)
'                          TblV = Split(StrS, ";")
'                         .Range(.Cells(j, 3), .Cells(j, .Cells(j, .Columns.Count).End(xlToLeft).Column)).ClearContents
'                          If UBound(TblV) + 3 > lc Then lc = UBound(TblV) + 3
'                          For i = 0 To UBound(TblV)
'                                .Cells(j, i + 3) = TblV(i)
'                          Next i
'
'              Next j
        End With
End Function

Function putAllPowertrainValue() As Boolean
        Dim TblV() As String
        Dim convTblV(3) As String
        Dim coLSup(3) As String
        Dim tblParc() As String
        Dim l As Integer
        Dim t As Long
        Dim tF As Long
        Dim df As Integer

        With ThisWorkbook.Worksheets("POWERTRAIN")
               convTblV(0) = Left(TabList(1), Len(TabList(1)) - 1)
               convTblV(0) = Right(convTblV(0), Len(convTblV(0)) - 1)
               TblV = Split(NotFound, ";")

               convTblV(1) = Left(TabList(2), Len(TabList(2)) - 1)
               convTblV(1) = Right(convTblV(1), Len(convTblV(1)) - 1)

               convTblV(2) = Left(TabList(3), Len(TabList(3)) - 1)
               convTblV(2) = Right(convTblV(2), Len(convTblV(2)) - 1)

               convTblV(3) = Left(TabList(4), Len(TabList(4)) - 1)
               convTblV(3) = Right(convTblV(3), Len(convTblV(3)) - 1)

               For t = 0 To UBound(TblV)
                    If Left(TblV(t), 2) = "1." Then
                            coLSup(0) = coLSup(0) & ";" & replace(TblV(t), "1.", "")
                    ElseIf Left(TblV(t), 2) = "2." Then
                            coLSup(1) = coLSup(1) & ";" & replace(TblV(t), "2.", "")
                    ElseIf Left(TblV(t), 2) = "3." Then
                            coLSup(2) = coLSup(2) & ";" & replace(TblV(t), "3.", "")
                    ElseIf Left(TblV(t), 2) = "4." Then
                            coLSup(3) = coLSup(3) & ";" & replace(TblV(t), "4.", "")
                    End If
               Next t
               
               t = 2
               tF = .Range("A65000").End(xlUp).row
               df = 1
               
               For t = t To tF

                        If .Cells(t, df) = "Engine type" Then
                            TblV = Split(convTblV(0), ";")
                            Call get_pwt_conf(t)
                            .Range(.Cells(t, 2).Address, .Cells(t + 1, 9).Address).ClearContents
                            .Cells(t, df + 1).Resize(1, UBound(TblV) + 1).Value = TblV
                            Call fill_pwt_conf(t)
'                            If Len(coLSup(0)) > 0 Then
'                                    tblParc = Split(coLSup(0), ";")
'                                    For l = 0 To UBound(tblParc)
'                                          tblParc(l) = val(tblParc(l)) - 1
'                                          If val(tblParc(l)) > 0 Then .Cells(t + 1, val(tblParc(l))) = ""
'                                    Next l
'                            End If

                            TblV = Split(convTblV(1), ";")
                            Call get_pwt_conf(t + 2)
                            .Range(.Cells(t + 2, 2).Address, .Cells(t + 3, 9).Address).ClearContents
                            .Cells(t + 2, df + 1).Resize(1, UBound(TblV) + 1).Value = TblV
                            Call fill_pwt_conf(t + 2)
'                            If Len(coLSup(1)) > 0 Then
'                                    tblParc = Split(coLSup(1), ";")
'                                    For l = 0 To UBound(tblParc)
'                                         tblParc(l) = val(tblParc(l)) - 1
'                                         If val(tblParc(l)) > 0 Then .Cells(t + 3, val(tblParc(l))) = ""
'                                    Next l
'                            End If

                            TblV = Split(convTblV(2), ";")
                            Call get_pwt_conf(t + 4)
                            .Range(.Cells(t + 4, 2).Address, .Cells(t + 5, 9).Address).ClearContents
                            .Cells(t + 4, df + 1).Resize(1, UBound(TblV) + 1).Value = TblV
                            Call fill_pwt_conf(t + 4)
'                            If Len(coLSup(2)) > 0 Then
'                                    tblParc = Split(coLSup(2), ";")
'                                    For l = 0 To UBound(tblParc)
'                                         tblParc(l) = val(tblParc(l)) - 1
'                                         If val(tblParc(l)) > 0 Then .Cells(t + 5, val(tblParc(l))) = ""
'                                    Next l
'                            End If

                            TblV = Split(convTblV(3), ";")
                            Call get_pwt_conf(t + 6)
                            .Range(.Cells(t + 6, 2).Address, .Cells(t + 7, 9).Address).ClearContents
                            .Cells(t + 6, df + 1).Resize(1, UBound(TblV) + 1).Value = TblV
                            Call fill_pwt_conf(t + 6)
'                            If Len(coLSup(3)) > 0 Then
'                                    tblParc = Split(coLSup(3), ";")
'                                    For l = 0 To UBound(tblParc)
'                                         tblParc(l) = val(tblParc(l)) - 1
'                                         If val(tblParc(l)) > 0 Then .Cells(t + 7, val(tblParc(l))) = ""
'                                    Next l
'                            End If

                             t = t + 7
                       End If
            Next t
           
        End With
End Function

Function get_pwt_conf(i)
    Dim j
    
    pwt_conf = vbNullString
    With ThisWorkbook.Worksheets("POWERTRAIN")
        For j = 2 To 9
            If pwt_conf = vbNullString Then pwt_conf = .Cells(i, j).Value & ":" & .Cells(i + 1, j).Value Else: pwt_conf = pwt_conf & ";" & .Cells(i, j).Value & ":" & .Cells(i + 1, j).Value
        Next
    End With
End Function

Function fill_pwt_conf(i)
    Dim el
    Dim j
    Dim data_conf

    data_conf = Split(pwt_conf, ";")
    With ThisWorkbook.Worksheets("POWERTRAIN")
        For j = 2 To 9
            For Each el In data_conf
                If .Cells(i, j).Value = Split(el, ":")(0) Then
                    .Cells(i + 1, j).Value = Split(el, ":")(1)
                    Exit For
                End If
            Next
        Next
    End With
End Function


Function putAllConfigValue() As Boolean
        Dim StrS As String
        Dim t As Long
        Dim tF As Long
        Dim derniereColonne As Integer, cm As Integer

        With ThisWorkbook.Worksheets("CONFIGURATIONS SEETINGS")
               .Outline.ShowLevels RowLevels:=2
               t = 4
               tF = 0
               derniereColonne = 30
               For cm = 1 To derniereColonne
                    If .Cells(.Rows.Count, cm).End(xlUp).row > tF Then tF = .Cells(.Rows.Count, cm).End(xlUp).row
               Next cm
             
               For t = t To tF

                        If .Cells(t, 2) = "Engine type" Then
                            StrS = Left(TabList(1), Len(TabList(1)) - 1)
                            StrS = Right(StrS, Len(StrS) - 1)
                             Call remplirSetting(t + 1, StrS, 2, 1)
                         End If
                         
                         If .Cells(t, 5) = "Gearbox type" Then
                            StrS = Left(TabList(2), Len(TabList(2)) - 1)
                            StrS = Right(StrS, Len(StrS) - 1)
                             Call remplirSetting(t + 1, StrS, 5, 2)
                             If TotCount(2) <> 0 Then t = t + TotCount(2)
                           End If
                           
                          If .Cells(t, 2) = "Number of gears" Then
                                StrS = Left(TabList(3), Len(TabList(3)) - 1)
                                StrS = Right(StrS, Len(StrS) - 1)
                                Call remplirSetting(t + 1, StrS, 2, 3)
                            End If
                            
                            If .Cells(t, 5) = "Area" Then
                                StrS = Left(TabList(4), Len(TabList(4)) - 1)
                                StrS = Right(StrS, Len(StrS) - 1)
                                Call remplirSetting(t + 1, StrS, 5, 4)
                                If TotCount(4) <> 0 Then t = t + TotCount(4)
                            End If
                           
            Next t
           .Outline.ShowLevels RowLevels:=1
        End With
End Function


Function getList(Rg As String)
     Dim c As Range
     getList = ";"
     With ThisWorkbook.Worksheets("CONFIGURATIONS")
         Set c = .Range(Rg)
         Set c = c.Offset(1, 0)
         While c.Value <> ""
             If Rg = "AREA" Then
                 getList = getList & c.Value & ";"
             ElseIf Rg = "ENGINE" Then
                getList = getList & c.Value & ";"
             ElseIf Rg = "GEARBOX" Then
                 getList = getList & c.Value & ";"
             ElseIf Rg = "NBGEAR" Then
                getList = getList & c.Value & ";"
             End If
             Set c = c.Offset(1, 0)
         Wend
    End With
End Function












