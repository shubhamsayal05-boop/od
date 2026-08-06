Attribute VB_Name = "CompteSumPriorityDyn"
Public PriorityListDyn As Object
Public colorListDyn As Object
Public PriorityListCountDyn As Object

Function initListDyn()
    If PriorityListDyn Is Nothing Then
        Set PriorityListDyn = CreateObject("Scripting.Dictionary")
        Set colorListDyn = CreateObject("Scripting.Dictionary")
        Set PriorityListCountDyn = CreateObject("Scripting.Dictionary")
    Else
        PriorityListDyn.RemoveAll
        colorListDyn.RemoveAll
        PriorityListCountDyn.RemoveAll
    End If
End Function

Function storeListDyn(keys As String, rRow As Long, StartR As Long)
        Dim subKey As String
        Dim valueConcat As String
        Dim tableCompte() As String
        Dim divisionMode As Variant
        Dim concatPriorite As Variant
        Dim concatIndice
        Dim indices As Variant
        Dim priorityRedPlus
        Dim priorityYELLOW
        Dim priorityKey
        Dim tableSub() As String
        Dim color As String
        Dim getTotAll
        'exemple index = {Resultat: 1, indice: 1,}
        'exempleKey = {sdv: 'coast', rowParam: 2, cellule: $A$2}
        'exempleValue = {comptePriorite: 1, compteIndice: 2}
       
        indices = IndiceAgrementByRow(CStr(replace(Split(keys, ";")(0), "sdv:", "")), rRow, 2)
        color = ThisWorkbook.sheets(CStr(replace(Split(keys, ";")(0), "sdv:", ""))).Cells(rRow, 74)
        If Len(keys) = 0 Then Exit Function
        If PriorityListDyn Is Nothing Or colorListDyn Is Nothing Then Exit Function
        
        If Not PriorityListDyn.Exists(keys) Then
            divisionMode = indices / 1
            valueConcat = "comptePriorite:" & 1 & " ; compteIndice:" & indices & ";compteDivision:" & divisionMode
            PriorityListDyn.Add key:=keys, Item:=valueConcat
        Else
            tableCompte = Split(PriorityListDyn(keys), ";")
            concatPriorite = 1 + toNum(Split(tableCompte(0), ":")(1))
           
            concatIndice = WorksheetFunction.Sum(toNum(Split(tableCompte(1), ":")(1)), toNum(indices))
            divisionMode = concatIndice / concatPriorite
            concatIndice = IIf(InStr(1, concatIndice, ".") <> 0, FormatNumber(concatIndice, Len(concatIndice) - InStr(1, concatIndice, ".")), concatIndice)
            
            valueConcat = "comptePriorite:" & concatPriorite _
                                & " ; compteIndice:" & concatIndice & ";compteDivision:" & divisionMode
                                
             PriorityListDyn.Remove keys
             PriorityListDyn.Add key:=keys, Item:=valueConcat
        End If
        
        
        If Not colorListDyn.Exists(keys) Then
             priorityKey = getPriorityFromCalculs(replace(Split(keys, ";")(0), "sdv:", ""), _
                               ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Range(replace(Split(keys, ";")(1), "resultat:", "")))
            
            valueConcat = "GREEN:0;YELLOW:0;RED:0;RED +:0;sdvPriority:" & priorityKey
            valueConcat = valueConcat & ";sdvTaux:" & getTauxFromCalculs(replace(Split(keys, ";")(0), "sdv:", ""))
            valueConcat = valueConcat & ";priority:" & ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Range(replace(Split(keys, ";")(1), "resultat:", ""))
            priorityRedPlus = ThisWorkbook.sheets("Calculs").Range("I1")
            priorityYELLOW = 1 / ThisWorkbook.sheets("Calculs").Range("I4")
            valueConcat = valueConcat & ";coefYELLOW:" & priorityYELLOW
            valueConcat = valueConcat & ";coefRedPlus:" & priorityRedPlus
            
            colorListDyn.Add key:=keys, Item:=valueConcat
       End If
       
       If colorListDyn.Exists(keys) Then
            priorityRedPlus = ThisWorkbook.sheets("Calculs").Range("I1")
            priorityYELLOW = 1 / ThisWorkbook.sheets("Calculs").Range("I4")
            tableCompte = Split(colorListDyn(keys), ";")
             For i = 0 To 3
                 If InStr(1, tableCompte(i), color & ":") <> 0 Then
                     concatPriorite = 1 + toNum(Split(tableCompte(i), ":")(1))
                     valueConcat = color & ":" & concatPriorite
                     If subKey = "" Then subKey = valueConcat Else subKey = subKey & ";" & valueConcat
                 Else
                     If subKey = "" Then subKey = tableCompte(i) Else subKey = subKey & ";" & tableCompte(i)
                 End If
             Next i
             tableSub = Split(subKey, ";")
            
             divisionMode = (toNum(Split(tableSub(1), ":")(1)) * priorityYELLOW) + (toNum(Split(tableSub(3), ":")(1)) * priorityRedPlus) + toNum(Split(tableSub(2), ":")(1))
             getTotAll = divisionMode

             subKey = subKey & ";" & tableCompte(4) & ";" & tableCompte(5) & ";" & tableCompte(6) & ";" & tableCompte(7) & ";" & tableCompte(8)
             
             subKey = subKey & ";" & IIf(InStr(1, divisionMode, ".") <> 0, "COEF:" & FormatNumber(divisionMode, Len(divisionMode) - InStr(1, divisionMode, ".")), "COEF:" & divisionMode)
             subKey = subKey & ";" & "SUMCOLOR:" & (toNum(Split(tableSub(0), ":")(1)) + toNum(Split(tableSub(1), ":")(1)) + toNum(Split(tableSub(2), ":")(1)) + toNum(Split(tableSub(3), ":")(1)))
             
             divisionMode = toNum(Split(tableCompte(4), ":")(1)) * (divisionMode / (toNum(Split(tableSub(0), ":")(1)) + toNum(Split(tableSub(1), ":")(1)) + toNum(Split(tableSub(2), ":")(1)) + toNum(Split(tableSub(3), ":")(1))))
             subKey = subKey & ";" & IIf(InStr(1, divisionMode, ".") <> 0, "TpbLigne:" & FormatNumber(divisionMode, Len(divisionMode) - InStr(1, divisionMode, ".")), "TpbLigne:" & divisionMode)
             subKey = subKey & ";TotLigne:" & getTotAll
             subKey = subKey & ";TotDec:" & getTotAll / (toNum(Split(tableSub(0), ":")(1)) + toNum(Split(tableSub(1), ":")(1)) + toNum(Split(tableSub(2), ":")(1)) + toNum(Split(tableSub(3), ":")(1)))
             colorListDyn.Remove keys
             colorListDyn.Add key:=keys, Item:=subKey
     End If
     
     '____
     subKey = ""
     If Not PriorityListCountDyn.Exists(keys) Then


            valueConcat = "YELLOW:0;RED:0;" _
                              & "P1:0;P2:0;P3:0;" _
                              & "P1RED:0;P2RED:0;P3RED:0;" _
                              & "P1YELLOW:0;P2YELLOW:0;P3YELLOW:0;" _
                              & "total:0;Start:0"

            PriorityListCountDyn.Add key:=keys, Item:=valueConcat
     End If

     If PriorityListCountDyn.Exists(keys) Then
            priorityKey = ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Range(replace(Split(keys, ";")(1), "resultat:", ""))
            tableCompte = Split(PriorityListCountDyn(keys), ";")
            If color = "RED +" Then color = "RED"
             subKey = ""
             For i = 0 To UBound(tableCompte) - 2

                 If i < 2 Then
                    If InStr(1, tableCompte(i), color & ":") <> 0 Then
                        concatPriorite = 1 + toNum(Split(tableCompte(i), ":")(1))
                        valueConcat = Split(tableCompte(i), ":")(0) & ":" & concatPriorite
                        If subKey = "" Then subKey = valueConcat Else subKey = subKey & ";" & valueConcat
                    Else
                        If subKey = "" Then subKey = tableCompte(i) Else subKey = subKey & ";" & tableCompte(i)
                    End If
                 ElseIf i > 1 And i < 5 Then
                    If InStr(1, replace(tableCompte(i), "P", ""), priorityKey & ":") <> 0 And Split(tableCompte(i), ":")(1) = "0" Then
                        concatPriorite = 1 + toNum(Split(tableCompte(i), ":")(1))
                        valueConcat = Split(tableCompte(i), ":")(0) & ":" & concatPriorite
                        If subKey = "" Then subKey = valueConcat Else subKey = subKey & ";" & valueConcat
                    Else
                        If subKey = "" Then subKey = tableCompte(i) Else subKey = subKey & ";" & tableCompte(i)
                    End If
                ElseIf i > 4 Then
                    If InStr(1, replace(tableCompte(i), "P", ""), priorityKey & color & ":") <> 0 And Split(tableCompte(i), ":")(1) = "0" Then
                        If color = "YELLOW" Then
                            If Split(tableCompte(i - 3), ":")(1) = "0" Then
                                concatPriorite = 1 + toNum(Split(tableCompte(i), ":")(1))
                                valueConcat = Split(tableCompte(i), ":")(0) & ":" & concatPriorite
                                If subKey = "" Then subKey = valueConcat Else subKey = subKey & ";" & valueConcat
                            Else
                                If subKey = "" Then subKey = tableCompte(i) Else subKey = subKey & ";" & tableCompte(i)
                            End If
                        ElseIf color = "RED" Then
                               If Split(tableCompte(i + 3), ":")(1) <> "0" Then
                                   tableCompte(i + 3) = replace(tableCompte(i + 3), tableCompte(i + 3), Split(tableCompte(i + 3), ":")(0) & ":0")
                                End If
                                concatPriorite = 1 + toNum(Split(tableCompte(i), ":")(1))
                                valueConcat = Split(tableCompte(i), ":")(0) & ":" & concatPriorite
                                If subKey = "" Then subKey = valueConcat Else subKey = subKey & ";" & valueConcat
                        End If

                    Else
                        If subKey = "" Then subKey = tableCompte(i) Else subKey = subKey & ";" & tableCompte(i)
                    End If

                 End If
             Next i
         
           
             concatPriorite = 1 + toNum(Split(tableCompte(UBound(tableCompte)), ":")(1))
             valueConcat = "total" & ":" & concatPriorite
             If subKey = "" Then subKey = valueConcat Else subKey = subKey & ";" & valueConcat
             subKey = subKey & ";Start:" & StartR
             PriorityListCountDyn.Remove keys
           
             PriorityListCountDyn.Add key:=keys, Item:=subKey
     End If

        
        
End Function

Function restoreListDyn(onglet As String)
    Dim k As Variant
    Dim countDiv
    Dim i As Integer
    Dim IsCount As Boolean
    
    i = 0
    IsCount = False
    If PriorityListDyn Is Nothing Then Exit Function
    If PriorityListDyn.Count = 0 Then Exit Function
    For Each k In PriorityListDyn.keys
           If InStr(1, k, "sdv:" & onglet & ";") <> 0 Then
                countDiv = countDiv + toNum(Split(replace(PriorityListDyn(k), "compteDivision:", ""), ";")(2))
                i = i + 1
                IsCount = True
           End If
    Next k
    If IsCount <> False Then
        countDiv = 1 + (countDiv / i)
        countDiv = Round(100 * (countDiv ^ ThisWorkbook.sheets("SETTINGS").Range("PUISS")), 1)
        ThisWorkbook.sheets(onglet).Range("BQ5") = countDiv
    End If
End Function
Function calculTPBDyn(onglet As String)
    Dim k As Variant
    Dim countDiv
    Dim i As Integer
    Dim IsCount As Boolean
    Dim countTotal(10) As Variant
    Dim calculTotal(14) As Variant
    Dim divByZero As Variant
    Dim plS As String
    Dim GpLs As Range
    
    i = 0
    IsCount = False
    
    
    If PriorityListCountDyn Is Nothing Then Exit Function
    If PriorityListCountDyn.Count = 0 Then Exit Function
    For Each k In PriorityListCountDyn.keys
           If InStr(1, k, "sdv:" & onglet & ";") <> 0 Then
              
                For i = 0 To UBound(Split(PriorityListCountDyn(k), ";")) - 2
                    
                   countTotal(i) = countTotal(i) + toNum(Split(Split(PriorityListCountDyn(k), ";")(i), ":")(1))
                Next i
               
                plS = replace(Split(PriorityListCountDyn(k), ";")(12), "Start:", "")
           End If
    Next k
    If plS = "" Then Exit Function
   Set GpLs = checkGetPlage(CLng(plS))

   ThisWorkbook.sheets(onglet).Range("BP8") = countTotal(2)
   ThisWorkbook.sheets(onglet).Range("BP9") = countTotal(3)
   ThisWorkbook.sheets(onglet).Range("BP10") = countTotal(4)
   
   ThisWorkbook.sheets(onglet).Range("BP14") = countTotal(8)
   ThisWorkbook.sheets(onglet).Range("BP15") = countTotal(9)
   ThisWorkbook.sheets(onglet).Range("BP16") = countTotal(10)
   
   ThisWorkbook.sheets(onglet).Range("BP17") = countTotal(5)
   ThisWorkbook.sheets(onglet).Range("BP18") = countTotal(6)
   ThisWorkbook.sheets(onglet).Range("BP19") = countTotal(7)
    
   divByZero = (countTotal(2) + countTotal(3) + countTotal(4))
   If divByZero <> 0 Then
        ThisWorkbook.sheets(onglet).Range("BQ8") = (countTotal(2) / (countTotal(2) + countTotal(3) + countTotal(4)) * 100)
        ThisWorkbook.sheets(onglet).Range("BQ9") = (countTotal(3) / (countTotal(2) + countTotal(3) + countTotal(4)) * 100)
        ThisWorkbook.sheets(onglet).Range("BQ10") = (countTotal(4) / (countTotal(2) + countTotal(3) + countTotal(4)) * 100)
   Else
        ThisWorkbook.sheets(onglet).Range("BQ8") = 0
        ThisWorkbook.sheets(onglet).Range("BQ9") = 0
        ThisWorkbook.sheets(onglet).Range("BQ10") = 0
        ThisWorkbook.sheets(onglet).Range("BQ14") = 0
        ThisWorkbook.sheets(onglet).Range("BQ15") = 0
        ThisWorkbook.sheets(onglet).Range("BQ16") = 0
        ThisWorkbook.sheets(onglet).Range("BQ17") = 0
        ThisWorkbook.sheets(onglet).Range("BQ18") = 0
        ThisWorkbook.sheets(onglet).Range("BQ19") = 0
   End If
   
   If countTotal(2) <> 0 Then
        ThisWorkbook.sheets(onglet).Range("BQ14") = (countTotal(8) / countTotal(2)) * 100
   Else
       ThisWorkbook.sheets(onglet).Range("BQ14") = 0
   End If
   
   If countTotal(3) <> 0 Then
        ThisWorkbook.sheets(onglet).Range("BQ15") = (countTotal(9) / countTotal(3)) * 100
   Else
       ThisWorkbook.sheets(onglet).Range("BQ15") = 0
   End If
   
   If countTotal(4) <> 0 Then
        ThisWorkbook.sheets(onglet).Range("BQ16") = (countTotal(10) / countTotal(4)) * 100
   Else
       ThisWorkbook.sheets(onglet).Range("BQ16") = 0
   End If
   
   
   If countTotal(2) <> 0 Then
        ThisWorkbook.sheets(onglet).Range("BQ17") = (countTotal(5) / countTotal(2)) * 100
   Else
       ThisWorkbook.sheets(onglet).Range("BQ17") = 0
   End If
   
   If countTotal(3) <> 0 Then
        ThisWorkbook.sheets(onglet).Range("BQ18") = (countTotal(6) / countTotal(3)) * 100
   Else
       ThisWorkbook.sheets(onglet).Range("BQ18") = 0
   End If
   
   If countTotal(4) <> 0 Then
        ThisWorkbook.sheets(onglet).Range("BQ19") = (countTotal(7) / countTotal(4)) * 100
   Else
       ThisWorkbook.sheets(onglet).Range("BQ19") = 0
   End If
   
    
'    ThisWorkbook.sheets(Onglet).Range("J17") = (countTotal(5) / (countTotal(2) + countTotal(3) + countTotal(4)) * 100)
'    ThisWorkbook.sheets(Onglet).Range("J18") = (countTotal(6) / (countTotal(2) + countTotal(3) + countTotal(4)) * 100)
'    ThisWorkbook.sheets(Onglet).Range("J19") = (countTotal(7) / (countTotal(2) + countTotal(3) + countTotal(4)) * 100)
'
  
   
  
   If Application.WorksheetFunction.CountIf(GpLs, 1) <> 0 Then
        ThisWorkbook.sheets(onglet).Range("BN20") = countTotal(2) / Application.WorksheetFunction.CountIf(GpLs, 1)
   Else
         ThisWorkbook.sheets(onglet).Range("BN20") = 0
   End If
   
    If Application.WorksheetFunction.CountIf(GpLs, 2) <> 0 Then
         ThisWorkbook.sheets(onglet).Range("BN21") = countTotal(3) / Application.WorksheetFunction.CountIf(GpLs, 2)
   Else
          ThisWorkbook.sheets(onglet).Range("BN21") = 0
   End If
  
   If Application.WorksheetFunction.CountIf(GpLs, 3) <> 0 Then
          ThisWorkbook.sheets(onglet).Range("BN22") = countTotal(4) / Application.WorksheetFunction.CountIf(GpLs, 3)
   Else
         ThisWorkbook.sheets(onglet).Range("BN22") = 0
   End If
   
   plS = GetCoverage(onglet)
   If Len(plS) > 0 Then
            ThisWorkbook.Worksheets(onglet).Range("BO20").formula = "=" & fixTheDecimalSeparator(ThisWorkbook.sheets(onglet).Range("BN20").Value) & " * 'POWERTRAIN'!" & Split(plS, ";")(0) & " * " & Application.WorksheetFunction.CountIf(GpLs, 1) & " + " & _
                fixTheDecimalSeparator(ThisWorkbook.sheets(onglet).Range("BN21").Value) & " * 'POWERTRAIN'!" & Split(plS, ";")(1) & " * " & Application.WorksheetFunction.CountIf(GpLs, 2) & " + " & _
                fixTheDecimalSeparator(ThisWorkbook.sheets(onglet).Range("BN22").Value) & " * 'POWERTRAIN'!" & Split(plS, ";")(2) & " * " & Application.WorksheetFunction.CountIf(GpLs, 3)
   End If
End Function

Sub calculTauxPointBasDyn()
    Dim k As Variant
    Dim tauxPts As Variant
    Dim sdvAct As String
    Dim cumulTotal As Variant
    Dim cumulPts As Variant
    Dim tableVal() As String
    Dim r As Range
    Dim foundT As Boolean
    Dim p1p2p3(2) As Variant
    Dim calculPercent As Variant
    Dim cumulTaux As Variant
    Dim taux, IR
    
    'A REDEFINIR APRES POINT RATING

     IR = ThisWorkbook.Worksheets("RATING").Rows("10:10").Find(What:="Tested vehicle", LookAt:=xlWhole).Column
    If colorListDyn Is Nothing Then
       ThisWorkbook.Worksheets("RATING").Cells(18, IR).Value = 0
        Exit Sub
    ElseIf colorListDyn.Count = 0 Then
        ThisWorkbook.Worksheets("RATING").Cells(18, IR).Value = 0
        Exit Sub
    End If
    
    sdvAct = ""
    tauxPts = 0
    foundT = False
    taux = 0

    For Each k In colorListDyn.keys
           tableVal = Split(colorListDyn(k), ";")
           If sdvAct = "" Then
                sdvAct = replace(Split(k, ";")(0), "sdv:", "")
                tauxPts = 0
                taux = toNum(Split(tableVal(5), ":")(1))
                p1p2p3(0) = GetP1P2P3(CStr(sdvAct), 1)
                p1p2p3(1) = GetP1P2P3(CStr(sdvAct), 2)
                p1p2p3(2) = GetP1P2P3(CStr(sdvAct), 3)
           End If
           
           If replace(Split(k, ";")(0), "sdv:", "") <> sdvAct Then
                sdvAct = replace(Split(k, ";")(0), "sdv:", "")
                cumulPts = cumulPts + (cumulTotal * taux)
                tauxPts = tauxPts + taux
                taux = toNum(Split(tableVal(5), ":")(1))
                cumulTotal = 0
                p1p2p3(0) = GetP1P2P3(CStr(sdvAct), 1)
                p1p2p3(1) = GetP1P2P3(CStr(sdvAct), 2)
                p1p2p3(2) = GetP1P2P3(CStr(sdvAct), 3)
           End If
            
            foundT = True
            If CStr(Split(tableVal(6), ":")(1)) = "1" Then
                calculPercent = (toNum(Split(tableVal(13), ":")(1)) * (toNum(Split(tableVal(4), ":")(1)))) / p1p2p3(0)
                cumulTotal = cumulTotal + (calculPercent)
            ElseIf CStr(Split(tableVal(6), ":")(1)) = "2" Then
                calculPercent = (toNum(Split(tableVal(13), ":")(1)) * (toNum(Split(tableVal(4), ":")(1)))) / p1p2p3(1)
                cumulTotal = cumulTotal + (calculPercent)
           ElseIf CStr(Split(tableVal(6), ":")(1)) = "3" Then
                calculPercent = (toNum(Split(tableVal(13), ":")(1)) * (toNum(Split(tableVal(4), ":")(1)))) / p1p2p3(2)
                cumulTotal = cumulTotal + (calculPercent)
           End If
    Next k

    If foundT = True Then
        cumulPts = cumulPts + (cumulTotal * taux)
        tauxPts = tauxPts + taux
    End If
    If tauxPts > 0 Then cumulPts = cumulPts / tauxPts
  
   ThisWorkbook.Worksheets("RATING").Cells(18, IR).Value = (cumulPts / 100)
      
    
End Sub
Sub restoreColorDyn()
    Dim k As Variant
    Dim countDiv
    Dim i As Integer
    Dim IsCount As Boolean

    i = 1

    If colorListDyn Is Nothing Then Exit Sub
    For Each k In colorListDyn.keys
           Range("A" & i) = k & ";" & colorListDyn(k)
           i = i + 1
    Next k

'calculTPB ("Converter release")
End Sub

Function clearListDyn()
    If Not PriorityListDyn Is Nothing Then
        PriorityListDyn.RemoveAll
        Set PriorityListDyn = Nothing
    End If
End Function
Function toNum(strVals As Variant)
    toNum = 0
   
         If InStr(1, strVals, ".") <> 0 Or InStr(1, strVals, ",") <> 0 Then
             toNum = CDbl(strVals)
         Else
             toNum = val(strVals)
         End If

End Function

Function getPriorityFromCalculs(SDV As String, priority As Variant)
    Dim colonne As String
    Dim r As Range
    Dim resultG
    
    colonne = ""
    getPriorityFromCalculs = 0
    If priority = 1 Then colonne = "D"
    If priority = 2 Then colonne = "E"
    If priority = 3 Then colonne = "F"
    
    If colonne = "" Then Exit Function
    With ThisWorkbook.sheets("Calculs")
        Set r = .Columns(2).Cells.Find(What:=SDV, LookAt:=xlWhole)
        If Not r Is Nothing Then resultG = .Range(colonne & r.row) * 100
        resultG = IIf(InStr(1, resultG, ".") <> 0, FormatNumber(resultG, Len(resultG) - InStr(1, resultG, ".")), resultG)
        getPriorityFromCalculs = resultG
    End With
    
    
End Function

Function getTauxFromCalculs(SDV As String)
    Dim r As Range
    Dim resultG
    getTauxFromCalculs = 0
  
    With ThisWorkbook.sheets("Calculs")
        Set r = .Columns(2).Cells.Find(What:=SDV, LookAt:=xlWhole)
        If Not r Is Nothing Then resultG = r.Offset(0, 1) * 100
        resultG = IIf(InStr(1, resultG, ".") <> 0, FormatNumber(resultG, Len(resultG) - InStr(1, resultG, ".")), resultG)
        getTauxFromCalculs = resultG
    End With
End Function

Function checkGetPlage(RStart As Long) As Range
    Dim i As Long
    Dim j As Long
    
    With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
        i = RStart
        While Len(.Cells(i, 10)) > 0
            i = i + 1
        Wend
        i = i - 1
        
        j = 10
        While Len(.Cells(RStart, j)) > 0
            j = j + 1
        Wend
        j = j - 1
        
       Set checkGetPlage = .Range(.Cells(RStart, 10), .Cells(i, j))
    End With
    
End Function



Function GetCoverage(onglet As String)
            Dim x As Long, D As Long
            
            With ThisWorkbook.Worksheets("POWERTRAIN")
               x = Formules.StartConFig
               If x = 0 Then Exit Function
               D = x
               While .Cells(D, 1) <> "SOMME"
                    If UCase(.Cells(D, 1)) = UCase(onglet) Then
                        GetCoverage = .Cells(D, 3).Address & ";" & .Cells(D, 4).Address & ";" & .Cells(D, 5).Address
                        Exit Function
                    End If
                    D = D + 1
              Wend
               
        End With
        
End Function

Function GetP1P2P3(onglet As String, p As Integer)
          
            
            With ThisWorkbook.Worksheets(onglet)
                If p = 1 Then GetP1P2P3 = .Range("BN20").Value
                If p = 2 Then GetP1P2P3 = .Range("BN21").Value
                If p = 3 Then GetP1P2P3 = .Range("BN22").Value
        End With
        
End Function



































