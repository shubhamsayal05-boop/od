Attribute VB_Name = "Rating_Priorisation"
Option Explicit
Private Str As String
Private k As Long
Private NP1R As Integer
Private NP1Y As Integer
Private NP1G As Integer
Private NP2R As Integer
Private NP2Y As Integer
Private NP2G As Integer
Private NP3R As Integer
Private NP3Y As Integer
Private NP3G As Integer
Private celResu(3) As Range
Sub Priorisations(ByVal onglet As String)
    Dim NEVENTS As Integer
    Dim LigneDebut As Long, Types As Long
    Dim colStart As Long, rowStart As Long
    Dim OperationLigne As String, OperationColonneDroite As String, OperationColonneGauche As String
    Dim parametreLigne As String, parametreColonneGauche As String, parametreColonneDroite As String
    Dim TablT(1) As String
    Dim pLigne As String
    Dim j As Long
    Dim CrLeft As Long, CrRight As Long, Cc As Long
    Dim GetColGauche As Boolean, GetColDroite As Boolean
    Dim isValeur As Boolean, TVal As Boolean
    Dim color(3)
    Dim keyPriorite As String
    
    On Error GoTo PqS
    Str = "Event Priority"
    With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
        NEVENTS = ThisWorkbook.sheets(onglet).Range("G8").Value
        If Len(NEVENTS) = 0 Then
            Exit Sub
        End If

        NP1R = 0
        NP1Y = 0
        NP1G = 0
        NP2R = 0
        NP2Y = 0
        NP2G = 0
        NP3R = 0
        NP3Y = 0
        NP3G = 0
        Set celResu(1) = Nothing
        Set celResu(2) = Nothing
        Set celResu(3) = Nothing
        
        LigneDebut = LigneSeeting(onglet)
        
        If LigneDebut = 0 Then Exit Sub
        Types = StartConFig(LigneDebut)
        If Types = 0 Then Exit Sub
      
        
        colStart = 10
        rowStart = Types - 23
        OperationLigne = ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Cells(Types + 3, 2).Value
        OperationColonneDroite = ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Cells(Types + 3, 4).Value
        OperationColonneGauche = ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Cells(Types + 3, 3).Value
        parametreLigne = ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Cells(Types + 1, 2).Value
        parametreColonneGauche = ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Cells(Types + 1, 3)
        parametreColonneDroite = ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Cells(Types + 1, 4)
       
        isValeur = False
'        If ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Cells(Types + 3, 6) = "X" Then isValeur = False Else isValeur = True
        If (colStart = 0 Or rowStart = 0) And isValeur = False Then Exit Sub
       
      For k = 7 To NEVENTS + 6
                    GetColGauche = False
                    GetColDroite = False
                    CrLeft = 0
                    CrRight = 0
                    Cc = colStart
                    TVal = True
                    
                  
                     keyPriorite = "sdv:" & onglet
                    
                    If isValeur = True Then
'                        Call putResutat(Onglet, val(ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Cells(Types + 3, 6)))
                    End If
                   'Ligne____________________________________
                   If checkLeftRight(parametreLigne, OperationLigne) = True And isValeur = False And checkIfColumnExists(parametreLigne, onglet) = True Then
                       pLigne = (ThisWorkbook.sheets(onglet).Cells(k, ThisWorkbook.sheets(onglet).Rows(6).Cells.Find(What:=parametreLigne, LookAt:=xlWhole).Column))
                      If Len(ThisWorkbook.sheets(onglet).Cells(k, ThisWorkbook.sheets(onglet).Rows(6).Cells.Find(What:=parametreLigne, LookAt:=xlWhole).Column)) <> 0 Then
                          Cc = allRowOperations(OperationLigne, colStart, rowStart, pLigne)
                     Else
                         Cc = 0
                     End If
                  ElseIf checkIfColumnExists(parametreLigne, onglet) = False Then
                        Cc = 0
                  End If
                  
                   'Deux Colonnes Presents____________________________________
                  If checkIfColumnExists(parametreColonneGauche, onglet) = True And checkLeftRight(parametreColonneGauche, OperationColonneGauche) = True And checkLeftRight(parametreColonneDroite, OperationColonneDroite) = True And isValeur = False Then
                     GetColGauche = True
                     GetColDroite = True
                     TablT(0) = (ThisWorkbook.sheets(onglet).Cells(k, ThisWorkbook.sheets(onglet).Rows(6).Cells.Find(What:=parametreColonneGauche, LookAt:=xlWhole).Column))
                     TablT(1) = (ThisWorkbook.sheets(onglet).Cells(k, ThisWorkbook.sheets(onglet).Rows(6).Cells.Find(What:=parametreColonneDroite, LookAt:=xlWhole).Column))
                     If Len(ThisWorkbook.sheets(onglet).Cells(k, ThisWorkbook.sheets(onglet).Rows(6).Cells.Find(What:=parametreColonneGauche, LookAt:=xlWhole).Column)) <> 0 And _
                        Len(ThisWorkbook.sheets(onglet).Cells(k, ThisWorkbook.sheets(onglet).Rows(6).Cells.Find(What:=parametreColonneDroite, LookAt:=xlWhole).Column)) <> 0 Then
                        If OperationColonneDroite = OperationColonneGauche Then
                            CrLeft = allColOperations(OperationColonneGauche, colStart, rowStart, pLigne, 0, TablT)
                            CrRight = CrLeft
                        Else
                            j = rowStart
                            Do Until Len(.Cells(j, colStart)) = 0 Or CrLeft <> 0
                                   If allColOperations(OperationColonneDroite, colStart, j, TablT(1), 1, , "OK") = allColOperations(OperationColonneGauche, colStart, j, TablT(0), 2, , "OK") Then
                                    CrLeft = allColOperations(OperationColonneDroite, colStart, j, TablT(1), 1, , "OK")
                                    CrRight = CrLeft
                                   End If
                                   j = j + 1
                            Loop
                           
                        End If
                     Else
                        CrLeft = 0
                        CrRight = CrLeft
                     End If
                   End If
                   
                  'Colonne Droite Presente____________________________________
                  If checkIfColumnExists(parametreColonneDroite, onglet) = True And checkLeftRight(parametreColonneDroite, OperationColonneDroite) = True And checkLeftRight(parametreColonneGauche, OperationColonneGauche) = False And isValeur = False Then
                     GetColDroite = True
                     pLigne = (ThisWorkbook.sheets(onglet).Cells(k, ThisWorkbook.sheets(onglet).Rows(6).Cells.Find(What:=parametreColonneDroite, LookAt:=xlWhole).Column))
                     If Len(ThisWorkbook.sheets(onglet).Cells(k, ThisWorkbook.sheets(onglet).Rows(6).Cells.Find(What:=parametreColonneDroite, LookAt:=xlWhole).Column)) <> 0 Then
                       CrRight = allColOperations(OperationColonneDroite, colStart, rowStart, pLigne, 1)
                     Else
                        CrRight = 0
                     End If
                  End If
                  
                   'Colonne Gauche Presente____________________________________
                  If checkIfColumnExists(parametreColonneGauche, onglet) = True And checkLeftRight(parametreColonneDroite, OperationColonneDroite) = False And checkLeftRight(parametreColonneGauche, OperationColonneGauche) = True And isValeur = False Then
                     GetColGauche = True
                     pLigne = (ThisWorkbook.sheets(onglet).Cells(k, ThisWorkbook.sheets(onglet).Rows(6).Cells.Find(What:=parametreColonneGauche, LookAt:=xlWhole).Column))
                     If Len(ThisWorkbook.sheets(onglet).Cells(k, ThisWorkbook.sheets(onglet).Rows(6).Cells.Find(What:=parametreColonneGauche, LookAt:=xlWhole).Column)) <> 0 Then
                       CrLeft = allColOperations(OperationColonneGauche, colStart, rowStart, pLigne, 2)
                     Else
                        CrLeft = 0
                     End If
                  End If
            
                    If isValeur = False Then
                         If GetColGauche = True And GetColDroite = True Then
                            If (checkLeftRight(parametreLigne, OperationLigne) = True And CrLeft = CrRight And CrRight <> 0 And Cc <> 0) Then
                                    Call putResutat(onglet, .Cells(CrLeft, Cc))
                                    keyPriorite = keyPriorite & ";" & " resultat:" & .Cells(CrLeft, Cc).Address
                                    Call CompteSumPriority.storeList(keyPriorite, k, rowStart)
                            ElseIf (checkLeftRight(parametreLigne, OperationLigne) = False And CrLeft = CrRight And CrRight <> 0) Then
                                    Call putResutat(onglet, .Cells(CrRight, colStart))
                                    keyPriorite = keyPriorite & ";" & " resultat:" & .Cells(CrRight, colStart).Address
                                    Call CompteSumPriority.storeList(keyPriorite, k, rowStart)
                            Else
                                    Call putResutat(onglet, 5)
                            End If
                            
                        ElseIf GetColDroite = True And GetColGauche = False Then
                            If (checkLeftRight(parametreLigne, OperationLigne) = True And Cc <> 0 And CrRight <> 0) Then
                                    Call putResutat(onglet, .Cells(CrRight, Cc))
                                     keyPriorite = keyPriorite & ";" & " resultat:" & .Cells(CrRight, Cc).Address
                                    Call CompteSumPriority.storeList(keyPriorite, k, rowStart)
                            ElseIf (checkLeftRight(parametreLigne, OperationLigne) = False And CrRight <> 0) Then
                                    Call putResutat(onglet, .Cells(CrRight, colStart))
                                    keyPriorite = keyPriorite & ";" & " resultat:" & .Cells(CrRight, colStart).Address
                                    Call CompteSumPriority.storeList(keyPriorite, k, rowStart)
                            Else
                                    Call putResutat(onglet, 5)
                            End If
                            
                        ElseIf GetColDroite = False And GetColGauche = True Then
                            If (checkLeftRight(parametreLigne, OperationLigne) = True And Cc <> 0 And CrLeft <> 0) Then
                                    Call putResutat(onglet, .Cells(CrLeft, Cc))
                                    keyPriorite = keyPriorite & ";" & " resultat:" & .Cells(CrLeft, Cc).Address
                                    Call CompteSumPriority.storeList(keyPriorite, k, rowStart)
                            ElseIf (checkLeftRight(parametreLigne, OperationLigne) = False And CrLeft <> 0) Then
                                    Call putResutat(onglet, .Cells(CrLeft, colStart))
                                    keyPriorite = keyPriorite & ";" & " resultat:" & .Cells(CrLeft, colStart).Address
                                    Call CompteSumPriority.storeList(keyPriorite, k, rowStart)
                            Else
                                    Call putResutat(onglet, 5)
                            End If
                        
                        ElseIf checkLeftRight(parametreColonneDroite, OperationColonneDroite) = False And checkLeftRight(parametreColonneGauche, OperationColonneGauche) = False Then
                            If Cc <> 0 Then
                                Call putResutat(onglet, .Cells(rowStart, Cc))
                                 keyPriorite = keyPriorite & ";" & " resultat:" & .Cells(rowStart, Cc).Address
                                Call CompteSumPriority.storeList(keyPriorite, k, rowStart)
                            Else
                                Call putResutat(onglet, 5)
                            End If
                        End If
                        
                     End If
                     
      Next k
    End With

    With ThisWorkbook.sheets(onglet)
'        .Range("I11") = NP1G
'        .Range("I14") = NP1Y
'        .Range("I17") = NP1R
'        .Range("I12") = NP2G
'        .Range("I15") = NP2Y
'        .Range("I18") = NP2R
'        .Range("I13") = NP3G
'        .Range("I16") = NP3Y
'        .Range("I19") = NP3R
    End With
    
    color(1) = RGB(192, 0, 0)
    color(2) = RGB(118, 147, 60)
    color(3) = RGB(31, 73, 125)
    For k = 1 To 3
            If Not celResu(k) Is Nothing Then celResu(k).Interior.color = color(k)
    Next k

PqS:
    If ERR.Number <> 0 Then
            MsgBox ERR.description, vbCritical, "ODRIV"
            ERR.Clear
            
    End If
End Sub
Function dcol(r As Range) As Integer
  dcol = r.Column
  Set r = r.Offset(0, 1)
  While r.Value <> "" And Len(r.Value) > 0
      dcol = dcol + 1
      Set r = r.Offset(0, 1)
  Wend

End Function

Function LigneSeeting(o As String) As Long
    
     With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
            If Not (.Columns("A:A").Find(What:=o, LookIn:=xlFormulas, LookAt _
                :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                False, SearchFormat:=False)) Is Nothing Then
                
                LigneSeeting = .Columns("A:A").Find(What:=o, LookIn:=xlFormulas, LookAt _
                :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                False, SearchFormat:=False).row + 1
                
           Else
                LigneSeeting = 0
           End If
    End With
End Function

Function StartConFig(RS As Long) As Long
    Dim lastr As Long
    Dim r As Range, c As Range
    Dim Engine As String, Gearbox As String, NbGear As String, Area As String
    Dim OK(4) As Boolean
    Dim i As Integer
    Dim p As Range
    Dim derniereColonne As Integer
    Dim cm As Integer
    
    
    StartConFig = 0
    With ThisWorkbook.sheets("HOME")
            Engine = .Range("Fuel")
            If InStr(1, .Range("Gears"), " ") <> 0 And .Range("Gears") <> "MANUAL GEARBOX" Then
                Gearbox = Left(.Range("Gears"), (InStr(1, .Range("Gears"), " ") - 1))
            Else
                Gearbox = .Range("Gears")
            End If
            NbGear = .Range("H23")
            Area = .Range("Area")
    End With
  
    OK(1) = False
    OK(2) = False
    OK(3) = False
    OK(4) = False
    With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
           
'            StartConFig = .Range("C" & Rs)
            Set r = .Range("A" & RS)
            lastr = 0
            .Outline.ShowLevels RowLevels:=2
            derniereColonne = 30
            For cm = 1 To derniereColonne
               If .Cells(.Rows.Count, cm).End(xlUp).row > lastr Then lastr = .Cells(.Rows.Count, cm).End(xlUp).row
            Next cm
            .Outline.ShowLevels RowLevels:=1
'            LastR = .Range("B65000").End(xlUp).Row
          
            While Len(r.Value) = 0 And r.row <= lastr 'And (Application.CountA(r.Row) > 0 Or Application.CountA(r.Row + 2) > 0)
                  
                If UCase(Left(.Range("B" & r.row), 6)) = "CONFIG" Then
                        Set c = .Range("B" & r.row + 1)
                        
                        While Application.CountA(.Range("B" & c.row & ":G" & c.row)) > 0 Or .Range("C" & c.row).Interior.color = 855309
                            If UCase(c.Value) = "ENGINE TYPE" Then OK(1) = checkConfigEnabled(c.row + 1, Engine, c.Column)
                            If UCase(c.Offset(0, 3).Value) = "GEARBOX TYPE" Then OK(2) = checkConfigEnabled(c.row + 1, Gearbox, c.Offset(0, 3).Column)
                            If UCase(c.Value) = "NUMBER OF GEARS" Then OK(3) = checkConfigEnabled(c.row + 1, NbGear, c.Column)
                            If UCase(c.Offset(0, 3).Value) = "AREA" Then OK(4) = checkConfigEnabled(c.row + 1, Area, c.Offset(0, 3).Column)
                            If c.Value = "Référence Ligne" Then
                                Set p = c
                            End If
                            Set c = c.Offset(1, 0)
                        Wend
                        If OK(1) = True And OK(2) = True And OK(3) = True And OK(4) = True Then
                            StartConFig = p.row
                            Exit Function
                        Else
                            OK(1) = False
                            OK(2) = False
                            OK(3) = False
                            OK(4) = False
                        End If
                       
                End If
                
                Set r = r.Offset(1, 0)
            Wend
    End With
End Function
Function checkConfigEnabled(i As Long, config As String, col As Integer) As Boolean
    Dim r As Range
    
    With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
        checkConfigEnabled = False
        Set r = .Cells(i, col + 1)
        
        While r.Interior.color = 855309
            If UCase(config) = UCase(r.Value) Then
                  If UCase(r.Offset(0, 1)) = "X" Then
                        checkConfigEnabled = True
                        Exit Function
                  End If
            End If
            Set r = r.Offset(1, 0)
        Wend
        
    End With
End Function
Function valGet(RS As Long) As String
    With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
            valGet = .Range("B" & RS + 10)
    End With
End Function
Function IsThrotle(RS As Long) As Boolean
    IsThrotle = False
    With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
            If .Range("B" & RS + 13) = "15-25-35-50-100" Then IsThrotle = True
    End With
End Function
Function opLigne(t As String, RS As Long) As String
     Dim StS As String
     opLigne = "X"
     
    StS = "INTERVALLE;SUPERIORITE;SUPERIORITE OU EGALITE;INFERIORITE;INFERIORITE OU EGALITE;EGALITE;ON-OFF;X"

     With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
            If t = "5-6-7-8-9 GEARBOX" Then
                If InStr(1, ";" & StS & ";", ";" & .Range("E" & (RS + 7)) & ";") <> 0 Then opLigne = .Range("E" & (RS + 7))
           ElseIf t = "MANUAL GEARBOX" Then
                If InStr(1, ";" & StS & ";", ";" & .Range("E" & (RS + 19)) & ";") <> 0 Then opLigne = .Range("E" & (RS + 19))
           ElseIf t = "ATDCT" Then
                If InStr(1, ";" & StS & ";", ";" & .Range("E" & (RS + 26)) & ";") <> 0 Then opLigne = .Range("E" & (RS + 26))
           ElseIf t = "ALL GEARBOX" Then
                If InStr(1, ";" & StS & ";", ";" & .Range("E" & (RS + 33)) & ";") <> 0 Then opLigne = .Range("E" & (RS + 33))
           ElseIf t = "ON-OFF" Then
                If InStr(1, ";" & StS & ";", ";" & .Range("E" & (RS + 40)) & ";") <> 0 Then opLigne = .Range("E" & (RS + 40))
           End If
    End With
End Function
Function opColonne(t As String, RS As Long) As String
     Dim StS As String
     opColonne = "X"
     
    StS = "EGALITE;INTERVALLE;X"

     With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
            If t = "5-6-7-8-9 GEARBOX" Then
                If InStr(1, ";" & StS & ";", ";" & .Range("F" & (RS + 7)) & ";") <> 0 Then opColonne = .Range("F" & (RS + 7))
           ElseIf t = "MANUAL GEARBOX" Then
                If InStr(1, ";" & StS & ";", ";" & .Range("F" & (RS + 19)) & ";") <> 0 Then opColonne = .Range("F" & (RS + 19))
           ElseIf t = "ATDCT" Then
                If InStr(1, ";" & StS & ";", ";" & .Range("F" & (RS + 26)) & ";") <> 0 Then opColonne = .Range("F" & (RS + 26))
           ElseIf t = "ALL GEARBOX" Then
                If InStr(1, ";" & StS & ";", ";" & .Range("F" & (RS + 33)) & ";") <> 0 Then opColonne = .Range("F" & (RS + 33))
           ElseIf t = "ON-OFF" Then
                 If InStr(1, ";" & StS & ";", ";" & .Range("F" & (RS + 40)) & ";") <> 0 Then opColonne = .Range("F" & (RS + 40))
           End If
    End With
End Function
Function paraLigne(t As String, RS As Long) As String
   
     paraLigne = "X"
     
    

     With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
            If t = "5-6-7-8-9 GEARBOX" Then
                If Len(.Range("B" & (RS + 7))) > 2 Then paraLigne = .Range("B" & (RS + 7))
           ElseIf t = "MANUAL GEARBOX" Then
                If Len(.Range("B" & (RS + 19))) > 2 Then paraLigne = .Range("B" & (RS + 19))
           ElseIf t = "ATDCT" Then
                If Len(.Range("B" & (RS + 26))) > 2 Then paraLigne = .Range("B" & (RS + 26))
           ElseIf t = "ALL GEARBOX" Then
                If Len(.Range("B" & (RS + 33))) > 2 Then paraLigne = .Range("B" & (RS + 33))
           ElseIf t = "ON-OFF" Then
                If Len(.Range("B" & (RS + 40))) > 2 Then paraLigne = .Range("B" & (RS + 40))
           End If
    End With
End Function
Function paraColonne(t As String, RS As Long) As String
     
     paraColonne = "X"
     
     With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
            If t = "5-6-7-8-9 GEARBOX" Then
                If Len(.Range("C" & (RS + 7))) > 2 Then paraColonne = .Range("C" & (RS + 7))
                If Len(.Range("D" & (RS + 7))) > 2 Then
                    If paraColonne = "X" Then paraColonne = .Range("D" & (RS + 7)) Else paraColonne = paraColonne & ";;" & .Range("D" & (RS + 7))
                End If
           ElseIf t = "MANUAL GEARBOX" Then
                 If Len(.Range("C" & (RS + 19))) > 2 Then paraColonne = .Range("C" & (RS + 19))
                If Len(.Range("D" & (RS + 19))) > 2 Then
                    If paraColonne = "X" Then paraColonne = .Range("D" & (RS + 19)) Else paraColonne = paraColonne & ";;" & .Range("D" & (RS + 19))
                End If
           ElseIf t = "ATDCT" Then
                 If Len(.Range("C" & (RS + 26))) > 2 Then paraColonne = .Range("C" & (RS + 26))
                If Len(.Range("D" & (RS + 26))) > 2 Then
                    If paraColonne = "X" Then paraColonne = .Range("D" & (RS + 26)) Else paraColonne = paraColonne & ";;" & .Range("D" & (RS + 26))
                End If
           ElseIf t = "ALL GEARBOX" Then
                If Len(.Range("C" & (RS + 33))) > 2 Then paraColonne = .Range("C" & (RS + 33))
                If Len(.Range("D" & (RS + 33))) > 2 Then
                    If paraColonne = "X" Then paraColonne = .Range("D" & (RS + 33)) Else paraColonne = paraColonne & ";;" & .Range("D" & (RS + 33))
                End If
            ElseIf t = "ON-OFF" Then
                If Len(.Range("C" & (RS + 40))) > 2 Then paraColonne = .Range("C" & (RS + 40))
                If Len(.Range("D" & (RS + 40))) > 2 Then
                    If paraColonne = "X" Then paraColonne = .Range("D" & (RS + 40)) Else paraColonne = paraColonne & ";;" & .Range("D" & (RS + 40))
                End If
           End If
    End With
End Function

Function Nshif()
    Dim counter As Integer
    Dim gea As Integer
    For counter = 1 To Len(ThisWorkbook.sheets("HOME").Range("Gears"))
        If IsNumeric(Mid(ThisWorkbook.sheets("HOME").Range("Gears"), counter, 1)) Then
            gea = Mid(ThisWorkbook.sheets("HOME").Range("Gears"), counter, 1)
        End If
    Next
    If Not isEmpty(gea) Then
            Nshif = gea
    Else
        If ThisWorkbook.sheets("HOME").Range("Gears") = "PHEV" Then
            Nshif = 8
        Else
            Nshif = ThisWorkbook.sheets("HOME").Range("Gears")
        End If
    End If
End Function

Function CStart(t As String, RS As Long, Optional Fuel As String) As Long
     CStart = 0
     With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
            If t = "5-6-7-8-9 GEARBOX" Then
                If Nshif = 5 Then
                    If IsNumeric(.Range("B" & (RS + 3))) Then CStart = .Range("B" & (RS + 3))
                ElseIf Nshif = 6 Then
                     If IsNumeric(.Range("C" & (RS + 3))) Then CStart = .Range("C" & (RS + 3))
                ElseIf Nshif = 7 Then
                     If IsNumeric(.Range("D" & (RS + 3))) Then CStart = .Range("D" & (RS + 3))
                ElseIf Nshif = 8 Then
                     If IsNumeric(.Range("E" & (RS + 3))) Then CStart = .Range("E" & (RS + 3))
                ElseIf Nshif = 9 Then
                     If IsNumeric(.Range("F" & (RS + 3))) Then CStart = .Range("F" & (RS + 3))
                End If
           ElseIf t = "MANUAL GEARBOX" Then
                If UCase(Fuel) = "GASOLINE" Then
                    If IsNumeric(.Range("B" & (RS + 15))) Then CStart = .Range("B" & (RS + 15))
                ElseIf UCase(Fuel) = "DIESEL" Then
                    If IsNumeric(.Range("C" & (RS + 15))) Then CStart = .Range("C" & (RS + 15))
                End If
           ElseIf t = "ATDCT" Then
                If IsNumeric(.Range("B" & (RS + 22))) Then CStart = .Range("B" & (RS + 22))
           ElseIf t = "ALL GEARBOX" Then
                If IsNumeric(.Range("B" & (RS + 29))) Then CStart = .Range("B" & (RS + 29))
            ElseIf t = "ON-OFF" Then
                If IsNumeric(.Range("B" & (RS + 36))) Then CStart = .Range("B" & (RS + 36))
           End If
    End With
End Function

Function RStart(t As String, RS As Long, Optional Zones As String) As Long
     RStart = 0
     With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
            If t = "5-6-7-8-9 GEARBOX" Then
                If Nshif = 5 Then
                    If IsNumeric(.Range("B" & (RS + 5))) Then RStart = .Range("B" & (RS + 5))
                ElseIf Nshif = 6 Then
                     If IsNumeric(.Range("C" & (RS + 5))) Then RStart = .Range("C" & (RS + 5))
                ElseIf Nshif = 7 Then
                     If IsNumeric(.Range("D" & (RS + 5))) Then RStart = .Range("D" & (RS + 5))
                ElseIf Nshif = 8 Then
                     If IsNumeric(.Range("E" & (RS + 5))) Then RStart = .Range("E" & (RS + 5))
                ElseIf Nshif = 9 Then
                     If IsNumeric(.Range("F" & (RS + 5))) Then RStart = .Range("F" & (RS + 5))
                End If
           ElseIf t = "MANUAL GEARBOX" Then
                If UCase(Zones) = "EUROPE" Then
                    If IsNumeric(.Range("B" & (RS + 17))) Then RStart = .Range("B" & (RS + 17))
                ElseIf UCase(Zones) = "CHINA" Then
                    If IsNumeric(.Range("C" & (RS + 17))) Then RStart = .Range("C" & (RS + 17))
                End If
           ElseIf t = "ATDCT" Then
                If IsNumeric(.Range("B" & (RS + 24))) Then RStart = .Range("B" & (RS + 24))
           ElseIf t = "ALL GEARBOX" Then
                If IsNumeric(.Range("B" & (RS + 31))) Then RStart = .Range("B" & (RS + 31))
           ElseIf t = "ON-OFF" Then
                 If IsNumeric(.Range("B" & (RS + 38))) Then RStart = .Range("B" & (RS + 38))
           End If
    End With
End Function

Function DetStart(RS As Long) As Long
         With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
                If .Range("B" & (RS + 17)) < .Range("C" & (RS + 17)) Then DetStart = .Range("B" & (RS + 17))
                If .Range("B" & (RS + 17)) > .Range("C" & (RS + 17)) Then DetStart = .Range("C" & (RS + 17))
    End With
End Function

Function allColOperations(Types As String, CStarts As Long, RStarts As Long, ValToCompar As String, GD As Integer, Optional t As Variant, Optional StopBoucle As String)
             allColOperations = 0
             
             On Error GoTo Ers
             
             With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
                 If Types = "EGALITE" Then
                          If GD = 0 Then
                                 If StopBoucle <> "OK" Then
                                      If CStr(.Cells(RStarts, CStarts - 2)) = CStr(t(0)) And CStr(.Cells(RStarts, CStarts - 1)) = CStr(t(1)) Then
                                         allColOperations = RStarts
                                     Else
                                         allColOperations = RStarts
                                         Do Until (CStr(.Cells(allColOperations, CStarts - 2)) = CStr(t(0)) And CStr(.Cells(allColOperations, CStarts - 1)) = CStr(t(1))) Or Len(.Cells(allColOperations, CStarts)) = 0
                                             allColOperations = allColOperations + 1
                                         Loop
                                         If Len(.Cells(allColOperations, CStarts)) = 0 Then allColOperations = 0
                                    End If
                                Else
                                    If CStr(.Cells(RStarts, CStarts - 2)) = CStr(t(0)) And CStr(.Cells(RStarts, CStarts - 1)) = CStr(t(1)) Then _
                                    allColOperations = RStarts Else allColOperations = 0
                                    
                                End If
                          Else
                                If CStr(.Cells(RStarts, CStarts - GD)) = CStr(ValToCompar) Then
                                     allColOperations = RStarts
                                Else
                                     allColOperations = RStarts
                                     Do Until CStr(.Cells(allColOperations, CStarts - GD)) = CStr(ValToCompar) Or Len(.Cells(allColOperations, CStarts - GD)) = 0
                                         allColOperations = allColOperations + 1
                                     Loop
                                     If Len(.Cells(allColOperations, CStarts - GD)) = 0 Then allColOperations = 0
                                End If
                         End If
                   
                 ElseIf Types = "INTERVALLE" Then
                         If GD = 0 Then
                                If StopBoucle <> "OK" Then
                                
                                        If checkIntervallok(.Cells(RStarts, CStarts - 2), t(0)) = True And checkIntervallok(.Cells(RStarts, CStarts - 1), t(1)) = True Then
                                             allColOperations = RStarts
                                            
                                        Else
                                           allColOperations = RStarts
                                            Do Until (checkIntervallok(.Cells(allColOperations, CStarts - 2), t(0)) = True And checkIntervallok(.Cells(allColOperations, CStarts - 1), t(1)) = True) Or Len(.Cells(allColOperations, CStarts)) = 0
                                                allColOperations = allColOperations + 1
                                            Loop
                                             If Len(.Cells(allColOperations, CStarts)) = 0 Then allColOperations = 0
                                       
                                       End If
                               Else
                                      If checkIntervallok(.Cells(RStarts, CStarts - 2), t(0)) = True And checkIntervallok(.Cells(RStarts, CStarts - 1), t(1)) = True Then _
                                        allColOperations = RStarts Else allColOperations = 0
                                
                               End If
                        Else
                        
                              If checkIntervallok(.Cells(RStarts, CStarts - GD), ValToCompar) = True Then
                                         allColOperations = RStarts
                              Else
                                        allColOperations = RStarts
                                        Do Until (checkIntervallok(.Cells(allColOperations, CStarts - GD), ValToCompar) = True) Or Len(.Cells(allColOperations, CStarts)) = 0
                                            allColOperations = allColOperations + 1
                                        Loop
                                        If Len(.Cells(allColOperations, CStarts)) = 0 Then allColOperations = 0
                               End If
                                
                           
                       End If
                                
                ElseIf Types = "INFERIORITE OU EGALITE" Then
                      If GD = 0 Then
                                If checkIOEok(.Cells(RStarts, CStarts - 2), t(0)) = True And checkIOEok(.Cells(RStarts, CStarts - 1), t(1)) = True Then
                                     allColOperations = RStarts
                                Else
                                    allColOperations = RStarts
                                    Do Until (checkIOEok(.Cells(allColOperations, CStarts - 2), t(0)) = True And _
                                               checkIOEok(.Cells(allColOperations, CStarts - 1), t(1))) _
                                                 Or Len(.Cells(allColOperations, CStarts)) = 0
                                        allColOperations = allColOperations + 1
                                    Loop
                                     If Len(.Cells(allColOperations, CStarts)) = 0 Then allColOperations = 0
                                End If
                               
                        Else
                            If checkIOEok(.Cells(RStarts, CStarts - GD), ValToCompar) Then
                                     allColOperations = RStarts
                            Else
                                    allColOperations = RStarts
                                    Do Until (checkIOEok(.Cells(allColOperations, CStarts - GD), ValToCompar)) Or Len(.Cells(allColOperations, CStarts)) = 0
                                        allColOperations = allColOperations + 1
                                    Loop
                                    If Len(.Cells(allColOperations, CStarts)) = 0 Then allColOperations = 0
                            End If
                       End If
  
                End If

          End With
          
Ers:
          If ERR.Number <> 0 Then
                allColOperations = 0
          End If
End Function
Function allRowOperations(Types As String, CStarts As Long, RStarts As Long, ValToCompar As String) As Long
        Dim j As Long
        allRowOperations = 0
        With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
             If Types = "SUPERIORITE OU EGALITE" Then
                For j = CStarts To dcol(.Cells(RStarts, CStarts))
                    If Abs(val(ValToCompar)) >= Abs(val(.Cells(RStarts - 2, j))) Then
                        allRowOperations = j
                    End If
               Next j
               
               
             ElseIf Types = "EGALITE" Then
                For j = CStarts To dcol(.Cells(RStarts, CStarts))
                    If ValToCompar = .Cells(RStarts - 2, j) Then
                        allRowOperations = j
                    End If
                Next j
                
                
             ElseIf Types = "SUPERIORITE" Then
                For j = CStarts To dcol(.Cells(RStarts, CStarts))
                    If Abs(val(ValToCompar)) > Abs(val(.Cells(RStarts - 2, j))) Then
                        allRowOperations = j
                    End If
                Next j
                
                
             ElseIf Types = "INFERIORITE" Then
                For j = CStarts To dcol(.Cells(RStarts, CStarts))
                    If Abs(val(ValToCompar)) < Abs(val(.Cells(RStarts - 2, j))) Then
                        allRowOperations = j
                    End If
                Next j
                
                
             ElseIf Types = "INFERIORITE OU EGALITE" Then
                For j = CStarts To dcol(.Cells(RStarts, CStarts))
                    If Abs(val(ValToCompar)) <= Abs(val(.Cells(RStarts - 2, j))) Then
                        allRowOperations = j
                    End If
                Next j
                
                
             ElseIf Types = "INTERVALLE" Then
                For j = CStarts To dcol(.Cells(RStarts, CStarts))
                    If Abs(val(ValToCompar)) >= Abs(val(.Cells(RStarts - 2, j))) And Abs(val(ValToCompar)) < Abs(val(.Cells(RStarts - 2, j + 1))) Then
                        allRowOperations = j
                    End If
                Next j
               
               
              ElseIf Types = "ON-OFF" Then
'                        If Abs(val(ValToCompar)) = 1 Then
'                            Call putResutat(Onglet, .Cells(RStarts, CStarts))
'                        ElseIf Abs(val(ValToCompar)) = 0 Then
'                            Call putResutat(Onglet, .Cells(RStarts, CStarts + 1))
'                        End If
                            
             End If
          End With
End Function
Function checkLeftRight(param As String, Operation As String) As Boolean
      If param <> "X" And Operation <> "X" And param <> "" And Operation <> "" Then checkLeftRight = True Else checkLeftRight = False
End Function
Function checkIfColumnExists(recherches As String, onglet As String) As Boolean
        checkIfColumnExists = True
        With ThisWorkbook.sheets(onglet)
                If .Rows(6).Cells.Find(What:=recherches, LookAt:=xlWhole) Is Nothing Then
                    checkIfColumnExists = False
'                    MsgBox "Attention " & recherches & " non retrouvé dans " & Onglet, vbCritical, "ODRIV"
                 End If
        End With
End Function
Function putResutat(onglet As String, Resu As Integer)
                Dim r As Range
                If Resu = 1 Or Resu = 2 Or Resu = 3 Then
                    Set r = ThisWorkbook.sheets(onglet).Cells(k, ThisWorkbook.sheets(onglet).Range("A6:BA6").Cells.Find(What:=Str, LookAt:=xlWhole).Column)
'                    MsgBox celResu(Resu).Address & "---" & r.Address
                     If celResu(Resu) Is Nothing Then Set celResu(Resu) = r Else Set celResu(Resu) = Union(celResu(Resu), r)
                     
                                               
                    ThisWorkbook.sheets(onglet).Cells(k, ThisWorkbook.sheets(onglet).Range("A6:BA6").Cells.Find(What:=Str, LookAt:=xlWhole).Column) = Resu
'                    ThisWorkbook.Sheets(onglet).Cells(k, ThisWorkbook.Sheets(onglet).Rows(6).Cells.Find(What:=Str, LookAt:=xlWhole).Column).Interior.Color = Color(Resu)
                Else
                    ThisWorkbook.sheets(onglet).Cells(k, ThisWorkbook.sheets(onglet).Range("A6:BA6").Cells.Find(What:=Str, LookAt:=xlWhole).Column) = "/"
                End If

                If ThisWorkbook.sheets(onglet).Cells(k, ThisWorkbook.sheets(onglet).Range("A6:BA6").Cells.Find(What:=Str, LookAt:=xlWhole).Column) = 1 Then
                    If ThisWorkbook.sheets(onglet).Cells(k, ThisWorkbook.sheets(onglet).Range("A6:BA6").Cells.Find(What:="Event Rating", LookAt:=xlWhole).Column) = "RED" Then
                        NP1R = NP1R + 1
                    ElseIf ThisWorkbook.sheets(onglet).Cells(k, ThisWorkbook.sheets(onglet).Range("A6:BA6").Cells.Find(What:="Event Rating", LookAt:=xlWhole).Column) = "YELLOW" Then
                        NP1Y = NP1Y + 1
                    ElseIf ThisWorkbook.sheets(onglet).Cells(k, ThisWorkbook.sheets(onglet).Range("A6:BA6").Cells.Find(What:="Event Rating", LookAt:=xlWhole).Column) = "GREEN" Then
                        NP1G = NP1G + 1
                    End If

                ElseIf ThisWorkbook.sheets(onglet).Cells(k, ThisWorkbook.sheets(onglet).Range("A6:BA6").Cells.Find(What:=Str, LookAt:=xlWhole).Column) = 2 Then
                    If ThisWorkbook.sheets(onglet).Cells(k, ThisWorkbook.sheets(onglet).Range("A6:BA6").Cells.Find(What:="Event Rating", LookAt:=xlWhole).Column) = "RED" Then
                        NP2R = NP2R + 1
                    ElseIf ThisWorkbook.sheets(onglet).Cells(k, ThisWorkbook.sheets(onglet).Range("A6:BA6").Cells.Find(What:="Event Rating", LookAt:=xlWhole).Column) = "YELLOW" Then
                        NP2Y = NP2Y + 1
                    ElseIf ThisWorkbook.sheets(onglet).Cells(k, ThisWorkbook.sheets(onglet).Range("A6:BA6").Cells.Find(What:="Event Rating", LookAt:=xlWhole).Column) = "GREEN" Then
                        NP2G = NP2G + 1
                    End If

                ElseIf ThisWorkbook.sheets(onglet).Cells(k, ThisWorkbook.sheets(onglet).Range("A6:BA6").Cells.Find(What:=Str, LookAt:=xlWhole).Column) = 3 Then
                    If ThisWorkbook.sheets(onglet).Cells(k, ThisWorkbook.sheets(onglet).Range("A6:BA6").Cells.Find(What:="Event Rating", LookAt:=xlWhole).Column) = "RED" Then
                        NP3R = NP3R + 1
                    ElseIf ThisWorkbook.sheets(onglet).Cells(k, ThisWorkbook.sheets(onglet).Range("A6:BA6").Cells.Find(What:="Event Rating", LookAt:=xlWhole).Column) = "YELLOW" Then
                        NP3Y = NP3Y + 1
                    ElseIf ThisWorkbook.sheets(onglet).Cells(k, ThisWorkbook.sheets(onglet).Range("A6:BA6").Cells.Find(What:="Event Rating", LookAt:=xlWhole).Column) = "GREEN" Then
                        NP3G = NP3G + 1
                    End If

                End If

End Function


Function checkIntervallU(v As String) As Boolean
    If InStr(1, v, "-") <> 0 Then checkIntervallU = True Else checkIntervallU = False
End Function

Function checkIntervallok(v As String, intVs As Variant) As Boolean
   Dim tabV() As String
   tabV = Split(v, "-")
   
   checkIntervallok = False
   If checkIntervallU(v) = False Then Exit Function
   If Abs(val(intVs)) >= Abs(val(tabV(0))) And Abs(val(intVs)) < Abs(val(tabV(1))) Then
        checkIntervallok = True
   End If
End Function


Function checkIOEok(v As String, intVs As Variant) As Boolean
   checkIOEok = False
   If checkIntervallU(v) = True Then Exit Function
   If Abs(val(intVs)) <= Abs(val(v)) Then
        checkIOEok = True
   End If
End Function





                  











