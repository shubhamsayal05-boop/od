Attribute VB_Name = "Rating_IndiceAgrement"
Option Explicit
Sub IndiceAgrement(ByVal onglet As String, Prt As String)
    If Prt = "driv" Then Call restoreList(onglet)
    If Prt = "dyn" Then Call restoreListDyn(onglet)
End Sub

Sub IndiceAgrementDefault(ByVal onglet As String)
    Dim NEVENTS As Variant
    Dim Ncrit As Variant
    Dim indices()
    Dim col As Integer, lig As Integer
    Dim IndiceSDV As Double
    Dim formule
    Dim Coeff1, Coeff2, Coeff3
    Dim SommeCriticities As Double
    Dim SommeC1s, SommeC2s
    Dim p
    Dim IndiceC1, IndiceC2
    Dim Indice
    Dim priority
    Dim k
    Dim waterline
    Dim Target, Criticity, Note, ZF
    Dim WaterlineT, targetT, noteT
    Dim Indiceoccurence, puissance
    Dim col_Criticity As Integer, col_Indice As Integer
    Dim prio As String
    Dim j As Integer, tabCol(3) As Integer
    
    NEVENTS = TotEventSheet(onglet)
    Ncrit = SDV2Ncrit(onglet)
    IndiceSDV = 0
    prio = ThisWorkbook.sheets("HOME").Range("Prestation").Value
    
    formule = ThisWorkbook.sheets("SETTINGS").Range("FORMS")
    Coeff1 = ThisWorkbook.sheets("SETTINGS").Range("COEF1")
    Coeff2 = ThisWorkbook.sheets("SETTINGS").Range("COEF2")
    Coeff3 = ThisWorkbook.sheets("SETTINGS").Range("COEF3")
    
    tabCol(0) = 14
    tabCol(1) = 15
    tabCol(2) = 73
    tabCol(3) = 74
    
    With ThisWorkbook.sheets(onglet)
            For j = 1 To 2
                If j = 1 Then
                    col_Indice = .Range("A6:BA6").Cells.Find(What:="Indice occurrencé", LookAt:=xlWhole).Column
                    col_Criticity = .Range("A5:BA5").Cells.Find(What:="Criticity", LookAt:=xlWhole).Column
                Else
                   col_Indice = .Range("BH6:GG6").Cells.Find(What:="Indice occurrencé", LookAt:=xlWhole).Column
                   col_Criticity = .Range("BH5:GG5").Cells.Find(What:="Criticity", LookAt:=xlWhole).Column
                End If
                col = col_Criticity + 1
                lig = 7
                
                For lig = lig To NEVENTS
                        ReDim indices(Ncrit)
                        Indice = 0
                        IndiceC1 = 0
                        IndiceC2 = 0
                        SommeCriticities = 0
                        SommeC1s = 0
                        SommeC2s = 0
                            
                       For col = col To Ncrit + col_Criticity
                            If j = 1 Then p = .Cells(lig, 14) Else p = .Cells(lig, 73)
                            If p <> 1 And p <> 2 And p <> 3 Then
                                p = 3
                            End If
                            If .Cells(lig, tabCol(1)) = "RED" Then
                                priority = ThisWorkbook.sheets("SETTINGS").Cells(p + 6, ThisWorkbook.sheets("SETTINGS").Range("RR").Column)
                            ElseIf .Cells(lig, tabCol(1)) = "YELLOW" Then
                                priority = ThisWorkbook.sheets("SETTINGS").Cells(p + 6, ThisWorkbook.sheets("SETTINGS").Range("OO").Column)
                            ElseIf .Cells(lig, tabCol(0)) = p And .Cells(lig, tabCol(1)) = "GREEN" Then
                               priority = ThisWorkbook.sheets("SETTINGS").Cells(p + 6, ThisWorkbook.sheets("SETTINGS").Range("GG").Column)
                            End If
                            
                             waterline = CSng(.Cells(3, col))
                             Target = CSng(.Cells(4, col))
                             Criticity = (3 - CSng(.Cells(5, col).Value)) / 2       'coeffs 1 et 0.5 selon les criticités àld 2 et 1
                      
                            If Criticity = 0 Then
                                indices(k) = 0
                            Else
                                    If IsNumeric(.Cells(lig, col).Value) Then
                                        Note = CSng(.Cells(lig, col).Value)
                                    Else
                                        Note = .Cells(lig, col)
                                    End If
                                    ZF = Coeff1 * waterline + Coeff2 * Target + Coeff3
                                    If Target <> 0 And ZF <> 0 And Target <> ZF Then
                                         WaterlineT = 10 * (waterline - ZF) / (Target - ZF)
                                    End If
                                    targetT = 10
                                    If .Cells(lig, col).Value <> "" And IsNumeric(Note) = True And Target <> 0 And ZF <> 0 And Target <> ZF Then
                                        noteT = 10 * (Note - ZF) / (Target - ZF)
                                    Else
                                        noteT = ""
                                    End If
                                    If Note < ZF And .Cells(lig, col).Value <> "" And IsNumeric(Note) Then
                                        indices(k) = -1 * Criticity
                                    ElseIf noteT >= 0 And noteT < WaterlineT Then
                                        indices(k) = Criticity * (2 * noteT - targetT - WaterlineT) / (targetT + WaterlineT)
                                    ElseIf noteT >= WaterlineT And noteT < targetT Then
                                        indices(k) = Criticity * (noteT - targetT) / (targetT + WaterlineT)
                                    Else
                                        indices(k) = 0
                                    End If
                           End If
                           
                            If Criticity = 1 Then
                                IndiceC1 = IndiceC1 + indices(k)
                            ElseIf Criticity = 0.5 Then
                                IndiceC2 = IndiceC2 + indices(k)
                            End If
                            
                            If waterline > 0 And Target > 0 And Note > 0 And Note <= 10 Then
                                If Criticity = 1 Then
                                    SommeC1s = SommeC1s + 1
                                ElseIf Criticity = 0.5 Then
                                    SommeC2s = SommeC2s + 1
                                End If
                            End If
                            
                       Next col
                       
                        If SommeC1s <> 0 And SommeC2s <> 0 Then
                            Indice = IndiceC1 + (IndiceC2 / SommeC2s)
                        ElseIf SommeC1s = 0 And SommeC2s <> 0 Then
                            Indice = IndiceC2 / SommeC2s
                        ElseIf SommeC1s <> 0 And SommeC2s = 0 Then
                            Indice = IndiceC1
                        Else
                            Indiceoccurence = 0
                        End If
                        If Indice < -1 Then
                            Indice = -1
                        End If
                        Indiceoccurence = Indice * priority / 100
                        ThisWorkbook.sheets(onglet).Cells(lig, col_Indice) = Indiceoccurence
                        IndiceSDV = IndiceSDV + Indiceoccurence
                        col = col_Criticity + 1
                Next lig
                 If lig > 7 Then
                    IndiceSDV = 1 + IndiceSDV / (lig - 6)
                    puissance = ThisWorkbook.sheets("SETTINGS").Range("PUISS")
                    If j = 1 Then
                        ThisWorkbook.sheets(onglet).Range("J5") = Round(100 * (IndiceSDV ^ puissance), 1)
                    Else
                        ThisWorkbook.sheets(onglet).Range("BQ5") = Round(100 * (IndiceSDV ^ puissance), 1)
                    End If
                End If
                Call HideC3(onglet)
            Next j
     End With
  
End Sub



Function IndiceAgrementByRow(ByVal onglet As String, ByVal rRow As Long, j As Integer)
    Dim NEVENTS As Variant
    Dim Ncrit As Variant
    Dim indices()
    Dim col As Integer, lig As Integer
    Dim IndiceSDV As Double
    Dim formule
    Dim Coeff1, Coeff2, Coeff3
    Dim SommeCriticities As Double
    Dim SommeC1s, SommeC2s
    Dim p
    Dim IndiceC1, IndiceC2
    Dim Indice
    Dim priority
    Dim k
    Dim waterline
    Dim Target, Criticity, Note, ZF
    Dim WaterlineT, targetT, noteT
    Dim Indiceoccurence, puissance
    Dim col_Criticity As Integer, col_Indice As Integer, numDec As Integer
    Dim prio As String
      
    NEVENTS = TotEventSheet(onglet)
    Ncrit = SDV2Ncrit(onglet)
    IndiceSDV = 0
    prio = ThisWorkbook.sheets("HOME").Range("Prestation").Value
    IndiceAgrementByRow = 0
    
    formule = ThisWorkbook.sheets("SETTINGS").Range("FORMS")
    Coeff1 = ThisWorkbook.sheets("SETTINGS").Range("COEF1")
    Coeff2 = ThisWorkbook.sheets("SETTINGS").Range("COEF2")
    Coeff3 = ThisWorkbook.sheets("SETTINGS").Range("COEF3")

    With ThisWorkbook.sheets(onglet)
                 If j = 1 Then
                    col_Indice = .Range("A6:BA6").Cells.Find(What:="Indice occurrencé", LookAt:=xlWhole).Column
                    col_Criticity = .Range("A5:BA5").Cells.Find(What:="Criticity", LookAt:=xlWhole).Column
                Else
                   col_Indice = .Range("BH6:GG6").Cells.Find(What:="Indice occurrencé", LookAt:=xlWhole).Column
                   col_Criticity = .Range("BH5:GG5").Cells.Find(What:="Criticity", LookAt:=xlWhole).Column
                End If
                col = col_Criticity + 1
                lig = rRow
                
                For lig = lig To lig
                        ReDim indices(Ncrit)
                        Indice = 0
                        IndiceC1 = 0
                        IndiceC2 = 0
                        SommeCriticities = 0
                        SommeC1s = 0
                        SommeC2s = 0
                            
                       For col = col To Ncrit + col_Criticity
                            If j = 1 Then p = .Cells(lig, 14) Else p = .Cells(lig, 73)
                            If p <> 1 And p <> 2 And p <> 3 Then
                                p = 3
                            End If
                            
                            If j = 1 Then
                                If .Cells(lig, 15) = "RED" Then
                                    priority = ThisWorkbook.sheets("SETTINGS").Cells(p + 6, ThisWorkbook.sheets("SETTINGS").Range("RR").Column)
                                ElseIf .Cells(lig, 15) = "YELLOW" Then
                                    priority = ThisWorkbook.sheets("SETTINGS").Cells(p + 6, ThisWorkbook.sheets("SETTINGS").Range("OO").Column)
                                ElseIf .Cells(lig, 14) = p And .Cells(lig, 15) = "GREEN" Then
                                   priority = ThisWorkbook.sheets("SETTINGS").Cells(p + 6, ThisWorkbook.sheets("SETTINGS").Range("GG").Column)
                                End If
                            Else
                                If .Cells(lig, 74) = "RED" Then
                                    priority = ThisWorkbook.sheets("SETTINGS").Cells(p + 6, ThisWorkbook.sheets("SETTINGS").Range("RR").Column)
                                ElseIf .Cells(lig, 74) = "YELLOW" Then
                                    priority = ThisWorkbook.sheets("SETTINGS").Cells(p + 6, ThisWorkbook.sheets("SETTINGS").Range("OO").Column)
                                ElseIf .Cells(lig, 73) = p And .Cells(lig, 74) = "GREEN" Then
                                   priority = ThisWorkbook.sheets("SETTINGS").Cells(p + 6, ThisWorkbook.sheets("SETTINGS").Range("GG").Column)
                                End If
                            End If
                            
                             waterline = CSng(.Cells(3, col))
                             Target = CSng(.Cells(4, col))
                             Criticity = (3 - CSng(.Cells(5, col).Value)) / 2       'coeffs 1 et 0.5 selon les criticités àld 2 et 1
                      
                            If Criticity = 0 Then
                                indices(k) = 0
                            Else
                                    If IsNumeric(.Cells(lig, col).Value) Then
                                        Note = CSng(.Cells(lig, col).Value)
                                    Else
                                        Note = .Cells(lig, col)
                                    End If
                                    ZF = Coeff1 * waterline + Coeff2 * Target + Coeff3
                                    If Target <> 0 And ZF <> 0 And Target <> ZF Then
                                         WaterlineT = 10 * (waterline - ZF) / (Target - ZF)
                                    End If
                                    targetT = 10
                                    If .Cells(lig, col).Value <> "" And IsNumeric(Note) = True And Target <> 0 And ZF <> 0 And Target <> ZF Then
                                        noteT = 10 * (Note - ZF) / (Target - ZF)
                                    Else
                                        noteT = ""
                                    End If
                                    If Note < ZF And .Cells(lig, col).Value <> "" And IsNumeric(Note) Then
                                        indices(k) = -1 * Criticity
                                    ElseIf noteT >= 0 And noteT < WaterlineT Then
                                        indices(k) = Criticity * (2 * noteT - targetT - WaterlineT) / (targetT + WaterlineT)
                                    ElseIf noteT >= WaterlineT And noteT < targetT Then
                                        indices(k) = Criticity * (noteT - targetT) / (targetT + WaterlineT)
                                    Else
                                        indices(k) = 0
                                    End If
                           End If
                           
                            If Criticity = 1 Then
                                IndiceC1 = IndiceC1 + indices(k)
                            ElseIf Criticity = 0.5 Then
                                IndiceC2 = IndiceC2 + indices(k)
                            End If
                            
                            If waterline > 0 And Target > 0 And Note > 0 And Note <= 10 Then
                                If Criticity = 1 Then
                                    SommeC1s = SommeC1s + 1
                                ElseIf Criticity = 0.5 Then
                                    SommeC2s = SommeC2s + 1
                                End If
                            End If
                            
                       Next col
                       
                        If SommeC1s <> 0 And SommeC2s <> 0 Then
                            Indice = IndiceC1 + (IndiceC2 / SommeC2s)
                        ElseIf SommeC1s = 0 And SommeC2s <> 0 Then
                            Indice = IndiceC2 / SommeC2s
                        ElseIf SommeC1s <> 0 And SommeC2s = 0 Then
                            Indice = IndiceC1
                        Else
                            Indiceoccurence = 0
                        End If
                        If Indice < -1 Then
                            Indice = -1
                        End If
                        Indiceoccurence = Indice * priority / 100
                        ThisWorkbook.sheets(onglet).Cells(rRow, col_Indice) = Indiceoccurence
                        
                        'ProgressTitle (Onglet & " : Couleur Points 'RED +'")
                        Call Color_RED_PLUS_BY_ROW(onglet, rRow, j)
                        If InStr(1, Indiceoccurence, ".") <> 0 Then
                            numDec = Len(Indiceoccurence) - InStr(1, Indiceoccurence, ".")
                            IndiceAgrementByRow = FormatNumber(Indiceoccurence, numDec)
                        Else
                             IndiceAgrementByRow = Indiceoccurence
                        End If
                Next lig
                
                
                 
     End With
  
End Function









