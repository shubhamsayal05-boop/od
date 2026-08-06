Attribute VB_Name = "Rating_NoteSDV"
Sub Note_SDV(onglet As String, prediction As Boolean)
    Dim cels As Range
    Dim NotEnough As Boolean, Milestone As Integer
    Dim tp As Variant, TPredic As Variant, NbPoint As Variant, NbPredic As Variant
    Dim t As Integer
    NotEnough = False
    P1_Rouge = False
    P1_YELLOW = False
    P2_Rouge = False
    P2_YELLOW = False
    Milestone = ThisWorkbook.sheets("HOME").Range("Milestone").Value

    With ThisWorkbook.sheets(onglet)
        tp = tauxPts(onglet, Milestone)
        NbPoint = NbMinPts(onglet, Milestone)
        TPredic = tauxPts(onglet, 4)
        NbPredic = NbMinPts(onglet, 4)
        
        .Range("K14") = tp(4)
        .Range("K15") = tp(5)
        .Range("K16") = tp(6)
        .Range("K17") = tp(1)
        .Range("K18") = tp(2)
        .Range("K19") = tp(3)
        .Range("K11") = IIf(100 - (.Range("K14") + .Range("K17")) < 0, 0, 100 - (.Range("K14") + .Range("K17")))
        .Range("K12") = IIf(100 - (.Range("K15") + .Range("K18")) < 0, 0, 100 - (.Range("K15") + .Range("K18")))
        .Range("K13") = IIf(100 - (.Range("K16") + .Range("K19")) < 0, 0, 100 - (.Range("K16") + .Range("K19")))
        
        'Nommer les nombres et pourcentages
        P1O = .Range("J14")
        P2O = .Range("J15")
        P3O = .Range("J16")
        P1R = .Range("J17")
        P2R = .Range("J18")
        P3R = .Range("J19")
        P1G = .Range("J11")
        P2G = .Range("J12")
        P3G = .Range("J13")
        NP1O = .Range("I14")
        NP2O = .Range("I15")
        NP3O = .Range("I16")
        NP1R = .Range("I17")
        NP2R = .Range("I18")
        NP3R = .Range("I19")
        NP1G = .Range("I11")
        NP2G = .Range("I12")
        NP3G = .Range("I13")
       
        If prediction = False Then
            target_P1O = .Range("K14")
            target_P2O = .Range("K15")
            target_P3O = .Range("K16")
            target_P1R = .Range("K17")
            target_P2R = .Range("K18")
            target_P3R = .Range("K19")
            target_P1G = .Range("K11")
            target_P2G = .Range("K12")
            target_P3G = .Range("K13")
            target_NP1O = NbPoint(4)
            target_NP2O = NbPoint(5)
            target_NP3O = NbPoint(6)
            target_NP1R = NbPoint(1)
            target_NP2R = NbPoint(2)
            target_NP3R = NbPoint(3)
        ElseIf prediction = True Then
            target_P1O = TPredic(4)
            target_P2O = TPredic(5)
            target_P3O = TPredic(6)
            target_P1R = TPredic(1)
            target_P2R = TPredic(2)
            target_P3R = TPredic(3)
            target_NP1O = NbPredic(4)
            target_NP2O = NbPredic(5)
            target_NP3O = NbPredic(6)
            target_NP1R = NbPredic(1)
            target_NP2R = NbPredic(2)
            target_NP3R = NbPredic(3)
        End If

        'Colorier les titres des graphes et les points dans l'onglet RATING

        If prediction = False Then
            'P1 *************************************************************
            t = 0
            If P1R + P1O + P1G <> 0 Then
                If P1R > target_P1R And NP1R >= target_NP1R Then
                    P1_Rouge = True
                    t = 3
                ElseIf P1O > target_P1O And P1O + P1R > target_P1O + target_P1R And NP1O >= target_NP1O Then
                    P1_YELLOW = True
                    t = 6
                Else
                   t = 10
                End If
                With ThisWorkbook.sheets("RATING")
                        .Cells(.Range(order(onglet, 0)).row, 7).Font.ColorIndex = t
               End With
            End If

            'P2 *************************************************************
           
            If P2R + P2O + P2G <> 0 Then
                If P2R > target_P2R And NP2R >= target_NP2R Then
                    P2_Rouge = True
                    t = 3
                ElseIf P2O > target_P2O And P2O + P2R > target_P2O + target_P2R And NP2O >= target_NP2O Then
                    P2_YELLOW = True
                    t = 6
                Else
                    t = 10
                End If
                With ThisWorkbook.sheets("RATING")
                        .Cells(.Range(order(onglet, 0)).row, 8).Font.ColorIndex = t
               End With
            End If

            'P3 *************************************************************
          
            If P3R + P3O + P3G <> 0 Then
                If P3R > target_P3R And NP3R >= target_NP3R Then
                   t = 3
                ElseIf P3O > target_P3O And P3O + P3R > target_P3O + target_P3R And NP3O >= target_NP3O Then
                   t = 6
                Else
                   t = 10
                End If
                With ThisWorkbook.sheets("RATING")
                      .Cells(.Range(order(onglet, 0)).row, 9).Font.ColorIndex = t
                End With
            End If

        ElseIf prediction = True Then
            'P1 *************************************************************
           With ThisWorkbook.sheets("RATING")
            If P1R + P1O + P1G <> 0 Then
                If P1R > target_P1R And NP1R >= target_NP1R Then
                    .Cells(.Range(order(onglet, 0)).row, 10).Font.ColorIndex = 3
                ElseIf P1O > target_P1O And P1O + P1R > target_P1O + target_P1R And NP1O >= target_NP1O Then
                   .Cells(.Range(order(onglet, 0)).row, 10).Font.ColorIndex = 6
                Else
                   .Cells(.Range(order(onglet, 0)).row, 10).Font.ColorIndex = 10
                End If
            End If

            'P2 *************************************************************
            If P2R + P2O + P2G <> 0 Then
                If P2R > target_P2R And NP2R >= target_NP2R Then
                    .Cells(.Range(order(onglet, 0)).row, 11).Font.ColorIndex = 3
                ElseIf P2O > target_P2O And P2O + P2R > target_P2O + target_P2R And NP2O >= target_NP2O Then
                   .Cells(.Range(order(onglet, 0)).row, 11).Font.ColorIndex = 6
                Else
                   .Cells(.Range(order(onglet, 0)).row, 11).Font.ColorIndex = 10
                End If
            End If

            'P3 *************************************************************
           
            If P3R + P3O + P3G <> 0 Then
                If P3R > target_P3R And NP3R >= target_NP3R Then
                    .Cells(.Range(order(onglet, 0)).row, 12).Font.ColorIndex = 3
                ElseIf P3O > target_P3O And P3O + P3R > target_P3O + target_P3R And NP3O >= target_NP3O Then
                    .Cells(.Range(order(onglet, 0)).row, 12).Font.ColorIndex = 6
                Else
                   .Cells(.Range(order(onglet, 0)).row, 12).Font.ColorIndex = 10
                End If
            End If
          End With
        End If

        'Variable booléenne 'Notenough' vraie si le nombre de points est insuffisant
        If ThisWorkbook.sheets(onglet).Range("G8") < OvMinPts(onglet) Then
            NotEnough = True
        End If

        'Note de la SDV
        If P1R > target_P1R And NP1R >= target_NP1R Then
            Note = "High Risk"
        ElseIf P2R > target_P2R And NP2R >= target_NP2R Then
            Note = "High Risk"
        ElseIf P3R > target_P3R And NP3R >= target_NP3R Then
            If P1O <= target_P1O Or P1O + P1R <= target_P1O + target_P1R Or NP1O < target_NP1O Then
                Note = "Medium Risk"
            Else
                Note = "High Risk"
            End If
        ElseIf P1O > target_P1O And P1O + P1R > target_P1O + target_P1R And NP1O >= target_NP1O Then
            Note = "Medium Risk"
        ElseIf P2O > target_P2O And P2O + P2R > target_P2O + target_P2R And NP2O >= target_NP2O Then
            Note = "Medium Risk"
        ElseIf P3O > target_P3O And P3O + P3R > target_P3O + target_P3R And NP3O >= target_NP3O Then
            If (P2O <= target_P2O And P1O <= target_P1O) Or (P1O + P1R <= target_P1O + target_P1R And P2O + P2R <= target_P2O + target_P2R) Or NP1O < target_NP1O Or NP2O < target_NP2O Then
                Note = "Low Risk"
            Else
                Note = "Medium Risk"
            End If
        Else
            If P1G + P2G + P3G <> 0 Then
                Note = "Low Risk"
            End If
        End If
        
        Set cels = ThisWorkbook.sheets("RATING").Range(order(onglet, 0))
        With ThisWorkbook.sheets("RATING")
            .Cells(.Range(order(onglet, 0)).row, 3).Font.Size = 2
        End With
        If prediction = False Then
            If NotEnough = True And Note <> "" Then
                .Range("D4") = Note & " /!\"
            ElseIf NotEnough = False And Note <> "" Then
                .Range("D4") = Note
            End If
            
            If Not colorGlobalDriv.Exists(UCase(onglet)) Then
                  colorGlobalDriv.Add key:=UCase(onglet), Item:=Note
             End If
        ElseIf prediction = True Then
'            If NotEnough = True And Note <> "" Then
'                ThisWorkbook.sheets("RATING").Range(Order(Onglet, 1)) = Note & " /!\"
'            ElseIf NotEnough = False And Note <> "" Then
'                ThisWorkbook.sheets("RATING").Range(Order(Onglet, 1)) = Note
'            End If
            If Not colorGlobalDrivPred.Exists(UCase(onglet)) Then
                  colorGlobalDrivPred.Add key:=UCase(onglet), Item:=Note
             End If
        End If
    End With

End Sub



