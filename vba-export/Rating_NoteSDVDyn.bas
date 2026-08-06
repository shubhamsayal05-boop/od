Attribute VB_Name = "Rating_NoteSDVDyn"
Sub Note_SDVDyn(onglet As String, prediction As Boolean)
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
        
        .Range("BR14") = tp(4)
        .Range("BR15") = tp(5)
        .Range("BR16") = tp(6)
        .Range("BR17") = tp(1)
        .Range("BR18") = tp(2)
        .Range("BR19") = tp(3)
        .Range("BR11") = IIf(100 - (.Range("BR14") + .Range("BR17")) < 0, 0, 100 - (.Range("BR14") + .Range("BR17")))
        .Range("BR12") = IIf(100 - (.Range("BR15") + .Range("BR18")) < 0, 0, 100 - (.Range("BR15") + .Range("BR18")))
        .Range("BR13") = IIf(100 - (.Range("BR16") + .Range("BR19")) < 0, 0, 100 - (.Range("BR16") + .Range("BR19")))
        
        'Nommer les nombres et pourcentages
        P1O = .Range("BQ14")
        P2O = .Range("BQ15")
        P3O = .Range("BQ16")
        P1R = .Range("BQ17")
        P2R = .Range("BQ18")
        P3R = .Range("BQ19")
        P1G = .Range("BQ11")
        P2G = .Range("BQ12")
        P3G = .Range("BQ13")
        
        NP1O = .Range("BP14")
        NP2O = .Range("BP15")
        NP3O = .Range("BP16")
        NP1R = .Range("BP17")
        NP2R = .Range("BP18")
        NP3R = .Range("BP19")
        NP1G = .Range("BP11")
        NP2G = .Range("BP12")
        NP3G = .Range("BP13")
       
        If prediction = False Then
            target_P1O = .Range("BR14")
            target_P2O = .Range("BR15")
            target_P3O = .Range("BR16")
            target_P1R = .Range("BR17")
            target_P2R = .Range("BR18")
            target_P3R = .Range("BR19")
            target_P1G = .Range("BR11")
            target_P2G = .Range("BR12")
            target_P3G = .Range("BR13")
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
                   
                        .Cells(.Range(order(onglet, 0)).row, .Range("colPD1").Column).Font.ColorIndex = t
                
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
                    
                       .Cells(.Range(order(onglet, 0)).row, .Range("colPD2").Column).Font.ColorIndex = t
               
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

                      .Cells(.Range(order(onglet, 0)).row, .Range("colPD3").Column).Font.ColorIndex = t
                   
                End With
            End If

        ElseIf prediction = True Then
            'P1 *************************************************************
           With ThisWorkbook.sheets("RATING")
            If P1R + P1O + P1G <> 0 Then
                If P1R > target_P1R And NP1R >= target_NP1R Then
                    .Cells(.Range(order(onglet, 0)).row, .Range("colPDD1").Column).Font.ColorIndex = 3
                ElseIf P1O > target_P1O And P1O + P1R > target_P1O + target_P1R And NP1O >= target_NP1O Then
                   .Cells(.Range(order(onglet, 0)).row, .Range("colPDD1").Column).Font.ColorIndex = 6
                Else
                   .Cells(.Range(order(onglet, 0)).row, .Range("colPDD1").Column).Font.ColorIndex = 10
                End If
            End If

            'P2 *************************************************************
            If P2R + P2O + P2G <> 0 Then
                If P2R > target_P2R And NP2R >= target_NP2R Then
                    .Cells(.Range(order(onglet, 0)).row, .Range("colPDD2").Column).Font.ColorIndex = 3
                ElseIf P2O > target_P2O And P2O + P2R > target_P2O + target_P2R And NP2O >= target_NP2O Then
                    .Cells(.Range(order(onglet, 0)).row, .Range("colPDD2").Column).Font.ColorIndex = 6
                Else
                    .Cells(.Range(order(onglet, 0)).row, .Range("colPDD2").Column).Font.ColorIndex = 10
                End If
            End If

            'P3 *************************************************************
           
            If P3R + P3O + P3G <> 0 Then
                If P3R > target_P3R And NP3R >= target_NP3R Then
                     .Cells(.Range(order(onglet, 0)).row, 21).Font.ColorIndex = 3
                ElseIf P3O > target_P3O And P3O + P3R > target_P3O + target_P3R And NP3O >= target_NP3O Then
                     .Cells(.Range(order(onglet, 0)).row, .Range("colPDD3").Column).Font.ColorIndex = 6
                Else
                      .Cells(.Range(order(onglet, 0)).row, .Range("colPDD3").Column).Font.ColorIndex = 10
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
        
   
        Set cels = ThisWorkbook.sheets("RATING").Range(order(onglet, 0)).Offset(0, 7)
        With ThisWorkbook.sheets("RATING")
             .Cells(.Range(order(onglet, 0)).row, 3).Font.Size = 2
        End With
        If prediction = False Then
            If NotEnough = True And Note <> "" Then
                .Range("BK4") = Note & " /!\"
            ElseIf NotEnough = False And Note <> "" Then
                .Range("BK4") = Note
            End If
            If Not colorGlobalDyn.Exists(UCase(onglet)) Then
                  colorGlobalDyn.Add key:=UCase(onglet), Item:=Note
             End If

        ElseIf prediction = True Then
'            If NotEnough = True And Note <> "" Then
'                ThisWorkbook.sheets("RATING").Range(Order(Onglet, 1)) = Note & " /!\"
'            ElseIf NotEnough = False And Note <> "" Then
'                ThisWorkbook.sheets("RATING").Range(Order(Onglet, 1)) = Note
'            End If
             If Not colorGlobalDynPred.Exists(UCase(onglet)) Then
                  colorGlobalDynPred.Add key:=UCase(onglet), Item:=Note
             End If
        End If
    End With

End Sub





