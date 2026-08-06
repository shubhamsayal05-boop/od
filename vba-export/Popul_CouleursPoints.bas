Attribute VB_Name = "Popul_CouleursPoints"
' Cellule Office

Option Explicit
Private pr123(3) As Integer
Private py123(3)  As Integer
Private pg123(3)  As Integer
Private pw123(3)  As Integer
Sub Couleur_Points(ByVal onglet As String, ByVal Ncrit As Integer)  'DETERMINATION DE LA COULEUR DE CHAQUE EVENEMENT
    Dim NEVENTS As Variant
    Dim col As Integer, ColDyn As Integer
    Dim lig As Integer, waterline As Single
    Dim Target As Single, prio As Integer
    Dim lrow As Long, lcol As Long
    Dim col_Drivability As Integer, col_Criticity As Integer
    Dim col_DrivabilityDyn As Integer, col_CriticityDyn As Integer
    Dim foundColor As String
    Dim rangeTry As Range
    
    For col = 1 To 3
        pr123(col) = 0
        py123(col) = 0
        pg123(col) = 0
        pw123(col) = 0
    Next col
    NEVENTS = TotEventSheet(onglet)
   
   With ThisWorkbook.sheets(onglet)
        If Not ThisWorkbook.sheets(onglet).Range("A6:BA6").Cells.Find(What:="Event Rating", LookAt:=xlWhole) Is Nothing Then
            col_Drivability = ThisWorkbook.sheets(onglet).Range("A6:BA6").Cells.Find(What:="Event Rating", LookAt:=xlWhole).Column
        End If
        If Not ThisWorkbook.sheets(onglet).Range("A5:BA5").Cells.Find(What:="Criticity", LookAt:=xlWhole) Is Nothing Then
            col_Criticity = ThisWorkbook.sheets(onglet).Range("A5:BA5").Cells.Find(What:="Criticity", LookAt:=xlWhole).Column
        End If
        
        
        col = col_Criticity + 1
        ColDyn = col_CriticityDyn + 1
        lig = 7
        For lig = lig To NEVENTS
              For col = col To Ncrit + col_Criticity
                  If lig = 7 Then
                      If WorksheetFunction.CountA(.Range(.Cells(7, col), .Cells(NEVENTS, col))) > 0 Then
                        .Range(.Cells(7, col), .Cells(NEVENTS, col)).TextToColumns Destination:=.Cells(7, col), DataType:=xlDelimited, _
                        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
                        :=Array(1, 1), TrailingMinusNumbers:=True
                      End If
                 End If
                 
                 waterline = CSng(.Cells(3, col))
                 Target = CSng(.Cells(4, col))
                 prio = CSng(.Cells(5, col))
                   
                  If .Cells(lig, col) <> "" Then
                     
                      foundColor = GetColor(.Cells(lig, col), waterline, Target, prio)
                      If foundColor <> "0;0;0" Then
                        .Cells(lig, col).Interior.color = RGB(Split(foundColor, ";")(0), Split(foundColor, ";")(1), Split(foundColor, ";")(2))
                      End If
                 End If
                 
            Next col
            Call affectColor(.Cells(lig, col_Drivability))
             For col = 1 To 3
                pr123(col) = 0
                py123(col) = 0
                pg123(col) = 0
                pw123(col) = 0
            Next col
            col = col_Criticity + 1
            '_______
            
         
        Next lig
        
        'Driv
       Call orderCol(onglet, 1)
       
        'Dyn
       
    End With

End Sub

Function GetColor(textVal As Range, waterline As Single, Target As Single, prio As Integer) As String
           GetColor = "0;0;0"
    If Application.International(xlDecimalSeparator) = "," Then
        textVal.Value = replace(textVal.Value, ".", ",")
    Else
        textVal.Value = replace(textVal.Value, ",", ".")
    End If

           If IsNumeric(textVal) = True Then
                        If CSng(textVal) < waterline - (Target - waterline) And IsNumeric(textVal) = True And textVal <> "" Then
                             GetColor = "222;0;0"
                             If prio = 1 Then pr123(1) = pr123(1) + 1
                             If prio = 2 Then pr123(2) = pr123(2) + 1
                             If prio = 3 Then pr123(3) = pr123(3) + 1
                         ElseIf CSng(textVal) < waterline And CSng(textVal) >= waterline - (Target - waterline) Then
                             GetColor = "246;110;96"
                             If prio = 1 Then pr123(1) = pr123(1) + 1
                             If prio = 2 Then pr123(2) = pr123(2) + 1
                             If prio = 3 Then pr123(3) = pr123(3) + 1
                         ElseIf CSng(textVal) < waterline + (Target - waterline) / 3 And CSng(textVal) >= waterline Then
                             GetColor = "255;222;0"
                             If prio = 1 Then py123(1) = py123(1) + 1
                             If prio = 2 Then py123(2) = py123(2) + 1
                             If prio = 3 Then py123(3) = py123(3) + 1
                         ElseIf CSng(textVal) < waterline + 2 * (Target - waterline) / 3 And CSng(textVal) >= waterline + (Target - waterline) / 3 Then
                             GetColor = "255;247;102"
                             If prio = 1 Then py123(1) = py123(1) + 1
                             If prio = 2 Then py123(2) = py123(2) + 1
                             If prio = 3 Then py123(3) = py123(3) + 1
                         ElseIf CSng(textVal) < Target And CSng(textVal) >= waterline + 2 * (Target - waterline) / 3 Then
                             GetColor = "207;231;71"
                             If prio = 1 Then py123(1) = py123(1) + 1
                             If prio = 2 Then py123(2) = py123(2) + 1
                             If prio = 3 Then py123(3) = py123(3) + 1
                         ElseIf CSng(textVal) >= Target And IsNumeric(textVal) = True Then
                             GetColor = "0;153;38"
                             If prio = 1 Then pg123(1) = pg123(1) + 1
                             If prio = 2 Then pg123(2) = pg123(2) + 1
                             If prio = 3 Then pg123(3) = pg123(3) + 1
                        End If
            Else
                If prio = 1 Then pw123(1) = pw123(1) + 1
                If prio = 2 Then pw123(2) = pw123(2) + 1
                If prio = 3 Then pw123(3) = pw123(3) + 1
            End If
End Function

Function affectColor(r As Range)
            Dim final_color As String
            If pr123(1) >= 1 Then
                final_color = "RED"
            ElseIf pr123(1) = 0 And py123(1) >= 1 Then
                If pr123(2) + py123(2) + pg123(2) + pw123(2) > 2 Then
                    If pr123(2) >= 0.5 * (pr123(2) + py123(2) + pg123(2) + pw123(2)) And pr123(2) + py123(2) + pg123(2) + pw123(2) > 0 Then
                        final_color = "RED"
                    Else
                        final_color = "YELLOW"
                    End If
                ElseIf pr123(2) + py123(2) + pg123(2) = 2 Then
                    If pr123(2) = 2 Then
                        final_color = "RED"
                    Else
                        final_color = "YELLOW"
                    End If
                ElseIf pr123(2) + py123(2) + pg123(2) = 1 Then
                    If pr123(2) = 1 Then
                        final_color = "RED"
                    Else
                        final_color = "YELLOW"
                    End If
                ElseIf pr123(2) + py123(2) + pg123(2) = 0 Then
                    final_color = "YELLOW"
                End If
            ElseIf pr123(1) = 0 And py123(1) = 0 Then
                If pr123(2) + py123(2) + pg123(2) + pw123(2) > 2 Then
                    If pr123(2) >= 0.5 * (pr123(2) + py123(2) + pg123(2) + pw123(2)) And pr123(2) + py123(2) + pg123(2) + pw123(2) > 0 Then
                        final_color = "RED"
                    ElseIf pr123(2) < 0.5 * (pr123(2) + py123(2) + pg123(2) + pw123(2)) And pr123(2) + py123(2) + pg123(2) + pw123(2) > 0 Then
                        If py123(2) + pr123(2) >= 0.5 * (pr123(2) + py123(2) + pg123(2) + pw123(2)) And pr123(2) + py123(2) + pg123(2) + pw123(2) > 0 Then
                            final_color = "YELLOW"
                        Else
                            final_color = "GREEN"
                        End If
                    Else
                        final_color = "GREEN"
                    End If
                ElseIf pr123(2) + py123(2) + pg123(2) = 2 Then
                    If pr123(2) = 2 Then
                        final_color = "RED"
                    ElseIf (pr123(2) = 1 And py123(2) = 1) Or py123(2) = 2 Then
                        final_color = "YELLOW"
                    Else
                        final_color = "GREEN"
                    End If
                ElseIf pr123(2) + py123(2) + pg123(2) = 1 Then
                    If pr123(2) + py123(2) = 1 Then
                        final_color = "YELLOW"
                    Else
                        final_color = "GREEN"
                    End If
                ElseIf pr123(2) + py123(2) + pg123(2) = 0 Then
                    final_color = "GREEN"
                End If
            Else
                final_color = "GREEN"
            End If

            With r
                .HorizontalAlignment = xlCenter
                .Font.color = RGB(255, 255, 255)
                If final_color = "GREEN" Then
                    .Value = final_color
                    .Interior.color = RGB(0, 127, 0)
                ElseIf final_color = "YELLOW" Then
                    .Value = final_color
                    .Interior.color = RGB(255, 247, 0)
                    .Font.color = RGB(0, 0, 0)
                ElseIf final_color = "RED" Then
                    .Value = final_color
                    .Interior.color = RGB(255, 0, 0)
                End If
            End With

End Function

















