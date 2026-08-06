Attribute VB_Name = "CreateNew"
Option Explicit

Function NewSDVSeetings(nom As String)
     Dim j As Long, i As Long
     Dim LR As Integer
     
     With ThisWorkbook.Worksheets("SETTINGS")
            For i = 1 To 10
                j = .Cells(.Rows.Count, i).End(xlUp).row
                If j > LR Then LR = j
            Next i
            While .Range("B" & LR + 1).Interior.color = 14277081
                    LR = LR + 1
            Wend
            ThisWorkbook.Worksheets("CONFIGURATIONS ARRAY").Rows("35:49").Copy Destination:=.Range("A" & LR + 2)
            .Range("A" & LR + 2) = nom
     End With
     
End Function

Function NewSDVRating(nom As String)
    Dim i As Long
    Dim j As Long
    With ThisWorkbook.Worksheets("RATING")
                    i = getLastRowRating
                    .Rows(i + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                    Application.DisplayAlerts = False
                    j = ThisWorkbook.Worksheets("totalPoint").Cells(1, ThisWorkbook.Worksheets("totalPoint").Columns.Count).End(xlToLeft).Column
                    While ThisWorkbook.Worksheets("totalPoint").Cells(1, j).Interior.color <> RGB(255, 255, 255)
                        j = j + 1
                    Wend
                    
                    ThisWorkbook.Worksheets("totalPoint").Range("S1:" & ThisWorkbook.Worksheets("totalPoint").Cells(1, j - 1).Address).Copy Destination:=.Range("B" & i + 1)
                    Application.DisplayAlerts = True
'                     j = val(replace(replace(.Range("E" & i).Name.Name, "SDV", ""), "ms", "")) + 1
'                    ActiveWorkbook.Names.Add Name:="SDV" & j & "ms", RefersTo:="=RATING!$E$" & i + 1
'                    ActiveWorkbook.Names.Add Name:="SDV" & j & "fin", RefersTo:="=RATING!$W$" & i + 1
                    .Range("D" & i + 1) = nom
                    .Range("AE" & i & ":AQ" & i).Copy
                    .Range("AE" & i + 1 & ":AQ" & i + 1).PasteSpecial Paste:=xlPasteFormats
                    Application.CutCopyMode = False
     End With
End Function

Function NewSDVPowertrain(nom As String)
    Dim lastr As Long
    Dim r As Range
    
    With ThisWorkbook.sheets("POWERTRAIN")
            Set r = .Range("A12")
            lastr = .Range("B65000").End(xlUp).row

            While Application.CountA(.Rows(r.row).EntireRow) > 0 And r.row <= .Range("B65000").End(xlUp).row
                If UCase(.Range("A" & r.row)) = "TITRE CONFIG" Then
                        Call NewRow(nom, r.row - 2)
                ElseIf r.row = .Range("B65000").End(xlUp).row Then
                        Call NewRow(nom, r.row - 1)
                End If
                Set r = r.Offset(1, 0)
            Wend
    End With
End Function
Function NewSDVCalcul(nom As String)
   Dim i As Long
    With ThisWorkbook.Worksheets("Calculs")
                     i = LasRsCalcul - 1
                     If i = 0 Then Exit Function
                    .Rows(i + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                    .Rows(i).Copy Destination:=.Range("A" & i + 1)
                    .Range("B" & i + 1) = nom
                    .Range("C" & i + 1 & ":F" & i + 1) = 0
                    'Range("L" & i + 1) = nom
        End With
    
End Function
Function NewSDVConfigurationSetting(nom As String)
        Dim r As Long
        Dim derniereColonne
        Dim cm As Integer

        With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
                .Outline.ShowLevels RowLevels:=2
                r = 0
                derniereColonne = 30
                For cm = 1 To derniereColonne
                   If .Cells(.Rows.Count, cm).End(xlUp).row > r Then r = .Cells(.Rows.Count, cm).End(xlUp).row
                Next cm
                
                .Outline.ShowLevels RowLevels:=1
                If .Range("A65000").End(xlUp).row = r Then
                     .Rows(.Range("A65000").End(xlUp).row).Copy Destination:=.Cells(r + 2, 1)
                Else
                    .Rows(.Range("A65000").End(xlUp).row).Copy Destination:=.Cells(r + 1, 1)
                End If
                r = .Range("A65000").End(xlUp).row
                .Cells(r, 1) = UCase(nom)
                
        End With


End Function
Function NewSDVDefinitionSDV(nom As String)
        Dim i As Long
        Dim o As Long
        
       With ThisWorkbook.sheets("DEFINITION SDV")
                .Outline.ShowLevels RowLevels:=2
                o = .Range("A" & .Range("A65000").End(xlUp).row)
'                .Range("A2:E3").Copy Destination:=ThisWorkbook.Worksheets("DEFINITION SDV").Cells(ThisWorkbook.Worksheets("DEFINITION SDV").Range("A65000").End(xlUp).Row + 1, 1)
                .Rows("2:3").Copy Destination:=ThisWorkbook.Worksheets("DEFINITION SDV").Cells(ThisWorkbook.Worksheets("DEFINITION SDV").Range("A65000").End(xlUp).row + 1, 1)
                i = ThisWorkbook.Worksheets("DEFINITION SDV").Range("A65000").End(xlUp).row
                .Range("A" & i - 1 & ":A" & i) = o + 1
                .Range("B" & i - 1) = nom
                .Outline.ShowLevels RowLevels:=1
      End With

End Function
Function NewSDVStructure(nom As String)
     Dim j As Integer
     Dim i As Long
     With ThisWorkbook.Worksheets("STRUCTURE")
                    i = .Range("B65000").End(xlUp).row
                    i = i + 1
                    j = i
                    While .Range("C" & i) > 0
                        i = i + 1
                    Wend
                    .Rows((j - 1) & ":" & (j + 2)).Copy Destination:=.Range("A" & i)
                    .Range("B" & i) = nom
        End With
    
End Function
Function NewRow(nom As String, i As Long)
    With ThisWorkbook.Worksheets("POWERTRAIN")
                    .Rows(i + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                    .Rows(i).Copy Destination:=.Range("A" & i + 1)
                    .Range("B" & i + 1 & ":E" & i + 1).Value = 0
                    .Range("G" & i + 1 & ":I" & i + 1).Value = 0
                    .Range("A" & i + 1) = nom
        End With
End Function

Function LasRsCalcul() As Long
    Dim r As Range
    LasRsCalcul = 0
     With ThisWorkbook.Worksheets("Calculs")
            Set r = .Range("B5")
            LasRsCalcul = .Range("B65000").End(xlUp).row
          
            While r.row <= LasRsCalcul
                  
                If Application.CountA(.Rows(r.row).EntireRow) = 0 Then
                        LasRsCalcul = r.row
                        Exit Function
                End If
                Set r = r.Offset(1, 0)
            Wend
        End With
End Function









